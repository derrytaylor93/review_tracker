# Review Tracker for LibreOffice Writer
#
# Tracks passages in a manuscript and keeps the line numbers cited in a
# response-to-reviewers document up to date.
#
# How it stores things (everything lives inside the .docx/.odt files):
#   * A tracked passage in the manuscript is a bookmark named RT_<id>
#     that spans exactly the text you selected. Its label is stored as a
#     custom document property with the same name.
#   * A citation in the response document is a bookmark named
#     RTref_<id>_<n> that spans the visible text, e.g. "Lines 45–51".
# Bookmarks are invisible, survive saving as .docx, and Word ignores them,
# so the files can be submitted as they are.
#
# Line numbers are read from LibreOffice's own page layout (the view cursor
# walks the document line by line), so they match the numbers LibreOffice
# shows in the margin with Tools > Line Numbering switched on.

import random
import re
import string

import uno
import unohelper
from com.sun.star.awt import XActionListener

TAG_PREFIX = "RT_"
REF_PREFIX = "RTref_"
TITLE = "Review Tracker"
DASH = "–"  # en dash

# Citation styles: (name shown in the dialog, singular prefix, plural prefix)
STYLES = [
    ("Lines 45–51", "Line ", "Lines "),
    ("lines 45–51", "line ", "lines "),
    ("L45–51", "L", "L"),
    ("45–51", "", ""),
]


# ── small helpers ────────────────────────────────────────────────────────────

def _new_id(existing, length=6):
    alphabet = string.ascii_lowercase + string.digits
    while True:
        tid = "".join(random.choice(alphabet) for _ in range(length))
        if tid not in existing:
            return tid


def _is_writer(doc):
    return doc is not None and hasattr(doc, "supportsService") and \
        doc.supportsService("com.sun.star.text.TextDocument")


def _bookmarks(doc, prefix):
    bms = doc.getBookmarks()
    return [bms.getByName(n) for n in bms.getElementNames() if n.startswith(prefix)]


def _tag_id(bookmark_name):
    """RT_abc123 -> abc123 ; RTref_abc123_x9z -> abc123"""
    if bookmark_name.startswith(REF_PREFIX):
        return bookmark_name[len(REF_PREFIX):].split("_")[0]
    return bookmark_name[len(TAG_PREFIX):]


def format_lines(style_index, start, end):
    _, one, many = STYLES[style_index]
    if end is None or end == start:
        return "%s%d" % (one, start)
    return "%s%d%s%d" % (many, start, DASH, end)


def detect_style(text):
    """Guess which citation style an existing citation uses, so updates keep it."""
    t = text.strip()
    if t.startswith("Line"):
        return 0
    if t.startswith("line"):
        return 1
    if re.match(r"L\s?\d", t):
        return 2
    if re.match(r"\d", t):
        return 3
    return 0


# ── labels (custom document properties) ─────────────────────────────────────

def _user_props(doc):
    return doc.getDocumentProperties().getUserDefinedProperties()


def get_label(doc, tag_name):
    props = _user_props(doc)
    if props.getPropertySetInfo().hasPropertyByName(tag_name):
        return props.getPropertyValue(tag_name)
    return ""


def set_label(doc, tag_name, label):
    props = _user_props(doc)
    if props.getPropertySetInfo().hasPropertyByName(tag_name):
        props.setPropertyValue(tag_name, label)
    else:
        props.addProperty(tag_name, 128, label)  # 128 = REMOVABLE


def _remove_label(doc, tag_name):
    props = _user_props(doc)
    if props.getPropertySetInfo().hasPropertyByName(tag_name):
        props.removeProperty(tag_name)


# ── line numbers from the layout ────────────────────────────────────────────

class LineMap:
    """Walks the document's layout once and answers 'which line is this on?'.

    Each visual line is identified by (page, y-position of its start). The
    walk follows the same lines LibreOffice numbers in the margin.
    """

    def __init__(self, doc):
        self.doc = doc
        self.vc = doc.getCurrentController().getViewCursor()
        self.lines = {}
        numbering = doc.getPropertyValue("LineNumberingProperties") \
            if doc.getPropertySetInfo().hasPropertyByName("LineNumberingProperties") else None
        count_empty = True
        if numbering is not None:
            try:
                count_empty = bool(numbering.CountEmptyLines)
            except Exception:
                pass
        self._walk(count_empty)

    def _key(self):
        vc = self.vc
        vc.gotoStartOfLine(False)
        return (vc.getPage(), vc.getPosition().Y)

    def _walk(self, count_empty):
        vc = self.vc
        vc.jumpToFirstPage()
        vc.jumpToStartOfPage()
        number = 0
        prev = None
        while True:
            key = self._key()
            if key == prev:
                break
            prev = key
            if not count_empty:
                vc.gotoEndOfLine(True)
                empty = vc.getString() == ""
                vc.gotoStartOfLine(False)
            else:
                empty = False
            if not empty:
                number += 1
            self.lines[key] = max(number, 1)
            if not vc.goDown(1, False):
                break

    def line_of(self, position):
        """Line number of a text position (an XTextRange), or None."""
        self.vc.gotoRange(position, False)
        return self.lines.get(self._key())

    def span_of(self, rng):
        start = self.line_of(rng.getStart())
        end = self.line_of(rng.getEnd())
        if start is not None and end is not None and end < start:
            end = start
        return start, end


class _PreserveView:
    """Remembers the user's cursor and puts it back after we move it around.

    (lockControllers() can't be used here: with the view locked the view cursor
    no longer reports layout positions.)
    """

    def __init__(self, doc):
        self.doc = doc

    def __enter__(self):
        vc = self.doc.getCurrentController().getViewCursor()
        try:
            self.saved = vc.getText().createTextCursorByRange(vc)
        except Exception:
            self.saved = None
        return self

    def __exit__(self, *exc):
        try:
            if self.saved is not None:
                self.doc.getCurrentController().getViewCursor().gotoRange(self.saved, False)
        except Exception:
            pass
        return False


def _line_maps():
    cache = {}

    def get(doc):
        key = id(doc)
        if key not in cache:
            cache[key] = (doc, LineMap(doc))
        return cache[key][1]
    return get


# ── core operations (no dialogs; used by the buttons and by the tests) ──────

def track_range(doc, rng, label):
    """Bookmark a range in the manuscript. Returns the bookmark name."""
    existing = {_tag_id(b.getName()) for b in _bookmarks(doc, TAG_PREFIX)}
    name = TAG_PREFIX + _new_id(existing)
    bm = doc.createInstance("com.sun.star.text.Bookmark")
    bm.setName(name)
    text = rng.getText()
    cursor = text.createTextCursorByRange(rng)
    text.insertTextContent(cursor, bm, True)
    set_label(doc, name, label)
    return name


def list_tags(doc, line_map=None):
    """[(bookmark name, label, start line, end line, snippet)] in document order."""
    tags = []
    bms = _bookmarks(doc, TAG_PREFIX)
    if not bms:
        return tags
    if line_map is None:
        with _PreserveView(doc):
            lm = LineMap(doc)
            spans = [lm.span_of(b.getAnchor()) for b in bms]
    else:
        spans = [line_map.span_of(b.getAnchor()) for b in bms]
    for b, (s, e) in zip(bms, spans):
        snippet = " ".join(b.getAnchor().getString().split())
        tags.append((b.getName(), get_label(doc, b.getName()), s, e, snippet))
    tags.sort(key=lambda t: (t[2] if t[2] is not None else 10**9))
    return tags


def remove_tag(doc, name):
    bms = doc.getBookmarks()
    if bms.hasByName(name):
        bm = bms.getByName(name)
        bm.getAnchor().getText().removeTextContent(bm)
    _remove_label(doc, name)


def _replace_bookmarked_text(doc, bm, new_text):
    name = bm.getName()
    anchor = bm.getAnchor()
    text = anchor.getText()
    cursor = text.createTextCursorByRange(anchor)
    text.removeTextContent(bm)
    cursor.setString(new_text)
    new_bm = doc.createInstance("com.sun.star.text.Bookmark")
    new_bm.setName(name)
    text.insertTextContent(cursor, new_bm, True)


def insert_ref(doc, position, tag_name, start, end, style_index):
    """Insert a citation such as 'Lines 45–51' at a text position."""
    tid = _tag_id(tag_name)
    existing = {b.getName() for b in _bookmarks(doc, REF_PREFIX)}
    while True:
        name = "%s%s_%s" % (REF_PREFIX, tid, _new_id(set(), 3))
        if name not in existing:
            break
    text = position.getText()
    cursor = text.createTextCursorByRange(position.getStart())
    label = format_lines(style_index, start, end)
    text.insertString(cursor, label, False)
    cursor.goLeft(len(label), True)
    bm = doc.createInstance("com.sun.star.text.Bookmark")
    bm.setName(name)
    text.insertTextContent(cursor, bm, True)
    return name, cursor


def find_manuscripts(desktop, exclude=None):
    """Open Writer documents that contain tracked passages."""
    docs = []
    comps = desktop.getComponents().createEnumeration()
    while comps.hasMoreElements():
        d = comps.nextElement()
        if _is_writer(d) and d is not exclude and _bookmarks(d, TAG_PREFIX):
            docs.append(d)
    return docs


def open_writer_docs(desktop):
    out = []
    comps = desktop.getComponents().createEnumeration()
    while comps.hasMoreElements():
        d = comps.nextElement()
        if _is_writer(d):
            out.append(d)
    return out


def update_refs(desktop, docs=None):
    """Refresh every citation in the given (default: all open) documents.

    Returns (n_updated, n_unchanged, missing_ids, docs_touched).
    """
    if docs is None:
        docs = open_writer_docs(desktop)
    manuscripts = find_manuscripts(desktop)
    where = {}  # tag id -> (manuscript doc, bookmark)
    for m in manuscripts:
        for b in _bookmarks(m, TAG_PREFIX):
            where[_tag_id(b.getName())] = (m, b)

    maps = {}
    updated = unchanged = 0
    missing = set()
    touched = []
    for doc in docs:
        refs = _bookmarks(doc, REF_PREFIX)
        if not refs:
            continue
        changed_here = False
        for ref in refs:
            tid = _tag_id(ref.getName())
            if tid not in where:
                missing.add(tid)
                continue
            m, tag_bm = where[tid]
            if id(m) not in maps:
                with _PreserveView(m):
                    lm = LineMap(m)
                    # A passage whose text was deleted leaves an empty bookmark
                    # behind; treat it as gone rather than citing that spot.
                    maps[id(m)] = {
                        _tag_id(b.getName()): (lm.span_of(b.getAnchor())
                                               if b.getAnchor().getString() else (None, None))
                        for b in _bookmarks(m, TAG_PREFIX)
                    }
            start, end = maps[id(m)][tid]
            if start is None:
                missing.add(tid)
                continue
            old = ref.getAnchor().getString()
            new = format_lines(detect_style(old), start, end)
            if new != old:
                _replace_bookmarked_text(doc, ref, new)
                updated += 1
                changed_here = True
            else:
                unchanged += 1
        if changed_here:
            touched.append(doc)
    return updated, unchanged, missing, touched


def finalise(doc):
    """Remove all Review Tracker bookmarks and labels from a document."""
    n = 0
    for b in _bookmarks(doc, REF_PREFIX) + _bookmarks(doc, TAG_PREFIX):
        name = b.getName()
        b.getAnchor().getText().removeTextContent(b)
        _remove_label(doc, name)
        n += 1
    return n


# ── user interface ──────────────────────────────────────────────────────────

def _ctx():
    return XSCRIPTCONTEXT.getComponentContext()  # noqa: F821 (provided by LibreOffice)


def _desktop():
    return XSCRIPTCONTEXT.getDesktop()  # noqa: F821


def _doc():
    return XSCRIPTCONTEXT.getDocument()  # noqa: F821


def _message(doc, text, kind="infobox", buttons=1):
    ctx = _ctx()
    tk = ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = doc.getCurrentController().getFrame().getContainerWindow()
    mtype = uno.Enum("com.sun.star.awt.MessageBoxType", kind.upper())
    box = tk.createMessageBox(parent, mtype, buttons, TITLE, text)
    result = box.execute()
    box.dispose()
    return result


class _Dialog:
    """Tiny helper for building dialogs in code (sizes are in 'app font' units)."""

    def __init__(self, title, width, height):
        ctx = _ctx()
        self.smgr = ctx.getServiceManager()
        self.ctx = ctx
        self.model = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        self.model.Title = title
        self.model.Width = width
        self.model.Height = height
        self.dialog = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        self.dialog.setModel(self.model)

    def add(self, kind, name, x, y, w, h, **props):
        m = self.model.createInstance("com.sun.star.awt.UnoControl%sModel" % kind)
        m.PositionX, m.PositionY, m.Width, m.Height = x, y, w, h
        for k, v in props.items():
            setattr(m, k, v)
        self.model.insertByName(name, m)
        return self.dialog.getControl(name)

    def button(self, name, label, x, y, w=50, h=14, kind=0, default=False):
        # kind: 0 = standard (needs a listener), 1 = OK, 2 = Cancel
        return self.add("Button", name, x, y, w, h, Label=label,
                        PushButtonType=kind, DefaultButton=default)

    def prepare(self):
        tk = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        self.dialog.createPeer(tk, None)

    def run(self):
        self.prepare()
        return self.dialog.execute()

    def close(self):
        self.dialog.dispose()


class _OnClick(unohelper.Base, XActionListener):
    def __init__(self, fn):
        self.fn = fn

    def actionPerformed(self, event):
        self.fn()

    def disposing(self, source):
        pass


def _set_list(dialog, name, items, select=0):
    """Fill a list box through its model (works before and after the dialog is shown)."""
    m = dialog.model.getByName(name)
    m.StringItemList = tuple(items)
    m.SelectedItems = (min(select, len(items) - 1),) if items else ()


def _get_list(dialog, name):
    sel = dialog.model.getByName(name).SelectedItems
    return sel[0] if sel else None


def _ask_text(prompt, default=""):
    d = _Dialog(TITLE, 240, 70)
    d.add("FixedText", "prompt", 8, 8, 224, 20, Label=prompt, MultiLine=True)
    edit = d.add("Edit", "value", 8, 30, 224, 14, Text=default)
    d.button("ok", "OK", 128, 50, kind=1, default=True)
    d.button("cancel", "Cancel", 182, 50, kind=2)
    d.prepare()
    edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(default)))
    edit.setFocus()
    ok = d.dialog.execute() == 1
    value = d.dialog.getControl("value").getText().strip()
    d.close()
    return value if ok else None


def _selection_range(doc):
    sel = doc.getCurrentController().getSelection()
    if sel is None or not hasattr(sel, "getCount") or sel.getCount() == 0:
        return None
    rng = sel.getByIndex(0)
    if not hasattr(rng, "getString") or rng.getString() == "":
        return None
    return rng


def _current_style(doc):
    """Default citation style: whatever the document's existing citations use."""
    refs = _bookmarks(doc, REF_PREFIX)
    return detect_style(refs[0].getAnchor().getString()) if refs else 0


def _describe(start, end):
    if start is None:
        return "?"
    return format_lines(0, start, end)


# ── buttons (these are the macros the toolbar and menu call) ────────────────

def track_selection(*args):
    """Track the selected text in the manuscript."""
    doc = _doc()
    rng = _selection_range(doc)
    if rng is None:
        _message(doc, "Select the passage you changed in the manuscript first, "
                      "then click Track Selection.")
        return
    snippet = " ".join(rng.getString().split())
    default = snippet[:50] + ("…" if len(snippet) > 50 else "")
    label = _ask_text("Give this passage a short label so you can find it later "
                      "(for example “R1.3 revised gaze definition”):", default)
    if label is None:
        return
    name = track_range(doc, rng, label or default)
    with _PreserveView(doc):
        start, end = LineMap(doc).span_of(doc.getBookmarks().getByName(name).getAnchor())
    _message(doc, "Tracking “%s” (%s).\n\nNow open your response document, "
                  "click where the line numbers should go and use Insert Line Reference."
             % (label or default, _describe(start, end)))


def tracked_passages(*args):
    """Show, jump to, rename or remove tracked passages in the current manuscript."""
    doc = _doc()
    tags = list_tags(doc)
    if not tags:
        _message(doc, "There are no tracked passages in this document yet. Select some "
                      "text and click Track Selection.")
        return

    d = _Dialog(TITLE + " – tracked passages", 320, 190)
    d.add("FixedText", "hint", 8, 6, 304, 10, Label="Tracked passages in this document:")
    d.add("ListBox", "list", 8, 18, 304, 146, Border=1)

    def items():
        return ["%s   ·   %s" % (_describe(s, e) if snippet else "text deleted",
                                   label or snippet[:60])
                for (_, label, s, e, snippet) in tags]

    def refresh(select=0):
        _set_list(d, "list", items(), select)

    def selected():
        pos = _get_list(d, "list")
        return pos if pos is not None and 0 <= pos < len(tags) else None

    def rename():
        pos = selected()
        if pos is None:
            return
        name, label, s, e, snip = tags[pos]
        new = _ask_text("New label:", label)
        if new:
            set_label(doc, name, new)
            tags[pos] = (name, new, s, e, snip)
            refresh(pos)

    def remove():
        pos = selected()
        if pos is None:
            return
        name = tags[pos][0]
        remove_tag(doc, name)
        del tags[pos]
        refresh(pos)

    refresh()
    d.button("goto", "Go to", 8, 170, kind=1, default=True)
    d.button("rename", "Rename…", 62, 170).addActionListener(_OnClick(rename))
    d.button("remove", "Stop tracking", 116, 170, w=60).addActionListener(_OnClick(remove))
    d.button("close", "Close", 262, 170, kind=2)
    result = d.run()
    pos = selected()
    d.close()
    if result == 1 and pos is not None:
        bm = doc.getBookmarks().getByName(tags[pos][0])
        doc.getCurrentController().select(bm.getAnchor())


def insert_line_reference(*args):
    """Insert the line numbers of a tracked passage at the cursor (response document)."""
    doc = _doc()
    desktop = _desktop()
    manuscripts = find_manuscripts(desktop, exclude=doc)
    if not manuscripts:
        _message(doc, "Open your manuscript in LibreOffice too, then try again.\n\n"
                      "(Review Tracker looks for an open document that has passages "
                      "tracked with Track Selection.)")
        return

    entries = []
    for m in manuscripts:
        title = m.getTitle()
        for (name, label, s, e, snippet) in list_tags(m):
            entries.append((m, name, s, e, label or snippet[:60], title))
    show_doc = len(manuscripts) > 1

    d = _Dialog(TITLE + " – insert line reference", 320, 200)
    d.add("FixedText", "hint", 8, 6, 304, 10, Label="Which tracked passage are you citing?")
    d.add("ListBox", "list", 8, 18, 304, 140, Border=1)
    _set_list(d, "list", [
        "%s   ·   %s%s" % (_describe(s, e), label, ("   [%s]" % t) if show_doc else "")
        for (_, _, s, e, label, t) in entries])
    d.add("FixedText", "fmt", 8, 166, 40, 10, Label="Format:")
    d.add("ListBox", "style", 48, 164, 90, 14, Dropdown=True, Border=1, LineCount=len(STYLES))
    _set_list(d, "style", [s[0] for s in STYLES], _current_style(doc))
    d.button("ok", "Insert", 208, 182, kind=1, default=True)
    d.button("cancel", "Cancel", 262, 182, kind=2)
    result = d.run()
    pos = _get_list(d, "list")
    style = _get_list(d, "style") or 0
    d.close()
    if result != 1 or pos is None:
        return

    m, name, s, e, _, _ = entries[pos]
    if s is None:
        _message(doc, "Couldn't work out the line number of that passage.")
        return
    vc = doc.getCurrentController().getViewCursor()
    _, cursor = insert_ref(doc, vc, name, s, e, style)
    vc.gotoRange(cursor.getEnd(), False)


def update_line_numbers(*args):
    """Refresh every line reference in all open documents."""
    doc = _doc()
    desktop = _desktop()
    if not find_manuscripts(desktop):
        _message(doc, "Open your manuscript in LibreOffice too, then try again.")
        return
    updated, unchanged, missing, touched = update_refs(desktop)
    if updated == 0 and unchanged == 0 and not missing:
        _message(doc, "No line references found in the open documents. Use Insert Line "
                      "Reference in your response document first.")
        return
    lines = ["Updated %d line reference%s (%d already up to date)."
             % (updated, "" if updated == 1 else "s", unchanged)]
    if touched:
        lines.append("Changed: " + ", ".join(t.getTitle() for t in touched) +
                     ". Remember to save.")
    if missing:
        lines.append("\n%d reference%s point to passages that were deleted, are no longer "
                     "tracked, or whose manuscript isn't open. They were left as they are."
                     % (len(missing), " seems to" if len(missing) == 1 else "s"))
    _message(doc, "\n".join(lines), kind="warningbox" if missing else "infobox")


def finalise_document(*args):
    """Strip all Review Tracker bookmarks from the current document (keeps the text)."""
    doc = _doc()
    answer = _message(doc, "Remove all Review Tracker bookmarks from “%s”?\n\n"
                           "The text and line numbers stay exactly as they are, but they "
                           "will no longer update. Usually you don't need this: the "
                           "bookmarks are invisible and safe to submit.\n\nTip: do this on "
                           "a copy." % doc.getTitle(),
                      kind="querybox", buttons=2)  # 2 = OK / Cancel
    if answer != 1:
        return
    n = finalise(doc)
    _message(doc, "Removed %d bookmark%s." % (n, "" if n == 1 else "s"))


g_exportedScripts = (
    track_selection,
    tracked_passages,
    insert_line_reference,
    update_line_numbers,
    finalise_document,
)
