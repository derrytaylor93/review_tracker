#!/usr/bin/env python3
"""
review_tracker.py  –  cross-platform line-number tracker for peer review

Workflow
--------
1. Enable "Line Numbers" in your manuscript (LibreOffice: Format → Line Numbering;
   Word: Layout → Page Setup → Line Numbers) then note the line number of your change.

2. Tag the change (the tool embeds a bookmark in the .docx so it moves with the text):
       python review_tracker.py tag manuscript.docx 42 "Revised species definition"

3. In your response document, type a placeholder wherever you want the line number:
       We revised the wording ({{line:RT_20240614_153042}}) to address this comment.

4. After further edits to the manuscript, sync — placeholders update, then stay in the
   document for future syncs:
       python review_tracker.py sync manuscript.docx response_to_reviewers.docx

5. Inspect all tags and their current line numbers:
       python review_tracker.py list manuscript.docx

Placeholder format
------------------
Initial  :  {{line:RT_20240614_153042}}
After sync:  Line 45{{line:RT_20240614_153042}}   ← ID is preserved for future syncs

You can also specify a line range; only the start line is tracked:
       python review_tracker.py tag manuscript.docx 42-48 "Methods paragraph"

Requirements
------------
    pip install python-docx
"""

import argparse
import json
import math
import re
import sys
import textwrap
from datetime import datetime
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Emu
from docx.enum.text import WD_COLOR_INDEX

# ── constants ────────────────────────────────────────────────────────────────

SIDECAR_SUFFIX = ".rt_tags.json"
TAG_PREFIX     = "RT_"
HIGHLIGHT      = WD_COLOR_INDEX.YELLOW

# Matches initial placeholder and already-synced single/range forms.
# Initial      : {{line:RT_xxx}}
# Synced single: Line 42{{line:RT_xxx}}
# Synced range : Lines 42-48{{line:RT_xxx}}
PLACEHOLDER_RE = re.compile(r'(?:Lines? \d+(?:-\d+)?)?\{\{line:([A-Za-z0-9_]+)\}\}')

EMU_PER_PT = 12700
EMU_PER_CM = 360000
A4_WIDTH   = int(21.0 * EMU_PER_CM)
MARGIN_DEF = int(2.54 * EMU_PER_CM)

# ── sidecar storage ───────────────────────────────────────────────────────────

def sidecar_path(manuscript: Path) -> Path:
    return manuscript.with_suffix(SIDECAR_SUFFIX)

def load_tags(manuscript: Path) -> dict:
    sp = sidecar_path(manuscript)
    return json.loads(sp.read_text(encoding="utf-8")) if sp.exists() else {}

def save_tags(manuscript: Path, tags: dict) -> None:
    sidecar_path(manuscript).write_text(
        json.dumps(tags, indent=2, ensure_ascii=False), encoding="utf-8"
    )

def new_tag_id() -> str:
    from random import randint
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{TAG_PREFIX}{ts}_{randint(0, 999):03d}"

# ── line-number calculation ───────────────────────────────────────────────────

def _text_width_pt(doc: Document) -> float:
    """Page text-column width in points, from the first section."""
    try:
        s = doc.sections[0]
        w  = s.page_width  or A4_WIDTH
        lm = s.left_margin  or MARGIN_DEF
        rm = s.right_margin or MARGIN_DEF
        return (w - lm - rm) / EMU_PER_PT
    except Exception:
        return (A4_WIDTH - 2 * MARGIN_DEF) / EMU_PER_PT   # ~467 pt for A4

def _font_pt(paragraph) -> float:
    for run in paragraph.runs:
        if run.font.size:
            return run.font.size / EMU_PER_PT
    try:
        sz = paragraph.style.font.size
        if sz:
            return sz / EMU_PER_PT
    except Exception:
        pass
    return 12.0

def compute_line_numbers(doc: Document) -> list[int]:
    """
    Return a list where element i is the first displayed line number of paragraph i
    (1-based, continuous). Matches what LibreOffice/Word show in the margin.
    """
    width = _text_width_pt(doc)
    starts = []
    current = 1
    for para in doc.paragraphs:
        starts.append(current)
        text = para.text
        if not text.strip():
            current += 1
            continue
        char_w  = _font_pt(para) * 0.5   # average char width ≈ 0.5 × font size
        cpl     = max(1, int(width / char_w))
        current += max(1, math.ceil(len(text) / cpl))
    return starts

def find_para_at_line(line_starts: list[int], target: int) -> int:
    """Return the 0-based paragraph index whose line range contains target."""
    for i in range(len(line_starts) - 1, -1, -1):
        if line_starts[i] <= target:
            return i
    return 0

# ── Word bookmark XML helpers ─────────────────────────────────────────────────

def _next_bookmark_id(doc: Document) -> int:
    ids = {
        int(el.get(qn("w:id"), 0))
        for el in doc.element.body.iter(qn("w:bookmarkStart"))
    }
    return max(ids, default=0) + 1

def insert_bookmark(doc: Document, para_index: int, name: str) -> None:
    """Embed a named bookmark at the start of the paragraph."""
    bm_id  = str(_next_bookmark_id(doc))
    para_p = doc.paragraphs[para_index]._p

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"),   bm_id)
    start.set(qn("w:name"), name)

    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bm_id)

    # Insert at the very beginning of the paragraph XML
    para_p.insert(0, start)
    para_p.append(end)

def find_bookmark_para(doc: Document, name: str) -> int:
    """Return the 0-based paragraph index that contains a named bookmark, or -1."""
    for i, para in enumerate(doc.paragraphs):
        for el in para._p.iter(qn("w:bookmarkStart")):
            if el.get(qn("w:name")) == name:
                return i
    return -1

# ── placeholder replacement ───────────────────────────────────────────────────

def _replace_in_text(text: str, current_lines: dict, missing: set) -> tuple[str, int]:
    """Replace all placeholders in a string. Returns (new_text, n_replaced)."""
    count = 0
    def _sub(m):
        nonlocal count
        tid = m.group(1)
        if tid in current_lines:
            count += 1
            return f"Line {current_lines[tid]}{{{{{{'line:{tid}'}}}}}}"  # preserve placeholder
        missing.add(tid)
        return m.group(0)
    # Build replacement using a plain function to avoid f-string nesting confusion
    def sub(m):
        nonlocal count
        tid = m.group(1)
        if tid in current_lines:
            count += 1
            val    = current_lines[tid]          # e.g. "42" or "42-48"
            plural = "Lines" if "-" in str(val) else "Line"
            return plural + " " + str(val) + "{{line:" + tid + "}}"
        missing.add(tid)
        return m.group(0)
    return PLACEHOLDER_RE.sub(sub, text), count

def _update_runs(para, current_lines: dict, missing: set) -> int:
    """Update placeholders run-by-run. Returns number of replacements."""
    total = 0
    for run in para.runs:
        if "{{line:" in run.text:
            new_text, n = _replace_in_text(run.text, current_lines, missing)
            run.text = new_text
            total += n
    return total

def _update_para_joining_runs(para, current_lines: dict, missing: set) -> int:
    """
    Fallback for when a placeholder is split across runs.
    Joins all run text, replaces, writes back into first run.
    """
    full = para.text
    if not PLACEHOLDER_RE.search(full):
        return 0
    new_text, n = _replace_in_text(full, current_lines, missing)
    if n == 0:
        return 0
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""
    return n

def replace_placeholders(doc: Document, current_lines: dict) -> tuple[int, set]:
    """Walk every paragraph and table cell. Returns (total_replaced, missing_ids)."""
    total   = 0
    missing: set = set()

    def process_para(para):
        nonlocal total
        n = _update_runs(para, current_lines, missing)
        if n == 0 and PLACEHOLDER_RE.search(para.text):
            n = _update_para_joining_runs(para, current_lines, missing)
        total += n

    for para in doc.paragraphs:
        process_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    process_para(para)

    return total, missing

# ── commands ──────────────────────────────────────────────────────────────────

def cmd_tag(manuscript: Path, line_spec: str, label: str) -> None:
    if not manuscript.exists():
        sys.exit(f"Error: '{manuscript}' not found.")

    # Accept "42" or "42-48"
    m = re.fullmatch(r"(\d+)(?:-(\d+))?", line_spec.strip())
    if not m:
        sys.exit("Error: line spec must be a number or range like '42' or '42-48'.")
    start_line = int(m.group(1))
    end_line   = int(m.group(2)) if m.group(2) else None

    if end_line is not None and end_line < start_line:
        sys.exit("Error: range end must be >= range start.")

    doc         = Document(str(manuscript))
    line_starts = compute_line_numbers(doc)
    total_lines = line_starts[-1]

    for ln, name in [(start_line, "start"), (end_line, "end")]:
        if ln is not None and (ln < 1 or ln > total_lines):
            sys.exit(f"Error: line {ln} ({name}) out of range "
                     f"(document has ~{total_lines} lines).")

    tag_id      = new_tag_id()
    start_para  = find_para_at_line(line_starts, start_line)
    end_para    = find_para_at_line(line_starts, end_line) if end_line else None

    # Embed bookmark(s) — they move with the text through future edits
    insert_bookmark(doc, start_para, tag_id)
    if end_para is not None and end_para != start_para:
        insert_bookmark(doc, end_para, tag_id + "_END")

    # Highlight the spanned paragraphs yellow
    span = range(start_para, (end_para or start_para) + 1)
    for i in span:
        para = doc.paragraphs[i]
        for run in para.runs:
            run.font.highlight_color = HIGHLIGHT
        if not para.runs:
            para.add_run("").font.highlight_color = HIGHLIGHT

    doc.save(str(manuscript))

    tags = load_tags(manuscript)
    tags[tag_id] = {
        "label":      label,
        "start_line": start_line,
        "end_line":   end_line,
    }
    save_tags(manuscript, tags)

    actual_start = line_starts[start_para]
    actual_end   = line_starts[end_para] if end_para is not None else None
    line_display = f"{actual_start}-{actual_end}" if actual_end else str(actual_start)

    print(f"\n  Tagged  : lines {line_display}")
    print(f"  Label   : {label}")
    print(f"  ID      : {tag_id}")
    print(f"\n  Paste this placeholder into your response document:")
    print(f"  {{{{line:{tag_id}}}}}\n")


def cmd_sync(manuscript: Path, response: Path) -> None:
    for p in (manuscript, response):
        if not p.exists():
            sys.exit(f"Error: '{p}' not found.")

    tags = load_tags(manuscript)
    if not tags:
        print("No tags in manuscript. Use 'tag' first.")
        return

    doc         = Document(str(manuscript))
    line_starts = compute_line_numbers(doc)

    # Resolve each tag's current line number(s) via its bookmark(s)
    # current_lines maps tag_id -> display string like "42" or "42-48"
    current_lines: dict[str, str] = {}
    for tag_id, info in tags.items():
        start_idx = find_bookmark_para(doc, tag_id)
        if start_idx < 0:
            print(f"  Warning: bookmark for '{tag_id}' not found in manuscript "
                  f"(was it deleted?). Skipping.")
            continue
        start_ln = line_starts[start_idx]

        end_idx = find_bookmark_para(doc, tag_id + "_END")
        if end_idx >= 0:
            end_ln = line_starts[end_idx]
            current_lines[tag_id] = f"{start_ln}-{end_ln}"
            tags[tag_id]["start_line"] = start_ln
            tags[tag_id]["end_line"]   = end_ln
        else:
            current_lines[tag_id] = str(start_ln)
            tags[tag_id]["start_line"] = start_ln

    if not current_lines:
        print("No bookmarks found — nothing to sync.")
        return

    resp_doc              = Document(str(response))
    n_replaced, missing   = replace_placeholders(resp_doc, current_lines)
    resp_doc.save(str(response))
    save_tags(manuscript, tags)

    print(f"\n  Sync complete — {n_replaced} reference(s) updated in '{response.name}'")
    if missing:
        print(f"\n  Warning: these IDs were not found in the manuscript tags:")
        for tid in sorted(missing):
            print(f"    {{{{line:{tid}}}}}")
    print()


def cmd_list(manuscript: Path) -> None:
    if not manuscript.exists():
        sys.exit(f"Error: '{manuscript}' not found.")

    tags = load_tags(manuscript)
    if not tags:
        print("No tags found. Use 'tag' to create one.")
        return

    doc         = Document(str(manuscript))
    line_starts = compute_line_numbers(doc)

    print(f"\n  Tags in '{manuscript.name}':\n")
    print(f"  {'ID':<32}  {'Lines':>9}  Label")
    print(f"  {'-'*32}  {'-'*9}  -----")
    for tag_id, info in tags.items():
        start_idx = find_bookmark_para(doc, tag_id)
        if start_idx >= 0:
            s = line_starts[start_idx]
            end_idx = find_bookmark_para(doc, tag_id + "_END")
            e = line_starts[end_idx] if end_idx >= 0 else None
            current = f"{s}-{e}" if e else str(s)
            old_s   = info.get("start_line", s)
            old_e   = info.get("end_line")
            old     = f"{old_s}-{old_e}" if old_e else str(old_s)
            changed = " *" if current != old else ""
        else:
            current = str(info.get("start_line", "?"))
            changed = " !"
        print(f"  {tag_id:<32}  {current:>9}{changed}  {info['label']}")

    print("\n  (* line has shifted since last sync;  ! bookmark missing)\n")


def cmd_untag(manuscript: Path, tag_id: str) -> None:
    if not manuscript.exists():
        sys.exit(f"Error: '{manuscript}' not found.")

    tags = load_tags(manuscript)
    if tag_id not in tags:
        sys.exit(f"Error: tag '{tag_id}' not found in sidecar.")

    info = tags.pop(tag_id)

    doc         = Document(str(manuscript))
    line_starts = compute_line_numbers(doc)
    para_idx    = find_bookmark_para(doc, tag_id)

    if para_idx >= 0:
        # Remove highlight
        para = doc.paragraphs[para_idx]
        for run in para.runs:
            run.font.highlight_color = None
        # Remove bookmark XML nodes
        body = doc.element.body
        for el in list(body.iter(qn("w:bookmarkStart"))):
            if el.get(qn("w:name")) == tag_id:
                bm_id = el.get(qn("w:id"))
                el.getparent().remove(el)
                # Remove matching bookmarkEnd
                for end in list(body.iter(qn("w:bookmarkEnd"))):
                    if end.get(qn("w:id")) == bm_id:
                        end.getparent().remove(end)
                        break
                break
        doc.save(str(manuscript))

    save_tags(manuscript, tags)
    print(f"  Removed tag '{tag_id}'  (label: \"{info['label']}\")\n")


def cmd_render(response: Path, output: Optional[Path]) -> None:
    """
    Produce a clean copy of the response document with all {{line:...}} tracking
    codes stripped, leaving only the human-readable line numbers.
    e.g.  Lines 42-48{{line:RT_xxx}}  →  Lines 42-48
    """
    if not response.exists():
        sys.exit(f"Error: '{response}' not found.")

    if output is None:
        output = response.with_stem(response.stem + "_FINAL")

    STRIP_RE = re.compile(r'\{\{line:[A-Za-z0-9_]+\}\}')

    doc   = Document(str(response))
    count = 0

    def clean_para(para):
        nonlocal count
        for run in para.runs:
            if "{{line:" in run.text:
                new_text = STRIP_RE.sub("", run.text)
                if new_text != run.text:
                    count   += len(STRIP_RE.findall(run.text))
                    run.text = new_text
        # fallback: split across runs
        if "{{line:" in para.text:
            full     = para.text
            new_full = STRIP_RE.sub("", full)
            if new_full != full and para.runs:
                para.runs[0].text = new_full
                for run in para.runs[1:]:
                    run.text = ""
                count += len(STRIP_RE.findall(full))

    for para in doc.paragraphs:
        clean_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    clean_para(para)

    doc.save(str(output))
    print(f"\n  Rendered '{output.name}' — {count} tracking code(s) removed.")
    print(f"  This file is safe to share with reviewers.\n")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        prog="review_tracker",
        description="Track line-number citations between a manuscript and a response document.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Examples
            --------
            # Tag the change at line 42
            python review_tracker.py tag manuscript.docx 42 "Revised species definition"

            # Tag a range (tracks the start line)
            python review_tracker.py tag manuscript.docx 42-48 "Rewrote methods paragraph"

            # Sync after further edits
            python review_tracker.py sync manuscript.docx response_to_reviewers.docx

            # See all tags and their current line numbers
            python review_tracker.py list manuscript.docx

            # Remove a tag
            python review_tracker.py untag manuscript.docx RT_20240614_153042_001
        """),
    )

    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("tag",   help="Tag a line in the manuscript")
    p.add_argument("manuscript", type=Path)
    p.add_argument("line",       type=str, help="Line number or range, e.g. 42 or 42-48")
    p.add_argument("label",      type=str, help="Short description of this change")

    p = sub.add_parser("sync",  help="Update {{line:ID}} placeholders in the response doc")
    p.add_argument("manuscript", type=Path)
    p.add_argument("response",   type=Path)

    p = sub.add_parser("list",  help="List all tags and their current line numbers")
    p.add_argument("manuscript", type=Path)

    p = sub.add_parser("untag", help="Remove a tag from the manuscript")
    p.add_argument("manuscript", type=Path)
    p.add_argument("tag_id",     type=str)

    p = sub.add_parser("render", help="Export clean response doc with tracking codes stripped")
    p.add_argument("response", type=Path)
    p.add_argument("output",   type=Path, nargs="?", default=None,
                   help="Output path (default: response_FINAL.docx)")

    args = parser.parse_args()

    if   args.command == "tag":    cmd_tag(args.manuscript, args.line, args.label)
    elif args.command == "sync":   cmd_sync(args.manuscript, args.response)
    elif args.command == "list":   cmd_list(args.manuscript)
    elif args.command == "untag":  cmd_untag(args.manuscript, args.tag_id)
    elif args.command == "render": cmd_render(args.response, args.output)


if __name__ == "__main__":
    main()
