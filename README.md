# Review Tracker

Keep the line numbers in your response to reviewers up to date while you keep revising the manuscript.

> *"We revised the definition on **Lines 42–48** to address this comment."*

Every time you edit the manuscript, those numbers drift. Review Tracker remembers which passage each citation points to and refreshes all of them with one click.

There are three versions:

- **LibreOffice extension.** Buttons inside LibreOffice Writer: highlight text, click, done. Line numbers come from LibreOffice's own page layout, so they match the numbers in the margin exactly.
- **Word add-in** (new, see [Word version](#word-version)). The same idea as a *Review Tracker* tab in Word's ribbon. It stores everything the same way, so a manuscript tracked in LibreOffice also works in Word and vice versa.
- **Command-line script** (`review_tracker.py`, the original version, documented [further down](#command-line-version)). Line numbers are estimated, so they can be slightly off.

---

## Suggested workflow

Line numbers drift in two directions. Your own edits move around, and so do the lines the reviewers pointed at: after a few changes, "line 120" in a reviewer's comment no longer points at the text they meant. Tracking both keeps everything findable.

**1. Before you change anything, track what the reviewers pointed at.**
Put the reviewer comments into your response document. Then go through the manuscript and track each passage a reviewer refers to, while the line numbers still match the version they read. First check a couple of numbers against the PDF the reviewers saw: Word and LibreOffice can lay out the same file slightly differently.

**2. Link each comment to its passage, but keep the reviewer's own numbers.**
Leave the reviewer's text exactly as written, so they recognise their comment, and insert a live reference next to it:

> *Reviewer 1.3: "The definition of gaze on line 120 is unclear." [now Lines 134–136]*

From then on you can always jump to what the reviewer meant, however much the manuscript changes.

**3. Name labels so the two kinds are easy to tell apart.**
For example `R1.3 reviewer` for the passage the reviewer pointed at, and `R1.3 change` for your edit. In the passage list, everything belonging to one comment then sits together.

**4. Make your revisions, tracking each change as you go.**
Highlight each revised passage, track it, and cite it in your reply:

> *We have rewritten the definition (Lines 134–139).*

**5. Update before you send.**
Click Update Line Numbers with both documents open, check that nothing is reported missing, and save. If you deleted a passage a reviewer pointed at, it is reported as missing once the deletion is accepted. That's a useful reminder to say in your reply that the passage was removed.

---

## LibreOffice extension

### Install (once)

1. Download [`libreoffice/ReviewTracker.oxt`](libreoffice/ReviewTracker.oxt) (click the file, then the download button).
2. In LibreOffice, open **Tools → Extensions… → Add…**, pick `ReviewTracker.oxt` and accept. (Double-clicking the file also works.)
3. Restart LibreOffice.

You'll now see a **Review Tracker** menu next to Help, and a **Review Tracker** toolbar with four buttons. If the toolbar is hidden, switch it on in **View → Toolbars → Review Tracker**. Nothing else needs installing: LibreOffice already includes Python.

Works with LibreOffice 7 or newer on Mac, Windows and Linux, with `.docx` and `.odt` files.

### Use

Turn on line numbers in the manuscript so you can see what the tool sees: **Tools → Line Numbering… → Show numbering**, interval 1.

1. **Track Selection** (in the manuscript). Highlight the passage you changed and click **Track Selection**. Give it a short label, e.g. `R1.3 gaze definition`. Highlight exactly what you want to cite: a sentence that spans lines 45–51 will be cited as "Lines 45–51".
2. **Insert Line Ref** (in the response document). Keep the manuscript open. Click where the line numbers should go, click **Insert Line Ref**, pick the passage and a format (`Lines 45–51`, `lines 45–51`, `L45–51` or `45–51`). The text is inserted like normal typing.
3. **Update Line Numbers** (anywhere). After more edits to the manuscript, click this with both documents open. Every citation in every open document is refreshed. Save the response document afterwards.
4. **Tracked Passages** (in the manuscript). See all tracked passages with their current lines, jump to one, rename it, or stop tracking it.

You can cite the same passage as many times as you like. You can also edit around a citation freely; just don't retype the number itself, or the link to the passage is lost (insert it again instead).

### Good to know

- **Everything is stored in your documents.** A tracked passage is an invisible bookmark in the manuscript; a citation is an invisible bookmark around the number in the response document. There are no extra files to keep track of, and the documents can be emailed, renamed or moved.
- **Safe to submit.** Bookmarks don't show up in print or PDF, and Word ignores them, so you can submit the files as they are. If you still want to remove them, use **Review Tracker → Remove Tracking from This Document…** (on a copy). The text stays exactly as it is.
- **Tracked changes count the way you see them.** Line numbers follow what's on screen. If deleted text is shown (Edit → Track Changes → Show), its lines are counted, just like the numbers in the margin. To cite the clean version, hide tracked changes before clicking Update Line Numbers.
- **Word gives slightly different numbers.** Word and LibreOffice lay out pages slightly differently, so the same file can have different line numbers in each. The extension matches LibreOffice. Make the PDF that reviewers see from LibreOffice, or check a few numbers if the journal converts your .docx itself.
- **Edits inside a tracked passage are fine.** The bookmark stretches and shrinks with the text. If you delete a tracked passage entirely, its citations are left as they are and Update Line Numbers tells you how many it couldn't find; track the new text and insert the citation again. (With Track Changes recording, deleted text stays in the document as struck-through text until you accept the deletion, so it is still cited until then.)

### For developers

The extension is plain Python in [`libreoffice/extension/Scripts/python/ReviewTracker.py`](libreoffice/extension/Scripts/python/ReviewTracker.py). Rebuild the `.oxt` after changing it:

```
python3 libreoffice/build_oxt.py
```

End-to-end test against a running LibreOffice (uses your own manuscript and response document; checks the numbers against a PDF export of the margin numbers, so `pdftotext` is needed):

```
soffice --headless --norestore --accept="socket,host=localhost,port=2002;urp;" &
python3 libreoffice/tests/test_core.py manuscript.docx response.docx /tmp/rt_test
```

---

## Word version

> **New and not yet tested in Word itself.** It was built and checked without access to Word. If something doesn't work, please open an issue with what you clicked and what happened.

### Install (once)

1. Download [`word/ReviewTracker.dotm`](word/ReviewTracker.dotm) (click the file, then the download button).
2. Put it in Word's **Startup** folder, so it loads every time Word starts:
   - **Mac:** in Finder choose **Go → Go to Folder…** and paste
     `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word`
     (on some versions the folder names have no `.localized`), then drag `ReviewTracker.dotm` into it.
     To see which folder your Word uses: **Word → Preferences → File Locations → Startup**.
   - **Windows:** press Win+R, paste `%appdata%\Microsoft\Word\STARTUP` and press Enter, then copy the file there.
3. Quit Word completely and open it again. You'll see a **Review Tracker** tab in the ribbon.

You don't open the `.dotm` file yourself; Word loads it in the background. If the tab appears but the buttons do nothing, check that macros from add-ins are allowed: **Word → Preferences → Security & Privacy → Macro Security** on Mac, or **File → Options → Trust Center** on Windows.

### Use

The buttons work like the LibreOffice version:

1. **Track Selection** (manuscript). Highlight the passage and click it, then give it a label.
2. **Insert Line Ref** (response document). Keep the manuscript open. A box lists the tracked passages with numbers; type the number you want. Tip: if you first click inside the passage in the manuscript, it is already filled in, so you only press Enter. The first time, you're also asked once how citations should look (`Lines 45–51`, `lines 45–51`, `L45–51` or `45–51`).
3. **Update Line Numbers** (anywhere). Refreshes every citation in every open document, then save.
4. **Next Passage** jumps to the next tracked passage (its label shows in the status bar). **Stop Tracking** forgets the passage the cursor is in. **Remove Tracking** strips all Review Tracker bookmarks from the document and keeps the text.

Turn on continuous line numbers so you can compare: **Layout → Line Numbers → Continuous**. Word's line numbers don't count lines inside tables or text boxes, but Review Tracker does, so numbers after a table in the main text can be off. (PeerJ wants tables in separate files, so this rarely matters.)

### If the ribbon tab doesn't appear

You can use the macros without the template. In Word, open **Tools → Macro → Visual Basic Editor** (Mac) or press Alt+F11 (Windows). Select **Normal** on the left, choose **File → Import File…** and pick [`word/ReviewTracker.bas`](word/ReviewTracker.bas). Then run `TrackSelection`, `InsertLineReference`, `UpdateLineNumbers`, `NextTrackedPassage`, `StopTrackingHere` or `RemoveTrackingFromDocument` from **Tools → Macro → Macros…**, or add them as buttons to the toolbar.

### For developers

The macros are in [`word/ReviewTracker.bas`](word/ReviewTracker.bas) and the ribbon is [`word/customUI14.xml`](word/customUI14.xml). `python3 word/build_dotm.py` rebuilds `ReviewTracker.dotm` from them without needing Word. It writes the VBA project file (`vbaProject.bin`) directly, and Word compiles the macros from source the first time it opens the template.

---

## Command-line version

The original script. It still works, but the LibreOffice extension above is easier to use and more accurate.

### How the command-line version works

1. You **tag** a change in the manuscript by specifying the line number. The tool embeds an invisible bookmark in the `.docx` file so it can track that passage even as surrounding text shifts.
2. You paste a **placeholder** (e.g. `{{line:RT_20260614_191649_513}}`) into your response document wherever you want to cite that line.
3. Whenever you want to update the citations, you run **sync**. The tool finds each bookmark's current line number and updates all the placeholders in your response document.
4. When you're ready to submit, **render** produces a clean copy with all the tracking codes stripped out.

---

### Installation

**Requirements:** Python 3.10+ and one library.

```
pip install python-docx
```

Download `review_tracker.py` and place it in the same folder as your manuscript and response document, or anywhere on your PATH.

---

### Workflow

#### Step 1 — Enable line numbers in your manuscript

Turn on continuous line numbering so you can read off the line numbers you want to tag.

- **LibreOffice Writer:** Format → Line Numbers → check *Show Line Numbers*, set *Interval* to 1
- **Microsoft Word:** Layout → Page Setup → Line Numbers → Continuous

#### Step 2 — Tag a change

Select the line number (or range) where you made a change, then run:

```
python review_tracker.py tag manuscript.docx 42 "Revised species definition"
```

For a range of lines:

```
python review_tracker.py tag manuscript.docx 42-48 "Rewrote methods paragraph"
```

The tool prints an ID and the placeholder to copy:

```
  Tagged  : lines 42-48
  Label   : Rewrote methods paragraph
  ID      : RT_20260614_191649_513

  Paste this placeholder into your response document:
  {{line:RT_20260614_191649_513}}
```

The tagged paragraphs are highlighted yellow in the manuscript so you can see at a glance what has been tagged.

#### Step 3 — Add placeholders to your response document

Open your response-to-reviewers document and paste the placeholder wherever you want the line number to appear:

> *We have rewritten the methods paragraph ({{line:RT_20260614_191649_513}}) to clarify the approach.*

You can use the same placeholder multiple times if you cite the same change in more than one place.

#### Step 4 — Sync after further edits

After making more changes to the manuscript, run:

```
python review_tracker.py sync manuscript.docx response_to_reviewers.docx
```

Every placeholder in the response document is updated to the current line number:

> *We have rewritten the methods paragraph (Lines 45–51{{line:RT_20260614_191649_513}}) to clarify the approach.*

The tracking code `{{line:...}}` stays embedded so you can keep re-syncing as you continue revising.

#### Step 5 — Render a clean copy for submission

When you are ready to submit, produce a clean copy with all tracking codes removed:

```
python review_tracker.py render response_to_reviewers.docx
```

This creates `response_to_reviewers_FINAL.docx`:

> *We have rewritten the methods paragraph (Lines 45–51) to clarify the approach.*

Your working copy is left untouched.

---

### All commands

| Command | What it does |
|---|---|
| `tag MANUSCRIPT LINE "label"` | Tag a line or range in the manuscript |
| `sync MANUSCRIPT RESPONSE` | Update all placeholders in the response document |
| `list MANUSCRIPT` | Show all tags and their current line numbers |
| `untag MANUSCRIPT ID` | Remove a tag and its highlight |
| `render RESPONSE [OUTPUT]` | Export a clean copy with tracking codes stripped |

#### `tag`

```
python review_tracker.py tag manuscript.docx 42 "label"
python review_tracker.py tag manuscript.docx 42-48 "label"
```

`LINE` can be a single line number (`42`) or a range (`42-48`). The label is for your own reference and appears in `list` output.

#### `sync`

```
python review_tracker.py sync manuscript.docx response_to_reviewers.docx
```

Updates the response document in place. Safe to run as many times as you like.

#### `list`

```
python review_tracker.py list manuscript.docx
```

Shows all tags, their current line numbers, and a `*` next to any that have shifted since the last sync.

#### `untag`

```
python review_tracker.py untag manuscript.docx RT_20260614_191649_513
```

Removes the bookmark and yellow highlight from the manuscript and deletes the tag from the sidecar file. Placeholders already in the response document are left unchanged (they will warn on the next sync).

#### `render`

```
python review_tracker.py render response_to_reviewers.docx
python review_tracker.py render response_to_reviewers.docx submission_response.docx
```

Strips all `{{line:...}}` codes, leaving only the human-readable line numbers. If no output path is given, saves as `response_to_reviewers_FINAL.docx`.

---

### Files

The tool creates one sidecar file next to your manuscript:

```
manuscript.docx
manuscript.rt_tags.json   ← tag metadata (keep this alongside the manuscript)
response_to_reviewers.docx
```

**Keep `manuscript.rt_tags.json` alongside `manuscript.docx`** — it stores the tag labels and last-known line numbers. The bookmarks embedded in the `.docx` are the source of truth for position tracking, but the sidecar is needed to know which IDs to look for.

If you share the manuscript with a co-author, send both files.

---

### Limitations

- Line numbers are estimated from the document's page width, margins, and font size. They should match what LibreOffice/Word display in the margin for standard documents, but may drift slightly in documents with complex formatting (tables, footnotes, multiple columns, mixed fonts).
- The tool works at the paragraph level — a bookmark is placed in the paragraph that contains your tagged line. If a tagged paragraph is deleted from the manuscript, the next sync will warn that the bookmark is missing.
- Placeholders must appear as plain text in the response document. Do not format part of a placeholder differently (e.g. making `{{line:` bold and the rest normal) or it may not be detected.

---

### Tips

- Run `sync` before `render` so the final document has the most current line numbers.
- Use descriptive labels when tagging — they show up in `list` output and make it easier to identify which change is which.
- You can tag the same line more than once (each gets a unique ID), but usually one tag per change is enough even if you cite it in multiple places in the response document.
