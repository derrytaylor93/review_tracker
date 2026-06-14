# Review Tracker

A command-line tool for tracking line-number citations between a manuscript and a response-to-reviewers document. When you revise your manuscript in response to peer review comments, line numbers shift — this tool keeps your citations up to date automatically.

Works on **Windows, Mac, and Linux** with LibreOffice or Microsoft Word (`.docx` files).

---

## The problem

When responding to peer reviewers, it's common to write things like:

> *"We revised the definition on **Lines 42–48** to address this comment."*

But as you continue revising the manuscript in response to other comments, those line numbers change. Manually hunting down and updating every citation in your response document is tedious and error-prone.

## How it works

1. You **tag** a change in the manuscript by specifying the line number. The tool embeds an invisible bookmark in the `.docx` file so it can track that passage even as surrounding text shifts.
2. You paste a **placeholder** (e.g. `{{line:RT_20260614_191649_513}}`) into your response document wherever you want to cite that line.
3. Whenever you want to update the citations, you run **sync**. The tool finds each bookmark's current line number and updates all the placeholders in your response document.
4. When you're ready to submit, **render** produces a clean copy with all the tracking codes stripped out.

---

## Installation

**Requirements:** Python 3.10+ and one library.

```
pip install python-docx
```

Download `review_tracker.py` and place it in the same folder as your manuscript and response document, or anywhere on your PATH.

---

## Workflow

### Step 1 — Enable line numbers in your manuscript

Turn on continuous line numbering so you can read off the line numbers you want to tag.

- **LibreOffice Writer:** Format → Line Numbers → check *Show Line Numbers*, set *Interval* to 1
- **Microsoft Word:** Layout → Page Setup → Line Numbers → Continuous

### Step 2 — Tag a change

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

### Step 3 — Add placeholders to your response document

Open your response-to-reviewers document and paste the placeholder wherever you want the line number to appear:

> *We have rewritten the methods paragraph ({{line:RT_20260614_191649_513}}) to clarify the approach.*

You can use the same placeholder multiple times if you cite the same change in more than one place.

### Step 4 — Sync after further edits

After making more changes to the manuscript, run:

```
python review_tracker.py sync manuscript.docx response_to_reviewers.docx
```

Every placeholder in the response document is updated to the current line number:

> *We have rewritten the methods paragraph (Lines 45–51{{line:RT_20260614_191649_513}}) to clarify the approach.*

The tracking code `{{line:...}}` stays embedded so you can keep re-syncing as you continue revising.

### Step 5 — Render a clean copy for submission

When you are ready to submit, produce a clean copy with all tracking codes removed:

```
python review_tracker.py render response_to_reviewers.docx
```

This creates `response_to_reviewers_FINAL.docx`:

> *We have rewritten the methods paragraph (Lines 45–51) to clarify the approach.*

Your working copy is left untouched.

---

## All commands

| Command | What it does |
|---|---|
| `tag MANUSCRIPT LINE "label"` | Tag a line or range in the manuscript |
| `sync MANUSCRIPT RESPONSE` | Update all placeholders in the response document |
| `list MANUSCRIPT` | Show all tags and their current line numbers |
| `untag MANUSCRIPT ID` | Remove a tag and its highlight |
| `render RESPONSE [OUTPUT]` | Export a clean copy with tracking codes stripped |

### `tag`

```
python review_tracker.py tag manuscript.docx 42 "label"
python review_tracker.py tag manuscript.docx 42-48 "label"
```

`LINE` can be a single line number (`42`) or a range (`42-48`). The label is for your own reference and appears in `list` output.

### `sync`

```
python review_tracker.py sync manuscript.docx response_to_reviewers.docx
```

Updates the response document in place. Safe to run as many times as you like.

### `list`

```
python review_tracker.py list manuscript.docx
```

Shows all tags, their current line numbers, and a `*` next to any that have shifted since the last sync.

### `untag`

```
python review_tracker.py untag manuscript.docx RT_20260614_191649_513
```

Removes the bookmark and yellow highlight from the manuscript and deletes the tag from the sidecar file. Placeholders already in the response document are left unchanged (they will warn on the next sync).

### `render`

```
python review_tracker.py render response_to_reviewers.docx
python review_tracker.py render response_to_reviewers.docx submission_response.docx
```

Strips all `{{line:...}}` codes, leaving only the human-readable line numbers. If no output path is given, saves as `response_to_reviewers_FINAL.docx`.

---

## Files

The tool creates one sidecar file next to your manuscript:

```
manuscript.docx
manuscript.rt_tags.json   ← tag metadata (keep this alongside the manuscript)
response_to_reviewers.docx
```

**Keep `manuscript.rt_tags.json` alongside `manuscript.docx`** — it stores the tag labels and last-known line numbers. The bookmarks embedded in the `.docx` are the source of truth for position tracking, but the sidecar is needed to know which IDs to look for.

If you share the manuscript with a co-author, send both files.

---

## Limitations

- Line numbers are estimated from the document's page width, margins, and font size. They should match what LibreOffice/Word display in the margin for standard documents, but may drift slightly in documents with complex formatting (tables, footnotes, multiple columns, mixed fonts).
- The tool works at the paragraph level — a bookmark is placed in the paragraph that contains your tagged line. If a tagged paragraph is deleted from the manuscript, the next sync will warn that the bookmark is missing.
- Placeholders must appear as plain text in the response document. Do not format part of a placeholder differently (e.g. making `{{line:` bold and the rest normal) or it may not be detected.

---

## Tips

- Run `sync` before `render` so the final document has the most current line numbers.
- Use descriptive labels when tagging — they show up in `list` output and make it easier to identify which change is which.
- You can tag the same line more than once (each gets a unique ID), but usually one tag per change is enough even if you cite it in multiple places in the response document.
