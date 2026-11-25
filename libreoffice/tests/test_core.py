"""End-to-end test of the Review Tracker core against a running LibreOffice.

Start LibreOffice first:
    soffice --headless --norestore --accept="socket,host=localhost,port=2002;urp;" &
Then:
    python3 libreoffice/tests/test_core.py MANUSCRIPT.docx RESPONSE.docx WORKDIR
"""
import os
import shutil
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "extension", "Scripts", "python"))
import ReviewTracker as RT  # noqa: E402


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(30):
        try:
            ctx = resolver.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except Exception:
            time.sleep(1)
    raise SystemExit("LibreOffice is not running")


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def load(desktop, path):
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0, ())


def save_docx(doc, path):
    doc.storeToURL(uno.systemPathToFileUrl(path), (prop("FilterName", "MS Word 2007 XML"), prop("Overwrite", True)))


def find(doc, needle):
    sd = doc.createSearchDescriptor()
    sd.SearchString = needle
    found = doc.findFirst(sd)
    assert found is not None, needle
    return found


def pdf_line(pdf_dir, doc_path, needle):
    """Line number LibreOffice prints in the margin for the line containing needle."""
    subprocess.run(["pdftotext", "-layout", doc_path, pdf_dir + "/out.txt"], check=True)
    for line in open(pdf_dir + "/out.txt", encoding="utf-8"):
        if needle in line:
            num = line.strip().split()[0]
            if num.isdigit():
                return int(num)
    return None


def export_pdf(doc, path):
    doc.storeToURL(uno.systemPathToFileUrl(path), (prop("FilterName", "writer_pdf_Export"),))


def main(ms_src, resp_src, work):
    os.makedirs(work, exist_ok=True)
    ms_path = os.path.join(work, "manuscript.docx")
    resp_path = os.path.join(work, "response.docx")
    shutil.copy(ms_src, ms_path)
    shutil.copy(resp_src, resp_path)
    desktop = connect()
    ok = True

    def check(cond, what):
        nonlocal ok
        print(("PASS " if cond else "FAIL ") + what)
        ok &= bool(cond)

    ms = load(desktop, ms_path)
    resp = load(desktop, resp_path)

    # 1. Track two passages: a phrase inside one line, and a multi-line sentence.
    a = RT.track_range(ms, find(ms, "Calculations were conducted using the ICC()"), "R1.2 ICC method")
    rng = find(ms, "The ability to read others")
    cur = rng.getText().createTextCursorByRange(rng.getStart())
    cur.gotoEndOfParagraph(True)
    b = RT.track_range(ms, cur, "Abstract rewritten")
    tags = {t[0]: t for t in RT.list_tags(ms)}
    check(tags[a][1] == "R1.2 ICC method", "label stored")
    print("   tags:", [(t[1], t[2], t[3]) for t in tags.values()])

    # Compare with the margin numbers LibreOffice prints in a PDF export.
    export_pdf(ms, work + "/ms.pdf")
    expect_a = pdf_line(work, work + "/ms.pdf", "Calculations were conducted using the ICC()")
    expect_b = pdf_line(work, work + "/ms.pdf", "The ability to read others")
    check(tags[a][2] == expect_a, "line of passage A matches margin (%s vs %s)" % (tags[a][2], expect_a))
    check(tags[b][2] == expect_b, "start of passage B matches margin (%s vs %s)" % (tags[b][2], expect_b))
    check(tags[b][3] > tags[b][2], "passage B spans several lines (%s-%s)" % (tags[b][2], tags[b][3]))

    # 2. Cite them in the response document.
    end = resp.getText().getEnd()
    resp.getText().insertString(end, "\nWe clarified the ICC calculation (", False)
    name_a, c = RT.insert_ref(resp, resp.getText().getEnd(), a, tags[a][2], tags[a][3], 0)
    resp.getText().insertString(resp.getText().getEnd(), ") and the abstract (", False)
    RT.insert_ref(resp, resp.getText().getEnd(), b, tags[b][2], tags[b][3], 1)
    resp.getText().insertString(resp.getText().getEnd(), ").", False)
    last = resp.getText().getEnd()
    para = resp.getText().createTextCursorByRange(last)
    para.gotoStartOfParagraph(True)
    print("   response: ..." + para.getString()[-80:])
    check("(Line %d)" % tags[a][2] in para.getString(), "citation A inserted")
    check("(lines %d–%d)" % (tags[b][2], tags[b][3]) in para.getString(), "citation B inserted, lower-case style")

    # 3. Edit the manuscript: add three paragraphs near the top.
    top = find(ms, "Abstract")
    t = top.getText()
    c = t.createTextCursorByRange(top.getStart())
    for _ in range(3):
        t.insertString(c, "New paragraph added during revision.", False)
        t.insertControlCharacter(c, 0, False)  # PARAGRAPH_BREAK
    updated, unchanged, missing, _ = RT.update_refs(desktop, [resp])
    check(updated == 2 and not missing, "update changed both citations (%d updated, missing %s)" % (updated, missing))
    export_pdf(ms, work + "/ms2.pdf")
    expect_a2 = pdf_line(work, work + "/ms2.pdf", "Calculations were conducted using the ICC()")
    para = resp.getText().createTextCursorByRange(resp.getText().getEnd())
    para.gotoStartOfParagraph(True)
    print("   response: ..." + para.getString()[-80:])
    check("(Line %d)" % expect_a2 in para.getString(), "citation A now matches margin line %s" % expect_a2)
    check(expect_a2 == expect_a + 3, "passage A moved down three lines")

    # 4. Running update again changes nothing.
    updated, unchanged, missing, _ = RT.update_refs(desktop, [resp])
    check(updated == 0 and unchanged == 2, "second update is a no-op")

    # 5. Save as .docx, reopen, and update again.
    ms.store()
    resp.store()
    ms.close(True)
    resp.close(True)
    ms = load(desktop, ms_path)
    resp = load(desktop, resp_path)
    tags2 = {t[0]: t for t in RT.list_tags(ms)}
    check(set(tags2) == {a, b}, "tracked passages survive saving as .docx")
    check(tags2.get(a, [None, None])[1] == "R1.2 ICC method", "labels survive saving as .docx")
    check(len(RT._bookmarks(resp, RT.REF_PREFIX)) == 2, "citations survive saving as .docx")
    t = ms.getText()
    c = t.createTextCursorByRange(find(ms, "Abstract").getStart())
    t.insertString(c, "One more paragraph.", False)
    t.insertControlCharacter(c, 0, False)
    updated, unchanged, missing, _ = RT.update_refs(desktop, [resp])
    para = resp.getText().createTextCursorByRange(resp.getText().getEnd())
    para.gotoStartOfParagraph(True)
    print("   response: ..." + para.getString()[-80:])
    check(updated == 2 and "(Line %d)" % (expect_a2 + 1) in para.getString(),
          "update works after reopening")

    # 6. Finalise strips the bookmarks but keeps the text.
    before = para.getString()
    n = RT.finalise(resp)
    para = resp.getText().createTextCursorByRange(resp.getText().getEnd())
    para.gotoStartOfParagraph(True)
    check(n == 2 and para.getString() == before and not RT._bookmarks(resp, RT.REF_PREFIX),
          "finalise removes bookmarks, keeps text")

    # 7. Remove a tag.
    RT.remove_tag(ms, a)
    check([t[0] for t in RT.list_tags(ms)] == [b], "stop tracking removes the passage")

    # 8. A passage whose text is deleted is reported, not cited at the empty spot.
    RT.insert_ref(resp, resp.getText().getEnd(), b, 1, 1, 0)
    #    (With Track Changes recording, deleted text stays in the document until the
    #    deletion is accepted, so switch recording off for this step.)
    ms.setPropertyValue("RecordChanges", False)
    c = ms.getText().createTextCursorByRange(ms.getBookmarks().getByName(b).getAnchor())
    c.setString("")
    updated, unchanged, missing, _ = RT.update_refs(desktop, [resp])
    check(missing == {RT._tag_id(b)} and updated == 0, "deleted passage is reported as missing")

    ms.close(True)
    resp.close(True)
    print("\nALL PASSED" if ok else "\nSOME CHECKS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:4]))
