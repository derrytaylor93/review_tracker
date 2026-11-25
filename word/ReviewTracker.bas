Attribute VB_Name = "ReviewTracker"
' Review Tracker for Microsoft Word (Mac and Windows)
'
' Tracks passages in a manuscript and keeps the line numbers cited in a
' response-to-reviewers document up to date.
'
' Storage is identical to the LibreOffice extension, so files work in both:
'   * a tracked passage is a bookmark named RT_<id> around the selected text,
'     with its label in a custom document property of the same name;
'   * a citation is a bookmark named RTref_<id>_<n> around text such as
'     "Lines 45-51" in the response document.
'
' Line numbers come from Word's own layout (Range.Information), so they match
' the numbers Word shows in the margin with continuous line numbering.

Option Explicit

Private Const TAG_PREFIX As String = "RT_"
Private Const REF_PREFIX As String = "RTref_"
Private Const APP_TITLE As String = "Review Tracker"

' Word constants, written as numbers so no extra library reference is needed
Private Const WD_FIRST_CHAR_LINE As Long = 10     ' wdFirstCharacterLineNumber
Private Const WD_PAGE_NUMBER As Long = 3          ' wdActiveEndPageNumber
Private Const WD_GOTO_PAGE As Long = 1            ' wdGoToPage
Private Const WD_GOTO_ABSOLUTE As Long = 1        ' wdGoToAbsolute
Private Const WD_STAT_PAGES As Long = 2           ' wdStatisticPages
Private Const WD_PRINT_VIEW As Long = 3           ' wdPrintView
Private Const PROP_STRING As Long = 4             ' msoPropertyTypeString

' Per-run cache of "lines before page k" for each manuscript
Private cacheNames As Collection
Private cacheData As Collection


' ===== ribbon buttons ========================================================
' (Callbacks take "As Object" so the template needs no Office library reference.)

Public Sub RT_OnTrack(control As Object)
    TrackSelection
End Sub

Public Sub RT_OnInsert(control As Object)
    InsertLineReference
End Sub

Public Sub RT_OnUpdate(control As Object)
    UpdateLineNumbers
End Sub

Public Sub RT_OnNext(control As Object)
    NextTrackedPassage
End Sub

Public Sub RT_OnStop(control As Object)
    StopTrackingHere
End Sub

Public Sub RT_OnRemoveAll(control As Object)
    RemoveTrackingFromDocument
End Sub


' ===== commands (can also be run from Tools > Macro > Macros) ===============

Public Sub TrackSelection()
    Dim doc As Document, rng As Range, snippet As String, label As String, name As String
    Set doc = ActiveDocument
    Set rng = Selection.Range
    If rng.Start = rng.End Then
        MsgBox "Select the passage you changed in the manuscript first, then click Track Selection.", _
               vbInformation, APP_TITLE
        Exit Sub
    End If
    snippet = Snip(rng.Text, 50)
    label = InputBox("Give this passage a short label so you can find it later " & _
                     "(for example " & ChrW(8220) & "R1.3 revised gaze definition" & ChrW(8221) & "):", _
                     APP_TITLE, snippet)
    If Len(Trim$(label)) = 0 Then Exit Sub   ' Cancel

    name = TAG_PREFIX & NewId(doc, 6)
    doc.Bookmarks.Add name, rng
    SetLabel doc, name, Trim$(label)

    ResetCache
    MsgBox "Tracking " & ChrW(8220) & Trim$(label) & ChrW(8221) & " (" & _
           Describe(doc, doc.Bookmarks(name).Range) & ")." & vbCr & vbCr & _
           "Now switch to your response document, click where the line numbers should go " & _
           "and click Insert Line Ref.", vbInformation, APP_TITLE
End Sub


Public Sub InsertLineReference()
    Dim target As Document, d As Document
    Dim tagNames As New Collection, tagDocs As New Collection, lines As New Collection
    Dim defaultPick As Long, i As Long, n As Long
    Set target = ActiveDocument
    ResetCache

    ' Gather tracked passages from the other open documents, in line order.
    For Each d In Documents
        If Not d Is target Then
            If CountBookmarks(d, TAG_PREFIX) > 0 Then CollectTags d, tagNames, tagDocs, lines
        End If
    Next d
    n = tagNames.Count
    If n = 0 Then
        MsgBox "Open your manuscript in Word too, then try again." & vbCr & vbCr & _
               "(Review Tracker looks for an open document with passages tracked " & _
               "using Track Selection.)", vbInformation, APP_TITLE
        Exit Sub
    End If

    ' Pre-select the passage the cursor is in, in the manuscript window.
    defaultPick = 0
    For i = 1 To n
        If CursorInside(tagDocs(i), tagNames(i)) Then defaultPick = i: Exit For
    Next i

    Dim pick As Long
    pick = AskForPassage(tagNames, tagDocs, lines, defaultPick)
    If pick = 0 Then Exit Sub

    Dim styleIndex As Long
    styleIndex = StyleForDocument(target)
    If styleIndex = 0 Then Exit Sub

    Dim tagDoc As Document, tagName As String, s As Long, e As Long, txt As String
    Set tagDoc = tagDocs(pick)
    tagName = tagNames(pick)
    SpanOf tagDoc, tagDoc.Bookmarks(tagName).Range, s, e
    txt = FormatLines(styleIndex, s, e)

    Dim r As Range, refName As String
    Set r = Selection.Range
    r.Collapse 1                                   ' wdCollapseStart
    r.Text = txt
    refName = REF_PREFIX & Mid$(tagName, Len(TAG_PREFIX) + 1) & "_" & NewId(target, 3)
    target.Bookmarks.Add refName, r
    Selection.SetRange r.End, r.End
End Sub


Public Sub UpdateLineNumbers()
    Dim d As Document, bm As Bookmark, names As New Collection
    Dim updated As Long, unchanged As Long, missing As Long
    Dim touched As String, i As Long
    ResetCache

    For Each d In Documents
        Set names = New Collection
        For Each bm In d.Bookmarks
            If Left$(bm.Name, Len(REF_PREFIX)) = REF_PREFIX Then names.Add bm.Name
        Next bm
        If names.Count > 0 Then
            Dim changedHere As Boolean, wasTracking As Boolean
            changedHere = False
            wasTracking = d.TrackRevisions
            d.TrackRevisions = False
            For i = 1 To names.Count
                Select Case UpdateOne(d, CStr(names(i)))
                    Case 1: updated = updated + 1: changedHere = True
                    Case 0: unchanged = unchanged + 1
                    Case Else: missing = missing + 1
                End Select
            Next i
            d.TrackRevisions = wasTracking
            If changedHere Then touched = touched & IIf(Len(touched) > 0, ", ", "") & d.Name
        End If
    Next d

    If updated + unchanged + missing = 0 Then
        MsgBox "No line references found in the open documents. Use Insert Line Ref in your " & _
               "response document first.", vbInformation, APP_TITLE
        Exit Sub
    End If
    Dim msg As String
    msg = "Updated " & updated & " line reference" & IIf(updated = 1, "", "s") & _
          " (" & unchanged & " already up to date)."
    If Len(touched) > 0 Then msg = msg & vbCr & "Changed: " & touched & ". Remember to save."
    If missing > 0 Then
        msg = msg & vbCr & vbCr & missing & " reference" & IIf(missing = 1, " points", "s point") & _
              " to passages that were deleted, are no longer tracked, or whose manuscript " & _
              "isn't open. They were left as they are."
        MsgBox msg, vbExclamation, APP_TITLE
    Else
        MsgBox msg, vbInformation, APP_TITLE
    End If
End Sub


Public Sub NextTrackedPassage()
    ' Jump to the next tracked passage after the cursor (wraps around).
    Dim doc As Document, bm As Bookmark, best As Bookmark, first As Bookmark
    Dim here As Long
    Set doc = ActiveDocument
    here = Selection.Start
    For Each bm In doc.Bookmarks
        If IsTag(bm.Name) Then
            If first Is Nothing Then
                Set first = bm
            ElseIf bm.Range.Start < first.Range.Start Then
                Set first = bm
            End If
            If bm.Range.Start > here Then
                If best Is Nothing Then
                    Set best = bm
                ElseIf bm.Range.Start < best.Range.Start Then
                    Set best = bm
                End If
            End If
        End If
    Next bm
    If first Is Nothing Then
        MsgBox "There are no tracked passages in this document yet. Select some text and " & _
               "click Track Selection.", vbInformation, APP_TITLE
        Exit Sub
    End If
    If best Is Nothing Then Set best = first
    best.Range.Select
    ResetCache
    Application.StatusBar = APP_TITLE & ": " & Describe(doc, best.Range) & "  " & _
                            ChrW(183) & "  " & GetLabel(doc, best.Name)
End Sub


Public Sub StopTrackingHere()
    Dim doc As Document, bm As Bookmark, name As String
    Set doc = ActiveDocument
    For Each bm In doc.Bookmarks
        If IsTag(bm.Name) Then
            If Selection.Start >= bm.Range.Start And Selection.Start <= bm.Range.End Then
                name = bm.Name
                Exit For
            End If
        End If
    Next bm
    If Len(name) = 0 Then
        MsgBox "Put the cursor inside a tracked passage first (Next Passage jumps to them).", _
               vbInformation, APP_TITLE
        Exit Sub
    End If
    If MsgBox("Stop tracking " & ChrW(8220) & GetLabel(doc, name) & ChrW(8221) & "?" & vbCr & vbCr & _
              "Citations of it in your response document will no longer update.", _
              vbOKCancel + vbQuestion, APP_TITLE) <> vbOK Then Exit Sub
    doc.Bookmarks(name).Delete
    RemoveLabel doc, name
End Sub


Public Sub RemoveTrackingFromDocument()
    Dim doc As Document, i As Long, n As Long, nm As String
    Set doc = ActiveDocument
    If MsgBox("Remove all Review Tracker bookmarks from " & ChrW(8220) & doc.Name & ChrW(8221) & "?" & _
              vbCr & vbCr & "The text and line numbers stay exactly as they are, but they will no " & _
              "longer update. Usually you don't need this: the bookmarks are invisible and safe " & _
              "to submit." & vbCr & vbCr & "Tip: do this on a copy.", _
              vbOKCancel + vbQuestion, APP_TITLE) <> vbOK Then Exit Sub
    For i = doc.Bookmarks.Count To 1 Step -1
        nm = doc.Bookmarks(i).Name
        If IsTag(nm) Or Left$(nm, Len(REF_PREFIX)) = REF_PREFIX Then
            doc.Bookmarks(i).Delete
            RemoveLabel doc, nm
            n = n + 1
        End If
    Next i
    MsgBox "Removed " & n & " bookmark" & IIf(n = 1, "", "s") & ".", vbInformation, APP_TITLE
End Sub


' ===== citations =============================================================

' Returns 1 if changed, 0 if already right, -1 if the passage can't be found.
Private Function UpdateOne(d As Document, refName As String) As Long
    Dim tagName As String, m As Document, s As Long, e As Long
    Dim r As Range, oldText As String, newText As String
    tagName = TAG_PREFIX & Split(Mid$(refName, Len(REF_PREFIX) + 1), "_")(0)
    Set m = FindManuscript(tagName)
    If m Is Nothing Then UpdateOne = -1: Exit Function
    If Len(m.Bookmarks(tagName).Range.Text) = 0 Then UpdateOne = -1: Exit Function

    SpanOf m, m.Bookmarks(tagName).Range, s, e
    Set r = d.Bookmarks(refName).Range
    oldText = r.Text
    newText = FormatLines(DetectStyle(oldText), s, e)
    If newText = oldText Then UpdateOne = 0: Exit Function
    r.Text = newText
    d.Bookmarks.Add refName, r
    UpdateOne = 1
End Function

Private Function FindManuscript(tagName As String) As Document
    Dim d As Document
    For Each d In Documents
        If d.Bookmarks.Exists(tagName) Then Set FindManuscript = d: Exit Function
    Next d
End Function

' Style 1 "Lines 45-51", 2 "lines 45-51", 3 "L45-51", 4 "45-51"
Private Function FormatLines(styleIndex As Long, s As Long, e As Long) As String
    Dim one As String, many As String
    Select Case styleIndex
        Case 2: one = "line ": many = "lines "
        Case 3: one = "L": many = "L"
        Case 4: one = "": many = ""
        Case Else: one = "Line ": many = "Lines "
    End Select
    If e <= s Then
        FormatLines = one & s
    Else
        FormatLines = many & s & ChrW(8211) & e
    End If
End Function

Private Function DetectStyle(txt As String) As Long
    Dim t As String
    t = Trim$(txt)
    If Left$(t, 4) = "Line" Then
        DetectStyle = 1
    ElseIf Left$(t, 4) = "line" Then
        DetectStyle = 2
    ElseIf t Like "L#*" Or t Like "L #*" Then
        DetectStyle = 3
    ElseIf t Like "#*" Then
        DetectStyle = 4
    Else
        DetectStyle = 1
    End If
End Function

' The style already used in this document, or ask once.
Private Function StyleForDocument(doc As Document) As Long
    Dim bm As Bookmark, answer As String
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, Len(REF_PREFIX)) = REF_PREFIX Then
            StyleForDocument = DetectStyle(bm.Range.Text)
            Exit Function
        End If
    Next bm
    answer = InputBox("How should line numbers look in this document?" & vbCr & vbCr & _
                      "1   Lines 45" & ChrW(8211) & "51" & vbCr & _
                      "2   lines 45" & ChrW(8211) & "51" & vbCr & _
                      "3   L45" & ChrW(8211) & "51" & vbCr & _
                      "4   45" & ChrW(8211) & "51" & vbCr & vbCr & _
                      "Type 1, 2, 3 or 4. (You'll only be asked once per document.)", APP_TITLE, "1")
    If Len(Trim$(answer)) = 0 Then StyleForDocument = 0: Exit Function
    If Not IsNumeric(answer) Then StyleForDocument = 1: Exit Function
    StyleForDocument = CLng(answer)
    If StyleForDocument < 1 Or StyleForDocument > 4 Then StyleForDocument = 1
End Function

Private Function AskForPassage(tagNames As Collection, tagDocs As Collection, lines As Collection, _
                               defaultPick As Long) As Long
    Dim n As Long, i As Long, first As Long, last As Long, prompt As String, answer As String
    Dim manyDocs As Boolean
    n = tagNames.Count
    For i = 2 To n
        If Not tagDocs(i) Is tagDocs(1) Then manyDocs = True
    Next i

    ' InputBox prompts are limited in length, so show at most 12 around the default.
    first = 1: last = n
    If n > 12 Then
        first = IIf(defaultPick > 0, defaultPick - 5, 1)
        If first < 1 Then first = 1
        last = first + 11
        If last > n Then last = n: first = n - 11
    End If
    prompt = "Which tracked passage are you citing? Type its number." & vbCr & vbCr
    For i = first To last
        prompt = prompt & i & "   " & lines(i) & "   " & ChrW(183) & "   " & _
                 Snip(GetLabel(tagDocs(i), CStr(tagNames(i))), 40)
        If manyDocs Then prompt = prompt & "  [" & tagDocs(i).Name & "]"
        prompt = prompt & vbCr
    Next i
    If n > 12 Then prompt = prompt & "(" & n & " passages in total)" & vbCr
    If defaultPick = 0 Then
        prompt = prompt & vbCr & "Tip: click inside a tracked passage in the manuscript and it " & _
                 "will be pre-selected here."
        defaultPick = 1
    End If
    answer = InputBox(Left$(prompt, 1000), APP_TITLE, CStr(defaultPick))
    If Len(Trim$(answer)) = 0 Or Not IsNumeric(answer) Then Exit Function
    If CLng(answer) < 1 Or CLng(answer) > n Then Exit Function
    AskForPassage = CLng(answer)
End Function

Private Sub CollectTags(d As Document, tagNames As Collection, tagDocs As Collection, lines As Collection)
    ' Adds d's tags to the collections, keeping them sorted by start line.
    Dim bm As Bookmark, s As Long, e As Long, i As Long, pos As Long
    For Each bm In d.Bookmarks
        If IsTag(bm.Name) Then
            SpanOf d, bm.Range, s, e
            pos = tagNames.Count + 1
            For i = 1 To tagNames.Count
                If tagDocs(i) Is d Then
                    If StartLineOf(tagDocs(i), CStr(tagNames(i))) > s Then pos = i: Exit For
                End If
            Next i
            If pos > tagNames.Count Then
                tagNames.Add bm.Name: tagDocs.Add d: lines.Add FormatLines(1, s, e)
            Else
                tagNames.Add bm.Name, Before:=pos: tagDocs.Add d, Before:=pos
                lines.Add FormatLines(1, s, e), Before:=pos
            End If
        End If
    Next bm
End Sub

Private Function StartLineOf(d As Document, tagName As String) As Long
    Dim s As Long, e As Long
    SpanOf d, d.Bookmarks(tagName).Range, s, e
    StartLineOf = s
End Function

Private Function CursorInside(d As Document, tagName As String) As Boolean
    Dim sel As Long, r As Range
    On Error GoTo nope
    sel = d.Windows(1).Selection.Start
    Set r = d.Bookmarks(tagName).Range
    CursorInside = (sel >= r.Start And sel <= r.End)
    Exit Function
nope:
    CursorInside = False
End Function


' ===== line numbers ==========================================================

Private Function Describe(doc As Document, r As Range) As String
    Dim s As Long, e As Long
    SpanOf doc, r, s, e
    Describe = FormatLines(1, s, e)
End Function

' Start and end line of a range, counted continuously from the start of the document.
Private Sub SpanOf(doc As Document, r As Range, ByRef s As Long, ByRef e As Long)
    Dim endPos As Long
    s = LineAt(doc, r.Start)
    endPos = r.End - 1
    If endPos < r.Start Then endPos = r.Start
    e = LineAt(doc, endPos)
    If e < s Then e = s
End Sub

Private Function LineAt(doc As Document, pos As Long) As Long
    Dim r As Range, page As Long, onPage As Long, before As Variant
    Set r = doc.Range(pos, pos)
    page = r.Information(WD_PAGE_NUMBER)
    onPage = r.Information(WD_FIRST_CHAR_LINE)
    before = LinesBeforePages(doc)
    If page >= LBound(before) And page <= UBound(before) Then
        LineAt = before(page) + onPage
    Else
        LineAt = onPage
    End If
End Function

' Array where element k = number of lines on pages 1..k-1.
' Word reports line numbers per page; this turns them into continuous numbers.
' (If Word already reports continuous numbers, all elements stay 0.)
Private Function LinesBeforePages(doc As Document) As Variant
    Dim key As String, i As Long
    If cacheNames Is Nothing Then ResetCache
    key = doc.FullName
    For i = 1 To cacheNames.Count
        If cacheNames(i) = key Then LinesBeforePages = cacheData(i): Exit Function
    Next i

    Dim pages As Long, k As Long, before() As Long, r As Range, continuous As Boolean
    On Error Resume Next
    If doc.Windows(1).View.Type <> WD_PRINT_VIEW Then doc.Windows(1).View.Type = WD_PRINT_VIEW
    On Error GoTo 0
    doc.Repaginate
    pages = doc.ComputeStatistics(WD_STAT_PAGES)
    ReDim before(1 To pages + 1)
    before(1) = 0

    ' Does Word number lines per page here (line 1 at the top of page 2)?
    continuous = False
    If pages > 1 Then
        Set r = doc.GoTo(What:=WD_GOTO_PAGE, Which:=WD_GOTO_ABSOLUTE, Count:=2)
        If r.Information(WD_FIRST_CHAR_LINE) > 1 Then continuous = True
    End If

    For k = 1 To pages
        If k = 1 Or continuous Then
            If k > 1 Then before(k) = 0
        Else
            ' last line on page k-1 = line of the character just before page k starts
            Set r = doc.GoTo(What:=WD_GOTO_PAGE, Which:=WD_GOTO_ABSOLUTE, Count:=k)
            If r.Start > 0 Then
                Set r = doc.Range(r.Start - 1, r.Start - 1)
                before(k) = before(k - 1) + r.Information(WD_FIRST_CHAR_LINE)
            Else
                before(k) = before(k - 1)
            End If
        End If
    Next k
    before(pages + 1) = before(pages)

    cacheNames.Add key
    cacheData.Add before
    LinesBeforePages = before
End Function

Private Sub ResetCache()
    Set cacheNames = New Collection
    Set cacheData = New Collection
End Sub


' ===== small helpers =========================================================

Private Function IsTag(nm As String) As Boolean
    IsTag = (Left$(nm, Len(TAG_PREFIX)) = TAG_PREFIX) And (Left$(nm, Len(REF_PREFIX)) <> REF_PREFIX)
End Function

Private Function CountBookmarks(d As Document, prefix As String) As Long
    Dim bm As Bookmark
    For Each bm In d.Bookmarks
        If prefix = TAG_PREFIX Then
            If IsTag(bm.Name) Then CountBookmarks = CountBookmarks + 1
        ElseIf Left$(bm.Name, Len(prefix)) = prefix Then
            CountBookmarks = CountBookmarks + 1
        End If
    Next bm
End Function

Private Function NewId(doc As Document, length As Long) As String
    Const ALPHABET As String = "abcdefghijklmnopqrstuvwxyz0123456789"
    Dim i As Long, s As String
    Randomize
    Do
        s = ""
        For i = 1 To length
            s = s & Mid$(ALPHABET, Int(Rnd() * Len(ALPHABET)) + 1, 1)
        Next i
    Loop While doc.Bookmarks.Exists(TAG_PREFIX & s) Or InStr(1, BookmarkNames(doc), "_" & s & "|") > 0
    NewId = s
End Function

Private Function BookmarkNames(doc As Document) As String
    Dim bm As Bookmark
    For Each bm In doc.Bookmarks
        BookmarkNames = BookmarkNames & bm.Name & "|"
    Next bm
End Function

Private Function Snip(txt As String, maxLen As Long) As String
    Dim t As String
    t = Replace(Replace(Replace(txt, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Trim$(t)
    If Len(t) > maxLen Then t = Left$(t, maxLen) & ChrW(8230)
    Snip = t
End Function

Private Function GetLabel(doc As Document, nm As String) As String
    On Error GoTo none
    GetLabel = doc.CustomDocumentProperties(nm).Value
    Exit Function
none:
    GetLabel = ""
End Function

Private Sub SetLabel(doc As Document, nm As String, label As String)
    On Error Resume Next
    doc.CustomDocumentProperties(nm).Delete
    On Error GoTo 0
    doc.CustomDocumentProperties.Add Name:=nm, LinkToContent:=False, Type:=PROP_STRING, Value:=label
End Sub

Private Sub RemoveLabel(doc As Document, nm As String)
    On Error Resume Next
    doc.CustomDocumentProperties(nm).Delete
End Sub
