Attribute VB_Name = "CitationLinker"
'==============================================================================
' CitationLinker.bas
'------------------------------------------------------------------------------
' Hyperlinks every legal authority in the active Word document, and removes
' those links again on demand. Detection is delegated to citation_extractor.py
' (your existing tool) through word_cite_bridge.py, so there is one source of
' truth for citation parsing.
'
' MACROS YOU RUN:
'   AddCitationLinks       - detect + hyperlink every authority (idempotent)
'   RemoveCitationLinks    - remove only the links this tool added (recommended)
'   RemoveAllHyperlinks    - remove EVERY hyperlink in the body (asks first)
'   ToggleCitationLinks    - Ctrl+Shift+H: remove this tool's links if any are
'                            present, otherwise apply them
'
' SETUP: edit the four Const lines below, then put word_cite_bridge.py and
' citation_extractor.py together in SCRIPT_DIR. See SETUP.md.
'
' Links added by this tool are tagged with a ScreenTip that begins with
' SCREENTIP_PREFIX, which is how RemoveCitationLinks finds them precisely.
'==============================================================================
Option Explicit

' ---- CONFIGURE THESE -------------------------------------------------------
Private Const PYTHON_EXE As String = "python"             ' or "py", or a full path to python.exe
Private Const SCRIPT_DIR As String = "C:\Users\ZCoderre\Apps\Workup Search"  ' folder holding the two .py files
Private Const REPO_JSON As String = ""                    ' full path to citation_repo.json, or "" to disable
Private Const SCREENTIP_PREFIX As String = "CiteLink:: "  ' tag identifying our links
' ----------------------------------------------------------------------------

' Result of normalizing a paragraph's raw text to the same plain text the
' bridge produced, plus a map from each normalized char to its raw index.
Private Type NormResult
    norm As String
    n As Long
    map() As Long          ' 0-based: map(j) = raw char index of normalized char j
End Type

Private Type CiteRow
    blk As Long
    s As Long
    e As Long
    typ As String
    url As String
    txt As String
End Type


'==============================================================================
' PUBLIC MACROS
'==============================================================================

Public Sub AddCitationLinks()
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub

    ' Re-running should not stack links, so clear ours first.
    RemoveCitationLinks_Quiet doc

    Dim n As Long
    n = doc.Paragraphs.Count
    If n = 0 Then
        MsgBox "The document has no paragraphs to process.", vbInformation, "Citation Linker"
        Exit Sub
    End If

    Dim prng() As Range
    Dim hasField() As Boolean
    Dim html() As String
    ReDim prng(1 To n)
    ReDim hasField(1 To n)
    ReDim html(1 To n)

    Dim p As Paragraph
    Dim i As Long
    Dim raw As String
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        Set prng(i) = p.Range
        hasField(i) = (p.Range.Fields.Count > 0) _
                   Or (p.Range.Footnotes.Count > 0) _
                   Or (p.Range.InlineShapes.Count > 0)
        raw = ParagraphRawText(p.Range)
        html(i) = "<p>" & EscapeHtml(raw) & "</p>"
    Next p

    Dim docHtml As String
    docHtml = Join(html, vbLf)

    ' Temp file paths.
    Dim tmpIn As String, tmpOut As String
    tmpIn = Environ$("TEMP") & "\citelink_in.html"
    tmpOut = Environ$("TEMP") & "\citelink_out.tsv"

    WriteUtf8File tmpIn, docHtml

    ' Run the bridge and wait.
    Dim cmd As String
    cmd = Q(PYTHON_EXE) & " " & Q(SCRIPT_DIR & "\word_cite_bridge.py") & _
          " " & Q(tmpIn) & " " & Q(tmpOut)
    If Len(REPO_JSON) > 0 Then cmd = cmd & " " & Q(REPO_JSON)

    Dim rc As Long
    rc = RunAndWait(cmd)
    If rc <> 0 Then
        MsgBox "The citation bridge did not run (exit code " & rc & ")." & vbCrLf & vbCrLf & _
               "Check PYTHON_EXE and SCRIPT_DIR at the top of the module." & vbCrLf & _
               "Command was:" & vbCrLf & cmd, vbExclamation, "Citation Linker"
        Exit Sub
    End If

    Dim tsv As String
    tsv = ReadUtf8File(tmpOut)
    If Len(Trim$(tsv)) = 0 Then
        MsgBox "No legal authorities were detected.", vbInformation, "Citation Linker"
        Exit Sub
    End If

    ' Parse rows.
    Dim lines() As String
    lines = Split(tsv, vbLf)

    Dim rows() As CiteRow
    ReDim rows(0 To UBound(lines))
    Dim cnt As Long
    Dim f() As String
    cnt = 0
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            f = Split(lines(i), vbTab)
            If UBound(f) >= 5 Then
                rows(cnt).blk = CLng(f(0))
                rows(cnt).s = CLng(f(1))
                rows(cnt).e = CLng(f(2))
                rows(cnt).typ = f(3)
                rows(cnt).url = f(4)
                rows(cnt).txt = f(5)
                cnt = cnt + 1
            End If
        End If
    Next i
    If cnt = 0 Then
        MsgBox "No legal authorities were detected.", vbInformation, "Citation Linker"
        Exit Sub
    End If
    ReDim Preserve rows(0 To cnt - 1)

    SortRows rows
    Dim keep() As CiteRow
    keep = FilterOverlaps(rows)

    ' Apply links in reverse document order so any positional shift from a
    ' hyperlink field only affects text to the right of spans not yet linked.
    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    Dim added As Long
    Dim curBlk As Long, hasN As NormResult
    curBlk = -1
    added = 0

    Dim k As Long
    For k = UBound(keep) To LBound(keep) Step -1
        Dim r As CiteRow
        r = keep(k)
        Dim paraIdx As Long
        paraIdx = r.blk + 1
        If paraIdx < 1 Or paraIdx > n Then GoTo NextK

        Dim placed As Boolean
        placed = False

        If Not hasField(paraIdx) Then
            If r.blk <> curBlk Then
                hasN = NormalizeAndMap(ParagraphRawText(prng(paraIdx)))
                curBlk = r.blk
            End If
            If r.s >= 0 And r.e >= 1 And r.e <= hasN.n And r.s < r.e Then
                Dim aStart As Long, aEnd As Long
                aStart = prng(paraIdx).Start + hasN.map(r.s)
                aEnd = prng(paraIdx).Start + hasN.map(r.e - 1) + 1
                If aEnd > aStart Then
                    Dim rng As Range
                    Set rng = ActiveDocument.Range(aStart, aEnd)
                    If AddLink(rng, r.url, r.typ) Then
                        added = added + 1
                        placed = True
                    End If
                End If
            End If
        End If

        If Not placed Then
            ' Fallback: locate the literal text inside the paragraph.
            If FindAndLink(prng(paraIdx), r.txt, r.url, r.typ) Then
                added = added + 1
            End If
        End If
NextK:
    Next k

CleanUp:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Stopped after an error: " & Err.Description, vbExclamation, "Citation Linker"
    Else
        MsgBox "Linked " & added & " citation" & IIf(added = 1, "", "s") & ".", _
               vbInformation, "Citation Linker"
    End If
End Sub


Public Sub RemoveCitationLinks()
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    Dim removed As Long
    removed = RemoveCitationLinks_Quiet(doc)
    MsgBox "Removed " & removed & " citation link" & IIf(removed = 1, "", "s") & ".", _
           vbInformation, "Citation Linker"
End Sub


' Toggle for the keyboard shortcut: if the document already has any of this
' tool's citation links, remove them; otherwise detect and apply them. A
' "mixed" document (some cites linked, some not) has citation links present, so
' it removes on this press and applies on the next.
Public Sub ToggleCitationLinks()
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub

    If HasCitationLinks(doc) Then
        RemoveCitationLinks
    Else
        AddCitationLinks
    End If
End Sub


Public Sub RemoveAllHyperlinks()
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub

    Dim total As Long
    total = doc.Hyperlinks.Count
    If total = 0 Then
        MsgBox "There are no hyperlinks in the body of this document.", _
               vbInformation, "Citation Linker"
        Exit Sub
    End If

    If MsgBox("Remove ALL " & total & " hyperlink" & IIf(total = 1, "", "s") & _
              " from the body, including any not added by this tool?", _
              vbYesNo + vbQuestion, "Citation Linker") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Dim i As Long, rng As Range
    For i = doc.Hyperlinks.Count To 1 Step -1
        Set rng = doc.Hyperlinks(i).Range
        doc.Hyperlinks(i).Delete
        ResetLinkFormatting rng
    Next i
    Application.ScreenUpdating = True

    MsgBox "Removed " & total & " hyperlink" & IIf(total = 1, "", "s") & ".", _
           vbInformation, "Citation Linker"
End Sub


' Quiet variant of RemoveAllHyperlinks for automated callers (e.g. the
' review-on-close flow). Removes EVERY hyperlink in the body with no
' confirmation and no result dialog. Returns the number removed.
Public Function RemoveAllHyperlinks_Quiet(ByVal doc As Document) As Long
    If doc Is Nothing Then Exit Function

    Dim removed As Long: removed = 0
    Dim i As Long, rng As Range
    Application.ScreenUpdating = False
    For i = doc.Hyperlinks.Count To 1 Step -1
        Set rng = doc.Hyperlinks(i).Range
        doc.Hyperlinks(i).Delete
        ResetLinkFormatting rng
        removed = removed + 1
    Next i
    Application.ScreenUpdating = True

    RemoveAllHyperlinks_Quiet = removed
End Function


'==============================================================================
' CORE HELPERS
'==============================================================================

' True if the document contains at least one hyperlink added by this tool
' (identified by the SCREENTIP_PREFIX tag). Used by ToggleCitationLinks.
Private Function HasCitationLinks(ByVal doc As Document) As Boolean
    Dim i As Long
    For i = 1 To doc.Hyperlinks.Count
        If Left$(doc.Hyperlinks(i).ScreenTip, Len(SCREENTIP_PREFIX)) = SCREENTIP_PREFIX Then
            HasCitationLinks = True
            Exit Function
        End If
    Next i
End Function


Private Function RemoveCitationLinks_Quiet(ByVal doc As Document) As Long
    Dim removed As Long
    Dim i As Long, rng As Range
    Application.ScreenUpdating = False
    For i = doc.Hyperlinks.Count To 1 Step -1
        If Left$(doc.Hyperlinks(i).ScreenTip, Len(SCREENTIP_PREFIX)) = SCREENTIP_PREFIX Then
            Set rng = doc.Hyperlinks(i).Range
            doc.Hyperlinks(i).Delete
            ResetLinkFormatting rng
            removed = removed + 1
        End If
    Next i
    Application.ScreenUpdating = True
    RemoveCitationLinks_Quiet = removed
End Function


Private Function AddLink(ByVal rng As Range, ByVal url As String, ByVal typ As String) As Boolean
    On Error GoTo Fail

    Dim h As Hyperlink
    Set h = ActiveDocument.Hyperlinks.Add(Anchor:=rng, Address:=url, _
        ScreenTip:=Left$(SCREENTIP_PREFIX & typ & " | " & url, 255))

    ' Word's Hyperlink style drops the case-name italic. Rather than try to
    ' preserve the prior formatting through the field boundary (fragile --
    ' anything applied to the first display character gets absorbed), re-derive
    ' the italic from citation structure: in a case cite the case name is
    ' everything to the left of the "(year)" date, or of ", supra".
    ItalicizeCaseName h.Range

    AddLink = True
    Exit Function
Fail:
    AddLink = False
End Function

' Italicize the case-name portion of a linked citation's display text: the run
' from the case name's first letter up to the "(year)" date or ", supra". Works
' directly on that run (via the display Characters, whose positions are the true
' text positions) rather than italicizing the whole span and clearing the tail
' -- which mis-handled a citation wrapped in outer parentheses, e.g.
' "(Gutierrez v. Tostado (2025) 18 Cal.5th 222, 231.)".
Private Sub ItalicizeCaseName(ByVal disp As Range)
    On Error Resume Next
    Dim s As String
    s = disp.text
    If Len(s) = 0 Then Exit Sub

    Dim tailStart As Long
    tailStart = CaseNameTailStart(s)   ' 1-based index where the non-italic tail begins
    If tailStart <= 1 Then
        ' No case name inside the link. It may be a supra cite whose short name
        ' sits just BEFORE the link (the linker anchors supra cites on the
        ' reporter). Italicize that preceding short name.
        ItalicizeSupraShortNameBefore disp
        Exit Sub
    End If

    Dim m As Long
    m = disp.Characters.count
    If tailStart > m + 1 Then tailStart = m + 1

    ' First letter of the case name: skip a leading outer "(", quote, or space,
    ' then any lowercase signal words ("see", "cf.", "see also"). A case short
    ' name always starts with a capital.
    Dim nameStart As Long: nameStart = 1
    Do While nameStart < tailStart
        If Mid$(s, nameStart, 1) Like "[A-Za-z]" Then Exit Do
        nameStart = nameStart + 1
    Loop
    Do While nameStart < tailStart
        If Mid$(s, nameStart, 1) Like "[a-z]" Then
            Do While nameStart < tailStart And Mid$(s, nameStart, 1) <> " ": nameStart = nameStart + 1
            Loop
            Do While nameStart < tailStart And Mid$(s, nameStart, 1) = " ": nameStart = nameStart + 1
            Loop
        Else
            Exit Do
        End If
    Loop

    ' Trim trailing spaces before the tail.
    Dim nameEnd As Long: nameEnd = tailStart - 1
    Do While nameEnd >= nameStart And Mid$(s, nameEnd, 1) = " ": nameEnd = nameEnd - 1
    Loop
    If nameEnd < nameStart Or nameStart > m Then Exit Sub
    If nameEnd > m Then nameEnd = m

    ' Italicize the case-name run as one range. Only when it starts at the very
    ' first display character do we extend the start one position back into the
    ' hidden field separator, so the field boundary doesn't absorb the italic on
    ' that first letter. (Characters(1).Start is the true text position; the
    ' Range's own .Start points into the field code and must not be used here.)
    Dim startPos As Long
    startPos = disp.Characters(nameStart).start
    If nameStart = 1 Then startPos = startPos - 1

    ActiveDocument.Range(startPos, disp.Characters(nameEnd).End).Font.Italic = True
End Sub

' Italicize the short name of a supra cite that sits just BEFORE a linked
' reporter, e.g. the document reads "Rappleyea, supra, " and then the linked
' "8 Cal.4th at p. 982". The short name is outside the hyperlink, so it is a
' plain document range (no field-boundary quirk). Only called when the in-link
' logic found nothing, so it never disturbs cites handled inside the link.
Private Sub ItalicizeSupraShortNameBefore(ByVal disp As Range)
    On Error Resume Next
    Dim linkStart As Long: linkStart = disp.start
    If linkStart < 8 Then Exit Sub

    Dim lookLen As Long: lookLen = 70
    If lookLen > linkStart Then lookLen = linkStart
    Dim base As Long: base = linkStart - lookLen
    Dim b As String: b = ActiveDocument.Range(base, linkStart).text
    If Len(b) = 0 Then Exit Sub

    ' The text right before the link must end with "..., supra" (ignoring any
    ' trailing spaces / comma the link itself doesn't include).
    Dim t As String: t = b
    Do While Len(t) > 0
        Dim last As String: last = Right$(t, 1)
        If last = " " Or last = "," Then t = Left$(t, Len(t) - 1) Else Exit Do
    Loop
    If Len(t) < 5 Then Exit Sub
    If LCase$(Right$(t, 5)) <> "supra" Then Exit Sub

    ' The comma separating the short name from ", supra".
    Dim supraPos As Long: supraPos = Len(t) - 4      ' 1-based start of "supra" in b
    Dim j As Long: j = supraPos - 1
    Do While j >= 1 And Mid$(b, j, 1) = " ": j = j - 1
    Loop
    If j >= 1 And Mid$(b, j, 1) = "," Then j = j - 1 Else Exit Sub
    Dim nameEnd As Long: nameEnd = j                 ' last char of the short name

    ' Walk back to the start of the short name: stop at "(", ";", or a sentence
    ' boundary ". ".
    Dim k As Long: k = nameEnd
    Do While k >= 1
        Dim ch As String: ch = Mid$(b, k, 1)
        If ch = "(" Or ch = ";" Then Exit Do
        If ch = " " And k >= 2 Then
            If Mid$(b, k - 1, 1) = "." Then Exit Do
        End If
        k = k - 1
    Loop
    Dim nameStart As Long: nameStart = k + 1

    ' Skip leading spaces and any lowercase signal words ("see", "cf.", etc.);
    ' a case short name always begins with a capital.
    Do
        Do While nameStart <= nameEnd And Mid$(b, nameStart, 1) = " ": nameStart = nameStart + 1
        Loop
        If nameStart > nameEnd Then Exit Sub
        Dim fc As String: fc = Mid$(b, nameStart, 1)
        If fc >= "a" And fc <= "z" Then
            Do While nameStart <= nameEnd And Mid$(b, nameStart, 1) <> " ": nameStart = nameStart + 1
            Loop
        Else
            Exit Do
        End If
    Loop
    If nameStart > nameEnd Then Exit Sub

    Dim absS As Long: absS = base + nameStart - 1
    Dim absE As Long: absE = base + nameEnd
    If absE > absS Then ActiveDocument.Range(absS, absE).Font.Italic = True
End Sub

' Return the 1-based character index where the non-italic citation tail begins:
' the comma of ", supra", else the "(" of the first four-digit "(year)". Returns
' 0 when neither is present (nothing to italicize).
Private Function CaseNameTailStart(ByVal s As String) As Long
    Dim p As Long
    p = InStr(1, s, ", supra", vbTextCompare)
    If p > 0 Then
        CaseNameTailStart = p
        Exit Function
    End If

    p = FindYearParen(s)
    If p > 0 Then
        CaseNameTailStart = p
        Exit Function
    End If

    ' No year and no supra. If this is still a case citation (has a "... v. ..."
    ' party separator), italicize the case name anyway: it runs from the start
    ' up to the court/docket parenthetical -- the first "(" -- e.g. "Pate v. BMW
    ' of North America, LLC (C.D.Cal., No. 2:21-cv-04915-KS)". With no such
    ' paren, italicize the whole span.
    If InStr(1, s, " v. ", vbTextCompare) > 0 Then
        p = InStr(1, s, "(")
        If p > 1 Then
            CaseNameTailStart = p
        Else
            CaseNameTailStart = Len(s) + 1
        End If
        Exit Function
    End If

    CaseNameTailStart = 0
End Function

' Index of the "(" that opens the date parenthetical -- the first parenthetical
' containing a 4-digit year (19xx/20xx). Handles "(1992)" (California) as well
' as "(C.D. Cal. 2021)" / "(9th Cir. 2019)" (federal: court + year). Returns 0
' when no parenthesized year is present.
Private Function FindYearParen(ByVal s As String) As Long
    Dim yearPos As Long
    yearPos = FindYearPos(s)
    If yearPos = 0 Then
        FindYearParen = 0
        Exit Function
    End If

    ' Walk left from the year to the "(" that opens its parenthetical. Stop if a
    ' ")" is reached first (the year is not inside parentheses).
    Dim i As Long
    For i = yearPos - 1 To 1 Step -1
        Dim c As String: c = Mid$(s, i, 1)
        If c = "(" Then
            FindYearParen = i
            Exit Function
        ElseIf c = ")" Then
            Exit For
        End If
    Next i
    FindYearParen = 0
End Function

' Position of the first standalone 4-digit year (19xx/20xx) in s, or 0.
Private Function FindYearPos(ByVal s As String) As Long
    Dim i As Long
    For i = 1 To Len(s) - 3
        Dim d1 As String, d2 As String, d3 As String, d4 As String
        d1 = Mid$(s, i, 1): d2 = Mid$(s, i + 1, 1)
        d3 = Mid$(s, i + 2, 1): d4 = Mid$(s, i + 3, 1)
        If d1 Like "#" And d2 Like "#" And d3 Like "#" And d4 Like "#" Then
            If (d1 = "1" And d2 = "9") Or (d1 = "2" And d2 = "0") Then
                Dim okBefore As Boolean, okAfter As Boolean
                okBefore = (i = 1)
                If Not okBefore Then okBefore = Not (Mid$(s, i - 1, 1) Like "#")
                okAfter = (i + 4 > Len(s))
                If Not okAfter Then okAfter = Not (Mid$(s, i + 4, 1) Like "#")
                If okBefore And okAfter Then
                    FindYearPos = i
                    Exit Function
                End If
            End If
        End If
    Next i
    FindYearPos = 0
End Function


Private Function FindAndLink(ByVal scope As Range, ByVal needle As String, _
                             ByVal url As String, ByVal typ As String) As Boolean
    On Error GoTo Fail
    If Len(needle) = 0 Or Len(needle) > 250 Then Exit Function
    Dim fr As Range
    Set fr = scope.Duplicate
    With fr.Find
        .ClearFormatting
        .Text = needle
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .MatchCase = True
        .Execute
    End With
    If fr.Find.Found Then
        FindAndLink = AddLink(fr, url, typ)
    End If
    Exit Function
Fail:
    FindAndLink = False
End Function


Private Sub ResetLinkFormatting(ByVal rng As Range)
    On Error Resume Next
    rng.Font.Underline = wdUnderlineNone
    rng.Font.ColorIndex = wdAuto
End Sub


' Paragraph text without the trailing paragraph mark, used for BOTH the HTML
' we send and the offset map we build, so the two stay consistent.
Private Function ParagraphRawText(ByVal r As Range) As String
    Dim t As String
    t = r.Text
    If Len(t) > 0 Then
        If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)
    End If
    ParagraphRawText = t
End Function


' Replicates citation_extractor's _normalize_ws(_strip_tags(...)) for tag-free,
' already-unescaped text: collapse whitespace runs to one space, trim ends,
' and record where each surviving character came from.
Private Function NormalizeAndMap(ByVal raw As String) As NormResult
    Dim res As NormResult
    Dim L As Long
    L = Len(raw)
    ReDim res.map(0 To L + 1)

    Dim sb As String
    Dim j As Long
    Dim inWs As Boolean, pendingStart As Long
    Dim i As Long, code As Long
    j = 0
    inWs = False
    pendingStart = 0

    For i = 1 To L
        code = AscW(Mid$(raw, i, 1))
        If IsWhitespaceCode(code) Then
            If Not inWs Then
                inWs = True
                pendingStart = i
            End If
        Else
            If inWs Then
                If j > 0 Then
                    sb = sb & " "
                    res.map(j) = pendingStart - 1   ' 0-based raw index of the run
                    j = j + 1
                End If
                inWs = False
            End If
            sb = sb & Mid$(raw, i, 1)
            res.map(j) = i - 1
            j = j + 1
        End If
    Next i

    res.norm = sb
    res.n = j
    NormalizeAndMap = res
End Function


Private Function IsWhitespaceCode(ByVal c As Long) As Boolean
    Select Case c
        Case 9, 10, 11, 12, 13, 32, 160
            IsWhitespaceCode = True
        Case 8192 To 8202, 8232, 8233, 8239, 8287, 12288
            IsWhitespaceCode = True
        Case Else
            IsWhitespaceCode = False
    End Select
End Function


Private Function EscapeHtml(ByVal s As String) As String
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    EscapeHtml = s
End Function


Private Function Q(ByVal s As String) As String
    Q = """" & s & """"
End Function


Private Function RunAndWait(ByVal cmd As String) As Long
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    RunAndWait = sh.Run(cmd, 0, True)   ' 0 = hidden window, True = wait
End Function


Private Sub WriteUtf8File(ByVal path As String, ByVal content As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile path, 2       ' adSaveCreateOverWrite
    st.Close
End Sub


Private Function ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText
    st.Close
End Function


'==============================================================================
' ROW ORDERING
'==============================================================================

' Insertion sort by (blk, s). Citation counts are small, so this is fine.
Private Sub SortRows(ByRef a() As CiteRow)
    Dim i As Long, jj As Long
    Dim key As CiteRow
    For i = LBound(a) + 1 To UBound(a)
        key = a(i)
        jj = i - 1
        Do While jj >= LBound(a)
            If (a(jj).blk > key.blk) Or _
               (a(jj).blk = key.blk And a(jj).s > key.s) Then
                a(jj + 1) = a(jj)
                jj = jj - 1
            Else
                Exit Do
            End If
        Loop
        a(jj + 1) = key
    Next i
End Sub


' Greedy filter: within a paragraph, drop any span that starts before the
' previous kept span ended. Word cannot nest a hyperlink inside another.
Private Function FilterOverlaps(ByRef a() As CiteRow) As CiteRow()
    Dim out() As CiteRow
    ReDim out(LBound(a) To UBound(a))
    Dim cnt As Long
    Dim curBlk As Long, lastEnd As Long
    cnt = 0
    curBlk = -1
    lastEnd = -1
    Dim i As Long
    For i = LBound(a) To UBound(a)
        If a(i).blk <> curBlk Then
            curBlk = a(i).blk
            lastEnd = -1
        End If
        If a(i).s >= lastEnd Then
            out(cnt) = a(i)
            cnt = cnt + 1
            lastEnd = a(i).e
        End If
    Next i
    If cnt = 0 Then
        ReDim out(0 To 0)
    Else
        ReDim Preserve out(0 To cnt - 1)
    End If
    FilterOverlaps = out
End Function
