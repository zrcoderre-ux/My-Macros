Attribute VB_Name = "HeadingFormat"
'==============================================================================
' HeadingFormat.bas
'------------------------------------------------------------------------------
' Two things a tentative ruling's headings always want, applied in one pass over
' the body:
'   1. KEEP WITH NEXT on every section heading, so a heading is never left
'      stranded at the foot of a page with its section starting overleaf.
'   2. The UNDERLINE on roman-numeral analysis headings: the title is underlined,
'      the numeral is not. "I. Reply Evidence" underlines "Reply Evidence" and
'      leaves "I. " alone.
' Lettered and numbered subheadings (A., B. / 1., 2.) and the all-caps labels
' take keep-with-next but no underline, which is the house rule.
'
' MACRO YOU RUN:
'   FormatHeadings - do both, everywhere in the body, and report what changed.
'
' WHAT COUNTS AS A HEADING. One per paragraph, and it must NOT end in closing
' punctuation -- a heading names a subject, it doesn't close a sentence, so a
' final ".", ",", ":", ";", ")" or quote mark is the tell that a paragraph is
' prose. Past that, a heading is one of:
'   - a ROMAN numeral label:   "I.", "II.", "IV." ...   (underlined title)
'   - a CAPITAL LETTER label:  "A.", "B.", "C." ...
'   - an ARABIC numeral label: "1.", "2.", "3." ...
'   - ALL CAPS with no label:  BACKGROUND, LEGAL STANDARD, ANALYSIS, CONCLUSION,
'                              REQUEST FOR JUDICIAL NOTICE, EVIDENTIARY
'                              OBJECTIONS
' A label has to be followed by a space (or be the whole line), so a sentence
' opening "Plaintiff v. Superior Court held ..." is not read as a label.
'
' NOTES:
'   - "I.", "V.", "X.", "C.", "D.", "M." and "L." are both single capital letters
'     and roman numerals. See LabelKind: one of those reads as a LETTER only when
'     it continues a lettered series (the previous lettered heading was the letter
'     before it), and as a roman numeral otherwise -- because roman is the top
'     level and the one a section opens with.
'   - Body only. The caption lives in the page header, which this never touches,
'     and paragraphs inside tables are skipped (an objections table's numbered
'     rows are not headings).
'   - A paragraph longer than MAX_HEADING_LEN is not a heading however it is
'     punctuated. That is the guard against a body paragraph that simply lost its
'     final period being underlined as though it were a heading.
'   - NATURE OF PROCEEDINGS carries its "Hearing on ..." text on the same line and
'     ends in a period, so it reads as prose and is left alone. It sits on page
'     one where nothing can orphan it.
'   - Underline is only ever applied to a roman heading's title and removed from
'     its numeral. No other underline in the document is touched -- including a
'     subheading the user underlined deliberately.
'   - Everything is one undo record: Ctrl+Z reverses the whole run.
'==============================================================================
Option Explicit

' What a heading is labelled with. The label decides the underline: a roman
' numeral heading carries an underlined title, every other kind carries none.
Private Const HEAD_NONE   As Long = 0
Private Const HEAD_CAPS   As Long = 1
Private Const HEAD_ROMAN  As Long = 2
Private Const HEAD_LETTER As Long = 3
Private Const HEAD_ARABIC As Long = 4

' A heading names its subject in a few words. The longest standing label,
' "REQUEST FOR JUDICIAL NOTICE", is 27 characters, so this is far above anything
' real: it is here so that a body paragraph which lost its final period is still
' too long to be mistaken for a heading.
Private Const MAX_HEADING_LEN As Long = 100

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub FormatHeadings()
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    If oDoc Is Nothing Then Exit Sub

    ' One named record so the whole pass reverses on a single Ctrl+Z. Closed in
    ' CleanUp whatever happens -- an undo record left open crashes the next run.
    Dim oUndo As UndoRecord
    Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "Format Headings"
    On Error GoTo CleanUp

    Application.ScreenUpdating = False

    Dim nHeads As Long, nKept As Long, nLined As Long
    Dim prevLetter As String
    prevLetter = ""

    Dim p As Paragraph
    For Each p In oDoc.content.Paragraphs
        If Not InTable(p) Then
            Dim raw As String
            raw = ParaText(p)

            Dim titleStart As Long
            Dim kind As Long
            kind = HeadingKind(raw, titleStart, prevLetter)

            If kind <> HEAD_NONE Then
                nHeads = nHeads + 1

                If p.KeepWithNext <> True Then
                    p.KeepWithNext = True
                    nKept = nKept + 1
                End If

                If kind = HEAD_ROMAN Then
                    If UnderlineTitle(oDoc, p, raw, titleStart) Then _
                        nLined = nLined + 1
                End If
            End If
        End If
    Next p

CleanUp:
    Dim eN As Long: eN = Err.Number
    Dim eD As String: eD = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    oUndo.EndCustomRecord
    On Error GoTo 0

    If eN <> 0 Then
        MsgBox "Format Headings stopped after an error:" & vbCrLf & vbCrLf & _
               "Error " & eN & ": " & eD & vbCrLf & vbCrLf & _
               "Whatever it had already changed is in one undo record, so " & _
               "Ctrl+Z puts the document back.", _
               vbExclamation, "Format Headings"
        Exit Sub
    End If

    If nHeads = 0 Then
        MsgBox "No headings found." & vbCrLf & vbCrLf & _
               "A heading is a line that does NOT end in closing punctuation " & _
               "and is either ALL CAPS or labelled ""I."" / ""A."" / ""1."". " & _
               "A heading that ends in a period reads as prose.", _
               vbInformation, "Format Headings"
        Exit Sub
    End If

    MsgBox "Found " & nHeads & " heading(s)." & vbCrLf & vbCrLf & _
           "Applied ""keep with next"" to " & nKept & " that did not have it." & _
           vbCrLf & _
           "Underlined the title of " & nLined & " roman-numeral heading(s), " & _
           "leaving the numeral itself alone." & vbCrLf & vbCrLf & _
           "Ctrl+Z reverses the whole run.", _
           vbInformation, "Format Headings"
End Sub

'==============================================================================
' DETECTION
'==============================================================================
' Which kind of heading this paragraph is, or HEAD_NONE. titleStart comes back as
' the 1-based index IN raw of the first character after the label and its
' following space -- where a roman heading's underline begins.
'
' prevLetter carries the last lettered label seen, so a series can be recognized
' as it goes; see LabelKind. It is passed ByRef and updated here.
Private Function HeadingKind(ByVal raw As String, ByRef titleStart As Long, _
                             ByRef prevLetter As String) As Long
    titleStart = 0

    ' Find the visible text without copying it out of position: the underline
    ' lands on real character offsets, so every index stays an index into raw.
    Dim s As Long: s = 1
    Do While s <= Len(raw)
        If Not IsWsChar(Mid$(raw, s, 1)) Then Exit Do
        s = s + 1
    Loop
    Dim e As Long: e = Len(raw)
    Do While e >= s
        If Not IsTrailingSkip(Mid$(raw, e, 1)) Then Exit Do
        e = e - 1
    Loop
    If e < s Then Exit Function                      ' empty paragraph

    Dim body As String: body = Mid$(raw, s, e - s + 1)
    If Len(body) > MAX_HEADING_LEN Then Exit Function

    ' A heading does not close. Every sentence-ending and clause-closing mark
    ' says the line is prose -- which is what keeps ordinary paragraphs, block
    ' quotes, and citation sentences out of this.
    If IsClosingPunct(Right$(body, 1)) Then Exit Function

    Dim titleAt As Long
    Dim lbl As String: lbl = LabelToken(body, titleAt)
    If Len(lbl) > 0 Then
        If titleAt > 0 Then titleStart = s + titleAt - 1
        HeadingKind = LabelKind(lbl, prevLetter)
        Exit Function
    End If

    ' No label: an all-caps line is one of the standing section labels, and a new
    ' section ends whatever lettered series was running.
    If IsAllCapsText(body) Then
        prevLetter = ""
        HeadingKind = HEAD_CAPS
    End If
End Function

' The "I." / "A." / "3." a heading opens with, without its period, or "" when the
' line opens with no label. titleAt comes back as the 1-based index in s of the
' first character after the label and the space following it -- 0 when the label
' is the whole line and there is no title to underline.
Private Function LabelToken(ByVal s As String, ByRef titleAt As Long) As String
    titleAt = 0

    Dim dot As Long: dot = InStr(1, s, ".")
    If dot < 2 Then Exit Function                    ' no period, or ".x"

    Dim tok As String: tok = Left$(s, dot - 1)
    If Len(tok) > 8 Then Exit Function               ' no real label is this long

    ' The label has to be followed by a space, or be the whole line. Without this
    ' an opening "Plaintiff v. Superior Court ..." reads as a label, and so does
    ' every abbreviation that starts a sentence.
    If dot < Len(s) Then
        If Not IsWsChar(Mid$(s, dot + 1, 1)) Then Exit Function
    End If

    If Not (IsRomanToken(tok) Or IsLetterToken(tok) Or IsDigits(tok)) Then Exit Function

    Dim k As Long: k = dot + 1
    Do While k <= Len(s)
        If Not IsWsChar(Mid$(s, k, 1)) Then Exit Do
        k = k + 1
    Loop
    If k <= Len(s) Then titleAt = k

    LabelToken = tok
End Function

' Which series a label belongs to, resolving the letters that are also numerals.
'
' "I.", "V.", "X.", "L.", "C.", "D." and "M." are single capital letters AND roman
' numerals, so "C." could be the third subheading of a lettered series or the
' hundredth section of a document nobody has ever written. It reads as a LETTER
' only when it continues a series -- the last lettered heading was the letter
' before it -- and as a ROMAN numeral otherwise, because roman is the top level
' and the one a section opens with. That resolves A., B., C. correctly (B. is
' unambiguous, so C. is seen to continue it) and leaves a bare "I." heading roman.
Private Function LabelKind(ByVal tok As String, ByRef prevLetter As String) As Long
    If IsDigits(tok) Then
        LabelKind = HEAD_ARABIC
        Exit Function
    End If

    Dim roman As Boolean: roman = IsRomanToken(tok)
    Dim letter As Boolean: letter = IsLetterToken(tok)

    If roman And Not letter Then                     ' "II.", "IV.", "XI."
        prevLetter = ""
        LabelKind = HEAD_ROMAN
        Exit Function
    End If

    If letter And Not roman Then                     ' "A.", "B.", "K."
        prevLetter = tok
        LabelKind = HEAD_LETTER
        Exit Function
    End If

    If Len(prevLetter) = 1 Then
        If Asc(tok) = Asc(prevLetter) + 1 Then
            prevLetter = tok
            LabelKind = HEAD_LETTER
            Exit Function
        End If
    End If

    prevLetter = ""
    LabelKind = HEAD_ROMAN
End Function

'==============================================================================
' FORMATTING
'==============================================================================
' Underline a roman heading's title and clear the underline from its numeral.
' Returns True only when something actually changed, so a second run reports
' nothing rather than claiming the same work twice.
Private Function UnderlineTitle(ByVal oDoc As Document, ByVal p As Paragraph, _
                                ByVal raw As String, ByVal titleStart As Long) As Boolean
    If titleStart < 2 Then Exit Function             ' no numeral ahead of it

    ' Stop the underline at the last visible character: an underline running out
    ' past the final word into trailing spaces is the sort of thing that shows up
    ' only once it is printed.
    Dim titleEnd As Long: titleEnd = Len(raw)
    Do While titleEnd >= titleStart
        If Not IsTrailingSkip(Mid$(raw, titleEnd, 1)) Then Exit Do
        titleEnd = titleEnd - 1
    Loop
    If titleEnd < titleStart Then Exit Function      ' a numeral with no title

    Dim body As Range
    Set body = p.Range.Duplicate
    If body.Characters.count < 2 Then Exit Function
    body.MoveEnd wdCharacter, -1                     ' drop the paragraph mark
    If body.Characters.count < titleEnd Then Exit Function

    Dim numeral As Range, title As Range
    Set numeral = oDoc.Range(body.Characters(1).start, body.Characters(titleStart).start)
    Set title = oDoc.Range(body.Characters(titleStart).start, body.Characters(titleEnd).End)

    Dim changed As Boolean: changed = False
    ' Font.Underline answers wdUndefined for a mixed range, which compares equal
    ' to neither -- so a half-underlined heading is corrected rather than skipped.
    If numeral.Font.Underline <> wdUnderlineNone Then
        numeral.Font.Underline = wdUnderlineNone
        changed = True
    End If
    If title.Font.Underline <> wdUnderlineSingle Then
        title.Font.Underline = wdUnderlineSingle
        changed = True
    End If

    UnderlineTitle = changed
End Function

'==============================================================================
' SMALL HELPERS
'==============================================================================
' The paragraph's text without its paragraph mark.
Private Function ParaText(ByVal p As Paragraph) As String
    Dim t As String
    On Error Resume Next
    t = p.Range.text
    On Error GoTo 0

    If Len(t) > 0 Then
        If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)
    End If
    ParaText = t
End Function

Private Function InTable(ByVal p As Paragraph) As Boolean
    On Error Resume Next
    InTable = p.Range.Information(wdWithInTable)
End Function

' Roman numeral characters only. Well-formedness isn't tested: "IIII." is not a
' numeral anyone writes, and reading it as one costs nothing.
Private Function IsRomanToken(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If InStr(1, "IVXLCDM", Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsRomanToken = True
End Function

Private Function IsLetterToken(ByVal s As String) As Boolean
    If Len(s) <> 1 Then Exit Function
    IsLetterToken = (s >= "A" And s <= "Z")
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) >= "0" And Mid$(s, i, 1) <= "9") Then Exit Function
    Next i
    IsDigits = True
End Function

' All caps: at least one letter, and no lowercase one. Digits, punctuation and
' spaces are neutral, so "PART II" and "REQUEST FOR JUDICIAL NOTICE" both pass.
Private Function IsAllCapsText(ByVal s As String) As Boolean
    Dim i As Long, c As String, sawUpper As Boolean
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "a" And c <= "z" Then Exit Function
        If c >= "A" And c <= "Z" Then sawUpper = True
    Next i
    IsAllCapsText = sawUpper
End Function

' The marks that end prose. A line closing with any of them is a sentence, not a
' heading -- including the closing bracket or quote that ends a citation.
Private Function IsClosingPunct(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    Select Case c
        Case ".", ",", ";", ":", "!", "?", ")", "]", "}", """", "'"
            IsClosingPunct = True
        Case ChrW$(8221), ChrW$(8217), ChrW$(8230), ChrW$(187)
            IsClosingPunct = True                    ' " ' ... >>
    End Select
End Function

Private Function IsWsChar(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    Select Case AscW(c)
        Case 9, 10, 11, 12, 13, 32, 160: IsWsChar = True
    End Select
End Function

' What may sit at the END of a line without being its last real character:
' whitespace, and the reference marks Word keeps in the text for a footnote or a
' comment. Without the marks, a sentence ending "... granted.[note]" ends in a
' mark rather than in the period, and reads as a heading. Kept separate from
' IsWsChar because a mark is NOT the space a label needs after its period.
Private Function IsTrailingSkip(ByVal c As String) As Boolean
    If IsWsChar(c) Then
        IsTrailingSkip = True
    ElseIf Len(c) > 0 Then
        Select Case AscW(c)
            Case 2, 5: IsTrailingSkip = True         ' note ref, comment ref
        End Select
    End If
End Function
