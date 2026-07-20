Attribute VB_Name = "modMain"
Option Explicit
' Internal plumbing (InitializeAppEvents, shared helpers). Keep it out of the
' Alt+F8 list; everything stays callable within the project -- including from the
' clsAppEvents / ThisDocument class modules.
Option Private Module

Public gAppEvents       As clsAppEvents
Public gSkipCloseChecks As Boolean          ' Set True by mail merge to suppress checks

Private Const HL_GREEN As Long = 1          ' maps to wdBrightGreen
Private Const HL_CYAN  As Long = 2          ' maps to wdTurquoise

' ============================================================
' INITIALIZE GLOBAL EVENT HANDLER
' ============================================================
Sub InitializeAppEvents()
    Set gAppEvents = New clsAppEvents
    Set gAppEvents.App = word.Application
End Sub
' ============================================================
' ONEDRIVE LOCATION GUARD
' Returns True if the document is saved inside the user's
' OneDrive folder. Returns False for unsaved documents or
' documents saved outside OneDrive (e.g., Downloads, Desktop).
' ============================================================
Public Function IsInOneDrive(Doc As Document) As Boolean
    Dim fullPath As String
    fullPath = ""
    On Error Resume Next
    fullPath = Doc.FullName
    On Error GoTo 0

    ' Unsaved documents have no real path � skip them
    If fullPath = "" Or Doc.Path = "" Then
        IsInOneDrive = False
        Exit Function
    End If

    ' Check for local OneDrive path (personal) or SharePoint/OneDrive for Business URL
    If InStr(1, fullPath, "\OneDrive", vbTextCompare) > 0 Then
        IsInOneDrive = True
    ElseIf InStr(1, fullPath, "sharepoint.com", vbTextCompare) > 0 Then
        IsInOneDrive = True
    ElseIf InStr(1, fullPath, "lacourts-my", vbTextCompare) > 0 Then
        IsInOneDrive = True
    End If
End Function
' ============================================================
' TITLE DATE GUARD
' Returns True only if the document's file name ends with a
' date in M.D.YYYY format (1-2 digit month and day, 4-digit
' year), optionally followed by a file extension -- e.g.
' "Ruling 6.25.2026.docx". Used to limit the close-review to
' dated work documents. Returns False for unsaved/untitled docs.
' ============================================================
Public Function TitleEndsWithDate(Doc As Document) As Boolean
    Dim nm As String
    nm = ""
    On Error Resume Next
    nm = Doc.Name
    On Error GoTo 0
    If nm = "" Then Exit Function

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "\d{1,2}\.\d{1,2}\.\d{4}(\.[A-Za-z]{2,5})?$"
    TitleEndsWithDate = re.Test(nm)
End Function
' ============================================================
' FAST-EXIT CHECK
' Runs checks in order, stops at the very first issue found,
' highlights that single character/word, and returns its Range
' plus a plain-English label. Returns Nothing if the document
' is clean. Apostrophe conversion always runs regardless.
' ============================================================
Public Function FindFirstIssue(ByVal Doc As Document, _
                                ByRef sLabel As String) As Range
    Dim r As Range

    ClearCheckHighlights Doc

    ' 1. Smart double quotes
    Set r = FindFirstUnmatchedPair(Doc, ChrW(8220), ChrW(8221), HL_GREEN)
    If Not r Is Nothing Then
        sLabel = "Unmatched smart double quote"
        Set FindFirstIssue = r
        GoTo RunApostrophes
    End If

    ' 2. Straight double quotes (odd count � highlight the last one)
    If (CountChar(Doc, Chr(34)) Mod 2) <> 0 Then
        Dim rng  As Range
        Dim last As Range
        Set rng = Doc.content
        With rng.Find
            .ClearFormatting
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .text = Chr(34)
            Do While .Execute
                Set last = rng.Duplicate
            Loop
        End With
        If Not last Is Nothing Then
            last.HighlightColorIndex = wdBrightGreen
            sLabel = "Unmatched straight double quote"
            Set FindFirstIssue = last
            GoTo RunApostrophes
        End If
    End If

    ' 3. Square brackets
    Set r = FindFirstUnmatchedPair(Doc, "[", "]", HL_GREEN)
    If Not r Is Nothing Then
        sLabel = "Unmatched square bracket"
        Set FindFirstIssue = r
        GoTo RunApostrophes
    End If

    ' 4. Curly braces
    Set r = FindFirstUnmatchedPair(Doc, "{", "}", HL_GREEN)
    If Not r Is Nothing Then
        sLabel = "Unmatched curly brace"
        Set FindFirstIssue = r
        GoTo RunApostrophes
    End If

    ' 5. Parentheses
    Set r = FindFirstUnmatchedPair(Doc, "(", ")", HL_GREEN)
    If Not r Is Nothing Then
        sLabel = "Unmatched parenthesis"
        Set FindFirstIssue = r
        GoTo RunApostrophes
    End If

    ' 6. Placeholder word "blank"
    Set r = FindFirstBlank(Doc)
    If Not r Is Nothing Then
        sLabel = "Placeholder word ""blank"""
        Set FindFirstIssue = r
        GoTo RunApostrophes
    End If

    ' 7. Double spaces
    Set r = FindFirstDoubleSpace(Doc)
    If Not r Is Nothing Then
        sLabel = "Double space"
        Set FindFirstIssue = r
        GoTo RunApostrophes
    End If

    ' No issues found
    Set FindFirstIssue = Nothing

RunApostrophes:
    ConvertStraightApostrophes Doc
End Function

' ============================================================
' FULL-RUN CHECK
' Used after the user chooses Yes (stay open) on a prior
' prompt. Runs every check and returns aggregate flags so
' clsAppEvents can show the full summary prompt.
' ============================================================
Public Sub RunAllDocumentChecks(ByVal Doc As Document, _
                                ByRef issues As Boolean, _
                                ByRef userHighlights As Boolean)
    issues = False
    userHighlights = False

    ClearCheckHighlights Doc

    ' Smart double quotes
    If CheckUnmatchedPairs(Doc, ChrW(8220), ChrW(8221), HL_GREEN) Then issues = True

    ' Straight double quotes (odd count � highlight the last one)
    If (CountChar(Doc, Chr(34)) Mod 2) <> 0 Then
        Dim rng  As Range
        Dim last As Range
        Set rng = Doc.content
        With rng.Find
            .ClearFormatting
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .text = Chr(34)
            Do While .Execute
                Set last = rng.Duplicate
            Loop
        End With
        If Not last Is Nothing Then
            last.HighlightColorIndex = wdBrightGreen
        End If
        issues = True
    End If

    ' Square brackets, curly braces, parentheses
    If CheckUnmatchedPairs(Doc, "[", "]", HL_GREEN) Then issues = True
    If CheckUnmatchedPairs(Doc, "{", "}", HL_GREEN) Then issues = True
    If CheckUnmatchedPairs(Doc, "(", ")", HL_GREEN) Then issues = True

    ' Double spaces
    If CheckDoubleSpaces(Doc) Then issues = True

    ' Apostrophe conversion (always runs, no prompt)
    ConvertStraightApostrophes Doc

    ' User highlight check (runs after macro colors are in place)
    If DocumentHasUserHighlights(Doc) Then userHighlights = True

End Sub

' ============================================================
' PAIR CHECKING � PARAGRAPH BY PARAGRAPH
' Both functions use the same collect ? sort ? stack algorithm
' scoped to one paragraph at a time. A ( in one paragraph and
' a ) in a different paragraph are each flagged as unmatched.
' ============================================================

' Full variant: highlights ALL unmatched characters in the document.
' Returns True if any are found. Used by RunAllDocumentChecks.
Private Function CheckUnmatchedPairs(Doc As Document, opener As String, _
                                     closer As String, color As Long) As Boolean
    Dim PARA        As Paragraph
    Dim paraRng     As Range
    Dim oRng        As Range
    Dim cRng        As Range
    Dim positions() As Long
    Dim types()     As Boolean      ' True = opener, False = closer
    Dim count       As Long
    Dim stack()     As Long
    Dim stackTop    As Long
    Dim i           As Long

    For Each PARA In Doc.Paragraphs
        Set paraRng = PARA.Range
        count = 0
        ReDim positions(0)
        ReDim types(0)

        ' Collect openers in this paragraph
        Set oRng = paraRng.Duplicate
        With oRng.Find
            .ClearFormatting
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .text = opener
            Do While .Execute
                ' Clamp to this paragraph: after a hit, Range.Find keeps
                ' going to the END OF THE STORY, so without this an opener
                ' here paired with a closer 40 paragraphs later and the
                ' cross-paragraph strays this checker exists to catch were
                ' never flagged (and every paragraph rescanned the rest of
                ' the document).
                If oRng.start >= paraRng.End Then Exit Do
                count = count + 1
                ReDim Preserve positions(count)
                ReDim Preserve types(count)
                positions(count) = oRng.start
                types(count) = True
            Loop
        End With

        ' Collect closers in this paragraph
        Set cRng = paraRng.Duplicate
        With cRng.Find
            .ClearFormatting
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .text = closer
            Do While .Execute
                If cRng.start >= paraRng.End Then Exit Do   ' clamp (see opener loop)
                count = count + 1
                ReDim Preserve positions(count)
                ReDim Preserve types(count)
                positions(count) = cRng.start
                types(count) = False
            Loop
        End With

        If count = 0 Then GoTo NextParaFull
        SortByPosition positions, types, count

        ' Stack-based match within this paragraph
        stackTop = 0
        ReDim stack(0)

        For i = 1 To count
            If types(i) Then
                ' Opener � push position
                stackTop = stackTop + 1
                ReDim Preserve stack(stackTop)
                stack(stackTop) = positions(i)
            Else
                ' Closer
                If stackTop > 0 Then
                    stackTop = stackTop - 1         ' Matched � pop
                Else
                    ' Unmatched closer � highlight it
                    Dim closeRng As Range
                    Set closeRng = Doc.Range(positions(i), positions(i) + 1)
                    closeRng.HighlightColorIndex = RGBToHighlightIndex(color)
                    CheckUnmatchedPairs = True
                End If
            End If
        Next i

        ' Anything left on the stack is an unmatched opener
        For i = 1 To stackTop
            Dim openRng As Range
            Set openRng = Doc.Range(stack(i), stack(i) + 1)
            openRng.HighlightColorIndex = RGBToHighlightIndex(color)
            CheckUnmatchedPairs = True
        Next i

NextParaFull:
    Next PARA
End Function

' Single-hit variant: highlights only the FIRST unmatched character found
' and returns its Range. Returns Nothing if the document is clean.
' Used by FindFirstIssue.
Private Function FindFirstUnmatchedPair(Doc As Document, opener As String, _
                                         closer As String, color As Long) As Range
    Dim PARA        As Paragraph
    Dim paraRng     As Range
    Dim oRng        As Range
    Dim cRng        As Range
    Dim positions() As Long
    Dim types()     As Boolean
    Dim count       As Long
    Dim stack()     As Long
    Dim stackTop    As Long
    Dim i           As Long

    For Each PARA In Doc.Paragraphs
        Set paraRng = PARA.Range
        count = 0
        ReDim positions(0)
        ReDim types(0)

        ' Collect openers in this paragraph
        Set oRng = paraRng.Duplicate
        With oRng.Find
            .ClearFormatting
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .text = opener
            Do While .Execute
                ' Clamp to this paragraph: after a hit, Range.Find keeps
                ' going to the END OF THE STORY, so without this an opener
                ' here paired with a closer 40 paragraphs later and the
                ' cross-paragraph strays this checker exists to catch were
                ' never flagged (and every paragraph rescanned the rest of
                ' the document).
                If oRng.start >= paraRng.End Then Exit Do
                count = count + 1
                ReDim Preserve positions(count)
                ReDim Preserve types(count)
                positions(count) = oRng.start
                types(count) = True
            Loop
        End With

        ' Collect closers in this paragraph
        Set cRng = paraRng.Duplicate
        With cRng.Find
            .ClearFormatting
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .text = closer
            Do While .Execute
                If cRng.start >= paraRng.End Then Exit Do   ' clamp (see opener loop)
                count = count + 1
                ReDim Preserve positions(count)
                ReDim Preserve types(count)
                positions(count) = cRng.start
                types(count) = False
            Loop
        End With

        If count = 0 Then GoTo NextParaFirst
        SortByPosition positions, types, count

        ' Stack-based match within this paragraph
        stackTop = 0
        ReDim stack(0)

        For i = 1 To count
            If types(i) Then
                ' Opener � push position
                stackTop = stackTop + 1
                ReDim Preserve stack(stackTop)
                stack(stackTop) = positions(i)
            Else
                ' Closer
                If stackTop > 0 Then
                    stackTop = stackTop - 1         ' Matched � pop
                Else
                    ' First unmatched closer � highlight and return immediately
                    Dim closeRng As Range
                    Set closeRng = Doc.Range(positions(i), positions(i) + 1)
                    closeRng.HighlightColorIndex = RGBToHighlightIndex(color)
                    Set FindFirstUnmatchedPair = closeRng
                    Exit Function
                End If
            End If
        Next i

        ' First unmatched opener is at the bottom of the stack (earliest in paragraph)
        If stackTop > 0 Then
            Dim openRng As Range
            Set openRng = Doc.Range(stack(1), stack(1) + 1)
            openRng.HighlightColorIndex = RGBToHighlightIndex(color)
            Set FindFirstUnmatchedPair = openRng
            Exit Function
        End If

NextParaFirst:
    Next PARA
End Function

' ============================================================
' SHARED SORT HELPER
' Bubble sort: orders positions array ascending so the stack
' algorithm always processes characters in document order.
' ============================================================
Private Sub SortByPosition(ByRef positions() As Long, _
                            ByRef types() As Boolean, _
                            ByVal count As Long)
    Dim i       As Long
    Dim j       As Long
    Dim tmpPos  As Long
    Dim tmpType As Boolean

    For i = 1 To count - 1
        For j = 1 To count - i
            If positions(j) > positions(j + 1) Then
                tmpPos = positions(j)
                positions(j) = positions(j + 1)
                positions(j + 1) = tmpPos
                tmpType = types(j)
                types(j) = types(j + 1)
                types(j + 1) = tmpType
            End If
        Next j
    Next i
End Sub

' ============================================================
' BLANK WORD CHECK
' Returns the first non-exempt "blank" hit highlighted in cyan.
' *blank* (asterisks on both sides) is the intentional-use
' marker and is skipped. Returns Nothing if no hits found.
' ============================================================
Private Function FindFirstBlank(Doc As Document) As Range
    Dim rng        As Range
    Dim charBefore As String
    Dim charAfter  As String
    Dim exempt     As Boolean

    Set rng = Doc.content
    With rng.Find
        .ClearFormatting
        .text = "blank"
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        Do While .Execute
            exempt = False
            If rng.start > 0 And rng.End < Doc.content.End Then
                charBefore = Doc.Range(rng.start - 1, rng.start).text
                charAfter = Doc.Range(rng.End, rng.End + 1).text
                If charBefore = "*" And charAfter = "*" Then exempt = True
            End If
            If Not exempt Then
                rng.HighlightColorIndex = wdTurquoise
                Set FindFirstBlank = rng.Duplicate
                Exit Function
            End If
        Loop
    End With
End Function

' ============================================================
' DOUBLE SPACE CHECK � FAST-EXIT
' Finds the first double space in Doc.Content (main body only;
' headers, footers, footnotes, and text boxes are excluded).
' Highlights the two-space run in bright green and returns the
' Range. Returns Nothing if no double spaces are found.
' ============================================================
Private Function FindFirstDoubleSpace(Doc As Document) As Range
    Dim rng As Range
    Set rng = Doc.content
    With rng.Find
        .ClearFormatting
        .text = "  "                ' two literal spaces
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        If .Execute Then
            rng.HighlightColorIndex = wdBrightGreen
            Set FindFirstDoubleSpace = rng.Duplicate
        End If
    End With
End Function

' ============================================================
' DOUBLE SPACE CHECK � FULL RUN
' Highlights ALL double-space runs in bright green.
' Returns True if any are found.
' ============================================================
Private Function CheckDoubleSpaces(Doc As Document) As Boolean
    Dim rng As Range
    Set rng = Doc.content
    With rng.Find
        .ClearFormatting
        .text = "  "                ' two literal spaces
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        Do While .Execute
            rng.HighlightColorIndex = wdBrightGreen
            CheckDoubleSpaces = True
        Loop
    End With
End Function

' ============================================================
' HIGHLIGHT WORD (full run � all occurrences)
' Used by RunAllDocumentChecks for the "blank" check.
' ============================================================
Private Function HighlightWord(Doc As Document, word As String, _
                                color As Long) As Boolean
    Dim rng        As Range
    Dim charBefore As String
    Dim charAfter  As String
    Dim exempt     As Boolean

    Set rng = Doc.content
    With rng.Find
        .ClearFormatting
        .text = word
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        Do While .Execute
            exempt = False
            If rng.start > 0 And rng.End < Doc.content.End Then
                charBefore = Doc.Range(rng.start - 1, rng.start).text
                charAfter = Doc.Range(rng.End, rng.End + 1).text
                If charBefore = "*" And charAfter = "*" Then exempt = True
            End If
            If Not exempt Then
                rng.HighlightColorIndex = RGBToHighlightIndex(color)
                HighlightWord = True
            End If
        Loop
    End With
End Function

' ============================================================
' COLOR MAP
' HL_GREEN and HL_CYAN are sentinel constants with no RGB
' meaning. This function maps them to WdColorIndex values.
' ============================================================
Private Function RGBToHighlightIndex(color As Long) As WdColorIndex
    Select Case color
        Case HL_GREEN: RGBToHighlightIndex = wdBrightGreen
        Case HL_CYAN:  RGBToHighlightIndex = wdTurquoise
        Case Else:     RGBToHighlightIndex = wdBrightGreen
    End Select
End Function

' ============================================================
' HIGHLIGHT CLEARING
' ============================================================

' Removes only the macro's own colors (bright green and cyan).
' Called at the start of every check run to clear prior results.
Public Sub ClearCheckHighlights(Doc As Document)
    Dim rng As Range
    For Each rng In Doc.content.Characters
        If rng.HighlightColorIndex = wdBrightGreen Or _
           rng.HighlightColorIndex = wdTurquoise Then
            rng.HighlightColorIndex = wdNoHighlight
        End If
    Next rng
End Sub

' Removes all highlight colors except yellow (the user's own color).
' Called when the user chooses No (close anyway) after a full-run prompt.
Public Sub ClearAllHighlightsExceptYellow(Doc As Document)
    Dim rng As Range
    For Each rng In Doc.content.Characters
        If rng.HighlightColorIndex <> wdYellow And _
           rng.HighlightColorIndex <> wdNoHighlight Then
            rng.HighlightColorIndex = wdNoHighlight
        End If
    Next rng
End Sub

' ============================================================
' USER HIGHLIGHT DETECTION
' Finds any highlight color that is not one of the macro's
' two colors (bright green, cyan). Yellow counts as a user
' highlight because the user uses it for their own reminders.
' ============================================================
Public Function DocumentHasUserHighlights(Doc As Document) As Boolean
    Dim rng As Range
    Set rng = Doc.content
    With rng.Find
        .ClearFormatting
        .text = ""
        .Highlight = True
        .Wrap = wdFindStop
        Do While .Execute
            If rng.HighlightColorIndex <> wdBrightGreen And _
               rng.HighlightColorIndex <> wdTurquoise Then
                DocumentHasUserHighlights = True
                Exit Function
            End If
        Loop
    End With
End Function

' ============================================================
' RESTORE INTENTIONAL BLANKS
' *blank* is the marker for an intentional use of the word.
' After HighlightWord flags all "blank" occurrences, this
' restores any that were wrapped in asterisks back to plain
' "blank" with no highlight. MatchWildcards = True is
' intentional and isolated to this one function.
' ============================================================
Private Sub RestoreIntentionalBlanks(Doc As Document)
    With Doc.content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "\*blank\*"
        .Replacement.text = "blank"
        .MatchCase = False
        .MatchWildcards = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' ============================================================
' APOSTROPHE CONVERSION
' Converts straight single quotes (Chr 39) to their curly
' forms, direction-aware: a quote at the start of a story or
' after whitespace/opening punctuation is an OPENING quote
' (ChrW 8216); everything else -- possessives, contractions,
' closers -- is a closing/apostrophe mark (ChrW 8217). The old
' blanket 8217 replacement turned 'quoted term' into two
' closing quotes. (Elisions like 'tis after a space still get
' 8216; that rarity is accepted.) Always runs on every close
' attempt regardless of whether issues were found.
' ============================================================
Private Sub ConvertStraightApostrophes(Doc As Document)
    Dim rng As Range: Set rng = Doc.content
    With rng.Find
        .ClearFormatting
        .text = Chr(39)
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
        Do While .Execute
            Dim bOpen As Boolean: bOpen = False
            If rng.start = 0 Then
                bOpen = True
            Else
                Select Case Doc.Range(rng.start - 1, rng.start).text
                    Case " ", vbCr, vbTab, Chr(11), ChrW(160), "(", "[", ChrW(8220), Chr(34)
                        bOpen = True
                End Select
            End If
            If bOpen Then
                rng.text = ChrW(8216)
            Else
                rng.text = ChrW(8217)
            End If
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub

' ============================================================
' CHAR COUNT HELPER
' Simple whole-document count of a single character.
' Used only for the straight double quote odd/even test.
' ============================================================
Private Function CountChar(Doc As Document, ch As String) As Long
    Dim rng As Range
    Dim n   As Long
    n = 0
    Set rng = Doc.content
    With rng.Find
        .ClearFormatting
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        .text = ch
        Do While .Execute
            n = n + 1
        Loop
    End With
    CountChar = n
End Function



