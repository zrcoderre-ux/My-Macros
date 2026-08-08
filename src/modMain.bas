Attribute VB_Name = "modMain"
Option Explicit
' Internal plumbing (InitializeAppEvents, shared helpers). Keep it out of the
' Alt+F8 list; everything stays callable within the project -- including from the
' clsAppEvents / ThisDocument class modules.
Option Private Module

Public gAppEvents       As clsAppEvents
Public gSkipCloseChecks As Boolean          ' Set True by mail merge to suppress checks

' True for as long as App_DocumentBeforeClose is on the stack. Word has already
' committed to closing the document by then, so anything that yields to the
' message pump mid-run (DoEvents) can let a queued close -- an impatient second
' click on the X -- re-enter a document that is half torn down. The long sweeps
' the close hook calls read this and skip their yields; run by hand from Alt+F8
' they still yield as before, which is where the yields earn their keep.
Public gInCloseReview   As Boolean

Private Const HL_GREEN As Long = 1          ' maps to wdBrightGreen
Private Const HL_CYAN  As Long = 2          ' maps to wdTurquoise

' The application settings the review borrows for the length of a run, so they
' can be put back exactly -- including on the error path.
Private Type ReviewState
    saved      As Boolean
    screen     As Boolean
    pagination As Boolean
    spell      As Boolean
    grammar    As Boolean
    autoSave   As Boolean
    autoSaved  As Boolean       ' False when AutoSaveOn couldn't be read
End Type

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
' issues means SOMETHING IS HIGHLIGHTED. Nothing sets it without having marked
' the text it is reporting -- "Document contains possible issues" over a document
' with no mark in it sends the user hunting for something that was never there.
' A finding that cannot be highlighted goes into notes instead, which the caller
' shows as its own message saying what was found and why it isn't marked.
Public Sub RunAllDocumentChecks(ByVal Doc As Document, _
                                ByRef issues As Boolean, _
                                ByRef userHighlights As Boolean, _
                                Optional ByRef notes As String)
    issues = False
    userHighlights = False
    notes = ""

    ' Everything below runs from inside DocumentBeforeClose. Background
    ' repagination, check-as-you-type and AutoSave each re-process the document
    ' after every edit -- and on a synced OneDrive file AutoSave also pushes it
    ' back to the server, from inside the event that is closing it. Borrow all
    ' five settings for the run -- those three plus spelling and screen redraw --
    ' and hand them back on every path, error included: this is the same
    ' suppression the de-anonymize close hook already does for the same reason.
    Dim st As ReviewState
    st = SuppressForReview(Doc)

    Dim eN As Long, eD As String
    On Error GoTo Fail

    ClearCheckHighlights Doc

    ' Smart double quotes
    If CheckUnmatchedPairs(Doc, ChrW(8220), ChrW(8221), HL_GREEN) Then issues = True

    ' Straight double quotes (odd count � highlight the last one). CountChar spans
    ' every reviewed story, so the search for the last one has to as well -- a
    ' header quote used to be counted but then flagged on the wrong character in
    ' the body, or on nothing at all.
    If (CountChar(Doc, Chr(34)) Mod 2) <> 0 Then
        Dim qStory  As Range
        Dim rng     As Range
        Dim last    As Range
        Dim qLastEnd As Long
        For Each qStory In ReviewStories(Doc)
            Set rng = qStory.Duplicate
            qLastEnd = -1
            With rng.Find
                .ClearFormatting
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .Wrap = wdFindStop
                .text = Chr(34)
                Do While .Execute
                    ' Only a real straight quote counts: Word hands curly ones
                    ' back for this search too (see CountChar).
                    If rng.text = Chr(34) Then Set last = rng.Duplicate
                    ' Nothing in this loop changes the document, so a Find that
                    ' ever stopped advancing would spin here forever.
                    If rng.End <= qLastEnd Then Exit Do
                    qLastEnd = rng.End
                Loop
            End With
        Next qStory
        If Not last Is Nothing Then
            last.HighlightColorIndex = wdBrightGreen
            issues = True
        Else
            ' Counted an odd number but could not locate one to flag. This is the
            ' one check whose count and whose highlight come from separate
            ' passes, so it is the one that could report an issue with nothing
            ' marked -- which is exactly how it was read: a warning, and then a
            ' document with no highlight anywhere in it. Say what was counted
            ' instead of borrowing the highlight dialog's wording.
            notes = notes & "- An odd number of straight quotation marks (""), " & _
                    "so one is unclosed -- but it could not be located to " & _
                    "highlight. Search for "" yourself." & vbCrLf
        End If
    End If

    ' Square brackets, curly braces, parentheses
    If CheckUnmatchedPairs(Doc, "[", "]", HL_GREEN) Then issues = True
    If CheckUnmatchedPairs(Doc, "{", "}", HL_GREEN) Then issues = True
    If CheckUnmatchedPairs(Doc, "(", ")", HL_GREEN) Then issues = True

    ' Placeholder word "blank" (cyan; *blank* marks intentional use and is
    ' skipped, then stripped back to plain "blank" by the restore pass)
    If HighlightWord(Doc, "blank", HL_CYAN) Then issues = True
    RestoreIntentionalBlanks Doc

    ' Double spaces
    If CheckDoubleSpaces(Doc) Then issues = True

    ' Leftover anonymizer fakes: every pseudonym-pool word and placeholder email
    ' domain still in the document, flagged pink -- even when mashed inside a
    ' larger word -- so a fake that slipped into the draft is caught before the
    ' document is shared. Covers headers and footers as well as the body: the
    ' running header carries the court identity and the party names, so a fake
    ' left there is the most damaging one to miss. (This used to pass
    ' bodyOnly:=True because the clearers only swept the body; they now sweep
    ' every reviewed story, so a header flag can no longer be stranded.) The pool
    ' is ~870 words, but a term absent from the story costs one in-memory InStr
    ' and no Find at all (DeAnonymize.TermMaybePresent), so the sweeps that
    ' actually run are the handful of words the document contains; screen redraw
    ' is already suspended for the whole run by SuppressForReview.
    If DeAnonymize.HighlightResidualPseudonyms(Doc) > 0 Then issues = True

    ' Apostrophe conversion (always runs, no prompt)
    ConvertStraightApostrophes Doc

    ' User highlight check (runs after macro colors are in place)
    If DocumentHasUserHighlights(Doc) Then userHighlights = True

    GoTo Cleanup

Fail:
    eN = Err.Number
    eD = Err.Description

Cleanup:
    RestoreAfterReview Doc, st
    If eN <> 0 Then
        ' Never let this escape into App_DocumentBeforeClose. An unhandled error
        ' inside a Word application event fires with the document already half
        ' closed, and Word has no safe way to unwind it.
        MsgBox "The document review hit an error and stopped:" & vbCrLf & vbCrLf & _
               "Error " & eN & ": " & eD & vbCrLf & vbCrLf & _
               "Anything found before that point is still highlighted.", _
               vbExclamation, "Document Check"
    End If
End Sub

' ============================================================
' REVIEW STATE
' Borrow the five application settings that turn a close-time
' review into a document-wide reprocessing storm, and hand them
' back exactly. Pagination, spell and grammar re-run after every
' edit and are not covered by ScreenUpdating; AutoSave on a
' synced OneDrive file also uploads, from inside the event that
' is closing the document; ScreenUpdating itself is the fifth.
' ============================================================
Private Function SuppressForReview(ByVal Doc As Document) As ReviewState
    Dim s As ReviewState
    On Error Resume Next

    s.screen = Application.ScreenUpdating
    s.pagination = Application.Options.Pagination
    s.spell = Application.Options.CheckSpellingAsYouType
    s.grammar = Application.Options.CheckGrammarAsYouType
    s.saved = True

    Application.ScreenUpdating = False
    Application.Options.Pagination = False
    Application.Options.CheckSpellingAsYouType = False
    Application.Options.CheckGrammarAsYouType = False

    ' AutoSaveOn is missing on older builds and raises on some document types,
    ' so track separately whether we actually took it.
    Err.Clear
    s.autoSave = Doc.AutoSaveOn
    If Err.Number = 0 Then
        Doc.AutoSaveOn = False
        s.autoSaved = (Err.Number = 0)
    End If

    On Error GoTo 0
    SuppressForReview = s
End Function

' Safe to call twice, and safe when SuppressForReview never ran (saved = False).
Private Sub RestoreAfterReview(ByVal Doc As Document, ByRef s As ReviewState)
    If Not s.saved Then Exit Sub
    On Error Resume Next
    If s.autoSaved Then Doc.AutoSaveOn = s.autoSave
    Application.Options.Pagination = s.pagination
    Application.Options.CheckSpellingAsYouType = s.spell
    Application.Options.CheckGrammarAsYouType = s.grammar
    Application.ScreenUpdating = s.screen
    s.saved = False
    On Error GoTo 0
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
    Dim story       As Range
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

    ' Every reviewed story, not just Doc.Paragraphs (the body). The collected
    ' positions are coordinates in the CURRENT story, so the ranges built from
    ' them below are cut from that story too -- Doc.Range(...) would resolve them
    ' against the body and highlight unrelated text for a header hit.
    For Each story In ReviewStories(Doc)
    For Each PARA In story.Paragraphs
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
                    ' Unmatched closer � highlight it, cut from THIS story
                    Dim closeRng As Range
                    Set closeRng = story.Duplicate
                    closeRng.SetRange positions(i), positions(i) + 1
                    closeRng.HighlightColorIndex = RGBToHighlightIndex(color)
                    CheckUnmatchedPairs = True
                End If
            End If
        Next i

        ' Anything left on the stack is an unmatched opener
        For i = 1 To stackTop
            Dim openRng As Range
            Set openRng = story.Duplicate
            openRng.SetRange stack(i), stack(i) + 1
            openRng.HighlightColorIndex = RGBToHighlightIndex(color)
            CheckUnmatchedPairs = True
        Next i

NextParaFull:
    Next PARA
    Next story
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
'
' BODY ONLY, deliberately -- unlike the other checks, which span the
' headers and footers too. A running header lays out its caption with
' runs of spaces on purpose, so flagging them is noise, not a defect.
' The checks that DO cover the header are the ones where a header hit
' means something: a leftover pseudonym, an unmatched bracket, a stray
' "blank".
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
    Dim story      As Range
    Dim rng        As Range
    Dim exempt     As Boolean

    For Each story In ReviewStories(Doc)
        Set rng = story.Duplicate
        With rng.Find
            .ClearFormatting
            .text = word
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .Wrap = wdFindStop
            Do While .Execute
                ' *blank* marks an intentional use. Probe within the match's OWN
                ' story -- the old Doc.Range(...) probe used body coordinates,
                ' which would read unrelated body text for a header match.
                exempt = (CharBeforeRange(rng) = "*" And CharAfterRange(rng) = "*")
                If Not exempt Then
                    rng.HighlightColorIndex = RGBToHighlightIndex(color)
                    HighlightWord = True
                End If
            Loop
        End With
    Next story
End Function

' ============================================================
' REVIEW STORIES
' Every story the close review inspects: the main body plus
' each section's NON-EMPTY headers and footers, and the
' footnote/endnote stories when present.
'
' The review used to read Doc.Content alone -- the body. A
' running header is a separate story, so nothing in it was ever
' checked: not a leftover pseudonym, not a double space, not an
' unmatched bracket. That is the worst place to miss one, since
' the header is where the court identity and the party names
' sit and it repeats on every page.
'
' Empty header/footer slots (most of a section's six) hold
' nothing but a paragraph mark and are dropped here so no check
' pays for them.
' ============================================================
Private Function ReviewStories(Doc As Document) As Collection
    Dim c As New Collection
    On Error Resume Next

    c.Add Doc.content

    Dim sec As Section, hf As HeaderFooter
    For Each sec In Doc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                If StoryHasText(hf.Range) Then c.Add hf.Range
            End If
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then
                If StoryHasText(hf.Range) Then c.Add hf.Range
            End If
        Next hf
    Next sec

    If Doc.Footnotes.count > 0 Then c.Add Doc.StoryRanges(wdFootnotesStory)
    If Doc.Endnotes.count > 0 Then c.Add Doc.StoryRanges(wdEndnotesStory)

    Set ReviewStories = c
End Function

' True when a story holds something other than paragraph/cell marks.
Private Function StoryHasText(rng As Range) As Boolean
    On Error Resume Next
    Dim t As String
    t = rng.text
    StoryHasText = (Len(Trim$(Replace(Replace(t, vbCr, ""), Chr$(7), ""))) > 0)
End Function

' The character immediately before / after a range, WITHIN the range's own
' story. Doc.Range(pos - 1, pos) cannot be used for this: its coordinates are
' the BODY's, so probing a header match that way reads unrelated body text.
Private Function CharBeforeRange(rng As Range) As String
    On Error Resume Next
    Dim p As Range: Set p = rng.Duplicate
    p.Collapse Direction:=wdCollapseStart
    If p.MoveStart(wdCharacter, -1) <> 0 Then CharBeforeRange = p.text
End Function

Private Function CharAfterRange(rng As Range) As String
    On Error Resume Next
    Dim p As Range: Set p = rng.Duplicate
    p.Collapse Direction:=wdCollapseEnd
    If p.MoveEnd(wdCharacter, 1) <> 0 Then CharAfterRange = p.text
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

' Removes only the macro's own colors (bright green, cyan, and pink).
' Called at the start of every check run to clear prior results.
' Uses a highlight-seeking Find (jumps between highlighted runs) instead of
' walking Doc.Content.Characters one COM call at a time, which froze Word
' for minutes on long documents.
' Sweeps every reviewed story, not just the body: the checks now flag issues in
' headers and footers, so their colors have to be clearable there too or a stale
' flag would be stranded where nothing can remove it.
Public Sub ClearCheckHighlights(Doc As Document)
    Dim story   As Range
    Dim rng     As Range
    Dim lastEnd As Long
    For Each story In ReviewStories(Doc)
        Set rng = story.Duplicate
        lastEnd = -1
        With rng.Find
            .ClearFormatting
            .text = ""
            .Highlight = True
            .Wrap = wdFindStop
            Do While .Execute
                If rng.HighlightColorIndex = wdBrightGreen Or _
                   rng.HighlightColorIndex = wdTurquoise Or _
                   rng.HighlightColorIndex = wdPink Then
                    rng.HighlightColorIndex = wdNoHighlight
                End If
                ' Guard against a zero-progress infinite loop.
                If rng.End <= lastEnd Then Exit Do
                lastEnd = rng.End
                rng.Collapse Direction:=wdCollapseEnd
                rng.End = story.End
                If rng.start >= rng.End Then Exit Do
            Loop
        End With
    Next story
End Sub

' Removes all highlight colors except yellow (the user's own color).
' Called when the user chooses No (close anyway) after a full-run prompt.
' Uses a highlight-seeking Find (jumps straight between highlighted runs) --
' the same rewrite ClearCheckHighlights got, for the same reason: this used to
' walk Doc.Content.Characters one COM call per character, freezing Word for the
' length of the document at the exact moment the user had asked to close it.
Public Sub ClearAllHighlightsExceptYellow(Doc As Document)
    Dim story   As Range
    Dim rng     As Range
    Dim lastEnd As Long
    For Each story In ReviewStories(Doc)
        Set rng = story.Duplicate
        lastEnd = -1
        With rng.Find
            .ClearFormatting
            .text = ""
            .Highlight = True
            .Wrap = wdFindStop
            Do While .Execute
                Dim hci As Long
                hci = rng.HighlightColorIndex
                If hci = wdUndefined Then
                    ' The found run mixes colors (e.g. yellow butted against
                    ' green). Resolve per character WITHIN this run only -- a
                    ' handful of characters -- so yellow inside the run survives
                    ' and the whole-story character walk never returns.
                    Dim ch As Range
                    For Each ch In rng.Characters
                        If ch.HighlightColorIndex <> wdYellow And _
                           ch.HighlightColorIndex <> wdNoHighlight Then
                            ch.HighlightColorIndex = wdNoHighlight
                        End If
                    Next ch
                ElseIf hci <> wdYellow And hci <> wdNoHighlight Then
                    rng.HighlightColorIndex = wdNoHighlight
                End If
                ' Guard against a zero-progress infinite loop.
                If rng.End <= lastEnd Then Exit Do
                lastEnd = rng.End
                rng.Collapse Direction:=wdCollapseEnd
                rng.End = story.End
                If rng.start >= rng.End Then Exit Do
            Loop
        End With
    Next story
End Sub

' ============================================================
' USER HIGHLIGHT DETECTION
' Finds any highlight color that is not one of the macro's
' own colors (bright green, cyan, pink). Yellow counts as a
' user highlight because the user uses it for their own
' reminders.
'
' This is the third highlight-seeking sweep in this module, and the only one
' that never modifies what it finds -- so it is the one that most needs the
' manual advance the other two carry. Its siblings remove the highlight, which
' is what moved them along; here a run painted one of the macro's own colors
' matches the criteria on every pass, and without the advance below the loop
' re-finds the same run until Word stops responding. That is exactly the state
' this function is called in: it runs LAST in RunAllDocumentChecks, after the
' checks have painted green, cyan and pink over the document.
' ============================================================
Public Function DocumentHasUserHighlights(Doc As Document) As Boolean
    Dim story   As Range
    Dim rng     As Range
    Dim lastEnd As Long
    For Each story In ReviewStories(Doc)
        Set rng = story.Duplicate
        lastEnd = -1
        With rng.Find
            .ClearFormatting
            .text = ""
            .Highlight = True
            .Wrap = wdFindStop
            Do While .Execute
                If rng.HighlightColorIndex <> wdBrightGreen And _
                   rng.HighlightColorIndex <> wdTurquoise And _
                   rng.HighlightColorIndex <> wdPink Then
                    DocumentHasUserHighlights = True
                    Exit Function
                End If
                ' Step past this run by hand, and guard against a zero-progress
                ' loop -- the same pair ClearCheckHighlights uses.
                If rng.End <= lastEnd Then Exit Do
                lastEnd = rng.End
                rng.Collapse Direction:=wdCollapseEnd
                rng.End = story.End
                If rng.start >= rng.End Then Exit Do
            Loop
        End With
    Next story
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
    Dim story As Range
    For Each story In ReviewStories(Doc)
        With story.Duplicate.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = "\*blank\*"
            .Replacement.text = "blank"
            .MatchCase = False
            .MatchWildcards = True
            .Wrap = wdFindStop      ' story-scoped: wdFindContinue would escape it
            .Execute Replace:=wdReplaceAll
        End With
    Next story
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
'
' TWO GATES, and they are the whole point of this pass being safe to run from
' inside DocumentBeforeClose.
'
' Word's Find does not honor the distinction this macro is built on. With
' AutoCorrect's "straight quotes with smart quotes" on -- the default -- a search
' for Chr(39) matches the CURLY forms as well. So on an ordinary ruling, where
' every apostrophe is already curly, this loop matched all of them and assigned
' Range.Text to each one: dozens of delete-and-reinsert edits that changed
' nothing, made from a live Find loop, while Word was already tearing the
' document down and AutoSave was pushing each one back to OneDrive. That is the
' only write the whole close review performs, and it scaled with the number of
' apostrophes in the document rather than with the number of straight ones --
' which on every document that reaches this review is zero.
'
' So: skip a story outright unless its text really holds a Chr(39), and re-read
' each hit before writing it, passing over the curly quotes Word hands back. A
' document with no straight quotes now leaves the review a pure read, and a
' document that has them is edited exactly where it needs to be.
' ============================================================
Private Sub ConvertStraightApostrophes(Doc As Document)
    Dim story As Range, rng As Range
    Dim bOpen As Boolean
    Dim prev  As String

    For Each story In ReviewStories(Doc)
        ' Gate 1: no straight quote in this story, nothing to convert.
        If Not StoryHasChar(story, Chr(39)) Then GoTo NextStory

        Set rng = story.Duplicate
        With rng.Find
            .ClearFormatting
            .text = Chr(39)
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
            Do While .Execute
                ' Gate 2: this hit may be a curly quote Word matched for us.
                If rng.text = Chr(39) Then
                    ' Probe the preceding character within the match's OWN story;
                    ' Doc.Range(...) would read body coordinates for a header hit.
                    bOpen = False
                    prev = CharBeforeRange(rng)
                    If Len(prev) = 0 Then
                        bOpen = True                 ' start of the story
                    Else
                        Select Case prev
                            Case " ", vbCr, vbTab, Chr(11), ChrW(160), "(", "[", ChrW(8220), Chr(34)
                                bOpen = True
                        End Select
                    End If
                    If bOpen Then
                        rng.text = ChrW(8216)
                    Else
                        rng.text = ChrW(8217)
                    End If
                End If
                ' Collapse whether or not the hit was ours, so a curly quote is
                ' stepped over rather than found again.
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With

NextStory:
    Next story
End Sub

' True when a story's text really contains ch. Word's Find cannot answer this for
' a quote character -- it treats the straight and curly forms as the same thing
' (see ConvertStraightApostrophes) -- so read the story once and ask the string.
'
' Fails OPEN, the same way the pseudonym scan's filter does: a story whose text
' can't be read, or that reads back empty when it plainly isn't, reports True and
' pays for a Find sweep. This gate is only ever an optimization, so a false True
' costs a scan that writes nothing, while a false False would silently stop
' converting a document's straight quotes.
Private Function StoryHasChar(rng As Range, ch As String) As Boolean
    Dim t      As String
    Dim readOK As Boolean

    On Error Resume Next
    t = rng.text
    readOK = (Err.Number = 0)
    On Error GoTo 0

    If Not readOK Then
        StoryHasChar = True
    ElseIf Len(t) = 0 And rng.End > rng.start Then
        StoryHasChar = True
    Else
        StoryHasChar = (InStr(1, t, ch, vbBinaryCompare) > 0)
    End If
End Function

' ============================================================
' CHAR COUNT HELPER
' Counts a single character across every reviewed story.
' Used only for the straight double quote odd/even test.
'
' Each hit is re-read before it is counted. A search for Chr(34) matches the
' curly quotes too (see ConvertStraightApostrophes), so this used to count every
' smart quote in the document as a straight one -- and any ruling with an odd
' number of them was reported as having an unmatched straight double quote,
' highlighted on a character the user never typed straight.
' ============================================================
Private Function CountChar(Doc As Document, ch As String) As Long
    Dim story   As Range
    Dim rng     As Range
    Dim n       As Long
    Dim lastEnd As Long
    n = 0
    For Each story In ReviewStories(Doc)
        Set rng = story.Duplicate
        lastEnd = -1
        With rng.Find
            .ClearFormatting
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .text = ch
            Do While .Execute
                If rng.text = ch Then n = n + 1
                ' Nothing here changes the document, so a Find that ever stopped
                ' advancing would spin forever.
                If rng.End <= lastEnd Then Exit Do
                lastEnd = rng.End
            Loop
        End With
    Next story
    CountChar = n
End Function



