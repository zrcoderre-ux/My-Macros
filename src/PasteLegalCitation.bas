Attribute VB_Name = "PasteLegalCitation"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    ' 32-bit fallback (not used in 64-bit Office installations)
    ' Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'===========================================================
' CONFIGURATION
' Adjust these two constants to match your environment.
'===========================================================
Private Const CITE_SERVER_URL  As String = "http://localhost:54321"
Private Const CITE_APP_PATH    As String = _
    "C:\Users\ZCoderre\Los Angeles Superior Court\" & _
    "Research Attorney and Law Clerk Unit - Zachary Coderre\app.py"
Private Const CITE_PYTHON_EXE  As String = "python"  ' or full path to python.exe

' Single-step undo (one Ctrl+Z unwinds the whole paste) is restored only on
' documents at or below this size -- character count of the main story. On a
' larger document the custom UndoRecord that spans the paste's hyperlink-field
' deletions can destabilise Word's undo stack and hard-crash it, so above this
' size the paste runs without the record (undo then takes several presses, but
' no crash). Roughly ~1,500 chars per double-spaced page, so 15,000 ~= 10 pages.
' Lower it if a large document still crashes; raise it if undo is off on
' documents you consider short.
Private Const MAX_UNDO_DOC_CHARS As Long = 15000

' Session-level flag: True once we have confirmed the server is reachable
' (or have exhausted the one startup attempt).  Prevents repeated delays.
Private g_ServerVerified As Boolean

' Diagnostic: the pipeline phase currently executing. Surfaced in the error
' handler so an unexpected runtime error (e.g. 4609 "Value out of range") can
' be traced to the step that raised it. Set as the pipeline advances.
Private g_Phase As String

'===========================================================
' Pleading-paste configuration (markers from pdf_linker.py)
'===========================================================
' Marker bracket size limit when scanning -- markers from pdf_linker.py
' top out around 30 chars; 80 is a comfortable ceiling.
Private Const MARKER_MAX_LEN   As Integer = 80
' Maximum markers we'll parse from a single paste.
Private Const MARKER_MAX_COUNT As Integer = 500

' Parallel-array storage for parsed pleading markers within one paste.
' Reused across the pleading pipeline; reset by ParseAllMarkers.
Private g_MarkerCount As Integer
Private g_MarkerPos()  As Long      ' 1-based char index in oRange.text where marker starts
Private g_MarkerLen()  As Integer   ' total marker character count (including brackets)
Private g_MarkerDoc()  As String    ' shortname or "" if compact form
Private g_MarkerPage() As Integer
Private g_MarkerLine() As Integer
Private g_MarkerPara() As Integer   ' 0 if no paragraph in marker

'===========================================================
' Main entry point
'===========================================================
Sub PasteLegalQuotation()

    Dim oDoc As Document
    Dim oSel As Selection
    Set oDoc = ActiveDocument
    Set oSel = Selection

    ' Main body only. Selection.Start in a footnote/header/text box is an offset
    ' into THAT story, but the whole pipeline addresses oDoc.Range -- the main
    ' text story -- so running anywhere else would read and mutate body text at
    ' unrelated offsets (the pre-paste detectors DELETE characters there).
    If oSel.StoryType <> wdMainTextStory Then
        MsgBox "PasteLegalQuotation only works in the main body of the " & _
               "document, not in a footnote, header/footer, or text box.", _
               vbInformation, "PasteLegalQuotation"
        Exit Sub
    End If

    ' Track Changes off for the run, restored in CleanUp. With revisions on,
    ' deletions don't shrink the story, so the lTailLen invariant and every
    ' snapshot offset drift and the output lands garbled.
    Dim bPrevTrack As Boolean
    bPrevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False

    ' Restore single-step undo (one Ctrl+Z unwinds the whole paste) -- but only
    ' when the document is small enough to be safe. A custom UndoRecord spanning
    ' this macro's edits (including the hyperlink-field deletions in
    ' HarvestAndRemoveHyperlinks) can destabilise Word's undo stack and hard-crash
    ' the app on a large document. So we open the record only under
    ' MAX_UNDO_DOC_CHARS; above that, the paste runs without it -- undo takes
    ' several presses, but no crash. The record is always closed in CleanUp.
    Dim bUseUndo As Boolean
    bUseUndo = (oDoc.content.End <= MAX_UNDO_DOC_CHARS)
    Dim oUndo As UndoRecord
    If bUseUndo Then
        Set oUndo = Application.UndoRecord
        oUndo.StartCustomRecord "Paste Legal Quotation"
    End If

    ' FIX 1: Wrap entire body in an error handler so an unexpected runtime
    '         error still lands on CleanUp and reports its phase.
    On Error GoTo CleanUp
    g_Phase = "pre-paste detection & paste"

    ' Detect and remove pre-paste opener marks before pasting
    ' so that position tracking is clean after the paste lands
    Dim lStart As Long
    Dim nPreDouble As Integer
    Dim nPreSingle As Integer
    Dim sPreOpeners As String
    Dim bProperNoun As Boolean
    Dim nFootnoteCount As Integer
    nFootnoteCount = 0
    lStart = oSel.start
    DetectPrePasteOpeners oDoc, lStart, nPreDouble, nPreSingle, sPreOpeners, bProperNoun

    ' Detect subdivision markers typed immediately before the cursor,
    ' e.g. the user typed "(a)(1)" before pressing Ctrl+Shift+V.
    ' Runs before paste so the markers are removed from the document
    ' before lStart is used for position tracking.
    ' Only used if bIsStatute turns out to be True after paste.
    Dim sSubdivision As String
    sSubdivision = DetectPrePasteSubdivision(oDoc, lStart)

    ' Detect open bracket immediately before the cursor (with optional
    ' trailing spaces). Signals parenthetical mode: the citation sentence
    ' wraps the quoted passage as  (Citation ["Passage"].)
    ' Runs after DetectPrePasteSubdivision so that a combined signal like
    ' "[(a)" is handled correctly: subdivision is stripped first, then the
    ' bracket is detected against the adjusted lStart.
    Dim bParenthetical As Boolean
    bParenthetical = DetectPrePasteParenthetical(oDoc, lStart)

    ' Detect "in"/"In" immediately before the cursor (peek only, no deletion).
    ' Signals textual-sentence mode: citation moves before the quote,
    ' loses its parens/period, and "the court held" is inserted after it.
    Dim bTextualSentence As Boolean
    bTextualSentence = DetectPrePasteTextual(oDoc, lStart)

    ' Detect "Defendant cites"/"Plaintiff cites" (and variants) immediately
    ' before the cursor (peek only, no deletion).
    ' Signals cites-sentence mode: citation loses its parens/period and
    ' "for the proposition that" is inserted between citation and passage.
    Dim bCitesSentence As Boolean
    bCitesSentence = DetectPrePasteCites(oDoc, lStart)

    ' Detect " " typed immediately before the cursor (subdivision
    ' signal already removed by DetectPrePasteSubdivision).
    ' Signals statutory-textual mode: no citation sentence is pasted;
    ' macro builds: CodeTextual section N[, subdivision (x)], provides
    ' "Passage." (CodeName,   N[, subd. (x)].)
    Dim bStatutoryTextual As Boolean
    bStatutoryTextual = False
    DetectPrePasteStatutoryTextual oDoc, lStart, bStatutoryTextual

    oSel.PasteAndFormat wdFormatOriginalFormatting

    ' lSelEnd tracks the end of the pasted block. Initialised from
    ' oSel.End once, then refreshed on demand from lTailLen (see below)
    ' before every range construction.
    Dim lSelEnd As Long
    lSelEnd = oSel.End

    ' lTailLen: invariant distance from end-of-pasted-block to
    ' end-of-document. Mutations inside the block (add or delete
    ' characters) change Content.End but not lTailLen. We recompute
    ' lSelEnd = Content.End - lTailLen before every range
    ' construction so it is always current without each mutating
    ' sub having to pass lSelEnd ByRef.
    Dim lTailLen As Long
    lTailLen = oDoc.content.End - lSelEnd

    ' If the paste landed as a NEW paragraph -- a paragraph mark now sits at the
    ' very start of the pasted block, separating it from the pre-existing text --
    ' advance lStart past that mark so the pipeline processes only the pasted
    ' content. Without this, FindQuoteEndByParagraph also sees the pre-existing
    ' paragraph and can return a passage boundary BEFORE lStart, making
    ' Range(lStart, lQuoteEnd) throw "Value out of range" (error 4608); it would
    ' also fold the separating mark into the passage as a visible pilcrow.
    Do While lStart < lSelEnd
        Dim oLeadChk As Range
        Set oLeadChk = oDoc.Range(lStart, lStart + 1)
        Dim nLead As Long: nLead = AscW(oLeadChk.text)
        Set oLeadChk = Nothing
        If nLead = 13 Or nLead = 11 Then
            lStart = lStart + 1
        Else
            Exit Do
        End If
    Loop

    Dim oRange As Range
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' --- Lexis+ hyperlink harvest (before Step 1) ---
    ' Extract URL and citation metadata from any Lexis+ hyperlink in the
    ' pasted range, POST to the citation repo server, then strip the
    ' hyperlink field so no live link remains in the Word document.
    HarvestAndRemoveHyperlinks oDoc, oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' -------------------------------------------------------
    ' Auto-detect statute vs. case from the citation sentence
    ' -------------------------------------------------------
    Dim bIsStatute As Boolean
    Dim bMultiSubsection As Boolean
    Dim bNumericSubParagraphs As Boolean
    bIsStatute = DetectStatutePaste(oRange)

    g_Phase = "Steps 1-6: paste cleanup (shapes, line numbers, pleading, headnotes, page/parallel cites)"

    ' Step 1: Remove Westlaw flag images (inline shapes)
    RemoveInlineShapes oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)


    ' Step 1b: Strip clipped pleading line numbers at paragraph starts.
    ' Self-gated -- only fires when 3+ ascending digit paragraph-starts
    ' are detected (always pleading-paste artifact, never real text).
    ' Runs before the pleading-paste branch so that the marker scan in
    ' Step 1c sees clean paragraph starts.
    RemovePleadingLineNumbers oDoc, oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 1c: Pleading-paper paste pipeline.
    ' Detects right-margin markers from pdf_linker.py and, if present,
    ' builds a single trailing citation summarizing the range, folds
    ' the PDF line-break paragraph marks into spaces, applies the
    ' alternating-quote conversion chain, and SKIPS the rest of the
    ' case / statute pipeline -- pleading pastes have none of the
    ' artifacts (Lexis headnotes, Westlaw page markers, parallel
    ' citations) that the main pipeline cleans up.
    '
    ' The presence of markers is itself the signal that this is a
    ' PDF paste, and this macro is the quote-paste macro (a separate
    ' macro handles non-quote pastes), so no further gating is needed.
    Dim lPleadEnd As Long
    lPleadEnd = lSelEnd
    If ProcessPleadingPaste(oDoc, lStart, lPleadEnd) Then
        ' Pleading paste handled. Apply final formatting and jump to
        ' the cleanup label so the undo record closes properly.
        Dim oPleadFinal As Range
        Set oPleadFinal = oDoc.Range(lStart, lPleadEnd)
        FixParagraphSpacing oPleadFinal
        FormatQuotation oPleadFinal
        FixSurroundingFont oDoc, lStart, lPleadEnd
        Set oPleadFinal = Nothing
        GoTo CleanUp
    End If

    ' Steps 2-6: Case-only artifact removal
    If Not bIsStatute Then

        ' Step 2: Remove Lexis+ headnotes
        RemoveLexisHeadnotes oRange
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)

        ' Step 3: Remove Lexis+ page markers [*7] [**12] [***19]
        RemovePageReferences oRange
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)

        ' Step 4: Remove Westlaw page markers *123 **123 (bold)
        RemoveWestlawPageNumbers oRange
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)

        ' Step 5: Remove Lexis+ parallel citations in brackets
        RemoveParallelCitationsManual oRange, lStart
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)

        ' Step 6: Remove Westlaw parallel citations after official cite
        RemoveWestlawParallelCitations oRange
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)

    Else

        ' Statutes still carry Lexis+ page markers; remove them
        RemovePageReferences oRange
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)

    End If

    g_Phase = "Steps 7-10: spacing, soft returns, statute normalize, subdivisions"

    ' Step 7: Replace non-breaking spaces except after
    ReplaceNonBreakingSpaces oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 8: Convert soft returns to hard paragraph returns
    ConvertSoftReturns oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 9: Reset paragraph spacing, reapply Normal style
    FixParagraphSpacing oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 9a: Normalize Lexis+ statute citations that arrive without
    ' proper punctuation or wrapping parens.
    ' (lTailLen auto-refresh handles lSelEnd update; lNormEnd plumbing
    ' retained for compatibility but no longer required.)
    If bIsStatute Then
        Dim lNormEnd As Long
        lSelEnd = oDoc.content.End - lTailLen
        lNormEnd = lSelEnd
        NormalizeLexisStatuteCitation oDoc, oRange, lNormEnd
        lSelEnd = lNormEnd
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
    End If

    ' Step 9b: Wrap Westlaw citation in parens   cases only
    ' (Statute citations are already parenthesised by Lexis+)
    If Not bIsStatute Then
        WrapWestlawCitation oDoc, oRange
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
    End If

    ' Step 9c: Remove trailing paragraph mark at end of pasted range
    RemoveTrailingParagraphMark oDoc, oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 9d: Remove spurious spaces inside words after apostrophes
    RemoveSpacesAfterApostrophes oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 10: Collapse multiple spaces to one
    CollapseMultipleSpaces oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 10b: Remove spaces between adjacent curly quote marks (Westlaw artifact).
    ' Skipped for statutes: the space between a closing single and closing double
    ' is meaningful punctuation, not an artifact (e.g. 'foo,' "bar").
    If Not bIsStatute Then
        RemoveSpacesBetweenQuotes oRange
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
    End If

    ' Step 10c: Remove extra spaces immediately after opening parenthesis
    RemoveSpacesAfterOpenParen oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 10d-pre: Detect multi-subsection paste (e.g. (a)...(b)...).
    If bIsStatute Then
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
        bMultiSubsection = DetectMultiSubsection(oRange)
        If bMultiSubsection Then sSubdivision = ""
    End If

    ' Step 10d: Extract leading subdivision markers from the passage start.
    ' Only runs for statute pastes and only if the user did not already
    ' supply markers via the pre-paste method.
    If bIsStatute And sSubdivision = "" And Not bMultiSubsection Then
        lSelEnd = oDoc.content.End - lTailLen
        sSubdivision = ExtractLeadingSubdivision(oDoc, lStart, lSelEnd)
        If sSubdivision <> "" Then
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
        End If
    End If

    ' Step 10e: Normalize numeric sub-paragraphs (e.g. (1)...(2)...) that follow
    ' an extracted leading letter subdivision -- FixNumericSubParagraphs inserts
    ' the missing space after each "(N)" so the passage reads "(1) text". The
    ' bNumericSubParagraphs flag it returns is no longer used to gate the
    ' citation insertion (that would drop the letter subdivision); it is kept
    ' only because the sub reports it.
    bNumericSubParagraphs = False
    If bIsStatute And sSubdivision <> "" Then
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
        FixNumericSubParagraphs oDoc, oRange, bNumericSubParagraphs
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
    End If

    g_Phase = "Steps 11-12: passage boundary, blank-paragraph merge, subdivision insert, citation-only"

    ' Step 11: Find passage end BEFORE removing blank paragraphs
    ' citation is still on its own paragraph here, most reliable boundary
    Dim lQuoteEnd As Long
    lQuoteEnd = FindQuoteEndByParagraph(oRange)
    ' Defensive: never let the passage boundary fall before the pasted block
    ' (would make Range(lStart, lQuoteEnd) throw error 4608). Fall back to the
    ' paren-depth boundary, which is measured within the pasted content.
    If lQuoteEnd < lStart Then lQuoteEnd = FindQuoteEnd(oRange)

    ' Step 11b: Remove Lexis+ publisher parenthetical   statutes only
    ' e.g. "(Deering, Lexis Advance through Ch. 6 ...)"
    ' Scoped to the citation portion so it cannot touch the passage.
    If bIsStatute Then
        Dim oCitationSearch As Range
        lSelEnd = oDoc.content.End - lTailLen
        Set oCitationSearch = oDoc.Range(lQuoteEnd, lSelEnd)
        RemoveLexisStatuteParenthetical oDoc, oCitationSearch
        Set oCitationSearch = Nothing
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
    End If

    ' Step 11c: Remove blank paragraphs between passage and citation
    RemoveBlankParagraphs oRange
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

   ' Step 11e: Replace paragraph marks inside the passage with [ ].
    ' RemoveBlankParagraphs already deleted the Chr(13) before the
    ' citation sentence; every Chr(13) still in [lStart, lQuoteEnd]
    ' is a mid-passage line break that becomes a visible pilcrow.
    ' Scoped to passage only so the citation sentence is untouched.
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lQuoteEnd)
    With oRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .text = Chr(13)
        .Replacement.text = " [" & Chr(182) & "] "
        .Execute Replace:=wdReplaceAll
    End With
    lQuoteEnd = lQuoteEnd + (oDoc.content.End - lTailLen - lSelEnd)
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 11d: Remove Lexis+ footnote markers while bold is still intact.
    ' Must run before step 12 which strips all bold from the block.
    ' The tag ([Fn. omitted.] / [Fns. omitted.]) is inserted after the
    ' outer closing quote mark, which does not exist yet at this point
    ' RemoveFootnotes stores the count and inserts the tag in a second pass
    ' that runs at step 18c, after WrapInDoubleQuotes.
    RemoveFootnotesPass1 oDoc, oRange, nFootnoteCount
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    ' Step 12: Strip bold from entire block
    oRange.Font.Bold = False

    ' Step 12b: Insert the extracted leading subdivision into the citation
    ' sentence, e.g. "(Code Civ. Proc., § 473, subd. (a).)". Runs after
    ' RemoveLexisStatuteParenthetical has cleaned the citation and after
    ' RemoveBlankParagraphs has merged the block onto one line.
    '
    ' This runs even when numeric sub-paragraphs ((1), (2), ...) are present:
    ' those stay in the passage, but the letter subdivision that
    ' ExtractLeadingSubdivision already removed from the passage must still land
    ' in the citation -- otherwise "(a)" is eaten without reappearing anywhere.
    ' (The multi-subsection case (a)...(b)... never reaches here: it leaves
    ' sSubdivision = "" and keeps every subdivision in the passage.)
    If bIsStatute And sSubdivision <> "" Then
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
        Dim lQEForSubd As Long
        lQEForSubd = FindQuoteEnd(oRange)
        InsertSubdivisionIntoCitation oDoc, lQEForSubd, oRange.End, sSubdivision
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
    End If

    ' --- Citation-only detection ---
    Dim bCitationOnly As Boolean
    bCitationOnly = False

    ' lQuoteEnd was set at Step 11 by FindQuoteEndByParagraph while the
    ' citation was still on its own paragraph -- that value is reliable.
    ' Do NOT recalculate here with FindQuoteEndByParagraph: after
    ' RemoveBlankParagraphs merges everything onto one paragraph it falls
    ' back to FindQuoteEnd (paren-depth scan), which mis-identifies the
    ' passage boundary when the passage ends with parens like "(c), below".
    ' RemoveBlankParagraphs only removes content after lQuoteEnd, so the
    ' position is still valid.
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)

    Dim sPassageText As String
    sPassageText = oDoc.Range(lStart, lQuoteEnd).text

    Dim sPassageStripped As String
    Dim iChar As Long
    Dim sTestChar As String
    sPassageStripped = ""
    For iChar = 1 To Len(sPassageText)
        sTestChar = Mid(sPassageText, iChar, 1)
        Select Case AscW(sTestChar)
            Case 32, 9, 11, 12, 13, 160
            Case Else
                sPassageStripped = sPassageStripped & sTestChar
        End Select
    Next iChar

    If Len(sPassageStripped) <= 1 Then
        bCitationOnly = True
        Dim oPassageDel As Range
        Set oPassageDel = oDoc.Range(lStart, lQuoteEnd)
        oPassageDel.Delete
        Set oPassageDel = Nothing
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
        lQuoteEnd = FindQuoteEnd(oRange)
    End If

    ' --- Quote-related steps: skipped for citation-only pastes ---

    g_Phase = "Steps 13-18: quote conversion, wrapping, capitalization, footnotes"

    If Not bCitationOnly Then

        ' Step 12c: Replace internal citations inside the passage with
        ' [citation], [citations], [Citation.], or [Citations.] depending
        ' on count and sentence position.  Cases only; statutes excluded.
        If Not bIsStatute Then
            ' Use FindQuoteEnd (paren-depth scan) not FindQuoteEndByParagraph,
            ' because RemoveBlankParagraphs has already merged the block onto
            ' one paragraph by this point -- the paragraph-boundary method
            ' falls back to the full range end and includes the final citation.
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
            Dim lQEForInternal As Long
            lQEForInternal = FindQuoteEnd(oRange)
            Set oRange = oDoc.Range(lStart, lQEForInternal)
            RemoveInternalCitations oDoc, oRange
            ' Collapse any double spaces left by citation removal
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
            CollapseMultipleSpaces oRange
            lSelEnd = oDoc.content.End - lTailLen
            lQuoteEnd = FindQuoteEnd(oDoc.Range(lStart, lSelEnd))
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
        End If

        ' Step 13: Insert pre-paste openers at passage start before swap.
        ' Nesting signals are accepted for cases; ignored for statutes
        ' but deleting them via DetectPrePasteOpeners is harmless.
        If Len(sPreOpeners) > 0 And Not bIsStatute Then
            Dim oPreInsert As Range
            Set oPreInsert = oDoc.Range(lStart, lStart)
            oPreInsert.InsertBefore sPreOpeners
            Set oPreInsert = Nothing
            lQuoteEnd = lQuoteEnd + Len(sPreOpeners)
        End If

        ' Steps 14-16: Quote conversion   diverges by paste type
        If bIsStatute Then

            ' Statute step A: Convert double quotes in passage to curly
            ' singles so they survive being wrapped in outer doubles.
            Dim oPassageS As Range
            Set oPassageS = oDoc.Range(lStart, lQuoteEnd)
            ConvertDoubleQuotesToSingles oPassageS

            ' Statute step B: Convert all straight singles to curly
            ' apostrophes unconditionally.
            Set oPassageS = oDoc.Range(lStart, lQuoteEnd)
            CurlyApostrophesStatute oPassageS
            Set oPassageS = Nothing

        Else

            ' Case step 14: Convert straight quotes to curly before swapping
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
            Call CurlyApostrophes(oRange)

            ' Case step 15: Swap internal quotes in passage only
            Dim oPassage As Range
            Set oPassage = oDoc.Range(lStart, lQuoteEnd)
            SwapSmartQuotes oPassage

            ' Case step 16: Balance quotes using same passage boundary
            Set oPassage = oDoc.Range(lStart, lQuoteEnd)
            BalanceNestedQuotes oPassage
            Set oPassage = Nothing

            ' Step 16b: Remove duplicate single quotes from passage edges.
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
            RemoveDuplicateOpenSingle oDoc, oRange
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
            RemoveDuplicateCloseSingle oDoc, oRange

        End If

        ' Step 17: Recompute passage end after any balance insertions
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
        lQuoteEnd = FindQuoteEnd(oRange)
        Dim oPassageWrap As Range
        Set oPassageWrap = oDoc.Range(lStart, lQuoteEnd)

        ' Step 18: Wrap passage in outer double smart quotes
        WrapInDoubleQuotes oPassageWrap

        ' Step 18b: Fix first letter capitalisation
        If Not bProperNoun Then
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
            lQuoteEnd = FindQuoteEnd(oRange)
            FixFirstLetterCapitalization oDoc, lStart, lQuoteEnd
        End If

        Set oPassageWrap = Nothing

        ' Step 18c: Insert footnote omission tag inside the citation paren.
        ' Pass 1 (step 11d) deleted inline markers and counted them.
        ' Pass 2 inserts ", fn. omitted." before the citation closing ")",
        ' removing the preceding period so it migrates to after "omitted".
        If nFootnoteCount > 0 Then
            lSelEnd = oDoc.content.End - lTailLen
            Set oRange = oDoc.Range(lStart, lSelEnd)
            RemoveFootnotesPass2 oDoc, oRange, nFootnoteCount
        End If

        ' Step 18d: Remove a duplicate closing double-quote if WrapInDoubleQuotes
        ' produced two consecutive U+201D at the passage end.
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
        RemoveDuplicateClosingQuote oDoc, oRange

        ' Step 18e: Remove a duplicate opening double-quote.
        lSelEnd = oDoc.content.End - lTailLen
        Set oRange = oDoc.Range(lStart, lSelEnd)
        RemoveDuplicateOpeningQuote oDoc, oRange

    End If

    g_Phase = "Steps 19-22: quote spacing, restructure sentence, italicize, format"

    ' Step 19: Ensure exactly one space between adjacent opposite-direction
    ' quotes and between quotes and adjacent non-exempt characters.
    lSelEnd = oDoc.content.End - lTailLen
    Set oRange = oDoc.Range(lStart, lSelEnd)
    EnsureSpacingAroundQuotes oDoc, oRange

    ' Step 19b: Restructure as textual sentence if "in"/"In" was detected.
    ' Input:  "Passage." (Citation.)  or just (Citation.) for citation-only
    ' Output: Citation, the court held "Passage."
    '         Citation                              <- citation-only
    ' Cases only; lEnd updated by reference.
    Dim lEnd As Long
    lSelEnd = oDoc.content.End - lTailLen
    lEnd = lSelEnd
    If bStatutoryTextual Then
        RestructureAsStatutoryTextual oDoc, lStart, lEnd
    ElseIf bTextualSentence And Not bIsStatute Then
        RestructureAsTextual oDoc, lStart, lEnd, bCitationOnly
    ElseIf bCitesSentence And Not bIsStatute Then
        ' Step 19b-ii: Restructure as cites-sentence.
        ' Input:  "Passage." (Citation.)  or just (Citation.) for citation-only
        ' Output: Citation for the proposition that "Passage."
        '         Citation                                      <- citation-only
        RestructureAsCites oDoc, lStart, lEnd, bCitationOnly
    ElseIf bParenthetical And Not bCitationOnly Then
        ' Step 19c: Restructure as parenthetical (skipped if textual-sentence).
        ' Input:  "Passage."  (Citation.)
        ' Output: (Citation ["Passage"].)      <- cases: period stripped
        '         (Citation ["Passage."].)     <- statutes: period kept
        RestructureAsParenthetical oDoc, lStart, lEnd, bIsStatute
    End If

    ' Step 20: Italicise case names   cases only.
    ' Uses lEnd (not lSelEnd) so it covers the restructured block.
    If Not bIsStatute And Not bStatutoryTextual Then
        Set oRange = oDoc.Range(lStart, lEnd)
        ItalicizeCaseNames oRange, oDoc
    End If

    ' Step 21: Format   Times New Roman 12pt, preserve italic, no bold.
    ' Uses lEnd so every character in the restructured block is covered.
    Set oRange = oDoc.Range(lStart, lEnd)
    FormatQuotation oRange

    ' Step 22: Fix font at the join points immediately before and after
    ' the pasted block. Word leaves the surrounding characters in Aptos
    ' Body when the paste lands next to existing text. Walk outward from
    ' lStart and lEnd through any run of non-Times-New-Roman characters
    ' within the same paragraph and reset them to Times New Roman 12pt.
    FixSurroundingFont oDoc, lStart, lEnd

CleanUp:
    ' FIX 1 (continued): capture the error first, then close the custom undo
    ' record if we opened one (small-document case), then report the error and
    ' the phase it occurred in.
    Dim lErrNum As Long
    Dim sErrDesc As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    On Error Resume Next
    oDoc.TrackRevisions = bPrevTrack
    On Error GoTo 0
    If bUseUndo Then
        On Error Resume Next
        oUndo.EndCustomRecord
        On Error GoTo 0
    End If
    If lErrNum <> 0 Then
        MsgBox "Error " & lErrNum & ": " & sErrDesc & vbCrLf & vbCrLf & _
               "During: " & g_Phase, vbExclamation, "PasteLegalQuotation"
    End If

    Set oRange = Nothing
    Set oUndo = Nothing

End Sub

'===========================================================
' Harvest Lexis+ hyperlinks from the pasted range:
'   1. Ensure the citation server is running (once per session).
'   2. For each hyperlink, extract URL + parse citation key.
'   3. POST to /cite_repo (silent on failure).
'   4. Remove the hyperlink field, preserving display text.
' Safe to call even if no hyperlinks are present.
'===========================================================
Private Sub HarvestAndRemoveHyperlinks(oDoc As Document, oRange As Range)

    ' Ensure server is reachable (at most one startup attempt per session)
    EnsureCiteServer

    ' Collect hyperlinks into an array first: iterating the live
    ' Hyperlinks collection while deleting members is unsafe.
    Dim nLinks As Integer
    nLinks = oRange.Hyperlinks.count
    If nLinks = 0 Then Exit Sub

    Const MAX_LINKS As Integer = 20
    Dim aURL(MAX_LINKS)  As String
    Dim aText(MAX_LINKS) As String
    Dim i As Integer
    For i = 1 To nLinks
        If i > MAX_LINKS Then Exit For
        Dim oHL As Hyperlink
        Set oHL = oRange.Hyperlinks(i)
        aURL(i) = oHL.Address
        aText(i) = oHL.Range.text
        Set oHL = Nothing
    Next i

    ' Remove all hyperlink fields (preserves text, just strips the field)
    ' Iterate in reverse so index shifts don't matter.
    Dim j As Integer
    For j = nLinks To 1 Step -1
        On Error Resume Next
        oRange.Hyperlinks(j).Delete
        On Error GoTo 0
    Next j

    ' POST each unique URL to the citation repo
    Dim k As Integer
    For k = 1 To nLinks
        If k > MAX_LINKS Then Exit For
        Dim sURL As String
        sURL = Trim(aURL(k))
        If Len(sURL) = 0 Then GoTo NextLink
        If InStr(1, sURL, "lexis", vbTextCompare) = 0 Then GoTo NextLink

        ' Parse citation key and metadata from the hyperlink display text
        Dim sKey      As String
        Dim sCaseName As String
        Dim sYear     As String
        Dim sVolume   As String
        Dim sReporter As String
        Dim sFirstPage As String
        ParseLexisCitationText aText(k), sKey, sCaseName, sYear, _
                               sVolume, sReporter, sFirstPage

        If Len(sKey) = 0 Then GoTo NextLink

        ' Build JSON payload
        Dim sJSON As String
        sJSON = "{" & _
            """type"":""case""," & _
            """key"":""" & JsonEscape(sKey) & """," & _
            """url"":""" & JsonEscape(sURL) & """," & _
            """source"":""lexis""," & _
            """added_by"":""macro""," & _
            """case_name"":""" & JsonEscape(sCaseName) & """," & _
            """year"":""" & JsonEscape(sYear) & """," & _
            """volume"":""" & JsonEscape(sVolume) & """," & _
            """reporter"":""" & JsonEscape(sReporter) & """," & _
            """first_page"":""" & JsonEscape(sFirstPage) & """" & _
            "}"

        PostToCiteRepo sJSON

NextLink:
    Next k

End Sub

'===========================================================
' Ensure the citation server is reachable.  Called once per
' Word session; subsequent calls return immediately.
' If the server does not respond to /ping, launches app.py
' via Shell, waits 2 seconds, then marks verified regardless
' so the delay never repeats.
'===========================================================
Private Sub EnsureCiteServer()

    If g_ServerVerified Then Exit Sub

    ' Try a quick ping
    If PingCiteServer() Then
        g_ServerVerified = True
        Exit Sub
    End If

    ' Server not running -- attempt to launch it. Build the quoted command with
    ' Chr(34): the previous doubled-quote literal was malformed (an odd quote
    ' run left the string open), so CITE_APP_PATH was passed as literal text
    ' and the server never actually launched.
    Dim sLaunchCmd As String
    sLaunchCmd = Chr(34) & CITE_PYTHON_EXE & Chr(34) & " " & _
                 Chr(34) & CITE_APP_PATH & Chr(34)
    On Error Resume Next
    Shell sLaunchCmd, vbHide
    On Error GoTo 0

    ' Wait 2 seconds for it to start
    Sleep 2000

    ' Mark verified regardless of outcome so delay never repeats
    g_ServerVerified = True

End Sub

'===========================================================
' Ping /ping on the citation server.
' Returns True if the server responded with HTTP 200.
' Uses WinHttp.WinHttpRequest (available on all Windows installs).
'===========================================================
Private Function PingCiteServer() As Boolean

    PingCiteServer = False
    On Error Resume Next

    Dim oHTTP As Object
    Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    oHTTP.SetTimeouts 800, 800, 800, 800  ' ms: resolve, connect, send, receive
    oHTTP.Open "GET", CITE_SERVER_URL & "/ping", False
    oHTTP.send

    If Err.Number = 0 And oHTTP.Status = 200 Then
        PingCiteServer = True
    End If

    Set oHTTP = Nothing
    On Error GoTo 0

End Function

'===========================================================
' POST a JSON string to /cite_repo.  Silent on any error.
'===========================================================
Private Sub PostToCiteRepo(sJSON As String)

    On Error Resume Next

    Dim oHTTP As Object
    Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    oHTTP.SetTimeouts 1000, 1000, 1000, 1000
    oHTTP.Open "POST", CITE_SERVER_URL & "/cite_repo", False
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.send sJSON

    Set oHTTP = Nothing
    On Error GoTo 0

End Sub

'===========================================================
' Parse a Lexis+ citation display string into its components.
' Input example:
'   (G.R. v. Intelligator (2010) 185 Cal.App.4th 606
'    [110 Cal.Rptr.3d 559].)
' Outputs:
'   sKey       = "185 Cal.App.4th 606"
'   sCaseName  = "G.R. v. Intelligator"
'   sYear      = "2010"
'   sVolume    = "185"
'   sReporter  = "Cal.App.4th"
'   sFirstPage = "606"
' Sets sKey = "" if no reporter citation is found.
'===========================================================
Private Sub ParseLexisCitationText(sText As String, _
                                    ByRef sKey As String, _
                                    ByRef sCaseName As String, _
                                    ByRef sYear As String, _
                                    ByRef sVolume As String, _
                                    ByRef sReporter As String, _
                                    ByRef sFirstPage As String)

    sKey = "": sCaseName = "": sYear = ""
    sVolume = "": sReporter = "": sFirstPage = ""

    ' Strip outer parens, brackets, and punctuation from the display text.
    ' Lexis+ wraps the whole sentence: (Case (Year) Vol Rep Page [parallel].)
    Dim s As String
    s = Trim(sText)
    ' Remove paragraph marks and normalize spaces
    Dim i2 As Long
    Dim sClean As String
    sClean = ""
    For i2 = 1 To Len(s)
        Dim nCC As Long
        nCC = AscW(Mid(s, i2, 1))
        If nCC = 13 Or nCC = 11 Or nCC = 10 Then
            sClean = sClean & " "
        ElseIf nCC = 160 Then
            sClean = sClean & " "
        Else
            sClean = sClean & Mid(s, i2, 1)
        End If
    Next i2
    ' Collapse multiple spaces
    Do While InStr(sClean, "  ") > 0
        sClean = Join(Split(sClean, "  "), " ")
    Loop
    s = Trim(sClean)

    ' --- Find (YYYY) year pattern ---
    ' Everything before it is the case name (stripped of leading "(").
    Dim iYear As Long
    iYear = -1
    Dim i As Long
    For i = 1 To Len(s) - 5
        If Mid(s, i, 1) = "(" Then
            Dim sFour As String
            sFour = Mid(s, i + 1, 4)
            If Mid(s, i + 5, 1) = ")" And IsAllDigits(sFour) Then
                Dim nYr As Long
                nYr = CLng(sFour)
                If nYr >= 1800 And nYr <= 2099 Then
                    iYear = i
                    sYear = sFour
                    Exit For
                End If
            End If
        End If
    Next i
    If iYear = -1 Then Exit Sub

    ' Case name: text before the year paren, strip leading "(" and trim
    Dim sNameRaw As String
    sNameRaw = Trim(Left(s, iYear - 1))
    If Left(sNameRaw, 1) = "(" Then sNameRaw = Trim(Mid(sNameRaw, 2))
    sCaseName = sNameRaw

    ' --- Find volume + reporter + first page after (YYYY) ---
    ' Scan forward from end of year paren: skip spaces, read digits (volume),
    ' skip space, match reporter, skip space, read digits (first page).
    Dim iAfterYear As Long
    iAfterYear = iYear + 6  ' position after "(YYYY)"

    ' Skip whitespace
    Do While iAfterYear <= Len(s) And Mid(s, iAfterYear, 1) = " "
        iAfterYear = iAfterYear + 1
    Loop

    ' Read volume digits
    Dim iVolStart As Long
    iVolStart = iAfterYear
    Do While iAfterYear <= Len(s)
        Dim nVC As Long
        nVC = AscW(Mid(s, iAfterYear, 1))
        If nVC >= 48 And nVC <= 57 Then
            iAfterYear = iAfterYear + 1
        Else
            Exit Do
        End If
    Loop
    If iAfterYear <= iVolStart Then Exit Sub  ' no digits found
    sVolume = Mid(s, iVolStart, iAfterYear - iVolStart)

    ' Skip one space between volume and reporter
    If Mid(s, iAfterYear, 1) = " " Then iAfterYear = iAfterYear + 1

    ' Match reporter from canonical list (longest match wins)
    Dim aRep(30) As String
    aRep(0) = "Cal.5th"
    aRep(1) = "Cal.4th"
    aRep(2) = "Cal.3d"
    aRep(3) = "Cal.2d"
    aRep(4) = "Cal.App.5th"
    aRep(5) = "Cal.App.4th"
    aRep(6) = "Cal.App.3d"
    aRep(7) = "Cal.App.2d"
    aRep(8) = "Cal.App.Supp."
    aRep(9) = "Cal.App."
    aRep(10) = "Cal.Rptr.3d"
    aRep(11) = "Cal.Rptr.2d"
    aRep(12) = "Cal.Rptr."
    aRep(13) = "Cal."
    aRep(14) = "U.S."
    aRep(15) = "S.Ct."
    aRep(16) = "L.Ed.2d"
    aRep(17) = "L.Ed."
    aRep(18) = "F.4th"
    aRep(19) = "F.3d"
    aRep(20) = "F.2d"
    aRep(21) = "F.Supp.3d"
    aRep(22) = "F.Supp.2d"
    aRep(23) = "F.Supp."
    aRep(24) = "P.3d"
    aRep(25) = "P.2d"
    aRep(26) = "P."
    aRep(27) = "B.R."
    aRep(28) = "Fed.Cl."
    aRep(29) = "N.Y.S.2d"
    aRep(30) = "N.Y.S."

    Dim r As Integer
    Dim sRepMatch As String
    sRepMatch = ""
    For r = 0 To 30
        Dim nRL As Long
        nRL = Len(aRep(r))
        If iAfterYear + nRL - 1 <= Len(s) Then
            If Mid(s, iAfterYear, nRL) = aRep(r) Then
                sRepMatch = aRep(r)
                iAfterYear = iAfterYear + nRL
                Exit For
            End If
        End If
    Next r
    If Len(sRepMatch) = 0 Then Exit Sub  ' no reporter matched
    sReporter = sRepMatch

    ' Skip one space between reporter and first page
    If iAfterYear <= Len(s) And Mid(s, iAfterYear, 1) = " " Then
        iAfterYear = iAfterYear + 1
    End If

    ' Read first-page digits
    Dim iPageStart As Long
    iPageStart = iAfterYear
    Do While iAfterYear <= Len(s)
        Dim nPC As Long
        nPC = AscW(Mid(s, iAfterYear, 1))
        If nPC >= 48 And nPC <= 57 Then
            iAfterYear = iAfterYear + 1
        Else
            Exit Do
        End If
    Loop
    If iAfterYear <= iPageStart Then Exit Sub  ' no page digits
    sFirstPage = Mid(s, iPageStart, iAfterYear - iPageStart)

    ' Build canonical key: "volume reporter firstpage"
    sKey = sVolume & " " & sReporter & " " & sFirstPage

End Sub

'===========================================================
' Escape a string for embedding in a JSON string literal.
' Handles backslash, double-quote, and common control chars.
' Note: VBA string literals have no escape sequences, so "\"
' is one backslash and "\\" is two -- each replacement below
' yields a two-character escape (backslash + letter/quote).
'===========================================================
Private Function JsonEscape(s As String) As String
    Dim r As String
    r = s
    r = Join(Split(r, "\"), "\\")
    r = Join(Split(r, Chr(34)), "\" & Chr(34))
    r = Join(Split(r, Chr(13)), "\r")
    r = Join(Split(r, Chr(10)), "\n")
    r = Join(Split(r, Chr(9)), "\t")
    JsonEscape = r
End Function

'===========================================================
' Return True if every character in s is an ASCII digit.
'===========================================================
Private Function IsAllDigits(s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        Dim n As Long
        n = AscW(Mid(s, i, 1))
        If n < 48 Or n > 57 Then
            IsAllDigits = False
            Exit Function
        End If
    Next i
    IsAllDigits = True
End Function

'===========================================================
' Remove Westlaw flag images (inline shapes) from pasted range.
'===========================================================
Private Sub RemoveInlineShapes(oRange As Range)

    Dim i As Integer
    For i = oRange.InlineShapes.count To 1 Step -1
        oRange.InlineShapes(i).Delete
    Next i

End Sub

'===========================================================
' Remove Lexis+ headnotes (bold numbered parentheticals)
' FIX 2: Added iteration cap to prevent infinite loop.
'===========================================================
Private Sub RemoveLexisHeadnotes(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Const MAX_ITER As Long = 500

    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim bFound As Boolean
        bFound = False

        Dim oFind As Range
        Set oFind = oRange.Duplicate

        With oFind.Find
            .ClearFormatting
            .Font.Bold = True
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = True
            .text = "\([0-9]@\)"
            .Execute
        End With

        If oFind.Find.found Then
            Dim oHeadnote As Range
            Set oHeadnote = oFind.Duplicate

            Do
                If oHeadnote.End >= oRange.End Then Exit Do
                Dim oNext As Range
                Set oNext = oDoc.Range(oHeadnote.End, oHeadnote.End + 1)
                Dim nCode As Long
                nCode = AscW(oNext.text)
                If nCode = 13 Or nCode = 11 Then Exit Do
                If oNext.Font.Bold <> True Then Exit Do
                oHeadnote.MoveEnd wdCharacter, 1
                Set oNext = Nothing
            Loop

            If oHeadnote.start > oRange.start Then
                Dim oPre As Range
                Set oPre = oDoc.Range(oHeadnote.start - 1, oHeadnote.start)
                If oPre.text = " " Then
                    oHeadnote.MoveStart wdCharacter, -1
                End If
                Set oPre = Nothing
            End If

            oHeadnote.Delete
            bFound = True
            Set oHeadnote = Nothing
        End If

        Set oFind = Nothing
        If Not bFound Then Exit Do

        Set oRange = oDoc.Range(oRange.start, oRange.End)
    Loop

End Sub

'===========================================================
' Remove Lexis+ page markers [*7] [**12] [***19]
'===========================================================
Private Sub RemovePageReferences(oRange As Range)

    With oRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .text = "\[\*@[0-9]@\] "
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
    End With

    With oRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .text = "\[\*@[0-9]@\]"
        .Replacement.text = ""
        .Execute Replace:=wdReplaceAll
    End With

End Sub

'===========================================================
' Remove Westlaw page numbers: *123 **123 ***123
' FIX 2: Added iteration cap to prevent infinite loop.
'===========================================================
Private Sub RemoveWestlawPageNumbers(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Const MAX_ITER As Long = 500

    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim sText As String
        sText = oRange.text

        Dim bFound As Boolean
        bFound = False

        Dim i As Long
        For i = 1 To Len(sText)
            If Mid(sText, i, 1) = "*" Then
                Dim j As Long
                j = i
                Do While j <= Len(sText) And Mid(sText, j, 1) = "*"
                    j = j + 1
                Loop
                If j <= Len(sText) Then
                    Dim nChar As Long
                    nChar = AscW(Mid(sText, j, 1))
                    If nChar >= 48 And nChar <= 57 Then
                        Dim k As Long
                        k = j
                        Do While k <= Len(sText)
                            nChar = AscW(Mid(sText, k, 1))
                            If nChar >= 48 And nChar <= 57 Then
                                k = k + 1
                            Else
                                Exit Do
                            End If
                        Loop
                        Dim lDelStart As Long
                        Dim lDelEnd As Long
                        lDelStart = oRange.start + i - 1
                        lDelEnd = oRange.start + k - 1
                        If i > 1 And Mid(sText, i - 1, 1) = " " Then
                            lDelStart = lDelStart - 1
                        End If
                        Dim oDel As Range
                        Set oDel = oDoc.Range(lDelStart, lDelEnd)
                        oDel.Delete
                        Set oDel = Nothing
                        bFound = True
                        Exit For
                    End If
                End If
            End If
        Next i

        If Not bFound Then Exit Do
        Set oRange = oDoc.Range(oRange.start, oRange.End)
    Loop

End Sub

'===========================================================
' Remove Lexis+ parallel citations in brackets
' FIX 2: Added iteration cap to prevent infinite loop.
'===========================================================
Private Sub RemoveParallelCitationsManual(oRange As Range, ByVal lRangeStart As Long)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Const MAX_ITER As Long = 500

    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim sText As String
        sText = oRange.text

        Dim i As Long
        Dim bFound As Boolean
        bFound = False

        For i = 1 To Len(sText)
            If Mid(sText, i, 1) = "[" Then
                Dim j As Long
                Dim nDepth As Integer
                nDepth = 1
                j = i + 1
                Do While j <= Len(sText) And nDepth > 0
                    If Mid(sText, j, 1) = "[" Then nDepth = nDepth + 1
                    If Mid(sText, j, 1) = "]" Then nDepth = nDepth - 1
                    j = j + 1
                Loop

                ' Unmatched "[" (a selection cut mid-bracket): the scan ran to
                ' the end of the block, so sBracket would span from the stray
                ' bracket THROUGH THE CITATION, and any reporter string in
                ' that tail deleted the whole span. Skip it instead.
                If nDepth <> 0 Then GoTo NextBracketScan

                Dim sBracket As String
                sBracket = Mid(sText, i, j - i)

                If IsParallelCitation(sBracket) Then
                    Dim lDelStart As Long
                    Dim lDelEnd As Long
                    lDelStart = lRangeStart + i - 1
                    lDelEnd = lRangeStart + j - 1

                    If i > 1 And Mid(sText, i - 1, 1) = " " Then
                        lDelStart = lDelStart - 1
                    End If

                    Dim oDel As Range
                    Set oDel = oDoc.Range(lDelStart, lDelEnd)
                    oDel.Delete
                    Set oDel = Nothing

                    Set oRange = oDoc.Range(lRangeStart, oRange.End)
                    bFound = True
                    Exit For
                End If
            End If
NextBracketScan:
        Next i

        If Not bFound Then Exit Do
    Loop

End Sub

'===========================================================
' Helper: does bracket contain an unofficial reporter?
'===========================================================
Private Function IsParallelCitation(s As String) As Boolean
    ' Parallel citations always start with a volume number, so the
    ' bracketed span only qualifies if its content (after the "[" and
    ' any leading whitespace) begins with a digit. This keeps editorial
    ' inserts like "[J.A. 123]" or "[A. Smith]" from tripping the bare
    ' "P."/"A."/"So." reporter tests below.
    IsParallelCitation = False
    Dim sContent As String
    sContent = s
    If Left(sContent, 1) = "[" Then sContent = Mid(sContent, 2)
    sContent = LTrim(sContent)
    If Len(sContent) = 0 Then Exit Function
    Dim nFirst As Long
    nFirst = AscW(Left(sContent, 1))
    If nFirst < 48 Or nFirst > 57 Then Exit Function

    If InStr(s, "L.Ed.2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "L.Ed.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "S.Ct.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "Cal. Rptr.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "Cal.Rptr.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "P.4th") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "P.3d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "P.2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "P.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "A.3d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "A.2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "A.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "N.E. 2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "N.E.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "N.W. 2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "N.W.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "N.Y.S. 2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "N.Y.S.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "S.E. 2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "S.E.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "So. 2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "So.") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "S.W. 3d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "S.W. 2d") > 0 Then IsParallelCitation = True: Exit Function
    If InStr(s, "S.W.") > 0 Then IsParallelCitation = True: Exit Function
    IsParallelCitation = False
End Function

'===========================================================
' Remove Westlaw parallel citations after official cite.
' FIX 2: Added iteration cap per reporter to prevent infinite loop.
'===========================================================
Private Sub RemoveWestlawParallelCitations(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim aReporters(25) As String
    aReporters(0) = "Cal. Rptr. 3d"
    aReporters(1) = "Cal. Rptr. 2d"
    aReporters(2) = "Cal. Rptr."
    aReporters(3) = "Cal.Rptr.3d"
    aReporters(4) = "Cal.Rptr.2d"
    aReporters(5) = "Cal.Rptr."
    aReporters(6) = "L.Ed.2d"
    aReporters(7) = "L.Ed."
    aReporters(8) = "P.4th"
    aReporters(9) = "P.3d"
    aReporters(10) = "P.2d"
    aReporters(11) = "S.Ct."
    aReporters(12) = "N.Y.S. 2d"
    aReporters(13) = "N.Y.S."
    aReporters(14) = "S.W. 3d"
    aReporters(15) = "S.W. 2d"
    aReporters(16) = "S.W."
    aReporters(17) = "S.E. 2d"
    aReporters(18) = "S.E."
    aReporters(19) = "So. 2d"
    aReporters(20) = "N.W. 2d"
    aReporters(21) = "N.W."
    aReporters(22) = "N.E. 2d"
    aReporters(23) = "N.E."
    aReporters(24) = "A.3d"
    aReporters(25) = "A.2d"

    Const MAX_ITER As Long = 500

    Dim r As Integer
    For r = 0 To 25

        Dim nIter As Long
        nIter = 0

        ' Progress cursor: where the next search starts. Advances past
        ' each non-qualifying match so the same hit is never re-found.
        Dim lSearchFrom As Long
        lSearchFrom = oRange.start

        Do
            nIter = nIter + 1
            If nIter > MAX_ITER Then Exit Do

            Dim oFind As Range
            Set oFind = oDoc.Range(lSearchFrom, oRange.End)

            With oFind.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
                .text = aReporters(r)
            End With

            oFind.Find.Execute

            If Not oFind.Find.found Then
                Set oFind = Nothing
                Exit Do
            End If

            ' --- Validate right side: reporter must be followed by digits ---
            Dim lAfterReporter As Long
            lAfterReporter = oFind.End

            Dim oRightCheck As Range
            Set oRightCheck = oDoc.Range(lAfterReporter, lAfterReporter + 1)
            If oRightCheck.text = " " Then
                lAfterReporter = lAfterReporter + 1
                Set oRightCheck = oDoc.Range(lAfterReporter, lAfterReporter + 1)
            End If

            Dim nRightFirst As Long
            nRightFirst = AscW(oRightCheck.text)
            Set oRightCheck = Nothing

            If nRightFirst < 48 Or nRightFirst > 57 Then
                ' Not a parallel cite here -- skip past this match and
                ' keep searching for later occurrences of the reporter.
                lSearchFrom = oFind.End
                Set oFind = Nothing
                GoTo NextReporterMatch
            End If

            ' --- Validate left side: comma then digits-only volume number ---
            Dim lDelStart As Long
            Dim lWalk As Long
            Dim bCommaFound As Boolean
            bCommaFound = False
            lDelStart = oFind.start
            lWalk = oFind.start - 1

            Do While lWalk >= oRange.start
                Dim oWalk As Range
                Set oWalk = oDoc.Range(lWalk, lWalk + 1)
                Dim cWalk As String
                cWalk = oWalk.text
                Set oWalk = Nothing

                If cWalk = "," Then
                    Dim sBetween As String
                    sBetween = Trim(oDoc.Range(lWalk + 1, oFind.start).text)
                    If IsCiteVolumeNumber(sBetween) Then
                        lDelStart = lWalk
                        bCommaFound = True
                    End If
                    Exit Do
                ElseIf cWalk = Chr(13) Or cWalk = Chr(11) Or _
                       cWalk = "." Or cWalk = ")" Or cWalk = "(" Then
                    Exit Do
                End If
                lWalk = lWalk - 1
            Loop

            If Not bCommaFound Then
                ' No qualifying volume/comma on the left -- skip past
                ' this match rather than re-finding it forever.
                lSearchFrom = oFind.End
                Set oFind = Nothing
            Else
                ' --- Walk right consuming page number digits only ---
                Dim lDelEnd As Long
                Dim lWalkR As Long
                lDelEnd = oFind.End
                lWalkR = oFind.End

                Dim oPageSpace As Range
                Set oPageSpace = oDoc.Range(lWalkR, lWalkR + 1)
                If AscW(oPageSpace.text) = 32 Then
                    lWalkR = lWalkR + 1
                End If
                Set oPageSpace = Nothing

                Do While lWalkR <= oRange.End
                    Dim oWalkR As Range
                    Set oWalkR = oDoc.Range(lWalkR, lWalkR + 1)
                    Dim nR As Long
                    nR = AscW(oWalkR.text)
                    Set oWalkR = Nothing
                    If nR >= 48 And nR <= 57 Then
                        lDelEnd = lWalkR + 1
                        lWalkR = lWalkR + 1
                    Else
                        Exit Do
                    End If
                Loop

                Dim oDel As Range
                Set oDel = oDoc.Range(lDelStart, lDelEnd)
                oDel.Delete
                Set oDel = Nothing
                Set oFind = Nothing
                Set oRange = oDoc.Range(oRange.start, oRange.End)
                ' Text shrank -- resume the search from the deletion point.
                lSearchFrom = lDelStart
            End If

NextReporterMatch:
        Loop
    Next r

End Sub

'===========================================================
' Replace non-breaking spaces Chr(160) with regular spaces
' except when immediately following
'===========================================================
Private Sub ReplaceNonBreakingSpaces(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim sText As String
    sText = oRange.text

    Dim i As Long
    For i = Len(sText) To 1 Step -1
        If AscW(Mid(sText, i, 1)) = 160 Then
            Dim sPrev As String
            If i > 1 Then sPrev = Mid(sText, i - 1, 1) Else sPrev = ""
            If sPrev <> ChrW(&HA7) Then
                Dim lPos As Long
                lPos = oRange.start + i - 1
                Dim oChar As Range
                Set oChar = oDoc.Range(lPos, lPos + 1)
                oChar.text = " "
                Set oChar = Nothing
            End If
        End If
    Next i

End Sub

'===========================================================
' Convert soft returns (Chr 11) to hard paragraph returns
'===========================================================
Private Sub ConvertSoftReturns(oRange As Range)

    With oRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .text = Chr(11)
        .Replacement.text = Chr(13)
        .Execute Replace:=wdReplaceAll
    End With

End Sub

'===========================================================
' Collapse multiple consecutive spaces to one
' FIX 2: Added iteration cap to prevent infinite loop.
'===========================================================
Private Sub CollapseMultipleSpaces(oRange As Range)

    Const MAX_ITER As Long = 500

    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim bFound As Boolean
        With oRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            .text = "  "
            .Replacement.text = " "
            bFound = .Execute(Replace:=wdReplaceAll)
        End With
        If Not bFound Then Exit Do
    Loop

End Sub

'===========================================================
' Remove spaces between adjacent curly quote marks (Westlaw artifact)
' FIX 2: Added iteration cap to prevent infinite loop.
'===========================================================
Private Sub RemoveSpacesBetweenQuotes(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Const MAX_ITER As Long = 500

    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim sText As String
        sText = oRange.text

        Dim bFound As Boolean
        bFound = False

        Dim i As Long
        For i = 2 To Len(sText) - 1
            If AscW(Mid(sText, i, 1)) = 32 Then
                Dim nPrev As Long
                Dim nNext As Long
                nPrev = AscW(Mid(sText, i - 1, 1))
                nNext = AscW(Mid(sText, i + 1, 1))
                Dim bPrevClose As Boolean, bPrevOpen As Boolean
                Dim bNextClose As Boolean, bNextOpen As Boolean
                bPrevClose = (nPrev = &H201D Or nPrev = &H2019)
                bPrevOpen = (nPrev = &H201C Or nPrev = &H2018)
                bNextClose = (nNext = &H201D Or nNext = &H2019)
                bNextOpen = (nNext = &H201C Or nNext = &H2018)
                If (bPrevClose And bNextClose) Or (bPrevOpen And bNextOpen) Then
                    Dim lDelPos As Long
                    lDelPos = oRange.start + i - 1
                    Dim oDel As Range
                    Set oDel = oDoc.Range(lDelPos, lDelPos + 1)
                    oDel.Delete
                    Set oDel = Nothing
                    bFound = True
                    Exit For
                End If
            End If
        Next i

        If Not bFound Then Exit Do
        Set oRange = oDoc.Range(oRange.start, oRange.End)
    Loop

End Sub

'===========================================================
' Remove spaces immediately following an opening parenthesis.
' e.g. "( the court" -> "(the court"
'===========================================================
Private Sub RemoveSpacesAfterOpenParen(oRange As Range)

    Const MAX_ITER As Long = 500
    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim bFound As Boolean
        With oRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            .text = "( "
            .Replacement.text = "("
            bFound = .Execute(Replace:=wdReplaceAll)
        End With
        If Not bFound Then Exit Do
    Loop

End Sub
Private Sub RemoveBlankParagraphs(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim sText As String
    sText = oRange.text

    Dim sTrimmed As String
    sTrimmed = RTrim(sText)

    Dim i As Long
    Dim nDepth As Integer
    nDepth = 0
    Dim nCitationStart As Long
    nCitationStart = -1

    For i = Len(sTrimmed) To 1 Step -1
        Dim c As String
        c = Mid(sTrimmed, i, 1)
        If c = ")" Then nDepth = nDepth + 1
        If c = "(" Then
            nDepth = nDepth - 1
            If nDepth = 0 Then
                nCitationStart = i
                Exit For
            End If
        End If
    Next i

    If nCitationStart <= 1 Then Exit Sub

    Dim lCitationPos As Long
    lCitationPos = oRange.start + nCitationStart - 1

    ' FIX 2: Added iteration cap to prevent infinite loop.
    Const MAX_ITER As Long = 500
    Dim nIter As Long
    nIter = 0

    Dim oCheck As Range
    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        If lCitationPos <= oRange.start Then Exit Do
        Set oCheck = oDoc.Range(lCitationPos - 1, lCitationPos)
        Dim nCode As Long
        nCode = AscW(oCheck.text)
        If nCode = 13 Or nCode = 11 Or nCode = 32 Then
            oCheck.Delete
            lCitationPos = lCitationPos - 1
        Else
            Exit Do
        End If
    Loop

    Set oCheck = Nothing

End Sub

'===========================================================
' Find passage end by paren depth counting (right-to-left).
' Returns position just before the outermost citation paren.
'===========================================================
Private Function FindQuoteEnd(oRange As Range) As Long

    Dim sText As String
    sText = oRange.text

    Dim sTrimmed As String
    sTrimmed = RTrim(sText)

    Dim i As Long
    Dim nDepth As Integer
    nDepth = 0
    Dim nCitationStart As Long
    nCitationStart = -1

    For i = Len(sTrimmed) To 1 Step -1
        Dim c As String
        c = Mid(sTrimmed, i, 1)
        If c = ")" Then nDepth = nDepth + 1
        If c = "(" Then
            nDepth = nDepth - 1
            If nDepth = 0 Then
                nCitationStart = i
                Exit For
            End If
        End If
    Next i

    Dim lPassageEnd As Long

    If nCitationStart > 1 Then
        Dim sPassagePart As String
        sPassagePart = Left(sTrimmed, nCitationStart - 1)
        sPassagePart = RTrim(sPassagePart)
        lPassageEnd = oRange.start + Len(sPassagePart)
    ElseIf nCitationStart = 1 Then
        ' The citation parenthesis opens at the very first character, so there
        ' is no quoted passage ahead of it -- this is a citation-only paste.
        ' Report an empty passage (boundary at the range start) so the
        ' citation-only branch fires and the passage-only steps do not run
        ' against the citation itself. In particular, the "replace paragraph
        ' marks with a pilcrow" step (Step 11e) would otherwise scan the whole
        ' citation and, when the paste sits at the end of the document, try to
        ' replace the final paragraph mark -- which raises error 4608 "Value out
        ' of range" because that mark cannot be deleted.
        lPassageEnd = oRange.start
    Else
        ' No parenthetical citation found at all: treat the whole range as the
        ' passage (unchanged fallback for bare / unparenthesized content).
        lPassageEnd = oRange.End
    End If

    FindQuoteEnd = lPassageEnd

End Function

'===========================================================
' Find passage end using Word paragraph objects.
' Called BEFORE RemoveBlankParagraphs while citation is
' still on its own paragraph   most reliable boundary
'===========================================================
Private Function FindQuoteEndByParagraph(oRange As Range) As Long

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim oPara As Paragraph
    Dim oParas() As Paragraph
    Dim nCount As Integer
    nCount = 0

    For Each oPara In oRange.Paragraphs
        Dim sParaText As String
        sParaText = oPara.Range.text

        Dim nStrip As Long
        nStrip = Len(sParaText)
        Do While nStrip > 0
            Dim nLast As Long
            nLast = AscW(Mid(sParaText, nStrip, 1))
            If nLast = 13 Or nLast = 11 Then
                nStrip = nStrip - 1
            Else
                Exit Do
            End If
        Loop
        sParaText = Trim(Left(sParaText, nStrip))

        If Len(sParaText) > 0 Then
            nCount = nCount + 1
            ReDim Preserve oParas(nCount - 1)
            Set oParas(nCount - 1) = oPara
        End If
    Next oPara

    If nCount < 2 Then
        FindQuoteEndByParagraph = FindQuoteEnd(oRange)
        Exit Function
    End If

    Dim oLastPassagePara As Paragraph
    Set oLastPassagePara = oParas(nCount - 2)

    Dim sLast As String
    sLast = oLastPassagePara.Range.text

    Dim nStripL As Long
    nStripL = Len(sLast)
    Do While nStripL > 0
        Dim nLastL As Long
        nLastL = AscW(Mid(sLast, nStripL, 1))
        If nLastL = 13 Or nLastL = 11 Then
            nStripL = nStripL - 1
        Else
            Exit Do
        End If
    Loop

    FindQuoteEndByParagraph = oLastPassagePara.Range.start + nStripL

End Function

'===========================================================
' Swap internal smart quotes (double <-> single).
' FIX 3: Reuse a single Range object via SetRange instead of
'         creating a new Range object for every character,
'         eliminating thousands of COM allocations.
'===========================================================
Private Sub SwapSmartQuotes(oRange As Range)
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    Dim lStart As Long
    Dim lEnd As Long
    Dim lPos As Long
    lStart = oRange.start
    lEnd = oRange.End

    ' Read the text once for fast character inspection
    Dim sText As String
    sText = oRange.text

    ' Allocate a single Range object and move it, rather than
    ' creating/destroying one per character
    Dim oChar As Range
    Set oChar = oDoc.Range(lStart, lStart + 1)

    For lPos = lEnd - 1 To lStart Step -1
        Dim nCode As Long
        nCode = AscW(Mid(sText, lPos - lStart + 1, 1))

        ' Only touch characters that need swapping
        If nCode = &H201C Or nCode = &H201D Or _
           nCode = &H2018 Or nCode = &H2019 Then

            oChar.SetRange lPos, lPos + 1

            Dim sReplace As String
            sReplace = ""
            Select Case nCode
                Case &H201C  ' " open double -> ' open single
                    sReplace = ChrW(&H2018)
                Case &H201D  ' " close double -> ' close single
                    sReplace = ChrW(&H2019)
                Case &H2018  ' ' open single -> " open double
                    sReplace = ChrW(&H201C)
                Case &H2019  ' U+2019 close single -> " close double
                             '   EXCEPTIONS: convert to straight apostrophe (Chr 39):
                             '   (1) immediately followed by s/S (statute's)
                             '   (2) immediately preceded by s/S AND no unmatched
                             '       U+2018 to the left (Plaintiffs', parties')
                    Dim nSwapPrev As Long
                    Dim nSwapNext As Long
                    Dim sSwapIdx As Long
                    sSwapIdx = lPos - lStart + 1
                    If sSwapIdx < Len(sText) Then
                        nSwapNext = AscW(Mid(sText, sSwapIdx + 1, 1))
                    Else
                        nSwapNext = 0
                    End If
                    If sSwapIdx > 1 Then
                        nSwapPrev = AscW(Mid(sText, sSwapIdx - 1, 1))
                    Else
                        nSwapPrev = 0
                    End If
                    If nSwapNext = 115 Or nSwapNext = 83 Then
                        sReplace = Chr(39)
                    ElseIf nSwapPrev = 115 Or nSwapPrev = 83 Then
                        Dim nSwapOpen As Long, nSwapClose As Long, kSwap As Long
                        nSwapOpen = 0
                        nSwapClose = 0
                        For kSwap = 1 To sSwapIdx - 1
                            Dim nKC As Long
                            nKC = AscW(Mid(sText, kSwap, 1))
                            If nKC = &H2018 Then
                                nSwapOpen = nSwapOpen + 1
                            ElseIf nKC = &H2019 Then
                                nSwapClose = nSwapClose + 1
                            End If
                        Next kSwap
                        If nSwapOpen > nSwapClose Then
                            sReplace = ChrW(&H201D)
                        Else
                            sReplace = Chr(39)
                        End If
                    Else
                        sReplace = ChrW(&H201D)
                    End If
            End Select
            If sReplace <> "" Then
                oChar.text = sReplace
            End If
        End If
    Next lPos

    Set oChar = Nothing
End Sub

'===========================================================
' Balance nested quote marks in passage.
' Curly right singles (8217) are always closing quote marks
' in Lexis+/Westlaw text ? never apostrophes.
' Straight singles (39) followed immediately by a letter are
' treated as apostrophes and not counted; all others are
' counted as open single quote marks.
' Straight doubles (34) are counted as open double quote marks.
'===========================================================
Private Sub BalanceNestedQuotes(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim lStart As Long
    Dim lEnd As Long
    Dim lPos As Long
    Dim sChar As String

    Dim nOpenDouble As Long
    Dim nCloseDouble As Long
    Dim nOpenSingle As Long
    Dim nCloseSingle As Long

    Dim nMissingOpenDouble As Long
    Dim nMissingOpenSingle As Long
    Dim nMissingCloseDouble As Long
    Dim nMissingCloseSingle As Long

    Dim i As Long

    lStart = oRange.start
    lEnd = oRange.End

    nOpenDouble = 0
    nCloseDouble = 0
    nOpenSingle = 0
    nCloseSingle = 0

    ' FIX 3: Read text once into a string for counting   no per-char Range objects needed.
    Dim sText As String
    sText = oRange.text

    For lPos = 1 To Len(sText)
        sChar = Mid(sText, lPos, 1)

        Select Case AscW(sChar)
            Case &H201C
                nOpenDouble = nOpenDouble + 1
            Case &H201D
                nCloseDouble = nCloseDouble + 1
            Case &H2018, &H201A
                nOpenSingle = nOpenSingle + 1
            Case &H2019
                nCloseSingle = nCloseSingle + 1
            Case 34
                ' straight double ? counted as an open double quote mark
                nOpenDouble = nOpenDouble + 1
            Case 39
                ' straight single ? treat as apostrophe only if immediately followed by a letter;
                ' otherwise count as an open single quote mark
                Dim nNextChar39 As Long
                If lPos < Len(sText) Then
                    nNextChar39 = AscW(Mid(sText, lPos + 1, 1))
                Else
                    nNextChar39 = 0
                End If
                If Not ((nNextChar39 >= 65 And nNextChar39 <= 90) Or _
                        (nNextChar39 >= 97 And nNextChar39 <= 122)) Then
                    nOpenSingle = nOpenSingle + 1
                End If
        End Select
    Next lPos

    If nOpenDouble > nCloseDouble Then
        nMissingCloseDouble = nOpenDouble - nCloseDouble
        nMissingOpenDouble = 0
    ElseIf nCloseDouble > nOpenDouble Then
        nMissingOpenDouble = nCloseDouble - nOpenDouble
        nMissingCloseDouble = 0
    Else
        nMissingOpenDouble = 0
        nMissingCloseDouble = 0
    End If

    If nOpenSingle > nCloseSingle Then
        nMissingCloseSingle = nOpenSingle - nCloseSingle
        nMissingOpenSingle = 0
    ElseIf nCloseSingle > nOpenSingle Then
        nMissingOpenSingle = nCloseSingle - nOpenSingle
        nMissingCloseSingle = 0
    Else
        nMissingOpenSingle = 0
        nMissingCloseSingle = 0
    End If

    Dim oInsert As Range
    Dim nShift As Long
    nShift = 0

    For i = 1 To nMissingOpenDouble
        Set oInsert = oDoc.Range(lStart + nShift, lStart + nShift)
        oInsert.InsertBefore ChrW(&H201C)
        nShift = nShift + 1
    Next i

    For i = 1 To nMissingOpenSingle
        Set oInsert = oDoc.Range(lStart + nShift, lStart + nShift)
        oInsert.InsertBefore ChrW(&H2018)
        nShift = nShift + 1
    Next i

    lEnd = lEnd + nShift

    Dim nShiftEnd As Long
    nShiftEnd = 0

    For i = 1 To nMissingCloseSingle
        Set oInsert = oDoc.Range(lEnd + nShiftEnd, lEnd + nShiftEnd)
        oInsert.InsertBefore ChrW(&H2019)
        nShiftEnd = nShiftEnd + 1
    Next i

    For i = 1 To nMissingCloseDouble
        Set oInsert = oDoc.Range(lEnd + nShiftEnd, lEnd + nShiftEnd)
        oInsert.InsertBefore ChrW(&H201D)
        nShiftEnd = nShiftEnd + 1
    Next i

    oRange.SetRange lStart, lEnd + nShiftEnd

End Sub

'===========================================================
' Detect and remove pre-paste opener signals typed before paste.
'===========================================================
Private Sub DetectPrePasteOpeners(oDoc As Document, ByRef lStart As Long, _
                                   ByRef nPreDouble As Integer, _
                                   ByRef nPreSingle As Integer, _
                                   ByRef sPreOpeners As String, _
                                   ByRef bProperNoun As Boolean)

    nPreDouble = 0
    nPreSingle = 0
    sPreOpeners = ""
    bProperNoun = False

    Dim lPos As Long
    lPos = lStart - 1

    Dim sCollect As String
    sCollect = ""

    Do While lPos >= 0
        Dim oLook As Range
        Set oLook = oDoc.Range(lPos, lPos + 1)
        Dim nLook As Long
        nLook = AscW(oLook.text)
        Set oLook = Nothing

        If nLook = &H201C Then
            nPreDouble = nPreDouble + 1
            sCollect = ChrW(&H201C) & sCollect
            Dim oDelD As Range
            Set oDelD = oDoc.Range(lPos, lPos + 1)
            oDelD.Delete
            Set oDelD = Nothing
            lStart = lStart - 1
            lPos = lPos - 1
        ElseIf nLook = &H2018 Then
            nPreSingle = nPreSingle + 1
            sCollect = ChrW(&H2018) & sCollect
            Dim oDelS As Range
            Set oDelS = oDoc.Range(lPos, lPos + 1)
            oDelS.Delete
            Set oDelS = Nothing
            lStart = lStart - 1
            lPos = lPos - 1
        ElseIf nLook = 94 Then
            bProperNoun = True
            Dim oDelC As Range
            Set oDelC = oDoc.Range(lPos, lPos + 1)
            oDelC.Delete
            Set oDelC = Nothing
            lStart = lStart - 1
            lPos = lPos - 1
        Else
            Exit Do
        End If
    Loop

    sPreOpeners = sCollect

End Sub

'===========================================================
' Detect and remove subdivision marker(s) typed in the
' document immediately before the paste cursor position.
'
' Reads a 20-character snapshot ending at the cursor, skips
' any trailing spaces/tabs, then walks backwards collecting
' "(x)" tokens where x is a single alphanumeric character.
' Tokens are prepended as collected so the result is already
' in left-to-right order. Deletes the tokens plus any
' trailing whitespace in one Range.Delete call and
' decrements lStart accordingly.
'
' Returns "" if no subdivision markers are found.
' Called before paste; result used only when bIsStatute = True.
'===========================================================
Private Function DetectPrePasteSubdivision(oDoc As Document, _
                                            ByRef lStart As Long) As String

    Dim lScanFrom As Long
    If lStart >= 20 Then lScanFrom = lStart - 20 Else lScanFrom = 0

    Dim sScan As String
    sScan = oDoc.Range(lScanFrom, lStart).text

    Dim i As Long
    i = Len(sScan)
    If i = 0 Then
        DetectPrePasteSubdivision = ""
        Exit Function
    End If

    ' Skip any trailing spaces or tabs the user may have typed after the
    ' marker   makes detection robust against accidental trailing spaces.
    Do While i >= 1
        Dim nWS As Long
        nWS = AscW(Mid(sScan, i, 1))
        If nWS = 32 Or nWS = 9 Then
            i = i - 1
        Else
            Exit Do
        End If
    Loop

    ' Collect "(x)" tokens walking left, prepending each so the result
    ' is in left-to-right order without needing to reverse an array.
    Dim sResult As String
    sResult = ""

    Do
        If i < 3 Then Exit Do
        If Mid(sScan, i, 1) <> ")" Then Exit Do

        Dim nMid As Long
        nMid = AscW(Mid(sScan, i - 1, 1))
        If Not ((nMid >= 65 And nMid <= 90) Or _
                (nMid >= 97 And nMid <= 122) Or _
                (nMid >= 48 And nMid <= 57)) Then Exit Do

        If Mid(sScan, i - 2, 1) <> "(" Then Exit Do

        sResult = "(" & ChrW(nMid) & ")" & sResult
        i = i - 3
    Loop

    If sResult = "" Then
        DetectPrePasteSubdivision = ""
        Exit Function
    End If

    Dim lDelStart As Long
    lDelStart = lScanFrom + i

    Dim oDel As Range
    Set oDel = oDoc.Range(lDelStart, lStart)
    oDel.Delete
    Set oDel = Nothing

    lStart = lDelStart
    DetectPrePasteSubdivision = sResult

End Function

'===========================================================
' Detect and remove an open bracket typed immediately before
' the paste cursor, with optional intervening spaces.
' Returns True if found and removed; False otherwise.
'===========================================================
Private Function DetectPrePasteParenthetical(oDoc As Document, _
                                              ByRef lStart As Long) As Boolean

    Dim lScanFrom As Long
    If lStart >= 10 Then lScanFrom = lStart - 10 Else lScanFrom = 0

    Dim sScan As String
    sScan = oDoc.Range(lScanFrom, lStart).text

    Dim i As Long
    i = Len(sScan)
    ' Test i before touching Mid: VBA's And is NOT short-circuiting, so when
    ' i = 0 -- an empty scan at the first paragraph (lStart = 0), or a run of
    ' only spaces before the cursor -- "Mid(sScan, 0, 1)" would raise error 5
    ' "Invalid procedure call or argument".
    Do While i >= 1
        If AscW(Mid(sScan, i, 1)) <> 32 Then Exit Do
        i = i - 1
    Loop

    ' Nested (not "i >= 1 And AscW(...)") so Mid is never called with i = 0 --
    ' VBA's And evaluates both sides, and Mid(sScan, 0, 1) raises error 5.
    DetectPrePasteParenthetical = False
    If i >= 1 Then
        If AscW(Mid(sScan, i, 1)) = 91 Then
            Dim lDelStart As Long
            lDelStart = lScanFrom + i - 1
            oDoc.Range(lDelStart, lStart).Delete
            lStart = lDelStart
            DetectPrePasteParenthetical = True
        End If
    End If

End Function

'===========================================================
' Detect "in" or "In" as the word immediately before lStart.
' Peek only -- does NOT remove the word or modify lStart.
' Returns True only for exactly "in" or "In" (not "IN",
' not part of a longer word like "begin").
'===========================================================
Private Function DetectPrePasteTextual(oDoc As Document, _
                                        ByVal lStart As Long) As Boolean

    DetectPrePasteTextual = False
    If lStart < 3 Then Exit Function  ' need at least " in"

    ' Read up to 10 characters before lStart -- wide enough that a run
    ' of trailing spaces cannot push the character preceding "in" out
    ' of the window (a 5-char window let "begin" + 3 spaces match).
    Dim nRead As Long
    nRead = 10
    If lStart < nRead Then nRead = lStart
    Dim sScan As String
    sScan = oDoc.Range(lStart - nRead, lStart).text

    ' Strip trailing spaces
    Dim i As Long
    i = Len(sScan)
    ' Test i before touching Mid: VBA's And is NOT short-circuiting, so when
    ' i = 0 -- an empty scan at the first paragraph (lStart = 0), or a run of
    ' only spaces before the cursor -- "Mid(sScan, 0, 1)" would raise error 5
    ' "Invalid procedure call or argument".
    Do While i >= 1
        If AscW(Mid(sScan, i, 1)) <> 32 Then Exit Do
        i = i - 1
    Loop
    If i < 2 Then Exit Function

    ' Check for "in" or "In" at positions i-1 and i
    Dim c1 As Long, c2 As Long
    c1 = AscW(Mid(sScan, i - 1, 1))
    c2 = AscW(Mid(sScan, i, 1))
    ' Must be exactly i=I followed by n=N
    If Not ((c1 = 105 Or c1 = 73) And c2 = 110) Then Exit Function  ' i/I, n

    ' The character before "in" must be a non-letter (space, para mark,
    ' punctuation) so "begin" does not trigger this.
    If i > 2 Then
        Dim cPre As Long
        cPre = AscW(Mid(sScan, i - 2, 1))
        Dim bIsLetter As Boolean
        bIsLetter = ((cPre >= 65 And cPre <= 90) Or (cPre >= 97 And cPre <= 122))
        If bIsLetter Then Exit Function
    End If

    DetectPrePasteTextual = True

End Function

'===========================================================
' Restructure pasted block for textual-sentence mode.
'
' Normal input:   "Passage." (Citation.)
' Output:         Citation, the court held "Passage."
'
' Citation-only:  (Citation.)
' Output:         Citation
'
' The citation body is extracted by finding the outermost (...)
' from the right and stripping its parens and terminal period.
'===========================================================
' Detect the " " statutory-textual pre-paste signal.
' Looks for a lone " " character immediately before the cursor
' (with optional trailing spaces, after subdivision removal).
' Removes the " " from the document and adjusts lStart.
' All data (code name, section, subdivision) is extracted later
' from the citation sentence already present in the paste.
'===========================================================
Private Sub DetectPrePasteStatutoryTextual(oDoc As Document, _
                                           ByRef lStart As Long, _
                                           ByRef bStatutoryTextual As Boolean)

    bStatutoryTextual = False
    If lStart < 1 Then Exit Sub

    ' Read up to 5 chars before cursor to allow for trailing spaces
    Dim nRead As Long
    nRead = 5
    If lStart < nRead Then nRead = lStart
    Dim sScan As String
    sScan = oDoc.Range(lStart - nRead, lStart).text

    ' Strip trailing spaces
    Dim i As Long
    i = Len(sScan)
    ' Test i before touching Mid: VBA's And is NOT short-circuiting, so when
    ' i = 0 -- an empty scan at the first paragraph (lStart = 0), or a run of
    ' only spaces before the cursor -- "Mid(sScan, 0, 1)" would raise error 5
    ' "Invalid procedure call or argument".
    Do While i >= 1
        If AscW(Mid(sScan, i, 1)) <> 32 Then Exit Do
        i = i - 1
    Loop
    If i < 1 Then Exit Sub

    ' Must be the section sign   (U+00A7)
    If AscW(Mid(sScan, i, 1)) <> &HA7 Then Exit Sub

    ' Found it -- remove from document (the   and any trailing spaces)
    Dim lDelStart As Long
    lDelStart = (lStart - nRead) + i - 1
    Dim oDel As Range
    Set oDel = oDoc.Range(lDelStart, lStart)
    oDel.Delete
    Set oDel = Nothing
    lStart = lDelStart

    bStatutoryTextual = True

End Sub

'===========================================================
' Restructure the pasted block as a statutory textual sentence.
' Input:  "Passage." (Civ. Code,   1794[, subd. (a)].)
'         (citation normalized + subdivision inserted by statute pipeline)
' Output: Civil Code section 1794[, subdivision (a)], provides "Passage."
' No trailing parenthetical. lRangeEnd updated.
'===========================================================
Private Sub RestructureAsStatutoryTextual(oDoc As Document, _
                                          ByVal lStart As Long, _
                                          ByRef lRangeEnd As Long)

    Dim oRng As Range
    Dim sFull As String
    Dim sCitBody As String
    Dim sCode As String
    Dim sCodeTxt As String
    Dim sSec As String
    Dim sSubDiv As String
    Dim sPass As String
    Dim sSubClause As String
    Dim sOut As String
    Dim sAfter As String
    Dim sSubRest As String
    Dim nDp As Integer
    Dim nCitS As Long
    Dim nSecS As Long
    Dim nSign As Long
    Dim nSubd As Long
    Dim nTDp As Integer
    Dim nAsc As Long
    Dim nPos As Long
    Dim nPos2 As Long
    Dim sC As String
    Dim oRpl As Range

    Set oRng = oDoc.Range(lStart, lRangeEnd)
    sFull = RTrim(oRng.text)
    If Len(sFull) = 0 Then Exit Sub

    ' Find outermost citation paren from the right
    nDp = 0
    nCitS = -1
    nPos = Len(sFull)
    Do While nPos >= 1
        sC = Mid(sFull, nPos, 1)
        If sC = ")" Then nDp = nDp + 1
        If sC = "(" Then
            nDp = nDp - 1
            If nDp = 0 Then
                nCitS = nPos
                nPos = 0
            End If
        End If
        nPos = nPos - 1
    Loop
    If nCitS = -1 Then Exit Sub

    ' Strip outer parens and terminal period
    sCitBody = Mid(sFull, nCitS + 1, Len(sFull) - nCitS - 1)
    sCitBody = RTrim(sCitBody)
    If Len(sCitBody) > 0 Then
        If Right(sCitBody, 1) = "." Then sCitBody = Left(sCitBody, Len(sCitBody) - 1)
    End If
    sCitBody = RTrim(sCitBody)

    ' Code name: before the section sign. With no section sign this may be a
    ' California Rules of Court cite ("... rule N"), which carries no "§" but
    ' should still restructure under the same "§" pre-paste hint, for
    ' consistency. Hand it off and stop.
    nSign = InStr(sCitBody, ChrW(&HA7))
    If nSign = 0 Then
        RestructureRuleOfCourtTextual oDoc, lStart, lRangeEnd, sFull, nCitS, sCitBody
        Exit Sub
    End If
    sCode = RTrim(Left(sCitBody, nSign - 1))
    If Len(sCode) > 0 Then
        If Right(sCode, 1) = "," Then sCode = RTrim(Left(sCode, Len(sCode) - 1))
    End If

    ' Section number: digits/dots/hyphens after section sign
    sAfter = LTrim(Mid(sCitBody, nSign + 1))
    sSec = ""
    nSecS = 1
    Do While nSecS <= Len(sAfter)
        nAsc = AscW(Mid(sAfter, nSecS, 1))
        If (nAsc >= 48 And nAsc <= 57) Or nAsc = 46 Or nAsc = 45 Then
            sSec = sSec & Mid(sAfter, nSecS, 1)
            nSecS = nSecS + 1
        Else
            nSecS = Len(sAfter) + 1
        End If
    Loop
    If Len(sSec) = 0 Then Exit Sub

    ' Subdivision: look for "subd. (x)" in citation body
    sSubDiv = ""
    nSubd = InStr(1, sCitBody, "subd. ", vbTextCompare)
    If nSubd > 0 Then
        sSubRest = Mid(sCitBody, nSubd + 6)
        nTDp = 0
        nPos2 = 1
        Do While nPos2 <= Len(sSubRest)
            sC = Mid(sSubRest, nPos2, 1)
            If sC = "(" Then nTDp = nTDp + 1
            If sC = ")" Then
                nTDp = nTDp - 1
                If nTDp = 0 Then
                    sSubDiv = Left(sSubRest, nPos2)
                    If nPos2 < Len(sSubRest) Then
                        If Mid(sSubRest, nPos2 + 1, 1) <> "(" Then
                            nPos2 = Len(sSubRest) + 1
                        End If
                    Else
                        nPos2 = Len(sSubRest) + 1
                    End If
                End If
            End If
            nPos2 = nPos2 + 1
        Loop
        If nTDp <> 0 Then sSubDiv = ""
    End If

    ' Map abbreviation to textual name
    sCodeTxt = StatuteCodeTextualName(sCode)
    If Len(sCodeTxt) = 0 Then sCodeTxt = sCode

    ' Passage: everything before the citation paren
    sPass = RTrim(Left(sFull, nCitS - 1))

    sSubClause = ""
    If Len(sSubDiv) > 0 Then sSubClause = ", subdivision " & sSubDiv

    sOut = sCodeTxt & " section " & sSec & sSubClause & " provides, " & sPass

    Set oRpl = oDoc.Range(lStart, lStart + Len(sFull))
    oRpl.text = sOut
    lRangeEnd = lStart + Len(sOut)

    Set oRpl = Nothing
    Set oRng = Nothing

End Sub

'===========================================================
' Restructure a California Rules of Court cite as a textual sentence -- the
' rule analog of RestructureAsStatutoryTextual, triggered by the same "§"
' pre-paste hint even though a rule cite carries no section sign.
'
'   Input:  "Passage." (Cal. Rules of Court, rule 3.1350(c).)
'   Output: California Rules of Court, rule 3.1350(c) provides, "Passage."
'
' Leaves the block untouched when the citation is not a Rules of Court cite,
' so a stray "§" hint on some other authority does nothing here. sFull, nCitS
' and sCitBody are the values the caller already computed (full block text,
' 1-based index of the citation's opening paren, and the citation body with
' its outer parens and terminal period already stripped).
'===========================================================
Private Sub RestructureRuleOfCourtTextual(oDoc As Document, _
                                          ByVal lStart As Long, _
                                          ByRef lRangeEnd As Long, _
                                          ByVal sFull As String, _
                                          ByVal nCitS As Long, _
                                          ByVal sCitBody As String)

    ' Require "Rules of Court" so the "rule N" pattern can't misfire on other
    ' authorities that happen to contain the word "rule".
    If InStr(1, sCitBody, "Rules of Court", vbTextCompare) = 0 Then Exit Sub

    Dim nRule As Long
    nRule = InStr(1, sCitBody, "rule ", vbTextCompare)
    If nRule = 0 Then Exit Sub

    ' Everything after "rule " is the rule number (with any inline subdivision,
    ' e.g. "3.1350(c)"). sCitBody already had its terminal period removed; strip
    ' any stray trailing comma/period defensively.
    Dim sRule As String
    sRule = Trim(Mid(sCitBody, nRule + Len("rule ")))
    Do While Len(sRule) > 0 And (Right(sRule, 1) = "," Or Right(sRule, 1) = ".")
        sRule = RTrim(Left(sRule, Len(sRule) - 1))
    Loop
    If Len(sRule) = 0 Then Exit Sub

    ' Passage: everything before the citation paren.
    Dim sPass As String
    sPass = RTrim(Left(sFull, nCitS - 1))

    Dim sOut As String
    sOut = "California Rules of Court, rule " & sRule & " provides, " & sPass

    Dim oRpl As Range
    Set oRpl = oDoc.Range(lStart, lStart + Len(sFull))
    oRpl.text = sOut
    lRangeEnd = lStart + Len(sOut)
    Set oRpl = Nothing

End Sub

'===========================================================
' Map a citation code abbreviation to its full textual name.
' e.g. "Civ. Code" -> "Civil Code"
' Returns empty string if not recognised (caller uses abbrev as fallback).
'===========================================================
Private Function StatuteCodeTextualName(sAbbrev As String) As String

    Const n As Integer = 22
    Dim aAbbrev(n) As String
    Dim aTextual(n) As String

    aAbbrev(0) = "Civ. Code": aTextual(0) = "Civil Code"
    aAbbrev(1) = "Pen. Code": aTextual(1) = "Penal Code"
    aAbbrev(2) = "Lab. Code": aTextual(2) = "Labor Code"
    aAbbrev(3) = "Veh. Code": aTextual(3) = "Vehicle Code"
    aAbbrev(4) = "Corp. Code": aTextual(4) = "Corporations Code"
    aAbbrev(5) = "Prob. Code": aTextual(5) = "Probate Code"
    aAbbrev(6) = "Fam. Code": aTextual(6) = "Family Code"
    aAbbrev(7) = "Gov. Code": aTextual(7) = "Government Code"
    aAbbrev(8) = "Health & Saf. Code": aTextual(8) = "Health and Safety Code"
    aAbbrev(9) = "Welf. & Inst. Code": aTextual(9) = "Welfare and Institutions Code"
    aAbbrev(10) = "Bus. & Prof. Code": aTextual(10) = "Business and Professions Code"
    aAbbrev(11) = "Com. Code": aTextual(11) = "Commercial Code"
    aAbbrev(12) = "Evid. Code": aTextual(12) = "Evidence Code"
    aAbbrev(13) = "Code Civ. Proc.": aTextual(13) = "Code of Civil Procedure"
    aAbbrev(14) = "Ins. Code": aTextual(14) = "Insurance Code"
    aAbbrev(15) = "Rev. & Tax. Code": aTextual(15) = "Revenue and Taxation Code"
    aAbbrev(16) = "U.S.C.": aTextual(16) = "United States Code"
    aAbbrev(17) = "Cal. U. Com. Code": aTextual(17) = "Commercial Code"
    aAbbrev(18) = "Cal. Civ. Code": aTextual(18) = "Civil Code"
    aAbbrev(19) = "Cal. Pen. Code": aTextual(19) = "Penal Code"
    aAbbrev(20) = "Cal. Lab. Code": aTextual(20) = "Labor Code"
    aAbbrev(21) = "Cal. Veh. Code": aTextual(21) = "Vehicle Code"
    aAbbrev(22) = "Cal. Health & Saf. Code": aTextual(22) = "Health and Safety Code"

    Dim i As Integer
    For i = 0 To n
        If StrComp(sAbbrev, aAbbrev(i), vbTextCompare) = 0 Then
            StatuteCodeTextualName = aTextual(i)
            Exit Function
        End If
    Next i

    StatuteCodeTextualName = ""

End Function

' lRangeEnd is updated to reflect the new block length.
'===========================================================
Private Sub RestructureAsTextual(oDoc As Document, _
                                  ByVal lStart As Long, _
                                  ByRef lRangeEnd As Long, _
                                  ByVal bCitationOnly As Boolean)

    Dim oRange As Range
    Set oRange = oDoc.Range(lStart, lRangeEnd)

    Dim sText As String
    sText = RTrim(oRange.text)
    If Len(sText) = 0 Then Exit Sub

    ' Find the outermost citation parenthesis by scanning right-to-left.
    Dim nDepth As Integer
    Dim nCitStart As Long
    nDepth = 0
    nCitStart = -1
    Dim k As Long
    For k = Len(sText) To 1 Step -1
        Dim ck As String
        ck = Mid(sText, k, 1)
        If ck = ")" Then nDepth = nDepth + 1
        If ck = "(" Then
            nDepth = nDepth - 1
            If nDepth = 0 Then
                nCitStart = k
                Exit For
            End If
        End If
    Next k
    If nCitStart = -1 Then Exit Sub

    ' Extract citation body: strip outer parens and terminal period.
    Dim sCitBody As String
    sCitBody = Mid(sText, nCitStart + 1, Len(sText) - nCitStart - 1)
    sCitBody = RTrim(sCitBody)
    If Len(sCitBody) > 0 And Right(sCitBody, 1) = "." Then
        sCitBody = Left(sCitBody, Len(sCitBody) - 1)
    End If
    sCitBody = RTrim(sCitBody)

    Dim sNew As String

    If bCitationOnly Then
        ' No passage -- just the bare citation
        sNew = sCitBody
    Else
        ' Extract the quoted passage: everything before nCitStart,
        ' trimmed of surrounding whitespace.
        Dim sPassage As String
        sPassage = RTrim(Left(sText, nCitStart - 1))
        ' sPassage is already wrapped in outer double quotes by WrapInDoubleQuotes.
        ' Build: Citation, the court held "Passage."
        sNew = sCitBody & ", the court held " & sPassage
    End If

    ' Replace the trimmed portion of the range with the new text.
    Dim oReplace As Range
    Set oReplace = oDoc.Range(lStart, lStart + Len(sText))
    oReplace.text = sNew
    lRangeEnd = lStart + Len(sNew)

    ' Italicize the case name. The citation body sits at lStart through
    ' lStart + Len(sCitBody). Scan forward for (YYYY) to locate the
    ' case name end, then apply italic to the text before it.
    ItalicizeCaseNameBare oDoc, lStart, lStart + Len(sCitBody)

    Set oReplace = Nothing
    Set oRange = Nothing

End Sub


'===========================================================
' Detect "Defendant cites" / "Plaintiff cites" (and recognised
' variants) as the phrase immediately before lStart.
' Peek only -- does NOT remove text or modify lStart.
'
' Recognised subject words (case-sensitive, must start with capital):
'   Defendant, Defendants, Plaintiff, Plaintiffs,
'   Defendant-Appellant, Defendant-Respondent,
'   Cross-Defendant, Cross-Defendants,
'   Cross-Complainant, Cross-Complainants,
'   Petitioner, Petitioners, Respondent, Respondents
'
' Recognised verb: "cites" or "cite" (with optional trailing spaces,
' up to 4, to tolerate accidental extra spaces before paste).
'===========================================================
Private Function DetectPrePasteCites(oDoc As Document, _
                                      ByVal lStart As Long) As Boolean

    DetectPrePasteCites = False
    If lStart < 6 Then Exit Function

    ' Read up to 60 characters before lStart to cover longest subject word
    ' ("Cross-Complainants") plus " cites " plus a few extra spaces.
    Dim nRead As Long
    nRead = 60
    If lStart < nRead Then nRead = lStart
    Dim sScan As String
    sScan = oDoc.Range(lStart - nRead, lStart).text

    ' Strip up to 4 trailing spaces to tolerate accidental extra spaces.
    Dim iEnd As Long
    iEnd = Len(sScan)
    Dim nStripped As Integer
    nStripped = 0
    Do While iEnd >= 1 And AscW(Mid(sScan, iEnd, 1)) = 32 And nStripped < 4
        iEnd = iEnd - 1
        nStripped = nStripped + 1
    Loop
    If iEnd < 4 Then Exit Function

    ' The text up to iEnd must end with "cites" or "cite". Nested Ifs so Mid is
    ' never called with a start < 1: iEnd can be exactly 4 here, and
    ' "iEnd >= 5 And Mid(sScan, iEnd - 4, 5)" would still evaluate Mid at index 0
    ' (VBA's And is not short-circuiting), raising error 5. iEnd >= 4 is already
    ' guaranteed above, so the "cite" branch's Mid start (iEnd - 3) is >= 1.
    Dim sVerb As String
    Dim nVerbLen As Integer
    nVerbLen = 0
    If iEnd >= 5 Then
        If Mid(sScan, iEnd - 4, 5) = "cites" Then nVerbLen = 5
    End If
    If nVerbLen = 0 Then
        If Mid(sScan, iEnd - 3, 4) = "cite" Then nVerbLen = 4
    End If
    If nVerbLen = 0 Then Exit Function

    ' The character before the verb must be a space (word boundary).
    Dim iVerbStart As Long
    iVerbStart = iEnd - nVerbLen + 1
    If iVerbStart < 2 Then Exit Function
    If AscW(Mid(sScan, iVerbStart - 1, 1)) <> 32 Then Exit Function

    ' Extract the word immediately before the verb (after its leading space).
    ' Walk left from iVerbStart-2 to find the start of the subject word.
    Dim iSubEnd As Long
    iSubEnd = iVerbStart - 2  ' last char of subject word
    If iSubEnd < 1 Then Exit Function

    Dim iSubStart As Long
    iSubStart = iSubEnd
    Do While iSubStart > 1
        Dim nC As Long
        nC = AscW(Mid(sScan, iSubStart - 1, 1))
        ' Allow letters and hyphens (for "Cross-Defendant" etc.)
        If (nC >= 65 And nC <= 90) Or (nC >= 97 And nC <= 122) Or nC = 45 Then
            iSubStart = iSubStart - 1
        Else
            Exit Do
        End If
    Loop

    Dim sSubject As String
    sSubject = Mid(sScan, iSubStart, iSubEnd - iSubStart + 1)

    ' Match against the recognised subject list.
    Dim aSubjects(15) As String
    aSubjects(0) = "Defendant"
    aSubjects(1) = "Defendants"
    aSubjects(2) = "Plaintiff"
    aSubjects(3) = "Plaintiffs"
    aSubjects(4) = "Defendant-Appellant"
    aSubjects(5) = "Defendant-Respondent"
    aSubjects(6) = "Cross-Defendant"
    aSubjects(7) = "Cross-Defendants"
    aSubjects(8) = "Cross-Complainant"
    aSubjects(9) = "Cross-Complainants"
    aSubjects(10) = "Petitioner"
    aSubjects(11) = "Petitioners"
    aSubjects(12) = "Respondent"
    aSubjects(13) = "Respondents"
    aSubjects(14) = "Plaintiff-Appellant"
    aSubjects(15) = "Plaintiff-Respondent"

    Dim i As Integer
    For i = 0 To 15
        If sSubject = aSubjects(i) Then
            DetectPrePasteCites = True
            Exit Function
        End If
    Next i

End Function

'===========================================================
' Restructure pasted block for cites-sentence mode.
'
' The subject phrase ("Defendant cites" etc.) is already in the
' document before lStart and is NOT touched.
'
' Normal input:   "Passage." (Citation.)
' Output:         Citation for the proposition that "Passage."
'
' Citation-only:  (Citation.)
' Output:         Citation
'
' lRangeEnd is updated to reflect the new block length.
'===========================================================
Private Sub RestructureAsCites(oDoc As Document, _
                                ByVal lStart As Long, _
                                ByRef lRangeEnd As Long, _
                                ByVal bCitationOnly As Boolean)

    Dim oRange As Range
    Set oRange = oDoc.Range(lStart, lRangeEnd)

    Dim sText As String
    sText = RTrim(oRange.text)
    If Len(sText) = 0 Then Exit Sub

    ' Find the outermost citation parenthesis by scanning right-to-left.
    Dim nDepth As Integer
    Dim nCitStart As Long
    nDepth = 0
    nCitStart = -1
    Dim k As Long
    For k = Len(sText) To 1 Step -1
        Dim ck As String
        ck = Mid(sText, k, 1)
        If ck = ")" Then nDepth = nDepth + 1
        If ck = "(" Then
            nDepth = nDepth - 1
            If nDepth = 0 Then
                nCitStart = k
                Exit For
            End If
        End If
    Next k
    If nCitStart = -1 Then Exit Sub

    ' Extract citation body: strip outer parens and terminal period.
    Dim sCitBody As String
    sCitBody = Mid(sText, nCitStart + 1, Len(sText) - nCitStart - 1)
    sCitBody = RTrim(sCitBody)
    If Len(sCitBody) > 0 And Right(sCitBody, 1) = "." Then
        sCitBody = Left(sCitBody, Len(sCitBody) - 1)
    End If
    sCitBody = RTrim(sCitBody)

    Dim sNew As String

    If bCitationOnly Then
        ' No passage -- just the bare citation, no proposition phrase.
        sNew = sCitBody
    Else
        ' Extract the quoted passage: everything before nCitStart, trimmed.
        Dim sPassage As String
        sPassage = RTrim(Left(sText, nCitStart - 1))
        ' Build: Citation for the proposition that "Passage."
        sNew = sCitBody & " for the proposition that " & sPassage
    End If

    ' Replace the trimmed portion of the range with the new text.
    Dim oReplace As Range
    Set oReplace = oDoc.Range(lStart, lStart + Len(sText))
    oReplace.text = sNew
    lRangeEnd = lStart + Len(sNew)

    ' Italicize the case name in the bare citation at the start.
    ItalicizeCaseNameBare oDoc, lStart, lStart + Len(sCitBody)

    Set oReplace = Nothing
    Set oRange = Nothing

End Sub
'===========================================================
' Returns True if the pasted block is a statute citation.
' Detection: the last non-empty paragraph contains the
' section symbol (U+00A7), which is always present in a
' Lexis+/Westlaw statute citation and never in a case cite.
'===========================================================
' Scan passage paragraphs for numeric sub-paragraph markers
' of the form (n) at the start of a paragraph, where n is a
' single digit (0-9).  Two jobs:
'   1. If two or more such markers appear in ascending numeric
'      order, set bNumericSubParagraphs = True.
'   2. For every paragraph-start (n) marker immediately followed
'      by a non-space character, insert a space after the closing
'      paren (Lexis+ omits this space).
'===========================================================
Private Sub FixNumericSubParagraphs(oDoc As Document, _
                                     oRange As Range, _
                                     ByRef bNumericSubParagraphs As Boolean)

    Dim oPara As Paragraph
    Dim sParaText As String
    Dim nStrip As Long
    Dim nTailC As Long
    Dim nFirst As Long
    Dim nMid As Long
    Dim nThird As Long
    Dim nFourth As Long
    Dim nLastDigit As Integer
    Dim nFound As Integer
    Dim lParaStart As Long
    Dim oIns As Range

    nLastDigit = -1
    nFound = 0
    bNumericSubParagraphs = False

    For Each oPara In oRange.Paragraphs

        sParaText = oPara.Range.text
        nStrip = Len(sParaText)
        Do While nStrip > 0
            nTailC = AscW(Mid(sParaText, nStrip, 1))
            If nTailC = 13 Or nTailC = 11 Then
                nStrip = nStrip - 1
            Else
                Exit Do
            End If
        Loop
        sParaText = Left(sParaText, nStrip)
        If Len(sParaText) < 3 Then GoTo NextNumPara

        nFirst = AscW(Mid(sParaText, 1, 1))
        nMid = AscW(Mid(sParaText, 2, 1))
        nThird = AscW(Mid(sParaText, 3, 1))

        If nFirst = 40 And nThird = 41 Then
            If nMid >= 48 And nMid <= 57 Then

                If nLastDigit = -1 Then
                    nLastDigit = nMid
                    nFound = 1
                ElseIf nMid > nLastDigit Then
                    nLastDigit = nMid
                    nFound = nFound + 1
                End If

                ' Space fix: insert space after ) if next char is not a space
                If Len(sParaText) >= 4 Then
                    nFourth = AscW(Mid(sParaText, 4, 1))
                Else
                    nFourth = 32
                End If
                If nFourth <> 32 And nFourth <> 160 Then
                    lParaStart = oPara.Range.start
                    Set oIns = oDoc.Range(lParaStart + 3, lParaStart + 3)
                    oIns.InsertAfter " "
                    Set oIns = Nothing
                End If

            End If
        End If

NextNumPara:
    Next oPara

    If nFound >= 2 Then bNumericSubParagraphs = True

End Sub

'===========================================================
' Detect a multi-subsection statute paste.
' Returns True when at least two paragraphs in oRange begin
' with a single lowercase letter in parentheses -- e.g. (a),
' (b), (c) -- and those letters appear in ascending alphabetical
' order (not necessarily consecutive or adjacent paragraphs).
'===========================================================
Private Function DetectMultiSubsection(oRange As Range) As Boolean

    Dim oPara As Paragraph
    Dim nLastLetter As Integer
    Dim nFound As Integer
    Dim sParaText As String
    Dim nFirst As Long
    Dim nMid As Long
    Dim nThird As Long
    Dim nStrip As Long
    Dim nTail As Long

    nLastLetter = 0
    nFound = 0

    For Each oPara In oRange.Paragraphs

        sParaText = oPara.Range.text
        nStrip = Len(sParaText)
        Do While nStrip > 0
            nTail = AscW(Mid(sParaText, nStrip, 1))
            If nTail = 13 Or nTail = 11 Then
                nStrip = nStrip - 1
            Else
                Exit Do
            End If
        Loop
        sParaText = Left(sParaText, nStrip)
        If Len(sParaText) < 3 Then GoTo NextPara

        nFirst = AscW(Mid(sParaText, 1, 1))
        nMid = AscW(Mid(sParaText, 2, 1))
        nThird = AscW(Mid(sParaText, 3, 1))

        If nFirst = 40 And nThird = 41 Then
            If nMid >= 97 And nMid <= 122 Then
                If nLastLetter = 0 Then
                    nLastLetter = nMid
                    nFound = 1
                ElseIf nMid > nLastLetter Then
                    nLastLetter = nMid
                    nFound = nFound + 1
                End If
            End If
        End If

NextPara:
    Next oPara

    DetectMultiSubsection = (nFound >= 2)

End Function

'===========================================================
Private Function DetectStatutePaste(oRange As Range) As Boolean

    Dim oPara As Paragraph
    Dim oLastPara As Paragraph
    Set oLastPara = Nothing

    For Each oPara In oRange.Paragraphs
        Dim sParaText As String
        sParaText = oPara.Range.text

        Dim nStrip As Long
        nStrip = Len(sParaText)
        Do While nStrip > 0
            Dim nLast As Long
            nLast = AscW(Mid(sParaText, nStrip, 1))
            If nLast = 13 Or nLast = 11 Then
                nStrip = nStrip - 1
            Else
                Exit Do
            End If
        Loop
        sParaText = Trim(Left(sParaText, nStrip))

        If Len(sParaText) > 0 Then Set oLastPara = oPara
    Next oPara

    If oLastPara Is Nothing Then
        DetectStatutePaste = False
        Exit Function
    End If

    ' Check only the LAST outermost (...) block of the last paragraph
    ' for the section sign.  Checking the whole paragraph text would
    ' cause a false positive when a case-law passage contains an
    ' embedded statute cite (e.g. "\u00a7 1793.22") but the citation sentence
    ' itself is a case citation.
    Dim sLP As String
    sLP = RTrim(oLastPara.Range.text)

    ' Scan right-to-left for the outermost closing paren
    Dim nD As Integer
    nD = 0
    Dim iDS As Long
    iDS = -1
    Dim iLP As Long
    For iLP = Len(sLP) To 1 Step -1
        Dim cLP As String
        cLP = Mid(sLP, iLP, 1)
        If cLP = ")" Then nD = nD + 1
        If cLP = "(" Then
            nD = nD - 1
            If nD = 0 Then
                iDS = iLP
                Exit For
            End If
        End If
    Next iLP

    Dim sCheck As String
    If iDS = -1 Then
        ' No outermost paren found -- check full text (statute may lack parens)
        sCheck = sLP
    Else
        ' Check only the citation paren block for the section sign
        sCheck = Mid(sLP, iDS)
    End If

    ' A statute paste has a section sign in its citation block. But a CASE
    ' string-cite can also carry a section sign when it cross-references a
    ' statute, e.g. "(Evid. Code, § 452, subd. (d); Sosinsky v. Grant (1992)
    ' 6 Cal.App.4th 1548, 1564-1565.)". Such a block always contains a case
    ' name ("... v. ..."); statute citations never do. Excluding " v. " keeps
    ' those on the case path (running them through the statute pipeline throws
    ' a "Value out of range" error on the mismatched structure).
    DetectStatutePaste = (InStr(sCheck, ChrW(&HA7)) > 0) And _
                         (InStr(sCheck, " v. ") = 0)

End Function

'===========================================================
' Ensure one space between closing quote and citation paren
'===========================================================
'===========================================================
' Ensure one space between opposite-direction adjacent quotes
' and between closing/opening quotes and non-exempt characters.
' Exempt after closing: another closing quote, ], )
' Exempt before opening: another opening quote, [, (
'===========================================================
Private Sub EnsureSpacingAroundQuotes(oDoc As Document, oRange As Range)
    Set oRange = oDoc.Range(oRange.start, oRange.End)
    Dim sText As String
    sText = oRange.text
    Dim nLen As Long
    Dim nAfterS As Long
    nLen = Len(sText)
    If nLen < 2 Then Exit Sub
    Dim aIns() As Long
    ReDim aIns(nLen)
    Dim nIns As Long
    nIns = 0
    Dim i As Long
    For i = 1 To nLen
        Dim nCur As Long
        nCur = AscW(Mid(sText, i, 1))
        Dim bCloseQ As Boolean, bOpenQ As Boolean
        bCloseQ = (nCur = &H201D Or nCur = &H2019)
        bOpenQ = (nCur = &H201C Or nCur = &H2018)
        If bCloseQ And i < nLen Then
            Dim nAfter As Long
            nAfter = AscW(Mid(sText, i + 1, 1))
            Dim bAfterExempt As Boolean
            ' Exempt trailing punctuation as well as quotes/brackets/space:
            ' a comma, period, semicolon, colon, question/exclamation mark,
            ' or dash directly after a closing quote is correct typography
            ' ("rule",  "rule".  "rule"--) and was getting a spurious space
            ' inserted before it ("rule" ,).
            bAfterExempt = (nAfter = &H201D Or nAfter = &H2019 _
                            Or nAfter = 93 Or nAfter = 41 _
                            Or nAfter = 32 Or nAfter = 160 _
                            Or nAfter = 44 Or nAfter = 46 _
                            Or nAfter = 59 Or nAfter = 58 _
                            Or nAfter = 63 Or nAfter = 33 _
                            Or nAfter = 45 Or nAfter = &H2013 Or nAfter = &H2014)
            ' Apostrophe exemption: U+2019 followed by s/S then space
            ' (possessive: association's ...) -- do not insert space.
            ' Requires space after s/S so a word starting with s is not triggered.
            If nCur = &H2019 And (nAfter = 115 Or nAfter = 83) Then
                If i + 2 <= nLen Then
                    nAfterS = AscW(Mid(sText, i + 2, 1))
                Else
                    nAfterS = 0
                End If
                If nAfterS = 32 Or nAfterS = 0 Then bAfterExempt = True
            End If
            If Not bAfterExempt Then
                aIns(nIns) = oRange.start + i
                nIns = nIns + 1
            End If
        End If
        If bOpenQ And i > 1 Then
            Dim nBefore As Long
            nBefore = AscW(Mid(sText, i - 1, 1))
            Dim bBeforeExempt As Boolean
            ' A dash before an opening quote ("--\"rule\"") is also correct
            ' typography; only genuine word-glued quotes need the space.
            bBeforeExempt = (nBefore = &H201C Or nBefore = &H2018 _
                             Or nBefore = 91 Or nBefore = 40 _
                             Or nBefore = 32 Or nBefore = 160 _
                             Or nBefore = 45 Or nBefore = &H2013 Or nBefore = &H2014)
            If Not bBeforeExempt Then
                aIns(nIns) = oRange.start + i - 1
                nIns = nIns + 1
            End If
        End If
    Next i
    If nIns = 0 Then Exit Sub
    Dim j As Long
    Dim lPrev As Long
    lPrev = -1
    For j = nIns - 1 To 0 Step -1
        If aIns(j) <> lPrev Then
            Dim oIns As Range
            Set oIns = oDoc.Range(aIns(j), aIns(j))
            oIns.InsertAfter " "
            Set oIns = Nothing
            lPrev = aIns(j)
        End If
    Next j
    Set oRange = oDoc.Range(oRange.start, oRange.End)
End Sub

Private Sub InsertSpaceBeforeCitation(oDoc As Document, lStart As Long)

    Dim oRange As Range
    Set oRange = oDoc.Range(lStart, oDoc.content.End)

    Dim sText As String
    sText = oRange.text

    Dim i As Long
    For i = 1 To Len(sText) - 1
        If AscW(Mid(sText, i, 1)) = &H201D And _
           AscW(Mid(sText, i + 1, 1)) = 40 Then
            Dim lInsert As Long
            lInsert = oRange.start + i
            Dim oInsert As Range
            Set oInsert = oDoc.Range(lInsert, lInsert)
            oInsert.InsertAfter " "
            Set oInsert = Nothing
            Exit For
        End If
    Next i

    Set oRange = Nothing

End Sub

'===========================================================
' Apply Times New Roman 12pt, remove bold, preserve italic.
' FIX 3: Reuse a single Range object via SetRange instead of
'         creating a new Range object for every character,
'         eliminating thousands of COM allocations.
'===========================================================
Private Sub FormatQuotation(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim sText As String
    sText = oRange.text

    ' Allocate one Range object and move it via SetRange
    Dim oChar As Range
    Set oChar = oDoc.Range(oRange.start, oRange.start + 1)

    Dim i As Long
    For i = 1 To Len(sText)
        Dim lPos As Long
        lPos = oRange.start + i - 1
        oChar.SetRange lPos, lPos + 1

        Dim bItalic As Boolean
        bItalic = (oChar.Font.Italic = True)

        oChar.Font.Name = "Times New Roman"
        oChar.Font.Size = 12
        oChar.Font.Bold = False
        ' Clear expanded/condensed character spacing and horizontal scaling that
        ' Lexis/Westlaw text can carry, so no character is left with a phantom
        ' extra-space effect. (Font.Position is left alone to preserve any
        ' intentional super/subscript.)
        oChar.Font.Spacing = 0
        oChar.Font.Scaling = 100
        If bItalic Then oChar.Font.Italic = True

        Set oChar = Nothing
        Set oChar = oDoc.Range(oRange.start, oRange.start + 1)
    Next i

    Set oChar = Nothing

End Sub

'===========================================================
' Italicize case names in the citation sentence.
'===========================================================
Private Sub ItalicizeCaseNames(oRange As Range, oDoc As Document)

    Dim sText As String
    sText = oRange.text

    Dim sTrimmed As String
    sTrimmed = RTrim(sText)

    Dim i As Long
    Dim nDepth As Integer
    nDepth = 0
    Dim nCitationStart As Long
    nCitationStart = -1

    For i = Len(sTrimmed) To 1 Step -1
        Dim c As String
        c = Mid(sTrimmed, i, 1)
        If c = ")" Then nDepth = nDepth + 1
        If c = "(" Then
            nDepth = nDepth - 1
            If nDepth = 0 Then
                nCitationStart = i
                Exit For
            End If
        End If
    Next i

    If nCitationStart = -1 Then Exit Sub

    Dim lCiteOpen As Long
    lCiteOpen = oRange.start + nCitationStart - 1

    Dim lYearOpen As Long
    lYearOpen = -1

    Dim lWalk As Long
    lWalk = lCiteOpen + 1

    Do While lWalk < oRange.End - 4
        Dim oW As Range
        Set oW = oDoc.Range(lWalk, lWalk + 1)
        Dim nW As Long
        nW = AscW(oW.text)
        Set oW = Nothing

        If nW = 40 Then
            Dim oYear As Range
            Set oYear = oDoc.Range(lWalk + 1, lWalk + 5)
            Dim sYear As String
            sYear = oYear.text
            Set oYear = Nothing

            Dim bAllDigits As Boolean
            bAllDigits = True
            Dim k As Integer
            For k = 1 To 4
                Dim nD As Long
                nD = AscW(Mid(sYear, k, 1))
                If nD < 48 Or nD > 57 Then
                    bAllDigits = False
                    Exit For
                End If
            Next k

            If bAllDigits Then
                lYearOpen = lWalk
                Exit Do
            End If
        End If

        lWalk = lWalk + 1
    Loop

    If lYearOpen = -1 Then Exit Sub

    Dim lItalicStart As Long
    Dim lItalicEnd As Long
    lItalicStart = lCiteOpen + 1
    lItalicEnd = lYearOpen

    Do While lItalicStart < lItalicEnd
        Dim oTrimL As Range
        Set oTrimL = oDoc.Range(lItalicStart, lItalicStart + 1)
        If oTrimL.text = " " Then
            lItalicStart = lItalicStart + 1
        Else
            Set oTrimL = Nothing
            Exit Do
        End If
        Set oTrimL = Nothing
    Loop

    Do While lItalicEnd > lItalicStart
        Dim oTrimR As Range
        Set oTrimR = oDoc.Range(lItalicEnd - 1, lItalicEnd)
        If oTrimR.text = " " Then
            lItalicEnd = lItalicEnd - 1
        Else
            Set oTrimR = Nothing
            Exit Do
        End If
        Set oTrimR = Nothing
    Loop

    If lItalicStart >= lItalicEnd Then Exit Sub

    Dim oItalic As Range
    Set oItalic = oDoc.Range(lItalicStart, lItalicEnd)
    oItalic.Font.Italic = True
    Set oItalic = Nothing

End Sub

'===========================================================
' Italicize the case name in a bare (non-parenthesized) citation.
' Used after RestructureAsTextual where the citation has no outer
' parens. Scans forward from lCiteStart for the first (YYYY) and
' italicizes all text from lCiteStart up to (but not including) it.
'===========================================================
Private Sub ItalicizeCaseNameBare(oDoc As Document, _
                                   ByVal lCiteStart As Long, _
                                   ByVal lCiteEnd As Long)

    If lCiteEnd <= lCiteStart Then Exit Sub

    Dim oCitRange As Range
    Set oCitRange = oDoc.Range(lCiteStart, lCiteEnd)
    Dim sCit As String
    sCit = oCitRange.text
    Set oCitRange = Nothing

    ' Scan forward for (YYYY) -- opening paren followed by exactly 4 digits
    Dim lYearOpen As Long
    lYearOpen = -1
    Dim lWalk As Long
    For lWalk = 1 To Len(sCit) - 4
        If AscW(Mid(sCit, lWalk, 1)) = 40 Then  ' "("
            Dim bAllDig As Boolean
            bAllDig = True
            Dim kD As Integer
            For kD = 1 To 4
                Dim nD As Long
                nD = AscW(Mid(sCit, lWalk + kD, 1))
                If nD < 48 Or nD > 57 Then
                    bAllDig = False
                    Exit For
                End If
            Next kD
            If bAllDig Then
                lYearOpen = lWalk
                Exit For
            End If
        End If
    Next lWalk

    If lYearOpen = -1 Then Exit Sub

    ' Italic region: lCiteStart to just before (YYYY), trimming spaces
    Dim lItalicStart As Long
    Dim lItalicEnd As Long
    lItalicStart = lCiteStart
    lItalicEnd = lCiteStart + lYearOpen - 1    ' position of the "("

    ' Trim trailing spaces
    Do While lItalicEnd > lItalicStart
        Dim oTrim As Range
        Set oTrim = oDoc.Range(lItalicEnd - 1, lItalicEnd)
        If oTrim.text = " " Then
            lItalicEnd = lItalicEnd - 1
        Else
            Set oTrim = Nothing
            Exit Do
        End If
        Set oTrim = Nothing
    Loop

    If lItalicEnd <= lItalicStart Then Exit Sub

    Dim oItalic As Range
    Set oItalic = oDoc.Range(lItalicStart, lItalicEnd)
    oItalic.Font.Italic = True
    Set oItalic = Nothing

End Sub

'===========================================================
' Reset paragraph spacing and reapply Normal style
'===========================================================
Private Sub FixParagraphSpacing(oRange As Range)

    Dim oPara As Paragraph
    For Each oPara In oRange.Paragraphs
        With oPara.Format
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 0
            .SpaceAfter = 8
        End With
        oPara.Style = ActiveDocument.Styles("Normal")
    Next oPara

End Sub

'===========================================================
' Wrap Westlaw citation paragraph in parentheses if not
' already wrapped.
'===========================================================
Private Sub WrapWestlawCitation(oDoc As Document, oRange As Range)

    Dim oPara As Paragraph
    Dim oLastPara As Paragraph
    Dim sText As String

    Set oLastPara = Nothing
    For Each oPara In oRange.Paragraphs
        sText = Trim(oPara.Range.text)
        If Len(sText) > 1 Then
            Set oLastPara = oPara
        End If
    Next oPara

    If oLastPara Is Nothing Then Exit Sub

    Dim sParaText As String
    sParaText = oLastPara.Range.text
    Do While Len(sParaText) > 0 And _
             (AscW(Right(sParaText, 1)) = 13 Or AscW(Right(sParaText, 1)) = 11)
        sParaText = Left(sParaText, Len(sParaText) - 1)
    Loop
    sParaText = Trim(sParaText)

    If Left(sParaText, 1) = "(" Then Exit Sub

    ' Guard: if the last outermost (...) in sParaText already contains a
    ' (YYYY) year paren, this is a Lexis+ citation already wrapped -- do not re-wrap.
    Dim iWW As Long, nWWDepth As Integer, nWWCiteStart As Long
    nWWDepth = 0
    nWWCiteStart = -1
    For iWW = Len(sParaText) To 1 Step -1
        Dim cWW As String
        cWW = Mid(sParaText, iWW, 1)
        If cWW = ")" Then nWWDepth = nWWDepth + 1
        If cWW = "(" Then
            nWWDepth = nWWDepth - 1
            If nWWDepth = 0 Then nWWCiteStart = iWW: Exit For
        End If
    Next iWW
    If nWWCiteStart > 0 Then
        Dim sWWInner As String
        sWWInner = Mid(sParaText, nWWCiteStart + 1, Len(sParaText) - nWWCiteStart - 1)
        Dim iWWY As Long
        For iWWY = 1 To Len(sWWInner) - 5
            If Mid(sWWInner, iWWY, 1) = "(" Then
                Dim sWWFour As String
                sWWFour = Mid(sWWInner, iWWY + 1, 4)
                Dim bWWDig As Boolean, kWW As Integer
                bWWDig = True
                For kWW = 1 To 4
                    If AscW(Mid(sWWFour, kWW, 1)) < 48 Or _
                       AscW(Mid(sWWFour, kWW, 1)) > 57 Then
                        bWWDig = False
                        Exit For
                    End If
                Next kWW
                If bWWDig And Mid(sWWInner, iWWY + 5, 1) = ")" Then
                    Dim nWWYear As Long
                    nWWYear = CLng(sWWFour)
                    If nWWYear >= 1800 And nWWYear <= 2099 Then Exit Sub
                End If
            End If
        Next iWWY
    End If

    Dim lParaStart As Long
    Dim lParaEnd As Long
    lParaStart = oLastPara.Range.start
    ' The Trim above stripped leading spaces from the MEASURED string but not
    ' from the range anchor; advance the anchor to match, or the ".)" lands
    ' short by the number of leading spaces and "(" goes before them.
    Dim sWWRaw As String
    sWWRaw = oLastPara.Range.text
    Dim nWWLead As Long
    nWWLead = 0
    Do While nWWLead < Len(sWWRaw)
        If AscW(Mid(sWWRaw, nWWLead + 1, 1)) = 32 Then
            nWWLead = nWWLead + 1
        Else
            Exit Do
        End If
    Loop
    lParaStart = lParaStart + nWWLead
    lParaEnd = lParaStart + Len(sParaText)

    Dim oEnd As Range
    Set oEnd = oDoc.Range(lParaEnd, lParaEnd)
    oEnd.InsertAfter ".)"
    Set oEnd = Nothing

    Dim oStart As Range
    Set oStart = oDoc.Range(lParaStart, lParaStart)
    oStart.InsertBefore "("
    Set oStart = Nothing

End Sub

'===========================================================
' Remove trailing paragraph mark at end of pasted range.
'===========================================================
Private Sub RemoveTrailingParagraphMark(oDoc As Document, oRange As Range)

    Dim lEnd As Long
    lEnd = oRange.End

    If lEnd <= oRange.start Then Exit Sub

    Dim oLast As Range
    Set oLast = oDoc.Range(lEnd - 1, lEnd)

    Dim nCode As Long
    nCode = AscW(oLast.text)

    If nCode = 13 Or nCode = 11 Then
        oLast.Delete
    End If

    Set oLast = Nothing

End Sub

'===========================================================
' Fix first letter capitalization per California Style Manual.
'===========================================================
Private Sub FixFirstLetterCapitalization(oDoc As Document, _
                                          lStart As Long, _
                                          lQuoteEnd As Long)

    Dim lFirstLetter As Long
    Dim sFirst As String
    Dim nFirst As Long
    lFirstLetter = lStart

    Do While lFirstLetter < lQuoteEnd
        Dim oFL As Range
        Set oFL = oDoc.Range(lFirstLetter, lFirstLetter + 1)
        nFirst = AscW(oFL.text)
        Set oFL = Nothing
        If nFirst = &H201C Or nFirst = &H2018 Or _
           nFirst = &H201D Or nFirst = &H2019 Then
            lFirstLetter = lFirstLetter + 1
        Else
            Exit Do
        End If
    Loop

    If lFirstLetter >= lQuoteEnd Then Exit Sub

    Dim oFirstChar As Range
    Set oFirstChar = oDoc.Range(lFirstLetter, lFirstLetter + 1)
    sFirst = oFirstChar.text
    nFirst = AscW(sFirst)
    Set oFirstChar = Nothing

    Dim bIsUpper As Boolean
    Dim bIsLower As Boolean
    bIsUpper = (nFirst >= 65 And nFirst <= 90)
    bIsLower = (nFirst >= 97 And nFirst <= 122)
    If Not bIsUpper And Not bIsLower Then Exit Sub

    Dim lLook As Long
    lLook = lStart - 1
    Dim sPreceding As String
    sPreceding = ""
    Dim bParagraphBoundary As Boolean
    bParagraphBoundary = False

    Do While lLook >= 0
        Dim oPre As Range
        Set oPre = oDoc.Range(lLook, lLook + 1)
        Dim nPre As Long
        nPre = AscW(oPre.text)
        Set oPre = Nothing
        If nPre = 13 Or nPre = 11 Then
            ' Paragraph mark ? treat as sentence start regardless of what
            ' preceded it in the previous paragraph
            bParagraphBoundary = True
            Exit Do
        ElseIf nPre = 32 Then
            lLook = lLook - 1
        ElseIf nPre = &H201C Or nPre = &H2018 Then
            lLook = lLook - 1
        ElseIf nPre = 41 Or nPre = &H201D Or nPre = &H2019 Or nPre = 93 Then
            lLook = lLook - 1
        Else
            sPreceding = ChrW(nPre)
            Exit Do
        End If
    Loop

    Dim nPreceding As Long
    If Len(sPreceding) > 0 Then nPreceding = AscW(sPreceding) Else nPreceding = 0

    Dim bSentenceStart As Boolean
    bSentenceStart = bParagraphBoundary Or _
                     (nPreceding = 46 Or nPreceding = 63 Or _
                      nPreceding = 33 Or nPreceding = 0)

    Dim bMidSentence As Boolean
    bMidSentence = Not bSentenceStart

    Dim lLastLetter As Long
    lLastLetter = lQuoteEnd - 1

    Dim nLast As Long
    nLast = 0

    Do While lLastLetter >= lStart
        Dim oLL As Range
        Set oLL = oDoc.Range(lLastLetter, lLastLetter + 1)
        Dim nLL As Long
        nLL = AscW(oLL.text)
        Set oLL = Nothing
        If nLL = &H201D Or nLL = &H2019 Or nLL = 32 Then
            lLastLetter = lLastLetter - 1
        Else
            nLast = nLL
            Exit Do
        End If
    Loop

    Dim bFullSentence As Boolean
    bFullSentence = (nLast = 46 Or nLast = 63 Or nLast = 33 Or nLast = 93)

    If bSentenceStart And bIsLower Then
        Dim sUpper As String
        sUpper = UCase(sFirst)
        Dim oBracket As Range
        Set oBracket = oDoc.Range(lFirstLetter, lFirstLetter + 1)
        oBracket.text = "[" & sUpper & "]"
        Set oBracket = Nothing

    ElseIf bMidSentence And bIsUpper And Not bFullSentence Then
        Dim sLower As String
        sLower = LCase(sFirst)
        Dim oBracket2 As Range
        Set oBracket2 = oDoc.Range(lFirstLetter, lFirstLetter + 1)
        oBracket2.text = "[" & sLower & "]"
        Set oBracket2 = Nothing
    End If

End Sub

'===========================================================
' Remove spurious spaces between apostrophe and s.
' FIX 2: Added iteration cap to prevent infinite loop.
'===========================================================
Private Sub RemoveSpacesAfterApostrophes(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Const MAX_ITER As Long = 500

    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim sText As String
        sText = oRange.text

        Dim bFound As Boolean
        bFound = False

        Dim i As Long
        For i = 1 To Len(sText) - 2
            Dim nThis As Long
            nThis = AscW(Mid(sText, i, 1))
            Dim nNext As Long
            nNext = AscW(Mid(sText, i + 1, 1))
            Dim nAfter As Long
            nAfter = AscW(Mid(sText, i + 2, 1))

            Dim nAfterS As Long
            If i + 3 <= Len(sText) Then
                nAfterS = AscW(Mid(sText, i + 3, 1))
            Else
                nAfterS = 0
            End If

            If (nThis = &H2019 Or nThis = 39) And _
               nNext = 32 And _
               (nAfter = 115 Or nAfter = 83) And _
               nAfterS = 32 Then

                Dim lDelPos As Long
                lDelPos = oRange.start + i
                Dim oDel As Range
                Set oDel = oDoc.Range(lDelPos, lDelPos + 1)
                oDel.Delete
                Set oDel = Nothing
                bFound = True
                Exit For
            End If
        Next i

        If Not bFound Then Exit Do
        Set oRange = oDoc.Range(oRange.start, oRange.End)
    Loop

End Sub

'===========================================================
' Wrap passage in outer double smart quotes
'===========================================================
Private Sub WrapInDoubleQuotes(oRange As Range)

    Dim oDoc As Document
    Dim lClosePos As Long
    Dim lOpenPos As Long
    Dim oEnd As Range
    Dim oStart As Range

    Set oDoc = ActiveDocument

    ' Insert closing quote first so its position isn't affected by the
    ' opening insertion that follows.
    lClosePos = oRange.End
    Set oEnd = oDoc.Range(lClosePos, lClosePos)
    oEnd.InsertAfter ChrW(&H201D)
    Set oEnd = oDoc.Range(lClosePos, lClosePos + 1)
    CleanQuoteFont oEnd
    Set oEnd = Nothing

    ' Insert opening quote and force clean Times New Roman -- do NOT inherit the
    ' passage's first-character font. When a Lexis headnote is stripped from the
    ' start of the copied text, that first character can carry the headnote's
    ' font AND expanded character spacing; inheriting it put the opening quote in
    ' the wrong font with a phantom-space effect (looks like an extra space that
    ' isn't there). Setting the font outright also defeats paragraph-mark
    ' inheritance at paragraph start.
    lOpenPos = oRange.start
    Set oStart = oDoc.Range(lOpenPos, lOpenPos)
    oStart.InsertBefore ChrW(&H201C)
    Set oStart = oDoc.Range(lOpenPos, lOpenPos + 1)
    CleanQuoteFont oStart
    Set oStart = Nothing

End Sub

' Normalize a just-inserted quote mark to clean Times New Roman 12pt with no
' inherited bold/italic and, crucially, no inherited advanced spacing
' (Spacing/Scaling/Position). A leftover font or expanded character spacing from
' stripped Lexis headnotes would otherwise leave the quote in the wrong font or
' produce a phantom-space effect next to it.
Private Sub CleanQuoteFont(ByVal oQuote As Range)
    With oQuote.Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
        .Italic = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
    End With
End Sub

'===========================================================
' Remove a duplicate closing outer double-quote mark.
' Scans right-to-left to find the outermost closing double
' quote (U+201D)   the one WrapInDoubleQuotes placed at the
' passage end. If the character immediately before it is
' also U+201D, that preceding one is the duplicate and is
' deleted, restoring the correct single closing quote.
'===========================================================
Private Sub RemoveDuplicateClosingQuote(oDoc As Document, oRange As Range)

    Dim sText As String
    sText = oRange.text

    ' RTrim so trailing paragraph marks don't push us past the quotes
    Dim sTrimmed As String
    sTrimmed = RTrim(sText)

    ' Scan right-to-left for the first U+201D   this is the wrap close quote
    Dim i As Long
    For i = Len(sTrimmed) To 1 Step -1
        If AscW(Mid(sTrimmed, i, 1)) = &H201D Then
            ' Found the outermost closing double quote.
            ' If the character immediately before it is also U+201D, delete that one.
            If i > 1 Then
                If AscW(Mid(sTrimmed, i - 1, 1)) = &H201D Then
                    ' Duplicate found at position i-1   delete it.
                    ' Navigate using Duplicate + MoveEnd + Collapse to avoid
                    ' computed oDoc.Range positions.
                    Dim oDup As Range
                    Set oDup = oRange.Duplicate
                    oDup.Collapse wdCollapseStart
                    oDup.MoveEnd wdCharacter, i - 1  ' End at the duplicate
                    oDup.Collapse wdCollapseEnd       ' zero-length at duplicate
                    oDup.MoveStart wdCharacter, -1    ' back 1: covers the duplicate
                    oDup.Delete
                    Set oDup = Nothing
                End If
            End If
            Exit For
        End If
    Next i

End Sub

'===========================================================
' Remove a duplicate opening outer double-quote (U+201C).
'===========================================================
Private Sub RemoveDuplicateOpeningQuote(oDoc As Document, oRange As Range)
    Dim sText As String
    sText = oRange.text
    Dim i As Long
    For i = 1 To Len(sText)
        If AscW(Mid(sText, i, 1)) = &H201C Then
            If i < Len(sText) Then
                If AscW(Mid(sText, i + 1, 1)) = &H201C Then
                    Dim oDup As Range
                    Set oDup = oRange.Duplicate
                    oDup.Collapse wdCollapseStart
                    oDup.MoveEnd wdCharacter, i + 1
                    oDup.Collapse wdCollapseEnd
                    oDup.MoveStart wdCharacter, -1
                    oDup.Delete
                    Set oDup = Nothing
                End If
            End If
            Exit For
        End If
    Next i
End Sub

'===========================================================
' Remove a duplicate opening single-quote (U+2018) at start.
'===========================================================
Private Sub RemoveDuplicateOpenSingle(oDoc As Document, oRange As Range)
    Dim sText As String
    sText = oRange.text
    Dim i As Long
    For i = 1 To Len(sText)
        Dim nC As Long
        nC = AscW(Mid(sText, i, 1))
        If nC = &H2018 Then
            If i < Len(sText) Then
                If AscW(Mid(sText, i + 1, 1)) = &H2018 Then
                    Dim oDup As Range
                    Set oDup = oRange.Duplicate
                    oDup.Collapse wdCollapseStart
                    oDup.MoveEnd wdCharacter, i + 1
                    oDup.Collapse wdCollapseEnd
                    oDup.MoveStart wdCharacter, -1
                    oDup.Delete
                    Set oDup = Nothing
                End If
            End If
            Exit For
        ElseIf nC <> &H201C Then
            Exit For
        End If
    Next i
End Sub

'===========================================================
' Remove a duplicate closing single-quote (U+2019) at end.
'===========================================================
Private Sub RemoveDuplicateCloseSingle(oDoc As Document, oRange As Range)
    Dim sText As String
    sText = RTrim(oRange.text)
    Dim i As Long
    For i = Len(sText) To 1 Step -1
        Dim nC As Long
        nC = AscW(Mid(sText, i, 1))
        If nC = &H2019 Then
            If i > 1 Then
                If AscW(Mid(sText, i - 1, 1)) = &H2019 Then
                    Dim oDup As Range
                    Set oDup = oRange.Duplicate
                    oDup.Collapse wdCollapseStart
                    oDup.MoveEnd wdCharacter, i - 1
                    oDup.Collapse wdCollapseEnd
                    oDup.MoveStart wdCharacter, -1
                    oDup.Delete
                    Set oDup = Nothing
                End If
            End If
            Exit For
        ElseIf nC <> &H201D Then
            Exit For
        End If
    Next i
End Sub


'===========================================================
' Fix the font of characters immediately surrounding the
' pasted block that Word left in Aptos Body.
'
' Walks left from lStart and right from lEnd, resetting
' non-Times-New-Roman characters to Times New Roman 12pt.
' Stops at a paragraph boundary so other paragraphs are
' never touched.
'===========================================================
' Fix Aptos Body font on characters surrounding the paste.
' Finds the paragraph(s) containing lStart and lEnd and
' resets any non-Times-New-Roman character to Times New
' Roman 12pt, preserving italic.
'===========================================================
Private Sub FixSurroundingFont(oDoc As Document, _
                                ByVal lStart As Long, _
                                ByVal lEnd As Long)

    Dim oStartPara As Paragraph
    Dim oEndPara   As Paragraph
    Dim oFix       As Range
    Dim oChar      As Range
    Dim lPos       As Long
    Dim lParaStart As Long
    Dim lParaEnd   As Long
    Dim bItalic    As Boolean

    ' Find the paragraph containing lStart
    Set oFix = oDoc.Range(lStart, lStart)
    Set oStartPara = oFix.Paragraphs(1)
    lParaStart = oStartPara.Range.start
    Set oStartPara = Nothing

    ' Find the paragraph containing lEnd
    Set oFix = oDoc.Range(lEnd, lEnd)
    Set oEndPara = oFix.Paragraphs(1)
    lParaEnd = oEndPara.Range.End
    Set oEndPara = Nothing
    Set oFix = Nothing

    ' Fix every non-Times-New-Roman character in the paragraph(s)
    For lPos = lParaStart To lParaEnd - 1
        Set oChar = oDoc.Range(lPos, lPos + 1)
        If oChar.Font.Name <> "Times New Roman" Then
            bItalic = (oChar.Font.Italic = True)
            oChar.Font.Name = "Times New Roman"
            oChar.Font.Size = 12
            If bItalic Then oChar.Font.Italic = True
        End If
        Set oChar = Nothing
    Next lPos

End Sub

'===========================================================
' Helper: word character check (letters and digits only
' excludes Unicode punctuation such as smart quotes)
'===========================================================
Private Function IsWordChar(c As String) As Boolean
    If Len(c) = 0 Then IsWordChar = False: Exit Function
    Dim n As Long
    n = AscW(c)
    IsWordChar = (n >= 65 And n <= 90) Or _
                 (n >= 97 And n <= 122) Or _
                 (n >= 48 And n <= 57)
End Function

'===========================================================
' Helper: is this character code a curly quote?
'===========================================================
Private Function IsQuoteChar(nCode As Long) As Boolean
    Select Case nCode
        Case &H2018, &H2019, &H201C, &H201D
            IsQuoteChar = True
        Case Else
            IsQuoteChar = False
    End Select
End Function

'===========================================================
' Helper: is this character code whitespace only (space, tab,
' paragraph mark, soft return)? Used for open-quote detection
' where a preceding close-quote must NOT trigger "open".
'===========================================================
Private Function IsSpacePara(nCode As Long) As Boolean
    Select Case nCode
        Case 32, 9, 13, 11
            IsSpacePara = True
        Case Else
            IsSpacePara = False
    End Select
End Function

'===========================================================
' Helper: is this character code a space or any quote mark
' (curly or straight)? Used to determine open/close context.
'===========================================================
Private Function IsSpaceOrQuote(nCode As Long) As Boolean
    Select Case nCode
        Case 32, 9, 13, 11  ' space, tab, paragraph mark, soft return
            IsSpaceOrQuote = True
        Case 34, 39  ' straight double, straight single
            IsSpaceOrQuote = True
        Case &H2018, &H2019, &H201C, &H201D  ' curly quotes
            IsSpaceOrQuote = True
        Case Else
            IsSpaceOrQuote = False
    End Select
End Function

'===========================================================
' Helper: does this string look like a citation volume number?
' Must be one or more digits only.
'===========================================================
Private Function IsCiteVolumeNumber(s As String) As Boolean
    Dim sTrimmed As String
    sTrimmed = Trim(s)
    If Len(sTrimmed) = 0 Then
        IsCiteVolumeNumber = False
        Exit Function
    End If
    Dim i As Integer
    For i = 1 To Len(sTrimmed)
        Dim n As Long
        n = AscW(Mid(sTrimmed, i, 1))
        If n < 48 Or n > 57 Then
            IsCiteVolumeNumber = False
            Exit Function
        End If
    Next i
    IsCiteVolumeNumber = True
End Function

'===========================================================
' Convert straight apostrophes to curly right singles.
' FIX 3: Reuse a single Range object via SetRange instead of
'         creating a new Range object for every character,
'         eliminating thousands of COM allocations.
'===========================================================
Private Sub CurlyApostrophes(oRange As Range)
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    Dim lPos As Long
    Dim lStart As Long
    Dim lEnd As Long
    lStart = oRange.start
    lEnd = oRange.End

    ' Read text once for fast scanning
    Dim sText As String
    sText = oRange.text

    ' Allocate one Range object; move it only when a match is found
    Dim oChar As Range
    Set oChar = oDoc.Range(lStart, lStart + 1)

    Dim nCode As Long
    Dim nPrev As Long
    Dim nNext As Long
    Dim sIdx As Long

    For lPos = lEnd - 1 To lStart Step -1
        sIdx = lPos - lStart + 1
        nCode = AscW(Mid(sText, sIdx, 1))

        Select Case nCode

            Case 39  ' straight single quote
                ' Read both neighbours from the snapshot.
                If sIdx > 1 Then
                    nPrev = AscW(Mid(sText, sIdx - 1, 1))
                Else
                    nPrev = 0
                End If
                If sIdx < Len(sText) Then
                    nNext = AscW(Mid(sText, sIdx + 1, 1))
                Else
                    nNext = 0
                End If

                ' Apostrophe condition 1: flanked by word characters on both sides
                '   (e.g. don't, it's, court's). Followed-by-letter alone is NOT
                '   enough ? that would wrongly flag opening quote marks like 'to.
                Dim bWordFlanked As Boolean
                bWordFlanked = (((nPrev >= 65 And nPrev <= 90) Or _
                                 (nPrev >= 97 And nPrev <= 122) Or _
                                 (nPrev >= 48 And nPrev <= 57)) And _
                                ((nNext >= 65 And nNext <= 90) Or _
                                 (nNext >= 97 And nNext <= 122)))

                ' Apostrophe condition 2: preceded by s/S, followed by a space,
                '   and the next non-space character after that space is a letter
                '   (e.g. "parties' agreement"). Counter-example: "United States'"
                '   with no following letter = not apostrophe.
                Dim bPossessiveS As Boolean
                bPossessiveS = False
                If (nPrev = 115 Or nPrev = 83) And nNext = 32 Then
                    Dim lScan As Long
                    lScan = sIdx + 2
                    Do While lScan <= Len(sText)
                        Dim nScan As Long
                        nScan = AscW(Mid(sText, lScan, 1))
                        If nScan = 32 Then
                            lScan = lScan + 1
                        ElseIf (nScan >= 65 And nScan <= 90) Or _
                               (nScan >= 97 And nScan <= 122) Then
                            bPossessiveS = True
                            Exit Do
                        Else
                            Exit Do
                        End If
                    Loop
                End If

                If bWordFlanked Or bPossessiveS Then
                    ' apostrophe ? do not convert
                Else
                    ' Open when preceded by a space or quote mark.
                    ' Close when followed by a space or quote mark.
                    oChar.SetRange lPos, lPos + 1
                    If IsSpaceOrQuote(nPrev) Or nPrev = 0 Then
                        oChar.text = ChrW(&H2018)  ' open single '
                    ElseIf IsSpaceOrQuote(nNext) Or nNext = 0 Then
                        oChar.text = ChrW(&H2019)  ' close single '
                    Else
                        oChar.text = ChrW(&H2019)  ' default close single '
                    End If
                End If

            Case 34  ' straight double quote
                ' Open when preceded by a space or quote mark.
                ' Close when followed by a space or quote mark.
                ' Preceding context takes priority when both sides are ambiguous.
                If sIdx > 1 Then
                    nPrev = AscW(Mid(sText, sIdx - 1, 1))
                Else
                    nPrev = 0
                End If
                If sIdx < Len(sText) Then
                    nNext = AscW(Mid(sText, sIdx + 1, 1))
                Else
                    nNext = 0
                End If

                oChar.SetRange lPos, lPos + 1
                ' OPEN only when preceded by whitespace or an OPEN quote mark.
                ' A " preceded by a close quote (e.g. the " in '"') is CLOSE.
                If IsSpacePara(nPrev) Or nPrev = 0 Or _
                   nPrev = &H201C Or nPrev = &H2018 Then
                    oChar.text = ChrW(&H201C)  ' open double "
                Else
                    oChar.text = ChrW(&H201D)  ' close double "
                End If

        End Select
    Next lPos

    Set oChar = Nothing
End Sub

'===========================================================
' Extract leading subdivision marker(s) from the start of
' the pasted content.
'
' Scans forward from lStart skipping whitespace/para marks,
' then collects consecutive "(x)" tokens where x is a single
' alphanumeric character. Removes all consumed characters and
' trailing whitespace before the actual passage text.
' Returns the combined subdivision string, e.g. "(a)(1)",
' or "" if no leading markers are found.
'===========================================================
Private Function ExtractLeadingSubdivision(oDoc As Document, _
                                            ByVal lStart As Long, _
                                            ByVal lRangeEnd As Long) As String

    Dim sResult As String
    Dim bFoundAny As Boolean
    Dim lPos As Long
    Dim bCrossedPara As Boolean
    Dim oWS As Range
    Dim nWS As Long
    Dim oO As Range
    Dim bO As Boolean
    Dim oM As Range
    Dim nM As Long
    Dim oC As Range
    Dim bC As Boolean

    sResult = ""
    bFoundAny = False
    lPos = lStart

    Do
        ' Track whether we cross a paragraph mark between tokens.
        ' If we have already found a token and the next one is on a new
        ' paragraph, stop -- paragraph-start markers belong in the passage.
        bCrossedPara = False
        Do While lPos < lRangeEnd
            Set oWS = oDoc.Range(lPos, lPos + 1)
            nWS = AscW(oWS.text)
            Set oWS = Nothing
            If nWS = 13 Or nWS = 11 Then
                bCrossedPara = True
                lPos = lPos + 1
            ElseIf nWS = 32 Or nWS = 9 Then
                lPos = lPos + 1
            Else
                Exit Do
            End If
        Loop

        ' Crossed a paragraph boundary after finding a token -- stop.
        If bCrossedPara And bFoundAny Then Exit Do

        If lPos > lRangeEnd - 3 Then Exit Do

        Set oO = oDoc.Range(lPos, lPos + 1)
        bO = (oO.text = "(")
        Set oO = Nothing
        If Not bO Then Exit Do

        Set oM = oDoc.Range(lPos + 1, lPos + 2)
        nM = AscW(oM.text)
        Set oM = Nothing
        If Not ((nM >= 65 And nM <= 90) Or _
                (nM >= 97 And nM <= 122) Or _
                (nM >= 48 And nM <= 57)) Then Exit Do

        Set oC = oDoc.Range(lPos + 2, lPos + 3)
        bC = (oC.text = ")")
        Set oC = Nothing
        If Not bC Then Exit Do

        sResult = sResult & "(" & ChrW(nM) & ")"
        bFoundAny = True
        lPos = lPos + 3
    Loop

    If Not bFoundAny Then
        ExtractLeadingSubdivision = ""
        Exit Function
    End If

    Do While lPos < lRangeEnd
        Dim oTr As Range
        Set oTr = oDoc.Range(lPos, lPos + 1)
        Dim nTr As Long
        nTr = AscW(oTr.text)
        Set oTr = Nothing
        If nTr = 32 Or nTr = 9 Or nTr = 13 Or nTr = 11 Then
            lPos = lPos + 1
        Else
            Exit Do
        End If
    Loop

    If lPos > lStart Then
        Dim oDel As Range
        Set oDel = oDoc.Range(lStart, lPos)
        oDel.Delete
        Set oDel = Nothing
    End If

    ExtractLeadingSubdivision = sResult

End Function

'===========================================================
' Insert ", subd. (x)(y)" into the citation sentence.
'
' Searches from lQuoteEnd to lRangeEnd for the period that
' sits directly inside the outermost citation parenthesis
' using paren-depth counting (right-to-left).
'
' Example:
'   (Corp. Code,   14007.)  + "(a)"
'   -> (Corp. Code,   14007, subd. (a).)
'===========================================================
Private Sub InsertSubdivisionIntoCitation(oDoc As Document, _
                                           lQuoteEnd As Long, _
                                           lRangeEnd As Long, _
                                           sSubdivision As String)

    Dim oSearch As Range
    Set oSearch = oDoc.Range(lQuoteEnd, lRangeEnd)

    Dim sText As String
    sText = oSearch.text

    Dim sTrimmed As String
    sTrimmed = RTrim(sText)

    Dim i As Long
    Dim nDepth As Integer
    nDepth = 0
    Dim lDotIdx As Long
    lDotIdx = 0

    For i = Len(sTrimmed) To 1 Step -1
        Dim c As String
        c = Mid(sTrimmed, i, 1)
        Select Case c
            Case ")": nDepth = nDepth + 1
            Case "(": nDepth = nDepth - 1
            Case "."
                If nDepth = 1 Then
                    lDotIdx = i
                    Exit For
                End If
        End Select
    Next i

    If lDotIdx = 0 Then
        Set oSearch = Nothing
        Exit Sub
    End If

    Dim lDotPos As Long
    lDotPos = oSearch.start + lDotIdx - 1

    Dim oInsert As Range
    Set oInsert = oDoc.Range(lDotPos, lDotPos)
    oInsert.InsertBefore ", subd. " & sSubdivision
    Set oInsert = Nothing

    Set oSearch = Nothing

End Sub

'===========================================================
' Remove Lexis+ publisher attribution parenthetical from
' statute citation text. Searches for "(Deering" and
' "(LexisNexis" and removes each match including the
' matching close paren and any immediately preceding space.
'
' Must be called on a range restricted to the citation
' portion so it cannot accidentally touch the passage.

'===========================================================
' Normalize Lexis+ statute citations that arrive with missing
' punctuation or without wrapping parentheses.
'
' Handles:
'   "Cal Civ Code   1793.2"  ->  "(Civ. Code,   1793.2.)"
'     - strips "Cal " prefix
'     - adds period: "Civ" -> "Civ."
'     - adds comma before   if missing
'     - wraps in parens with terminal period if unwrapped
'   "(Com. Code,   xyz.)"  ->  "(Cal. U. Com. Code,   xyz.)"
'     - renames to California Uniform Commercial Code form
'
' Operates on the last non-empty paragraph (the citation sentence).
'===========================================================
Private Sub NormalizeLexisStatuteCitation(oDoc As Document, oRange As Range, _
                                           ByRef lNormEnd As Long)

    ' Find the last non-empty paragraph
    Dim oPara As Paragraph
    Dim oLastPara As Paragraph
    Set oLastPara = Nothing
    For Each oPara In oRange.Paragraphs
        If Len(Trim(oPara.Range.text)) > 1 Then
            Set oLastPara = oPara
        End If
    Next oPara
    If oLastPara Is Nothing Then Exit Sub

    Dim sCit As String
    sCit = oLastPara.Range.text
    ' Strip trailing paragraph marks
    Do While Len(sCit) > 0
        Dim nT As Long
        nT = AscW(Right(sCit, 1))
        If nT = 13 Or nT = 11 Then
            sCit = Left(sCit, Len(sCit) - 1)
        Else
            Exit Do
        End If
    Loop
    sCit = Trim(sCit)
    If Len(sCit) = 0 Then Exit Sub

    Dim sNew As String
    sNew = sCit

    ' Fix 1: "Com. Code" -> "Cal. U. Com. Code"
    ' Must run before Cal-strip so we don't remove a Cal prefix that belongs.
    Dim lComCode As Long
    lComCode = InStr(1, sNew, "Com. Code", vbTextCompare)
    If lComCode > 0 Then
        If InStr(1, Left(sNew, lComCode - 1), "Cal. U.", vbBinaryCompare) = 0 Then
            sNew = Left(sNew, lComCode - 1) & "Cal. U. Com. Code" & _
                   Mid(sNew, lComCode + Len("Com. Code"))
        End If
    End If

    ' Fix 2: Strip leading "Cal " prefix (no period -- Lexis+ artifact).
    ' Cal. U. Com. Code already has its prefix; Left(sNew,4) = "Cal." not "Cal "
    If Left(sNew, 4) = "Cal " Then
        sNew = Mid(sNew, 5)
    End If

    ' Fix 3: "Civ Code" -> "Civ. Code" (missing period after abbreviation)
    Dim lCivCode As Long
    lCivCode = InStr(1, sNew, "Civ Code", vbBinaryCompare)
    If lCivCode > 0 Then
        sNew = Left(sNew, lCivCode + 2) & "." & Mid(sNew, lCivCode + 3)
    End If

    ' Fix 4: Add comma before section sign if missing
    Dim lSec As Long
    lSec = InStr(sNew, ChrW(&HA7))
    If lSec > 2 Then
        If Mid(sNew, lSec - 1, 1) = " " And Mid(sNew, lSec - 2, 1) <> "," Then
            sNew = Left(sNew, lSec - 2) & "," & Mid(sNew, lSec - 1)
        End If
    End If

    ' Fix 5: Wrap in parens with terminal period if unwrapped
    If Left(sNew, 1) <> "(" Then
        If Right(sNew, 1) <> "." Then sNew = sNew & "."
        sNew = "(" & sNew & ")"
    End If

    ' Apply only if text changed
    If sNew <> sCit Then
        Dim lParaStart As Long
        lParaStart = oLastPara.Range.start
        ' The Trim above stripped leading spaces from the MEASURED string but
        ' not from the range anchor; advance the anchor to match, or the
        ' replacement covers the spaces and drops the citation's last chars.
        Dim sNLRaw As String
        sNLRaw = oLastPara.Range.text
        Dim nNLLead As Long
        nNLLead = 0
        Do While nNLLead < Len(sNLRaw)
            If AscW(Mid(sNLRaw, nNLLead + 1, 1)) = 32 Then
                nNLLead = nNLLead + 1
            Else
                Exit Do
            End If
        Loop
        lParaStart = lParaStart + nNLLead
        Dim oReplace As Range
        Set oReplace = oDoc.Range(lParaStart, lParaStart + Len(sCit))
        oReplace.text = sNew
        ' Update lNormEnd to reflect characters added by normalization
        lNormEnd = lNormEnd + (Len(sNew) - Len(sCit))
        Set oReplace = Nothing
    End If

End Sub

'===========================================================
Private Sub RemoveLexisStatuteParenthetical(oDoc As Document, oRange As Range)

    Const MAX_ITER As Long = 50

    Dim aSignals(1) As String
    aSignals(0) = "(Deering"
    aSignals(1) = "(LexisNexis"

    Dim s As Integer
    For s = 0 To UBound(aSignals)

        Dim nIter As Long
        nIter = 0

        Do
            nIter = nIter + 1
            If nIter > MAX_ITER Then Exit Do

            Dim sText As String
            sText = oRange.text

            Dim lFound As Long
            lFound = InStr(sText, aSignals(s))
            If lFound = 0 Then Exit Do

            Dim nDepth As Integer
            nDepth = 1
            Dim j As Long
            j = lFound + 1

            Do While j <= Len(sText) And nDepth > 0
                Dim c As String
                c = Mid(sText, j, 1)
                If c = "(" Then nDepth = nDepth + 1
                If c = ")" Then nDepth = nDepth - 1
                j = j + 1
            Loop

            If nDepth <> 0 Then Exit Do

            Dim lDelStart As Long
            Dim lDelEnd As Long
            lDelStart = oRange.start + lFound - 1
            lDelEnd = oRange.start + j - 1

            If lFound > 1 And Mid(sText, lFound - 1, 1) = " " Then
                lDelStart = lDelStart - 1
            End If

            Dim oDel As Range
            Set oDel = oDoc.Range(lDelStart, lDelEnd)
            oDel.Delete
            Set oDel = Nothing

            Set oRange = oDoc.Range(oRange.start, oRange.End)

        Loop

    Next s

End Sub

'===========================================================
' Convert double quote marks in the passage to curly single
' quote marks. Used for statute pastes because the passage
' will be wrapped in outer double quotes.
'
' Curly open double  (U+201C) -> curly open single  (U+2018)
' Curly close double (U+201D) -> curly close single (U+2019)
' Straight double    (34)     -> curly single by context
'===========================================================
Private Sub ConvertDoubleQuotesToSingles(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim sText As String
    sText = oRange.text

    Dim oChar As Range
    Set oChar = oDoc.Range(oRange.start, oRange.start + 1)

    Dim lPos As Long
    For lPos = oRange.End - 1 To oRange.start Step -1

        Dim sIdx As Long
        sIdx = lPos - oRange.start + 1
        Dim nCode As Long
        nCode = AscW(Mid(sText, sIdx, 1))

        Select Case nCode
            Case &H201C
                oChar.SetRange lPos, lPos + 1
                oChar.text = ChrW(&H2018)
            Case &H201D
                oChar.SetRange lPos, lPos + 1
                oChar.text = ChrW(&H2019)
            Case 34
                Dim nPrev34 As Long
                Dim nNext34 As Long
                If sIdx > 1 Then nPrev34 = AscW(Mid(sText, sIdx - 1, 1)) Else nPrev34 = 0
                If sIdx < Len(sText) Then nNext34 = AscW(Mid(sText, sIdx + 1, 1)) Else nNext34 = 0
                oChar.SetRange lPos, lPos + 1
                If IsSpaceOrQuote(nPrev34) Or nPrev34 = 0 Then
                    oChar.text = ChrW(&H2018)
                Else
                    oChar.text = ChrW(&H2019)
                End If
        End Select

    Next lPos

    Set oChar = Nothing

End Sub

'===========================================================
' Convert all straight apostrophes/singles (char 39) to
' curly right singles (U+2019) unconditionally.
' Used for statute pastes only: every straight single in
' statute text is a contraction or possessive, never an
' opening quotation mark.
'===========================================================
Private Sub CurlyApostrophesStatute(oRange As Range)

    Dim oDoc As Document
    Set oDoc = ActiveDocument

    Dim sText As String
    sText = oRange.text

    Dim oChar As Range
    Set oChar = oDoc.Range(oRange.start, oRange.start + 1)

    Dim lPos As Long
    For lPos = oRange.End - 1 To oRange.start Step -1
        Dim sIdx As Long
        sIdx = lPos - oRange.start + 1
        If AscW(Mid(sText, sIdx, 1)) = 39 Then
            oChar.SetRange lPos, lPos + 1
            oChar.text = ChrW(&H2019)
        End If
    Next lPos

    Set oChar = Nothing

End Sub

'===========================================================
' Restructure a completed paste block into parenthetical form.
'
' Expected input:   "Passage text."  (Citation text.)
' Output (case):    (Citation text ["Passage text"].)
' Output (statute): (Citation text ["Passage text."].)
'
' For cases the trailing sentence-ending punctuation is
' stripped. For statutes it is kept.
'
' lRangeEnd is updated by reference to the end of the newly
' inserted text.
'===========================================================
Private Sub RestructureAsParenthetical(oDoc As Document, _
                                        ByVal lStart As Long, _
                                        ByRef lRangeEnd As Long, _
                                        ByVal bIsStatute As Boolean)

    Dim oRange As Range
    Set oRange = oDoc.Range(lStart, lRangeEnd)

    Dim sText As String
    sText = oRange.text

    Dim sTrimmed As String
    sTrimmed = RTrim(sText)
    If Len(sTrimmed) = 0 Then Exit Sub

    ' Safety: first character must be an open double quote
    If AscW(Left(sTrimmed, 1)) <> &H201C Then Exit Sub

    ' Locate the outermost citation parenthesis (right-to-left)
    Dim nDepth As Integer
    Dim nCitationStart As Long
    nDepth = 0: nCitationStart = -1
    Dim i As Long
    Dim c As String
    For i = Len(sTrimmed) To 1 Step -1
        c = Mid(sTrimmed, i, 1)
        If c = ")" Then nDepth = nDepth + 1
        If c = "(" Then
            nDepth = nDepth - 1
            If nDepth = 0 Then nCitationStart = i: Exit For
        End If
    Next i
    If nCitationStart = -1 Then Exit Sub

    ' Locate the passage's closing double quote (U+201D): the LAST
    ' one before the citation parenthesis. The passage can contain
    ' internal U+201D marks (SwapSmartQuotes converts nested single
    ' quotes to close-doubles), so the first one found scanning
    ' forward may be internal and would truncate the passage.
    Dim lCloseQuoteIdx As Long
    lCloseQuoteIdx = 0
    For i = nCitationStart - 1 To 2 Step -1
        If AscW(Mid(sTrimmed, i, 1)) = &H201D Then
            lCloseQuoteIdx = i
            Exit For
        End If
    Next i
    If lCloseQuoteIdx = 0 Then Exit Sub

    ' Extract passage (between the outer quote marks)
    Dim sPassage As String
    sPassage = Mid(sTrimmed, 2, lCloseQuoteIdx - 2)

    ' Strip trailing sentence-ending punctuation for cases
    If Not bIsStatute And Len(sPassage) > 0 Then
        Dim nLast As Long
        nLast = AscW(Right(sPassage, 1))
        Select Case nLast
            Case 46, 33, 63, 59
                sPassage = Left(sPassage, Len(sPassage) - 1)
        End Select
    End If

    ' Extract citation body (between outer parens, strip trailing period)
    Dim sCitationBody As String
    sCitationBody = Mid(sTrimmed, nCitationStart + 1, _
                        Len(sTrimmed) - nCitationStart - 1)
    sCitationBody = RTrim(sCitationBody)
    If Len(sCitationBody) > 0 And Right(sCitationBody, 1) = "." Then
        sCitationBody = Left(sCitationBody, Len(sCitationBody) - 1)
    End If
    sCitationBody = RTrim(sCitationBody)

    ' Build the restructured string
    Dim sNew As String
    sNew = "(" & sCitationBody & " [" & ChrW(&H201C) & _
           sPassage & ChrW(&H201D) & "].)"

    ' Replace the trimmed portion of the range
    Dim oReplace As Range
    Set oReplace = oDoc.Range(lStart, lStart + Len(sTrimmed))
    oReplace.text = sNew

    lRangeEnd = lStart + Len(sNew)

    Set oReplace = Nothing
    Set oRange = Nothing

End Sub

'===========================================================
' PASS 1   Scan the passage for bold 1-2 digit footnote
' markers, delete each one (absorbing any preceding space),
' and return the count in nFootnoteCount (ByRef).
'
' Must run BEFORE bold is stripped (step 12).
' The closing outer quote does not exist yet at this point,
' so the tag is NOT inserted here   that is Pass 2.
'
' Detection rules:
'   - Bold 1-2 digit number anywhere in the passage
'   - Rejected if immediately preceded by "(" (headnote)
'   - Rejected if part of a 3+ digit run

'===========================================================
' Replace internal citations inside the quoted passage with
' [citation], [citations], [Citation.], or [Citations.]
'
' Rules:
'   - Replaces parenthetical citation blocks ( ... ) that contain
'     a legal citation: full case cite, Id., Ibid., supra, or any
'     of those preceded by an introductory signal (See, Cf., etc.)
'   - End-of-sentence citations (terminal period before closing paren)
'     produce [Citation.] / [Citations.]  (capital C, with period)
'   - Mid-sentence citations (no terminal period)
'     produce [citation] / [citations]    (lowercase c, no period)
'   - Multiple citations separated by semicolons -> plural form
'   - Does NOT replace statute/code citations (blocks containing
'     section sign or known code-name patterns)
'   - Does NOT replace (Id.) / (Ibid.) that follow a statute cite
'     (tracks the type of the most recent citation seen)
'===========================================================
Private Sub RemoveInternalCitations(oDoc As Document, oRange As Range)

    ' MAX_ITER must exceed the character count of the longest possible passage
    ' (every non-paren character costs one iteration).
    Const MAX_ITER As Long = 5000
    Dim nIter As Long

    Dim sText As String
    sText = oRange.text
    Dim nLen As Long
    nLen = Len(sText)

    ' Storage for up to 50 replacement regions applied right-to-left
    Const MAX_REGIONS As Integer = 50
    Dim aStart(MAX_REGIONS) As Long
    Dim aEnd(MAX_REGIONS)   As Long
    Dim aRepl(MAX_REGIONS)  As String
    Dim nRegions As Integer
    nRegions = 0

    ' Track whether the last citation seen was a statute, so that a
    ' following (Id.)/(Ibid.) is not replaced.
    Dim bLastCiteWasStatute As Boolean
    bLastCiteWasStatute = False

    Dim iPos As Long
    iPos = 1

    nIter = 0
    Do While iPos <= nLen
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do

        Dim c As String
        c = Mid(sText, iPos, 1)

        If c <> "(" Then
            iPos = iPos + 1
        Else
            ' Find matching close paren
            Dim iClose As Long
            iClose = -1
            Dim nDepth As Integer
            nDepth = 1
            Dim jPos As Long
            jPos = iPos + 1
            Do While jPos <= nLen And nDepth > 0
                Dim d As String
                d = Mid(sText, jPos, 1)
                If d = "(" Then nDepth = nDepth + 1
                If d = ")" Then
                    nDepth = nDepth - 1
                    If nDepth = 0 Then iClose = jPos
                End If
                jPos = jPos + 1
            Loop

            If iClose = -1 Then
                iPos = iPos + 1
            Else
                Dim sInner As String
                sInner = Mid(sText, iPos + 1, iClose - iPos - 1)
                Dim sTrimInner As String
                sTrimInner = Trim(sInner)

                Dim bIsStatuteCite As Boolean
                Dim bIsLegalCite As Boolean
                Dim bIsIdOrIbid As Boolean
                Dim bIsTreatise As Boolean
                Dim bCapSignal As Boolean
                bIsStatuteCite = False
                bIsLegalCite = False
                bIsIdOrIbid = False
                bIsTreatise = False
                bCapSignal = False

                ' Treatise check comes FIRST: secondary-source cites like
                ' (Rest.2d Torts, sec. 314) or (Witkin ... sec. 123) often
                ' contain a section symbol but should be replaced like a
                ' case cite, not preserved like a statute.
                bIsTreatise = InternalCiteIsTreatise(sTrimInner)

                If Not bIsTreatise Then
                    ' Page references ("p. 268", "pp. 100-105") and "supra"
                    ' back-references are both definitive signals that a cite
                    ' is to a case or treatise, NOT a statute. Statutes are
                    ' cited by section number and never have page numbers;
                    ' nor are they back-referenced with "supra" (codes are
                    ' named on every reference). So either signal forces the
                    ' classifier off the statute path -- the cite then falls
                    ' through to InternalCiteIsLegal for normal replacement.
                    If InStr(1, sTrimInner, "supra", vbTextCompare) > 0 Or _
                       HasPageReference(sTrimInner) Then
                        bIsStatuteCite = False
                    Else
                        bIsStatuteCite = InternalCiteIsStatute(sTrimInner)
                    End If
                End If

                If Not bIsStatuteCite And Not bIsTreatise Then
                    bIsIdOrIbid = InternalCiteIsIdOrIbid(sTrimInner)
                    If Not bIsIdOrIbid Then
                        bIsLegalCite = InternalCiteIsLegal(sTrimInner)
                    End If
                End If

                If Not bIsStatuteCite Then
                    ' Detect capital signal for end-of-sentence override.
                    ' StripInternalSignal sets bCapSignal as a side effect.
                    Dim bCapDummy4 As Boolean
                    Dim sSignalStripped As String
                    sSignalStripped = StripInternalSignal(sTrimInner, bCapSignal)
                End If

                If bIsStatuteCite Then
                    bLastCiteWasStatute = True
                    iPos = iClose + 1

                ElseIf bIsTreatise Then
                    If nRegions < MAX_REGIONS Then
                        aStart(nRegions) = oRange.start + iPos - 1
                        aEnd(nRegions) = oRange.start + iClose
                        aRepl(nRegions) = BuildCiteReplacement(sTrimInner, bCapSignal, False)
                        nRegions = nRegions + 1
                    End If
                    bLastCiteWasStatute = False
                    iPos = iClose + 1

                ElseIf bIsIdOrIbid Then
                    If bLastCiteWasStatute Then
                        iPos = iClose + 1
                    Else
                        If nRegions < MAX_REGIONS Then
                            aStart(nRegions) = oRange.start + iPos - 1
                            aEnd(nRegions) = oRange.start + iClose
                            aRepl(nRegions) = BuildCiteReplacement(sTrimInner, bCapSignal, True)
                            nRegions = nRegions + 1
                        End If
                        bLastCiteWasStatute = False
                        iPos = iClose + 1
                    End If

                ElseIf bIsLegalCite Then
                    If nRegions < MAX_REGIONS Then
                        aStart(nRegions) = oRange.start + iPos - 1
                        aEnd(nRegions) = oRange.start + iClose
                        aRepl(nRegions) = BuildCiteReplacement(sTrimInner, bCapSignal, False)
                        nRegions = nRegions + 1
                    End If
                    bLastCiteWasStatute = False
                    iPos = iClose + 1

                Else
                    iPos = iClose + 1
                End If
            End If
        End If
    Loop

    Dim m As Integer
    For m = nRegions - 1 To 0 Step -1
        Dim oReplace As Range
        Set oReplace = oDoc.Range(aStart(m), aEnd(m))
        oReplace.text = aRepl(m)
        Set oReplace = Nothing
    Next m

    ' Pass 2: catch bare (unwrapped) citation sentences that escape the
    ' (...)-block scanner above (e.g., federal case cites that end in a year
    ' paren but are not themselves wrapped in outer parens, and bare LEXIS/WL
    ' string cites that have no parens at all).
    RemoveBareCitationSentences oDoc, oRange

End Sub

'===========================================================
' Phase 2 of internal-cite detection: bare (unwrapped) citation sentences.
' Catches federal cases ending in a year paren but not themselves wrapped,
' bare LEXIS/WL string cites, and "See ... v. ..." string cites.
' Tight criteria; designed to leave editorial parentheticals and prose
' that merely mentions a case name alone.
'===========================================================
Private Sub RemoveBareCitationSentences(oDoc As Document, oRange As Range)
    If oRange Is Nothing Then Exit Sub

    Dim sText As String
    sText = oRange.text
    Dim nLen As Long
    nLen = Len(sText)
    If nLen < 10 Then Exit Sub

    ' Collect sentence (start,end) positions in sText. end is inclusive of
    ' the sentence-terminating "." plus any trailing closing quotes/brackets,
    ' but excludes trailing whitespace.
    Dim aSStart() As Long, aSEnd() As Long
    Dim nSent As Integer
    ReDim aSStart(0 To 500), aSEnd(0 To 500)
    nSent = 0

    Dim curStart As Long
    curStart = 1
    Dim i As Long
    i = 1
    Do While i <= nLen
        If Mid(sText, i, 1) = "." Then
            If IsSentenceBoundary(sText, i, nLen) Then
                Dim k As Long
                k = i + 1
                Do While k <= nLen
                    Dim ck As String
                    ck = Mid(sText, k, 1)
                    If ck = """" Or ck = "'" Or ck = ")" Or ck = "]" Or _
                       ck = ChrW(8217) Or ck = ChrW(8221) Then
                        k = k + 1
                    Else
                        Exit Do
                    End If
                Loop
                If nSent < 500 Then
                    aSStart(nSent) = curStart
                    aSEnd(nSent) = k - 1
                    nSent = nSent + 1
                End If
                Do While k <= nLen
                    If Mid(sText, k, 1) <> " " Then Exit Do
                    k = k + 1
                Loop
                curStart = k
                i = k
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    If curStart <= nLen And nSent < 500 Then
        aSStart(nSent) = curStart
        aSEnd(nSent) = nLen
        nSent = nSent + 1
    End If

    ' Classify each sentence; collect replacements.
    Dim aReplStart() As Long, aReplEnd() As Long, aRepl() As String
    Dim nRepl As Integer
    ReDim aReplStart(0 To 500), aReplEnd(0 To 500), aRepl(0 To 500)
    nRepl = 0

    Dim s As Integer
    For s = 0 To nSent - 1
        Dim sSent As String
        sSent = Mid(sText, aSStart(s), aSEnd(s) - aSStart(s) + 1)
        If IsBareCitationSentence(sSent) Then
            Dim sR As String
            sR = BuildBareCiteReplacement(sSent)
            aReplStart(nRepl) = oRange.start + aSStart(s) - 1
            aReplEnd(nRepl) = oRange.start + aSEnd(s)
            aRepl(nRepl) = sR
            nRepl = nRepl + 1
        End If
    Next s

    Dim m As Integer
    For m = nRepl - 1 To 0 Step -1
        Dim oReplBare As Range
        Set oReplBare = oDoc.Range(aReplStart(m), aReplEnd(m))
        oReplBare.text = aRepl(m)
        Set oReplBare = Nothing
    Next m
End Sub

'===========================================================
' Is sText(pos) a "." marking a sentence end?
' Tight rules to avoid false positives on legal abbreviations:
'   - "." preceded by single letter (e.g. "v.", "L.", "n.") -> NO
'   - "." preceded by short multi-letter legal abbreviation
'     ("Inc.", "Cal.", "Cir.", "Ed.", "Ct." ...) -> NO
'   - "." preceded by ")" (e.g. "(1948).") -> YES
'   - "." preceded by digit (e.g. "*5.", "1.") -> YES
'===========================================================
Private Function IsSentenceBoundary(sText As String, _
                                    pos As Long, _
                                    nLen As Long) As Boolean
    If pos > nLen Then Exit Function
    If Mid(sText, pos, 1) <> "." Then Exit Function

    Dim k As Long
    k = pos + 1
    Do While k <= nLen
        Dim ck As String
        ck = Mid(sText, k, 1)
        If ck = """" Or ck = "'" Or ck = ")" Or ck = "]" Or _
           ck = ChrW(8217) Or ck = ChrW(8221) Then
            k = k + 1
        Else
            Exit Do
        End If
    Loop
    If k > nLen Then Exit Function
    If Mid(sText, k, 1) <> " " Then Exit Function

    If k + 1 > nLen Then
        IsSentenceBoundary = True
        Exit Function
    End If

    Dim chNext As String
    chNext = Mid(sText, k + 1, 1)
    Dim nA As Long
    nA = AscW(chNext)
    If Not ((nA >= 65 And nA <= 90) Or (nA >= 48 And nA <= 57)) Then Exit Function

    If pos >= 2 Then
        Dim chBefore As String
        chBefore = Mid(sText, pos - 1, 1)
        If chBefore = ")" Then
            IsSentenceBoundary = True
            Exit Function
        End If
        Dim nb As Long
        nb = AscW(chBefore)
        If nb >= 48 And nb <= 57 Then
            IsSentenceBoundary = True
            Exit Function
        End If
    End If

    Dim sWord As String
    sWord = WordBeforeDot(sText, pos)
    If Len(sWord) = 0 Then
        IsSentenceBoundary = True
        Exit Function
    End If
    If IsLegalAbbreviation(sWord) Then Exit Function
    IsSentenceBoundary = True
End Function

Private Function WordBeforeDot(sText As String, pos As Long) As String
    Dim j As Long
    j = pos - 1
    Do While j >= 1
        Dim ch As String
        ch = Mid(sText, j, 1)
        Dim nA As Long
        nA = AscW(ch)
        If (nA >= 65 And nA <= 90) Or (nA >= 97 And nA <= 122) Then
            j = j - 1
        Else
            Exit Do
        End If
    Loop
    If j + 1 <= pos - 1 Then
        WordBeforeDot = Mid(sText, j + 1, (pos - 1) - j)
    Else
        WordBeforeDot = ""
    End If
End Function

Private Function IsLegalAbbreviation(sWord As String) As Boolean
    If Len(sWord) = 0 Then Exit Function
    If Len(sWord) = 1 Then
        Dim nA As Long
        nA = AscW(sWord)
        If (nA >= 65 And nA <= 90) Or (nA >= 97 And nA <= 122) Then
            IsLegalAbbreviation = True
            Exit Function
        End If
    End If
    Select Case sWord
        Case "pp", "nn", "fn", "fns", "subd", "subds", "supra", "cert", "reh"
            IsLegalAbbreviation = True
        Case "Inc", "Co", "Corp", "Ltd", "LLC", "LLP", "LP", "Bros"
            IsLegalAbbreviation = True
        Case "Cir", "Ed", "Eds", "Ct", "Op", "Vol", "Sec", "Misc", _
             "Stat", "Supp", "Univ", "No", "Nos", "Reg", "ed"
            IsLegalAbbreviation = True
        Case "Jr", "Sr", "St", "Mr", "Mrs", "Ms", "Dr"
            IsLegalAbbreviation = True
        Case "Ala", "Ariz", "Ark", "Cal", "Colo", "Conn", "Del", "Fla", _
             "Ga", "Haw", "Ill", "Ind", "Iowa", "Kan", "Ky", "La", "Md", _
             "Mass", "Mich", "Minn", "Miss", "Mo", "Mont", "Neb", "Nev", _
             "Okla", "Or", "Ore", "Pa", "Tenn", "Tex", "Va", "Vt", "Wash", _
             "Wis", "Wyo"
            IsLegalAbbreviation = True
        Case "Fed", "F2d", "F3d", "Pub", "DDC"
            IsLegalAbbreviation = True
        Case "Civ", "Pen", "Lab", "Veh", "Prob", "Bus", "Prof", "Wel", _
             "Inst", "Gov", "Fam", "Educ", "Health", "Saf"
            IsLegalAbbreviation = True
    End Select
End Function

'===========================================================
' Citation-shape detection.
'===========================================================
Private Function IsBareCitationSentence(sSent As String) As Boolean
    Dim s As String
    s = Trim(sSent)
    If Len(s) < 6 Then Exit Function
    If Left(s, 3) = "Id." Or Left(s, 5) = "Ibid." Or _
       Left(s, 3) = "Id," Or Left(s, 3) = "Id " Then
        Exit Function
    End If

    Dim sStripped As String
    sStripped = StripBareCiteSignal(s)
    sStripped = LTrim(sStripped)
    If Len(sStripped) < 5 Then Exit Function

    Dim bStart As Boolean
    bStart = False
    If StartsWithCaseName(sStripped) Then bStart = True
    If Not bStart Then
        If StartsWithInReExParte(sStripped) Then bStart = True
    End If
    If Not bStart Then
        If StartsWithBareReporter(sStripped) Then bStart = True
    End If
    If Not bStart Then Exit Function

    If ContainsBareCitationMarker(sStripped) Then
        IsBareCitationSentence = True
    End If
End Function

Private Function StripBareCiteSignal(s As String) As String
    Dim arr(0 To 13) As String
    arr(0) = "See also "
    arr(1) = "See, e.g., "
    arr(2) = "See e.g., "
    arr(3) = "But see, e.g., "
    arr(4) = "But see "
    arr(5) = "But cf. "
    arr(6) = "Cf. "
    arr(7) = "Cf., "
    arr(8) = "Accord, "
    arr(9) = "Accord "
    arr(10) = "Compare "
    arr(11) = "Contra "
    arr(12) = "E.g., "
    arr(13) = "See "
    Dim i As Integer
    For i = 0 To 13
        If Len(s) >= Len(arr(i)) Then
            If Left(s, Len(arr(i))) = arr(i) Then
                StripBareCiteSignal = Mid(s, Len(arr(i)) + 1)
                Exit Function
            End If
        End If
    Next i
    StripBareCiteSignal = s
End Function

Private Function StartsWithCaseName(s As String) As Boolean
    If Len(s) < 5 Then Exit Function
    Dim nA As Long
    nA = AscW(Mid(s, 1, 1))
    If nA < 65 Or nA > 90 Then Exit Function

    Dim nSearch As Long
    nSearch = Len(s)
    If nSearch > 200 Then nSearch = 200
    Dim p As Long
    p = InStr(1, Left(s, nSearch), " v. ", vbBinaryCompare)
    If p = 0 Then Exit Function
    If p + 4 > Len(s) Then Exit Function
    Dim chAfter As String
    chAfter = Mid(s, p + 4, 1)
    Dim nb As Long
    nb = AscW(chAfter)
    If nb >= 65 And nb <= 90 Then StartsWithCaseName = True
End Function

Private Function StartsWithInReExParte(s As String) As Boolean
    If Len(s) >= 6 Then
        If Left(s, 6) = "In re " Then
            StartsWithInReExParte = True
            Exit Function
        End If
    End If
    If Len(s) >= 9 Then
        If Left(s, 9) = "Ex parte " Then
            StartsWithInReExParte = True
            Exit Function
        End If
    End If
    If Len(s) >= 8 Then
        If Left(s, 8) = "Ex rel. " Then
            StartsWithInReExParte = True
            Exit Function
        End If
    End If
    If Len(s) >= 17 Then
        If Left(s, 17) = "In the Matter of " Then
            StartsWithInReExParte = True
            Exit Function
        End If
    End If
End Function

Private Function StartsWithBareReporter(s As String) As Boolean
    If Len(s) < 8 Then Exit Function
    Dim nA As Long
    nA = AscW(Mid(s, 1, 1))
    If nA < 48 Or nA > 57 Then Exit Function

    Dim i As Long
    i = 1
    Do While i <= Len(s)
        Dim nD As Long
        nD = AscW(Mid(s, i, 1))
        If nD < 48 Or nD > 57 Then Exit Do
        i = i + 1
    Loop
    If i = 1 Or i > Len(s) Then Exit Function
    If Mid(s, i, 1) <> " " Then Exit Function

    If InStr(1, s, " LEXIS ", vbBinaryCompare) > 0 Then
        StartsWithBareReporter = True
        Exit Function
    End If
    If InStr(1, s, " WL ", vbBinaryCompare) > 0 Then
        StartsWithBareReporter = True
        Exit Function
    End If
End Function

Private Function ContainsBareCitationMarker(s As String) As Boolean
    Dim nCount As Integer
    If HasLegalYearParen(s, nCount) Then
        ContainsBareCitationMarker = True
        Exit Function
    End If
    If InStr(1, s, " LEXIS ", vbBinaryCompare) > 0 Then
        ContainsBareCitationMarker = True
        Exit Function
    End If
    If InStr(1, s, " WL ", vbBinaryCompare) > 0 Then
        ContainsBareCitationMarker = True
        Exit Function
    End If
    If InStr(1, s, "supra", vbTextCompare) > 0 Then
        ContainsBareCitationMarker = True
        Exit Function
    End If
    If InStr(1, s, "at p.", vbBinaryCompare) > 0 Or _
       InStr(1, s, "at pp.", vbBinaryCompare) > 0 Then
        ContainsBareCitationMarker = True
        Exit Function
    End If
End Function

Private Function BuildBareCiteReplacement(sSent As String) As String
    Dim sT As String
    sT = Trim(sSent)
    Dim bPlural As Boolean
    bPlural = False
    If InStr(1, sT, ";", vbBinaryCompare) > 0 Then bPlural = True
    If Not bPlural Then
        Dim nY As Integer
        Call HasLegalYearParen(sT, nY)
        If nY > 1 Then bPlural = True
    End If
    If Not bPlural Then
        Dim p As Long, nV As Integer
        p = 1
        nV = 0
        Do
            p = InStr(p, sT, " v. ", vbBinaryCompare)
            If p = 0 Then Exit Do
            nV = nV + 1
            p = p + 4
        Loop
        If nV > 1 Then bPlural = True
    End If

    Dim bEnd As Boolean
    Dim chLast As String
    chLast = Right(sT, 1)
    If chLast = "." Or chLast = """" Or chLast = ChrW(8221) Or chLast = "]" Then
        bEnd = True
    Else
        bEnd = False
    End If

    If bEnd Then
        If bPlural Then
            BuildBareCiteReplacement = "[Citations.]"
        Else
            BuildBareCiteReplacement = "[Citation.]"
        End If
    Else
        If bPlural Then
            BuildBareCiteReplacement = "[citations]"
        Else
            BuildBareCiteReplacement = "[citation]"
        End If
    End If
End Function

'===========================================================
' Return True if sInner represents a statute/code citation.
'===========================================================
Private Function InternalCiteIsStatute(sInner As String) As Boolean
    If InStr(sInner, ChrW(&HA7)) > 0 Then
        InternalCiteIsStatute = True
        Exit Function
    End If

    Dim aPatterns(19) As String
    aPatterns(0) = "Cal. Civ. Code"
    aPatterns(1) = "Cal. Code Regs."
    aPatterns(2) = "Cal. Health & Saf. Code"
    aPatterns(3) = "Cal. Ins. Code"
    aPatterns(4) = "Cal. Lab. Code"
    aPatterns(5) = "Cal. Pen. Code"
    aPatterns(6) = "Cal. Prob. Code"
    aPatterns(7) = "Cal. Rev. & Tax. Code"
    aPatterns(8) = "Cal. Veh. Code"
    aPatterns(9) = "Pen. Code"
    aPatterns(10) = "Lab. Code"
    aPatterns(11) = "Civ. Code"
    aPatterns(12) = "Prob. Code"
    aPatterns(13) = "Veh. Code"
    aPatterns(14) = "Corp. Code"
    aPatterns(15) = "Welf. & Inst. Code"
    aPatterns(16) = "U.S.C."
    aPatterns(17) = "C.F.R."
    aPatterns(18) = "Fed. Reg."
    aPatterns(19) = "Cal. Bus. & Prof. Code"

    Dim i As Integer
    For i = 0 To 19
        If InStr(1, sInner, aPatterns(i), vbTextCompare) > 0 Then
            InternalCiteIsStatute = True
            Exit Function
        End If
    Next i

    InternalCiteIsStatute = False
End Function

'===========================================================
' Return True if sInner is an Id. or Ibid. cite (after stripping
' any introductory signal).
'===========================================================
Private Function InternalCiteIsIdOrIbid(sInner As String) As Boolean
    Dim sCk As String
    Dim bCapDummy As Boolean
    sCk = LTrim(StripInternalSignal(sInner, bCapDummy))

    ' Match both uppercase (Ibid./Id.) and lowercase (ibid./id.)
    If Left(sCk, 5) = "Ibid." Or Left(sCk, 5) = "ibid." Or _
       sCk = "Ibid" Or sCk = "ibid" Then
        InternalCiteIsIdOrIbid = True
        Exit Function
    End If

    If Left(sCk, 3) = "Id." Or Left(sCk, 3) = "id." Then
        InternalCiteIsIdOrIbid = True
        Exit Function
    End If

    InternalCiteIsIdOrIbid = False
End Function

'===========================================================
' Return True if sInner contains a legal citation to replace:
' a supra cite, or a full citation containing a (YYYY) year paren.
'===========================================================
Private Function InternalCiteIsLegal(sInner As String) As Boolean
    Dim sCk As String
    Dim bCapDummy2 As Boolean
    sCk = LTrim(StripInternalSignal(sInner, bCapDummy2))

    If Len(sCk) = 0 Then
        InternalCiteIsLegal = False
        Exit Function
    End If

    If InStr(1, sCk, "supra", vbTextCompare) > 0 Then
        InternalCiteIsLegal = True
        Exit Function
    End If

    ' Page reference (" p. <num>" or " pp. <num>") is a case/treatise marker.
    If HasPageReference(sCk) Then
        InternalCiteIsLegal = True
        Exit Function
    End If

    ' Full citation: look for any year paren -- (YYYY) Cal-style, or
    ' (<court abbrev> YYYY) Bluebook/federal style.
    Dim nYearParens As Integer
    nYearParens = 0
    If HasLegalYearParen(sCk, nYearParens) Then
        InternalCiteIsLegal = True
        Exit Function
    End If

    InternalCiteIsLegal = False
End Function

'===========================================================
' Shared year-paren detector. Recognizes:
'   (YYYY)                 -- California style
'   (<court> YYYY)         -- Bluebook / federal style, e.g.
'                              (9th Cir. 2020), (2d Cir. 2020),
'                              (Cal. 2020), (D. Cal. 2020),
'                              (N.D. Cal. 2020), (Fed. Cir. 2020),
'                              (S.D.N.Y. 2020)
'   (<edition>th ed. YYYY) -- treatise editions, e.g. (5th ed. 2008)
' Plausible year range 1800-2099.
'
' On exit, nCount holds the number of year parens found.
' Returns True if at least one year paren is found.
'
' Conservative: a paren whose closing four digits are NOT preceded by
' a recognized court abbreviation or "ed." is rejected, so things like
' (footnote omitted) and (emphasis added) never match.
'===========================================================
Private Function HasLegalYearParen(sCk As String, _
                                    ByRef nCount As Integer) As Boolean
    nCount = 0
    Dim nL As Long
    nL = Len(sCk)
    If nL < 6 Then
        HasLegalYearParen = False
        Exit Function
    End If

    Dim i As Long
    For i = 1 To nL - 5
        If Mid(sCk, i, 1) = "(" Then
            ' Find matching close paren (depth 1; nested parens unlikely here)
            Dim j As Long, nDepth As Integer
            nDepth = 1
            j = i + 1
            Do While j <= nL And nDepth > 0
                Dim ch As String
                ch = Mid(sCk, j, 1)
                If ch = "(" Then nDepth = nDepth + 1
                If ch = ")" Then nDepth = nDepth - 1
                If nDepth = 0 Then Exit Do
                j = j + 1
            Loop
            If j > nL Then
                ' No close paren found; stop scanning this i
            Else
                ' Inner text between ( at i and ) at j (exclusive of parens)
                Dim sInside As String
                If j - i - 1 > 0 Then
                    sInside = Mid(sCk, i + 1, j - i - 1)
                Else
                    sInside = ""
                End If

                ' The year, if any, is the LAST 4 chars of sInside (possibly
                ' preceded by space + court abbreviation).
                Dim sTrim As String
                sTrim = RTrim(sInside)
                Dim nT As Long
                nT = Len(sTrim)
                If nT >= 4 Then
                    Dim sLast4 As String
                    sLast4 = Mid(sTrim, nT - 3, 4)
                    Dim bDigits As Boolean
                    bDigits = True
                    Dim k As Integer
                    For k = 1 To 4
                        Dim nD As Long
                        nD = AscW(Mid(sLast4, k, 1))
                        If nD < 48 Or nD > 57 Then
                            bDigits = False
                            Exit For
                        End If
                    Next k
                    If bDigits Then
                        Dim nYear As Long
                        nYear = CLng(sLast4)
                        If nYear >= 1800 And nYear <= 2099 Then
                            ' Determine prefix portion before the 4-digit year.
                            Dim sPrefix As String
                            If nT = 4 Then
                                sPrefix = ""
                            Else
                                sPrefix = RTrim(Left(sTrim, nT - 4))
                            End If

                            Dim bAccept As Boolean
                            bAccept = False
                            If Len(sPrefix) = 0 Then
                                ' Pure (YYYY) -- California style.
                                bAccept = True
                            ElseIf IsRecognizedYearPrefix(sPrefix) Then
                                bAccept = True
                            End If

                            If bAccept Then
                                nCount = nCount + 1
                            End If
                        End If
                    End If
                End If

                ' Continue scanning after the close paren.
                i = j
            End If
        End If
    Next i

    HasLegalYearParen = (nCount > 0)
End Function

'===========================================================
' Detect a page reference: " p. <digit>" or " pp. <digit>", with the
' leading space required so we don't match "Corp." or similar. Page
' references are exclusive to case and treatise citations; statutes
' use section/subdivision numbers, not pages.
'===========================================================
Private Function HasPageReference(s As String) As Boolean
    Dim n As Long
    n = Len(s)
    If n < 5 Then Exit Function

    Dim i As Long
    For i = 1 To n - 3
        If Mid(s, i, 1) = " " Then
            ' Try " p. <digit>"
            If i + 3 <= n Then
                If Mid(s, i + 1, 3) = "p. " Then
                    If i + 4 <= n Then
                        Dim nA As Long
                        nA = AscW(Mid(s, i + 4, 1))
                        If nA >= 48 And nA <= 57 Then
                            HasPageReference = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            ' Try " pp. <digit>"
            If i + 4 <= n Then
                If Mid(s, i + 1, 4) = "pp. " Then
                    If i + 5 <= n Then
                        Dim nb As Long
                        nb = AscW(Mid(s, i + 5, 1))
                        If nb >= 48 And nb <= 57 Then
                            HasPageReference = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

'===========================================================
' Recognize the text that precedes a 4-digit year inside a citation paren
' as a court abbreviation or treatise-edition indicator. Conservative:
' if we don't recognize it, return False so non-citation parens like
' (footnote omitted) never qualify.
'
' Accepts:
'   - any string ending in "Cir." (e.g. "9th Cir.", "2d Cir.", "Fed. Cir.")
'   - any string ending in "ed." (e.g. "5th ed.", "3d ed.")
'   - explicit court abbreviations: Cal., U.S., D.C., D. Cal., N.D. Cal.,
'     S.D. Cal., E.D. Cal., C.D. Cal., S.D.N.Y., E.D.N.Y., N.D.N.Y.,
'     N.D. Ill., D. Mass., D. Or., D. Nev., D. Ariz., W.D. Wash.,
'     N.D. Tex., S.D. Tex., D. Colo., D. Md., D.D.C., 1st Cir. ... 11th Cir.,
'     Fed. Cir., D.C. Cir.
'===========================================================
Private Function IsRecognizedYearPrefix(sPrefix As String) As Boolean
    Dim s As String
    s = Trim(sPrefix)
    If Len(s) = 0 Then
        IsRecognizedYearPrefix = False
        Exit Function
    End If

    ' Any circuit abbreviation: ends with "Cir."
    If Len(s) >= 4 Then
        If StrComp(Right(s, 4), "Cir.", vbTextCompare) = 0 Then
            IsRecognizedYearPrefix = True
            Exit Function
        End If
    End If

    ' Treatise edition: ends with "ed."
    If Len(s) >= 3 Then
        If StrComp(Right(s, 3), "ed.", vbTextCompare) = 0 Then
            IsRecognizedYearPrefix = True
            Exit Function
        End If
    End If

    ' Explicit court abbreviations (case-insensitive exact match)
    Dim aCourts(31) As String
    aCourts(0) = "Cal."
    aCourts(1) = "U.S."
    aCourts(2) = "D.C."
    aCourts(3) = "D. Cal."
    aCourts(4) = "N.D. Cal."
    aCourts(5) = "S.D. Cal."
    aCourts(6) = "E.D. Cal."
    aCourts(7) = "C.D. Cal."
    aCourts(8) = "S.D.N.Y."
    aCourts(9) = "E.D.N.Y."
    aCourts(10) = "N.D.N.Y."
    aCourts(11) = "W.D.N.Y."
    aCourts(12) = "N.D. Ill."
    aCourts(13) = "C.D. Ill."
    aCourts(14) = "S.D. Ill."
    aCourts(15) = "D. Mass."
    aCourts(16) = "D. Or."
    aCourts(17) = "D. Nev."
    aCourts(18) = "D. Ariz."
    aCourts(19) = "W.D. Wash."
    aCourts(20) = "E.D. Wash."
    aCourts(21) = "N.D. Tex."
    aCourts(22) = "S.D. Tex."
    aCourts(23) = "E.D. Tex."
    aCourts(24) = "W.D. Tex."
    aCourts(25) = "D. Colo."
    aCourts(26) = "D. Md."
    aCourts(27) = "D.D.C."
    aCourts(28) = "N.D. Ga."
    aCourts(29) = "M.D. Fla."
    aCourts(30) = "S.D. Fla."
    aCourts(31) = "D. Minn."

    Dim i As Integer
    For i = 0 To 31
        If StrComp(s, aCourts(i), vbTextCompare) = 0 Then
            IsRecognizedYearPrefix = True
            Exit Function
        End If
    Next i

    IsRecognizedYearPrefix = False
End Function

'===========================================================
' Return True if sInner is a treatise / secondary-source citation.
' Treatises are replaced like cases (per California Style Manual
' treatment in [citation] omission), NOT preserved like statutes.
'
' Recognizes (case-insensitive substring match):
'   Witkin              -- Cal. Procedure, Summary of Cal. Law, etc.
'   Rest.2d / Rest.3d / Restatement
'   Rutter / Cal. Prac. Guide
'   C.J.S. / Corpus Juris
'   Am.Jur. / Am. Jur. / American Jurisprudence
'   A.L.R. / ALR
'   B.E. Witkin (older signature)
'   McLane, Vorbeck, Bender, Matthew Bender (publisher names)
'   LaFave, Wright & Miller, Moore's Federal Practice (federal treatises)
'===========================================================
Private Function InternalCiteIsTreatise(sInner As String) As Boolean
    Dim sCk As String
    Dim bCapDummyT As Boolean
    sCk = LTrim(StripInternalSignal(sInner, bCapDummyT))

    If Len(sCk) = 0 Then
        InternalCiteIsTreatise = False
        Exit Function
    End If

    Dim aSigs(19) As String
    aSigs(0) = "Witkin"
    aSigs(1) = "Rest.2d"
    aSigs(2) = "Rest.3d"
    aSigs(3) = "Rest.4th"
    aSigs(4) = "Restatement"
    aSigs(5) = "Rutter"
    aSigs(6) = "Cal. Prac. Guide"
    aSigs(7) = "C.J.S."
    aSigs(8) = "Corpus Juris"
    aSigs(9) = "Am.Jur."
    aSigs(10) = "Am. Jur."
    aSigs(11) = "American Jurisprudence"
    aSigs(12) = "A.L.R."
    aSigs(13) = "Matthew Bender"
    aSigs(14) = "Moore's Federal Practice"
    aSigs(15) = "Wright & Miller"
    aSigs(16) = "LaFave"
    aSigs(17) = "Nimmer"
    aSigs(18) = "Couch on Insurance"
    aSigs(19) = "Cal. Jur."

    Dim i As Integer
    For i = 0 To 19
        If InStr(1, sCk, aSigs(i), vbTextCompare) > 0 Then
            InternalCiteIsTreatise = True
            Exit Function
        End If
    Next i

    InternalCiteIsTreatise = False
End Function

'===========================================================
' Strip a leading introductory signal from sInner and return
' the remainder.  Handles: See, See also, See e.g., Cf., But see,
' Accord, E.g., Compare, Contra (case-insensitive).
'===========================================================
Private Function StripInternalSignal(sInner As String, _
                                      ByRef bCapSignal As Boolean) As String
    ' Signals 0-12 are capital (force end-of-sentence).
    ' Signal 13 ("cf. ") is lowercase (mid-sentence stays mid-sentence).
    Dim aSignals(13) As String
    aSignals(0) = "See also "
    aSignals(1) = "See, e.g., "
    aSignals(2) = "See, e.g.,"
    aSignals(3) = "See e.g., "
    aSignals(4) = "See e.g.,"
    aSignals(5) = "See "
    aSignals(6) = "Cf. "
    aSignals(7) = "But see "
    aSignals(8) = "Accord "
    aSignals(9) = "E.g., "
    aSignals(10) = "E.g.,"
    aSignals(11) = "Compare "
    aSignals(12) = "Contra "
    aSignals(13) = "cf. "   ' lowercase: does NOT force end-of-sentence

    bCapSignal = False
    Dim i As Integer
    For i = 0 To 13
        Dim nSig As Long
        nSig = Len(aSignals(i))
        If nSig > 0 And Len(sInner) >= nSig Then
            If StrComp(Left(sInner, nSig), aSignals(i), vbTextCompare) = 0 Then
                If i <= 12 Then bCapSignal = True
                StripInternalSignal = Mid(sInner, nSig + 1)
                Exit Function
            End If
        End If
    Next i

    StripInternalSignal = sInner
End Function

'===========================================================
' Build the replacement token for a citation paren.
'
' End-of-sentence (trailing period inside paren):
'   [Citation.] or [Citations.]
' Mid-sentence (no trailing period):
'   [citation] or [citations]
'
' Plural when: sInner contains ";" OR more than one (YYYY) year paren.
'===========================================================
Private Function BuildCiteReplacement(sInner As String, _
                                       bCapSignal As Boolean, _
                                       bIsIdOrIbid As Boolean) As String
    ' End-of-sentence determination:
    '   1. Capital introductory signal always forces end-of-sentence.
    '   2. For Id./Ibid.: uppercase first letter of the abbreviation
    '      signals end-of-sentence (Id./Ibid. vs id./ibid.).
    '   3. For all other citations: trailing "." before ")" is the marker.
    Dim sTrimR As String
    sTrimR = RTrim(sInner)
    Dim bEndSentence As Boolean
    If bCapSignal Then
        bEndSentence = True
    ElseIf bIsIdOrIbid Then
        ' Strip any lowercase signal that was already removed before
        ' this call; check first letter of the remaining abbreviation.
        Dim bCapDummy3 As Boolean
        Dim sStripped As String
        sStripped = LTrim(StripInternalSignal(sInner, bCapDummy3))
        Dim nFirst As Long
        nFirst = AscW(Left(sStripped, 1))
        ' Uppercase I = 73
        bEndSentence = (nFirst = 73)
    Else
        bEndSentence = (Len(sTrimR) > 0 And Right(sTrimR, 1) = ".")
    End If

    ' Plural test: semicolon between cites OR more than one year paren
    ' (Cal-style (YYYY) or federal-style (<court> YYYY)).
    Dim bPlural As Boolean
    bPlural = (InStr(sInner, ";") > 0)

    If Not bPlural Then
        Dim nYears As Integer
        nYears = 0
        Call HasLegalYearParen(sInner, nYears)
        If nYears > 1 Then bPlural = True
    End If

    If bEndSentence Then
        If bPlural Then
            BuildCiteReplacement = "[Citations.]"
        Else
            BuildCiteReplacement = "[Citation.]"
        End If
    Else
        If bPlural Then
            BuildCiteReplacement = "[citations]"
        Else
            BuildCiteReplacement = "[citation]"
        End If
    End If
End Function


'===========================================================
Private Sub RemoveFootnotesPass1(oDoc As Document, _
                                  oRange As Range, _
                                  ByRef nFootnoteCount As Integer)

    nFootnoteCount = 0

    ' Work only on the passage portion   stop before the citation paren.
    ' At this point blank paragraphs have been removed but the outer
    ' quote marks have NOT been added yet, so use FindQuoteEnd to locate
    ' the passage boundary.
    Dim lPassageEnd As Long
    lPassageEnd = FindQuoteEnd(oRange)
    If lPassageEnd <= oRange.start Then lPassageEnd = oRange.End

    Dim oFind   As Range
    Dim oChk2   As Range
    Dim oChk3   As Range
    Dim oPre    As Range
    Dim oDigits As Range

    Dim lSearchFrom As Long
    lSearchFrom = oRange.start

    Const MAX_ITER As Long = 500
    Dim nIter As Long
    nIter = 0

    Do
        nIter = nIter + 1
        If nIter > MAX_ITER Then Exit Do
        If lSearchFrom >= lPassageEnd Then Exit Do

        ' Find next bold digit in the passage
        Set oFind = oDoc.Range(lSearchFrom, lPassageEnd)

        With oFind.Find
            .ClearFormatting
            .Font.Bold = True
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = True
            .text = "[0-9]"
        End With
        oFind.Find.Execute

        If Not oFind.Find.found Then
            Set oFind = Nothing
            Exit Do
        End If

        ' oFind = exactly one bold digit

        Dim nDigits As Integer
        nDigits = 1

        ' Try to extend to a second bold digit
        If oFind.End < lPassageEnd Then
            Set oChk2 = oFind.Duplicate
            oChk2.MoveEnd wdCharacter, 1

            Dim n2 As Long
            n2 = AscW(Right(oChk2.text, 1))

            If n2 >= 48 And n2 <= 57 Then
                Set oDigits = oChk2.Duplicate
                oDigits.Collapse wdCollapseEnd
                oDigits.MoveStart wdCharacter, -1
                Dim bSecondBold As Boolean
                bSecondBold = (oDigits.Font.Bold = True)
                Set oDigits = Nothing

                If bSecondBold Then
                    nDigits = 2

                    If oChk2.End < lPassageEnd Then
                        Set oChk3 = oChk2.Duplicate
                        oChk3.MoveEnd wdCharacter, 1
                        Dim n3 As Long
                        n3 = AscW(Right(oChk3.text, 1))
                        If n3 >= 48 And n3 <= 57 Then
                            nDigits = 0
                        End If
                        Set oChk3 = Nothing
                    End If

                    If nDigits = 2 Then
                        oFind.MoveEnd wdCharacter, 1
                    End If
                End If
            End If
            Set oChk2 = Nothing
        End If

        If nDigits = 0 Then
            lSearchFrom = oFind.End
            Set oFind = Nothing

        Else
            ' Reject headnotes: "(" immediately before the digit(s)
            Dim bHeadnote As Boolean
            bHeadnote = False

            If oFind.start > oRange.start Then
                Set oPre = oFind.Duplicate
                oPre.Collapse wdCollapseStart
                oPre.MoveStart wdCharacter, -1
                If oPre.text = "(" Then bHeadnote = True
                Set oPre = Nothing
            End If

            Dim bSectionNum As Boolean
            bSectionNum = False
            If Not bHeadnote And oFind.start > oRange.start Then
                Dim nLookMax As Integer
                nLookMax = 15
                If oFind.start - oRange.start < nLookMax Then _
                    nLookMax = CInt(oFind.start - oRange.start)
                Dim oLook As Range
                Set oLook = oFind.Duplicate
                oLook.Collapse wdCollapseStart
                oLook.MoveStart wdCharacter, -nLookMax
                Dim sLook As String
                sLook = oLook.text
                Set oLook = Nothing
                Dim sLookT As String
                sLookT = sLook
                Do While Len(sLookT) > 0
                    Dim nTail As Long
                    nTail = AscW(Right(sLookT, 1))
                    If (nTail >= 48 And nTail <= 57) _
                       Or nTail = 32 Or nTail = 160 Then
                        sLookT = Left(sLookT, Len(sLookT) - 1)
                    Else
                        Exit Do
                    End If
                Loop
                If Len(sLookT) > 0 Then
                    If AscW(Right(sLookT, 1)) = 167 Then
                        bSectionNum = True
                    ElseIf Len(sLookT) >= 7 Then
                        If LCase(Right(sLookT, 7)) = "section" Then
                            bSectionNum = True
                        End If
                    End If
                End If
            End If

            If bHeadnote Or bSectionNum Then
                lSearchFrom = oFind.End
                Set oFind = Nothing

            Else
                ' Absorb any preceding space or Chr(160)
                If oFind.start > oRange.start Then
                    Set oPre = oFind.Duplicate
                    oPre.Collapse wdCollapseStart
                    oPre.MoveStart wdCharacter, -1
                    Dim nPreChar As Long
                    nPreChar = AscW(oPre.text)
                    Set oPre = Nothing
                    If nPreChar = 32 Or nPreChar = 160 Then
                        oFind.MoveStart wdCharacter, -1
                    End If
                End If

                Dim nDelChars As Integer
                nDelChars = Len(oFind.text)
                oFind.Delete
                Set oFind = Nothing

                nFootnoteCount = nFootnoteCount + 1
                lPassageEnd = lPassageEnd - nDelChars
                Set oRange = oDoc.Range(oRange.start, oRange.End)
                ' lSearchFrom stays   after deletion it points at next char
            End If
        End If

    Loop

    Set oFind = Nothing
    Set oChk2 = Nothing
    Set oChk3 = Nothing
    Set oPre = Nothing
    Set oDigits = Nothing

End Sub

'===========================================================
' PASS 2   Insert [Fn. omitted.] or [Fns. omitted.] before
' the closing outer double-quote mark (U+201D).
'
' Runs after WrapInDoubleQuotes (step 18) so the outer
' closing quote exists as the insertion target.
' nFootnoteCount comes from Pass 1.
'===========================================================
' PASS 2 -- Append ", fn. omitted." or ", fns. omitted."
' inside the citation closing parenthesis.
' Finds the last ")", removes any period immediately before it,
' inserts the tag so the terminal period migrates.
'===========================================================
Private Sub RemoveFootnotesPass2(oDoc As Document, _
                                  oRange As Range, _
                                  ByVal nFootnoteCount As Integer)
    If nFootnoteCount = 0 Then Exit Sub
    Set oRange = oDoc.Range(oRange.start, oRange.End)
    Dim sText As String
    sText = oRange.text
    Dim lLastParen As Long
    lLastParen = 0
    Dim ci As Long
    For ci = Len(sText) To 1 Step -1
        If Mid(sText, ci, 1) = ")" Then
            lLastParen = ci
            Exit For
        End If
    Next ci
    If lLastParen = 0 Then Exit Sub
    Dim sTag As String
    If nFootnoteCount = 1 Then
        sTag = ", fn. omitted."
    Else
        sTag = ", fns. omitted."
    End If
    Dim oInsert As Range
    Set oInsert = oRange.Duplicate
    oInsert.Collapse wdCollapseStart
    oInsert.MoveEnd wdCharacter, lLastParen - 1
    oInsert.Collapse wdCollapseEnd
    Dim oPeriodChk As Range
    Set oPeriodChk = oInsert.Duplicate
    oPeriodChk.MoveStart wdCharacter, -1
    If oPeriodChk.text = "." Then
        oPeriodChk.Delete
        Set oRange = oDoc.Range(oRange.start, oRange.End)
        sText = oRange.text
        lLastParen = 0
        For ci = Len(sText) To 1 Step -1
            If Mid(sText, ci, 1) = ")" Then
                lLastParen = ci
                Exit For
            End If
        Next ci
        If lLastParen = 0 Then
            Set oPeriodChk = Nothing: Set oInsert = Nothing: Exit Sub
        End If
        Set oInsert = oRange.Duplicate
        oInsert.Collapse wdCollapseStart
        oInsert.MoveEnd wdCharacter, lLastParen - 1
        oInsert.Collapse wdCollapseEnd
    End If
    Set oPeriodChk = Nothing
    oInsert.InsertBefore sTag
    Set oInsert = Nothing
    Set oRange = oDoc.Range(oRange.start, oRange.End)
End Sub
'===========================================================
' PLEADING-PAPER PASTE PIPELINE
' ===========================================================
' Triggered when one or more right-margin markers from
' pdf_linker.py are detected in the pasted range.
' Outputs a single trailing citation summarizing the range.
'
' Marker formats produced by pdf_linker.py:
'   [Britton Decl.|p2:3]      -- full self-describing
'   [Britton Decl.|p2:3 7]    -- with paragraph (  = U+00B6)
'   [p2:3]                    -- compact fallback (no shortname)
'   [p2:3 7]                  -- compact with paragraph
'
' Quote-mode signal: a curly opening (U+201C) or closing (U+201D)
' double-quote typed immediately before the cursor BEFORE pasting.
' Word's autocorrect produces U+201C in most contexts; we accept
' either as the trigger to be forgiving.
'
' Citation format (per Zachary's spec, May 2026):
'   No paragraph:   "(Doc at p. 2:3.)"
'                   "(Doc at p. 2:3-5.)"
'                   "(Doc at pp. 2:27-3:5.)"
'   Paragraph mode: "(Doc   7.)"
'                   "(Doc     7-9.)"   (   = U+00B6, doubled for plural)
' Paragraph mode trumps page/line when EVERY marker carries a
' paragraph component; otherwise falls back to page/line.
'===========================================================

'===========================================================
' Strip clipped pleading line numbers at paragraph starts.
'
' Pleading-paper line numbers occasionally end up in the
' clipboard when the user drag-selects 3+ lines and Adobe's
' selection band catches the gutter digit on inner lines.
'
' Detection signal: 3 or more paragraph-starts in the pasted
' range match the pattern "<1-2 digits><space><capital letter>"
' AND the numbers form a strictly ascending sequence (within
' a run; resets across page breaks are tolerated).
' Real text essentially never produces this pattern; pleading
' paste artifacts always do.
'
' If the signal fires, every matching paragraph-start in the
' range has its leading digits+space stripped. If it does not
' fire, nothing is touched -- so an isolated sentence like
' "3 cars were destroyed." is always safe.
'===========================================================
Private Sub RemovePleadingLineNumbers(oDoc As Document, oRange As Range)

    Dim sText As String
    sText = oRange.text
    If Len(sText) < 6 Then Exit Sub

    ' First pass: find every paragraph-start that matches the
    ' clipped pattern. Record the position and the parsed number.
    Const MAX_HITS As Integer = 200
    Dim aHitPos(MAX_HITS) As Long       ' 1-based index of digit start in sText
    Dim aHitLen(MAX_HITS) As Integer    ' chars to delete (digits + space)
    Dim aHitNum(MAX_HITS) As Integer    ' parsed line number
    Dim nHits As Integer
    nHits = 0

    Dim i As Long
    For i = 1 To Len(sText) - 2
        Dim bIsParaStart As Boolean
        If i = 1 Then
            bIsParaStart = True
        Else
            Dim nPrev As Long
            nPrev = AscW(Mid(sText, i - 1, 1))
            bIsParaStart = (nPrev = 13 Or nPrev = 11)
        End If
        If bIsParaStart Then
            ' Read 1-2 digits.
            Dim nDigits As Integer
            nDigits = 0
            Dim k As Long
            k = i
            Do While k <= Len(sText) And nDigits < 2
                Dim nC As Long
                nC = AscW(Mid(sText, k, 1))
                If nC >= 48 And nC <= 57 Then
                    nDigits = nDigits + 1
                    k = k + 1
                Else
                    Exit Do
                End If
            Loop
            ' Reject if a third digit follows (3+ digit number is
            ' never a pleading line number).
            Dim bThirdDigit As Boolean
            bThirdDigit = False
            If nDigits = 2 And k <= Len(sText) Then
                Dim nNext As Long
                nNext = AscW(Mid(sText, k, 1))
                If nNext >= 48 And nNext <= 57 Then bThirdDigit = True
            End If
            If nDigits >= 1 And Not bThirdDigit And k + 1 <= Len(sText) Then
                ' Must be: digits, exactly one space, then capital A-Z.
                If AscW(Mid(sText, k, 1)) = 32 Then
                    Dim nAfter As Long
                    nAfter = AscW(Mid(sText, k + 1, 1))
                    If nAfter >= 65 And nAfter <= 90 Then
                        If nHits < MAX_HITS Then
                            aHitPos(nHits) = i
                            aHitLen(nHits) = nDigits + 1   ' digits + the space
                            aHitNum(nHits) = CInt(Mid(sText, i, nDigits))
                            nHits = nHits + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' Need at least 3 hits to consider stripping.
    If nHits < 3 Then Exit Sub

    ' Verify the numbers are strictly ascending. Pleading line
    ' numbers run 1, 2, 3, ... 28 down a page; multi-page pastes
    ' restart at 1 on each page, so be lenient: allow a reset to
    ' any smaller number (page break) but require ascending runs
    ' of at least 3 between resets.
    Dim nLongestRun As Integer
    Dim nCurrentRun As Integer
    nLongestRun = 1
    nCurrentRun = 1
    Dim h As Integer
    For h = 1 To nHits - 1
        If aHitNum(h) > aHitNum(h - 1) Then
            nCurrentRun = nCurrentRun + 1
            If nCurrentRun > nLongestRun Then nLongestRun = nCurrentRun
        Else
            nCurrentRun = 1
        End If
    Next h

    If nLongestRun < 3 Then Exit Sub

    ' Signal fires. Strip every recorded hit, working
    ' right-to-left so earlier positions stay valid.
    Dim m As Integer
    For m = nHits - 1 To 0 Step -1
        Dim lDelStart As Long
        Dim lDelEnd As Long
        lDelStart = oRange.start + aHitPos(m) - 1
        lDelEnd = lDelStart + aHitLen(m)
        Dim oDel As Range
        Set oDel = oDoc.Range(lDelStart, lDelEnd)
        oDel.Delete
        Set oDel = Nothing
    Next m

End Sub

'===========================================================
' Try to parse the inside of a bracket as a marker.
' Accepted forms (case-sensitive 'p'):
'   "Doc Name|pN:N"
'   "Doc Name|pN:N N"        (   = U+00B6)
'   "pN:N"
'   "pN:N N"
' On success, fills sDoc/nPage/nLine/nPara and returns True.
' nPara is 0 when no paragraph component is present.
' sDoc is "" for the compact form (no shortname).
'===========================================================
Private Function TryParseMarkerInner(sInner As String, _
                                      ByRef sDoc As String, _
                                      ByRef nPage As Integer, _
                                      ByRef nLine As Integer, _
                                      ByRef nPara As Integer) As Boolean

    TryParseMarkerInner = False
    sDoc = "": nPage = 0: nLine = 0: nPara = 0

    If Len(sInner) < 4 Then Exit Function

    Dim sCoords As String
    Dim lPipe As Long
    lPipe = InStr(sInner, "|")
    If lPipe > 0 Then
        sDoc = Trim(Left(sInner, lPipe - 1))
        sCoords = Mid(sInner, lPipe + 1)
    Else
        sDoc = ""
        sCoords = sInner
    End If

    ' sCoords must start with lowercase 'p' followed by digits, then ':',
    ' then digits, then optionally ChrW(&HB6) followed by digits.
    If Len(sCoords) < 4 Then Exit Function
    If Left(sCoords, 1) <> "p" Then Exit Function

    Dim k As Long
    k = 2  ' position after 'p'

    ' Read page digits
    Dim sPg As String
    sPg = ""
    Do While k <= Len(sCoords)
        Dim nP As Long
        nP = AscW(Mid(sCoords, k, 1))
        If nP >= 48 And nP <= 57 Then
            sPg = sPg & Mid(sCoords, k, 1)
            k = k + 1
        Else
            Exit Do
        End If
    Loop
    If Len(sPg) = 0 Then Exit Function
    If k > Len(sCoords) Then Exit Function
    If Mid(sCoords, k, 1) <> ":" Then Exit Function
    k = k + 1

    ' Read line digits
    Dim sLn As String
    sLn = ""
    Do While k <= Len(sCoords)
        Dim nL As Long
        nL = AscW(Mid(sCoords, k, 1))
        If nL >= 48 And nL <= 57 Then
            sLn = sLn & Mid(sCoords, k, 1)
            k = k + 1
        Else
            Exit Do
        End If
    Loop
    If Len(sLn) = 0 Then Exit Function

    ' Optional paragraph component
    If k <= Len(sCoords) Then
        If AscW(Mid(sCoords, k, 1)) = &HB6 Then
            k = k + 1
            Dim sPa As String
            sPa = ""
            Do While k <= Len(sCoords)
                Dim nA As Long
                nA = AscW(Mid(sCoords, k, 1))
                If nA >= 48 And nA <= 57 Then
                    sPa = sPa & Mid(sCoords, k, 1)
                    k = k + 1
                Else
                    Exit Do
                End If
            Loop
            If Len(sPa) > 0 Then nPara = CInt(sPa)
        End If
    End If

    nPage = CInt(sPg)
    nLine = CInt(sLn)
    TryParseMarkerInner = True

End Function

'===========================================================
' Scan oRange for markers and populate the g_Marker* arrays.
' Returns the number of markers parsed.
'===========================================================
Private Function ParseAllMarkers(oRange As Range) As Integer

    g_MarkerCount = 0
    ReDim g_MarkerPos(MARKER_MAX_COUNT - 1)
    ReDim g_MarkerLen(MARKER_MAX_COUNT - 1)
    ReDim g_MarkerDoc(MARKER_MAX_COUNT - 1)
    ReDim g_MarkerPage(MARKER_MAX_COUNT - 1)
    ReDim g_MarkerLine(MARKER_MAX_COUNT - 1)
    ReDim g_MarkerPara(MARKER_MAX_COUNT - 1)

    Dim sText As String
    sText = oRange.text
    If Len(sText) < 5 Then
        ParseAllMarkers = 0
        Exit Function
    End If

    Dim i As Long
    For i = 1 To Len(sText) - 4
        If g_MarkerCount >= MARKER_MAX_COUNT Then Exit For
        If Mid(sText, i, 1) = "[" Then
            ' Find matching close within MARKER_MAX_LEN chars
            Dim j As Long
            Dim iClose As Long
            iClose = 0
            Dim jMax As Long
            jMax = i + MARKER_MAX_LEN
            If jMax > Len(sText) Then jMax = Len(sText)
            For j = i + 1 To jMax
                If Mid(sText, j, 1) = "]" Then
                    iClose = j
                    Exit For
                End If
            Next j
            If iClose > 0 Then
                Dim sInner As String
                sInner = Mid(sText, i + 1, iClose - i - 1)
                Dim sDoc As String
                Dim nPage As Integer
                Dim nLine As Integer
                Dim nPara As Integer
                If TryParseMarkerInner(sInner, sDoc, nPage, nLine, nPara) Then
                    g_MarkerPos(g_MarkerCount) = i
                    g_MarkerLen(g_MarkerCount) = iClose - i + 1
                    g_MarkerDoc(g_MarkerCount) = sDoc
                    g_MarkerPage(g_MarkerCount) = nPage
                    g_MarkerLine(g_MarkerCount) = nLine
                    g_MarkerPara(g_MarkerCount) = nPara
                    g_MarkerCount = g_MarkerCount + 1
                End If
            End If
        End If
    Next i

    ParseAllMarkers = g_MarkerCount

End Function

'===========================================================
' If any parsed marker has a non-empty shortname, propagate
' it to all markers that lack one. Per spec: a single full
' marker anywhere in the paste defines the doc for all.
'===========================================================
Private Sub PromoteShortname()
    If g_MarkerCount = 0 Then Exit Sub
    Dim sDoc As String
    sDoc = ""
    Dim i As Integer
    For i = 0 To g_MarkerCount - 1
        If Len(g_MarkerDoc(i)) > 0 Then
            sDoc = g_MarkerDoc(i)
            Exit For
        End If
    Next i
    If Len(sDoc) = 0 Then Exit Sub
    For i = 0 To g_MarkerCount - 1
        If Len(g_MarkerDoc(i)) = 0 Then g_MarkerDoc(i) = sDoc
    Next i
End Sub

'===========================================================
' Return True if sToken appears in sText as a standalone
' word -- i.e. its neighbours on each side are either the
' string boundary or a non-letter character. Case-insensitive.
'
' Used by the complaint-shortname normalizer to distinguish
' "FAC" the abbreviation from "FAC" embedded inside a longer
' alphabetic run (which shouldn't occur in derived shortnames
' but is cheap insurance).
'===========================================================
Private Function ContainsWholeWord(ByVal sText As String, _
                                    ByVal sToken As String) As Boolean
    ContainsWholeWord = False
    If Len(sText) = 0 Or Len(sToken) = 0 Then Exit Function

    Dim sUpperText As String
    Dim sUpperTok As String
    sUpperText = UCase(sText)
    sUpperTok = UCase(sToken)

    Dim k As Long
    k = 1
    Do
        Dim lHit As Long
        lHit = InStr(k, sUpperText, sUpperTok)
        If lHit = 0 Then Exit Function

        Dim bLeftOk As Boolean
        Dim bRightOk As Boolean
        If lHit = 1 Then
            bLeftOk = True
        Else
            bLeftOk = Not IsLetterChar(Mid(sUpperText, lHit - 1, 1))
        End If
        Dim lAfter As Long
        lAfter = lHit + Len(sUpperTok)
        If lAfter > Len(sUpperText) Then
            bRightOk = True
        Else
            bRightOk = Not IsLetterChar(Mid(sUpperText, lAfter, 1))
        End If

        If bLeftOk And bRightOk Then
            ContainsWholeWord = True
            Exit Function
        End If
        k = lHit + 1
    Loop
End Function

Private Function IsLetterChar(ByVal sCh As String) As Boolean
    If Len(sCh) = 0 Then
        IsLetterChar = False
        Exit Function
    End If
    Dim n As Long
    n = AscW(sCh)
    IsLetterChar = (n >= 65 And n <= 90) Or (n >= 97 And n <= 122)
End Function

'===========================================================
' Normalize a complaint-document shortname for citation use.
'
' The shortname in the marker is derived directly from the PDF
' filename by pdf_linker.py -- so casing and form depend on
' however the user named the file. We can get any of:
'   "Complaint", "complaint", "COMPLAINT"  -> "Compl."
'   "Compl", "Compl.", "compl."            -> "Compl."
'   "FAC", "Fac", "fac"                    -> "FAC"
'   "SAC", "Sac", "sac"                    -> "SAC"
'   "TAC", "Tac", "tac"                    -> "TAC"
'   "Second Amended Complaint"             -> "SAC"
'   "First Amended Complaint"              -> "FAC"
'   "Third Amended Complaint"              -> "TAC"
'
' The amended-complaint abbreviations take priority over the
' bare "Complaint" branch -- a filename like "FAC.pdf" or
' "Plaintiff_FAC.pdf" must map to "FAC", not "Compl."
'
' Anything that doesn't match is returned unchanged. Declaration
' shortnames ("Britton Decl.") and other pleading types are
' untouched.
'===========================================================
Private Function NormalizeComplaintShortname(ByVal sDoc As String) As String

    NormalizeComplaintShortname = sDoc
    If Len(sDoc) = 0 Then Exit Function

    ' Check amended-complaint tokens first (priority over bare "Compl").
    ' The "Xst/nd/rd Amended Complaint" full-spelling forms must
    ' precede the bare "FAC/SAC/TAC" check because a filename like
    ' "First Amended Complaint.pdf" contains neither FAC nor SAC nor
    ' TAC as whole words but does contain "Complaint".
    If InStr(1, sDoc, "First Amended Complaint", vbTextCompare) > 0 Then
        NormalizeComplaintShortname = "FAC"
        Exit Function
    End If
    If InStr(1, sDoc, "Second Amended Complaint", vbTextCompare) > 0 Then
        NormalizeComplaintShortname = "SAC"
        Exit Function
    End If
    If InStr(1, sDoc, "Third Amended Complaint", vbTextCompare) > 0 Then
        NormalizeComplaintShortname = "TAC"
        Exit Function
    End If

    If ContainsWholeWord(sDoc, "FAC") Then
        NormalizeComplaintShortname = "FAC"
        Exit Function
    End If
    If ContainsWholeWord(sDoc, "SAC") Then
        NormalizeComplaintShortname = "SAC"
        Exit Function
    End If
    If ContainsWholeWord(sDoc, "TAC") Then
        NormalizeComplaintShortname = "TAC"
        Exit Function
    End If

    ' Fall through to plain complaint. "Compl" matches both
    ' "Compl.", "Compl" and "Complaint" (all variants).
    If InStr(1, sDoc, "Compl", vbTextCompare) > 0 Then
        NormalizeComplaintShortname = "Compl."
        Exit Function
    End If

End Function

'===========================================================
' Build the trailing citation string from the parsed markers.
' Returns just the citation text, e.g. "(Britton Decl. at p. 2:3.)"
' or "(Britton Decl.   7-9.)".
'
' Rules:
'   - If every marker has a paragraph component, paragraph cite
'     trumps page/line: "(Doc   N.)" or "(Doc     N-M.)"
'   - Otherwise page/line:
'       single line:   "(Doc at p. 2:3.)"
'       same page:     "(Doc at p. 2:3-5.)"
'       spans pages:   "(Doc at pp. 2:27-3:5.)"
'   - When the doc shortname is empty, omit it: "(at p. 2:3.)"
'     Shouldn't occur in practice -- pdf_linker.py always emits
'     full markers; compact is only the overflow fallback.
'===========================================================
Private Function BuildPleadingCitation() As String

    BuildPleadingCitation = ""
    If g_MarkerCount = 0 Then Exit Function

    ' Take the doc from the first marker (PromoteShortname has run).
    Dim sDoc As String
    sDoc = g_MarkerDoc(0)

    ' Normalize complaint-style shortnames so casing is consistent
    ' regardless of how the PDF file was named. Non-complaint
    ' shortnames (declarations, separate statements, motions) are
    ' returned unchanged.
    sDoc = NormalizeComplaintShortname(sDoc)

    ' Determine paragraph coverage: do all markers have a paragraph?
    Dim bAllHavePara As Boolean
    bAllHavePara = True
    Dim i As Integer
    For i = 0 To g_MarkerCount - 1
        If g_MarkerPara(i) = 0 Then
            bAllHavePara = False
            Exit For
        End If
    Next i

    Dim sCit As String

    If bAllHavePara Then
        ' Paragraph mode. Find min/max paragraph numbers.
        Dim nMinP As Integer, nMaxP As Integer
        nMinP = g_MarkerPara(0)
        nMaxP = g_MarkerPara(0)
        For i = 1 To g_MarkerCount - 1
            If g_MarkerPara(i) < nMinP Then nMinP = g_MarkerPara(i)
            If g_MarkerPara(i) > nMaxP Then nMaxP = g_MarkerPara(i)
        Next i

        If nMinP = nMaxP Then
            sCit = ChrW(&HB6) & " " & CStr(nMinP)
        Else
            sCit = ChrW(&HB6) & ChrW(&HB6) & " " & _
                   CStr(nMinP) & "-" & CStr(nMaxP)
        End If

        If Len(sDoc) > 0 Then
            BuildPleadingCitation = "(" & sDoc & " " & sCit & ".)"
        Else
            BuildPleadingCitation = "(" & sCit & ".)"
        End If
        Exit Function
    End If

    ' Page/line mode. Find min and max (page,line) pairs.
    Dim nFirstPg As Integer, nFirstLn As Integer
    Dim nLastPg As Integer, nLastLn As Integer
    nFirstPg = g_MarkerPage(0): nFirstLn = g_MarkerLine(0)
    nLastPg = g_MarkerPage(0): nLastLn = g_MarkerLine(0)

    For i = 1 To g_MarkerCount - 1
        Dim nPg As Integer, nLn As Integer
        nPg = g_MarkerPage(i): nLn = g_MarkerLine(i)
        ' Compare (page, line) lexicographically
        If nPg < nFirstPg Or (nPg = nFirstPg And nLn < nFirstLn) Then
            nFirstPg = nPg: nFirstLn = nLn
        End If
        If nPg > nLastPg Or (nPg = nLastPg And nLn > nLastLn) Then
            nLastPg = nPg: nLastLn = nLn
        End If
    Next i

    Dim sRange As String
    If nFirstPg = nLastPg And nFirstLn = nLastLn Then
        sRange = "at p. " & CStr(nFirstPg) & ":" & CStr(nFirstLn)
    ElseIf nFirstPg = nLastPg Then
        sRange = "at p. " & CStr(nFirstPg) & ":" & _
                 CStr(nFirstLn) & "-" & CStr(nLastLn)
    Else
        sRange = "at pp. " & CStr(nFirstPg) & ":" & CStr(nFirstLn) & _
                 "-" & CStr(nLastPg) & ":" & CStr(nLastLn)
    End If

    If Len(sDoc) > 0 Then
        BuildPleadingCitation = "(" & sDoc & " " & sRange & ".)"
    Else
        BuildPleadingCitation = "(" & sRange & ".)"
    End If

End Function

'===========================================================
' Delete all parsed markers from oRange. Walks the markers
' right-to-left so earlier positions stay valid as deletions
' shrink the range. Also absorbs a single trailing space
' after each marker so adjacent text doesn't end up with
' double spaces.
'
' After this returns the g_Marker* arrays are stale -- do
' not reuse them.
'===========================================================
Private Sub StripAllMarkers(oDoc As Document, oRange As Range)

    If g_MarkerCount = 0 Then Exit Sub

    Dim lRangeStart As Long
    lRangeStart = oRange.start

    Dim sText As String
    sText = oRange.text

    Dim i As Integer
    For i = g_MarkerCount - 1 To 0 Step -1
        Dim lDelStart As Long
        Dim lDelEnd As Long
        lDelStart = lRangeStart + g_MarkerPos(i) - 1
        lDelEnd = lDelStart + g_MarkerLen(i)
        ' Absorb trailing space if present in the original text
        Dim iAfter As Long
        iAfter = g_MarkerPos(i) + g_MarkerLen(i)
        If iAfter <= Len(sText) Then
            If AscW(Mid(sText, iAfter, 1)) = 32 Then
                lDelEnd = lDelEnd + 1
            End If
        End If
        Dim oDel As Range
        Set oDel = oDoc.Range(lDelStart, lDelEnd)
        oDel.Delete
        Set oDel = Nothing
    Next i

    g_MarkerCount = 0

End Sub

'===========================================================
' Convert every paragraph mark (chr 13) and vertical tab
' (chr 11, soft line break) inside [lStart, lEnd] to a single
' ASCII space, then collapse any resulting runs of 2+ spaces
' down to a single space. Trailing paragraph marks at the very
' end of the range are stripped rather than replaced -- they
' belong to the surrounding document, not the pasted block.
'
' This is the workhorse for the pleading-paste branch: a PDF
' selection that spans N visible lines lands in Word as N
' paragraphs, but those line breaks are artifacts of the PDF's
' fixed-line layout, not real paragraph boundaries. After
' StripAllMarkers removes the right-margin citation markers,
' this sub welds the lines back into a single passage.
'
' Real paragraph breaks (multi-paragraph pastes) are added back
' manually by the user via a bracketed pilcrow.
'
' lEnd is updated in place: each replacement is char-for-char
' so it does not shift, but the trailing-mark trim and the
' double-space collapse do, so we recover lEnd from the
' tail-length invariant.
'===========================================================
Private Sub ConvertParaMarksToSpaces(oDoc As Document, _
                                      ByVal lStart As Long, _
                                      ByRef lEnd As Long)

    If lEnd <= lStart Then Exit Sub

    Dim lTailLen As Long
    lTailLen = oDoc.content.End - lEnd

    ' Pass 1: trim trailing paragraph/line marks off the range so
    ' we don't fold them into the passage. Done before the body
    ' replacement so it's a single shrink at the tail rather than
    ' a replace-then-strip dance.
    Do While lEnd > lStart
        Dim oTail As Range
        Set oTail = oDoc.Range(lEnd - 1, lEnd)
        Dim nTail As Long
        nTail = AscW(oTail.text)
        Set oTail = Nothing
        If nTail = 13 Or nTail = 11 Then
            Dim oDelTail As Range
            Set oDelTail = oDoc.Range(lEnd - 1, lEnd)
            oDelTail.Delete
            Set oDelTail = Nothing
            lEnd = oDoc.content.End - lTailLen
        Else
            Exit Do
        End If
    Loop

    If lEnd <= lStart Then Exit Sub

    ' Pass 2: walk the range text and replace every interior CR/VT
    ' with a single space. Use Find/Replace on the range -- it's
    ' cheaper than character-by-character ranges and Word handles
    ' the position bookkeeping for us.
    Dim oBody As Range
    Set oBody = oDoc.Range(lStart, lEnd)
    With oBody.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .text = Chr(13)
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
    End With
    Set oBody = Nothing
    lEnd = oDoc.content.End - lTailLen

    Set oBody = oDoc.Range(lStart, lEnd)
    With oBody.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .text = Chr(11)
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
    End With
    Set oBody = Nothing
    lEnd = oDoc.content.End - lTailLen

    ' Pass 3: collapse any double-spaces created by joining a
    ' line that ended with a space to the next line. Loop until
    ' stable to handle 3+ space runs (rare but cheap to cover).
    Do
        Set oBody = oDoc.Range(lStart, lEnd)
        Dim bFound As Boolean
        With oBody.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            .text = "  "
            .Replacement.text = " "
            bFound = .Execute(Replace:=wdReplaceAll)
        End With
        Set oBody = Nothing
        lEnd = oDoc.content.End - lTailLen
        If Not bFound Then Exit Do
    Loop

End Sub

'===========================================================
' If [lStart, lEnd] begins with "N. " (and optionally any
' run of leading whitespace before N), where N matches the
' supplied paragraph number, delete that prefix.
'
' This is used by the pleading-paste branch when in paragraph
' mode: the cited passage is "(Compl.   4.)" and the pasted
' passage starts with "4. Defendant is..." -- the leading
' "4. " is redundant with the citation and would read awkwardly.
'
' Only fires when:
'   - nParaNum > 0 (we're in paragraph mode), and
'   - the passage actually starts with that exact number
'     followed by a period and at least one whitespace char.
'
' lEnd is updated in place via the tail-length invariant.
'===========================================================
Private Sub StripLeadingParagraphNumber(oDoc As Document, _
                                         ByVal lStart As Long, _
                                         ByRef lEnd As Long, _
                                         ByVal nParaNum As Integer)

    If nParaNum <= 0 Then Exit Sub
    If lEnd <= lStart Then Exit Sub

    Dim lTailLen As Long
    lTailLen = oDoc.content.End - lEnd

    Dim oBody As Range
    Set oBody = oDoc.Range(lStart, lEnd)
    Dim sText As String
    sText = oBody.text
    Set oBody = Nothing

    If Len(sText) < 3 Then Exit Sub

    ' Skip any leading whitespace (space, NBSP, tab) before the digits.
    ' Don't skip CR/VT -- ConvertParaMarksToSpaces has already folded
    ' those, so any remaining at this point shouldn't exist, but if
    ' they did we want to leave them alone so we don't accidentally
    ' eat a real paragraph boundary.
    Dim k As Long
    k = 1
    Do While k <= Len(sText)
        Dim nC As Long
        nC = AscW(Mid(sText, k, 1))
        If nC = 32 Or nC = 160 Or nC = 9 Then
            k = k + 1
        Else
            Exit Do
        End If
    Loop
    If k > Len(sText) Then Exit Sub

    ' Read digits at position k.
    Dim sNum As String
    sNum = ""
    Do While k <= Len(sText)
        Dim nD As Long
        nD = AscW(Mid(sText, k, 1))
        If nD >= 48 And nD <= 57 Then
            sNum = sNum & Mid(sText, k, 1)
            k = k + 1
        Else
            Exit Do
        End If
    Loop
    If Len(sNum) = 0 Then Exit Sub
    If CInt(sNum) <> nParaNum Then Exit Sub

    ' Must be followed by a period.
    If k > Len(sText) Then Exit Sub
    If Mid(sText, k, 1) <> "." Then Exit Sub
    k = k + 1

    ' Must be followed by at least one space-like character.
    If k > Len(sText) Then Exit Sub
    Dim nFollow As Long
    nFollow = AscW(Mid(sText, k, 1))
    If Not (nFollow = 32 Or nFollow = 160 Or nFollow = 9) Then Exit Sub

    ' Eat all trailing whitespace after the period so the passage
    ' starts cleanly with the next word.
    Do While k <= Len(sText)
        Dim nW As Long
        nW = AscW(Mid(sText, k, 1))
        If nW = 32 Or nW = 160 Or nW = 9 Then
            k = k + 1
        Else
            Exit Do
        End If
    Loop

    ' Delete characters 1 .. k-1 from the range (the leading "N. ").
    Dim lDelStart As Long
    Dim lDelEnd As Long
    lDelStart = lStart
    lDelEnd = lStart + (k - 1)
    Dim oDel As Range
    Set oDel = oDoc.Range(lDelStart, lDelEnd)
    oDel.Delete
    Set oDel = Nothing

    lEnd = oDoc.content.End - lTailLen

End Sub

'===========================================================
' Return the minimum paragraph number across all parsed markers,
' or 0 if any marker lacks a paragraph component. Used to decide
' whether to strip a redundant "N. " prefix from the passage.
'
' Must be called BEFORE StripAllMarkers, which clears the arrays.
'===========================================================
Private Function GetMinMarkerPara() As Integer
    GetMinMarkerPara = 0
    If g_MarkerCount = 0 Then Exit Function
    Dim i As Integer
    Dim nMin As Integer
    nMin = g_MarkerPara(0)
    For i = 1 To g_MarkerCount - 1
        If g_MarkerPara(i) = 0 Then
            GetMinMarkerPara = 0
            Exit Function
        End If
        If g_MarkerPara(i) < nMin Then nMin = g_MarkerPara(i)
    Next i
    If nMin = 0 Then Exit Function
    GetMinMarkerPara = nMin
End Function

'===========================================================
' Append " " + citation text at the end of the paste range.
' Trims trailing whitespace / paragraph marks first so the
' citation lands cleanly. Returns the new end position.
'===========================================================
Private Function AppendPleadingCitation(oDoc As Document, _
                                         ByVal lStart As Long, _
                                         ByVal lEnd As Long, _
                                         ByVal sCit As String) As Long

    Dim lTrimmedEnd As Long
    lTrimmedEnd = lEnd
    Do While lTrimmedEnd > lStart
        Dim oCh As Range
        Set oCh = oDoc.Range(lTrimmedEnd - 1, lTrimmedEnd)
        Dim n As Long
        n = AscW(oCh.text)
        Set oCh = Nothing
        If n = 13 Or n = 11 Or n = 32 Or n = 160 Then
            Dim oDelTr As Range
            Set oDelTr = oDoc.Range(lTrimmedEnd - 1, lTrimmedEnd)
            oDelTr.Delete
            Set oDelTr = Nothing
            lTrimmedEnd = lTrimmedEnd - 1
        Else
            Exit Do
        End If
    Loop

    Dim oIns As Range
    Set oIns = oDoc.Range(lTrimmedEnd, lTrimmedEnd)
    oIns.InsertAfter " " & sCit
    Set oIns = Nothing

    AppendPleadingCitation = lTrimmedEnd + Len(sCit) + 1

End Function

'===========================================================
' Main pleading-pipeline entry point.
' Returns True if the paste was handled as a pleading paste
' (caller skips the case/statute pipeline).
' Returns False if no markers were found (caller proceeds normally).
'
' The presence of right-margin markers is itself the signal
' that this is a PDF paste, and this macro is the quote-paste
' macro (a separate macro handles the non-quote case). So we
' unconditionally:
'   1. strip the markers,
'   2. fold every paragraph mark inside the range to a space
'      (PDF line breaks are layout artifacts, not real para
'      breaks -- the user manually adds a bracketed pilcrow
'      for true paragraph boundaries in multi-paragraph pastes),
'   3. run the alternating-quote conversion chain, and
'   4. append the summarising citation.
'
' lStart, lEnd updated in place.
'===========================================================
Private Function ProcessPleadingPaste(oDoc As Document, _
                                       ByRef lStart As Long, _
                                       ByRef lEnd As Long) As Boolean

    ProcessPleadingPaste = False

    Dim oRange As Range
    Set oRange = oDoc.Range(lStart, lEnd)

    If ParseAllMarkers(oRange) = 0 Then Exit Function
    PromoteShortname

    Dim sCit As String
    sCit = BuildPleadingCitation()
    If Len(sCit) = 0 Then Exit Function

    ' Establish the tail-length invariant from this point forward.
    ' Same pattern as the main macro uses: any mutation inside
    ' [lStart, lEnd] preserves doc tail length, so lEnd recovers
    ' automatically via lEnd = oDoc.Content.End - lTailLen.
    Dim lTailLen As Long
    lTailLen = oDoc.content.End - lEnd

    ' Capture the minimum paragraph number (if paragraph mode applies)
    ' before StripAllMarkers wipes the arrays. Used a few lines down
    ' to remove a redundant leading "N. " from the passage.
    Dim nLeadPara As Integer
    nLeadPara = GetMinMarkerPara()

    ' Strip every marker before any other text mutation.
    StripAllMarkers oDoc, oRange
    lEnd = oDoc.content.End - lTailLen

    ' Fold PDF line-break paragraph marks to spaces and collapse
    ' the doubles. After this the passage is one contiguous line.
    ConvertParaMarksToSpaces oDoc, lStart, lEnd

    ' If we're in paragraph mode and the passage opens with the
    ' cited paragraph number (e.g. "4. Defendant is..."), strip
    ' that "N. " -- the citation already says "  N" so the prefix
    ' is redundant and reads awkwardly in the quoted passage.
    StripLeadingParagraphNumber oDoc, lStart, lEnd, nLeadPara

    ' Apply alternating-quote conversion to passage. This is the
    ' same chain the case-law path uses: convert straight quotes
    ' to curly, swap nested quote directions, balance unmatched
    ' quotes, then wrap the whole thing in outer doubles.
    Set oRange = oDoc.Range(lStart, lEnd)
    CurlyApostrophes oRange
    lEnd = oDoc.content.End - lTailLen

    Set oRange = oDoc.Range(lStart, lEnd)
    SwapSmartQuotes oRange
    lEnd = oDoc.content.End - lTailLen

    Set oRange = oDoc.Range(lStart, lEnd)
    BalanceNestedQuotes oRange
    lEnd = oDoc.content.End - lTailLen

    Set oRange = oDoc.Range(lStart, lEnd)
    WrapInDoubleQuotes oRange
    lEnd = oDoc.content.End - lTailLen

    ' Append the citation.
    lEnd = AppendPleadingCitation(oDoc, lStart, lEnd, sCit)

    ProcessPleadingPaste = True

End Function






