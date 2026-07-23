Attribute VB_Name = "WrapLegalCitation"
' =============================================================================
' WrapLegalCitation  (integrated with WrapCitations patterns)
' Wraps California legal citations in parentheses.
'
' PASS 1 � ITALIC-ANCHOR CITATIONS (original WrapLegalCitation logic)
'   [italic case name] [non-italic: (year) volume Cal./Cal.App.Xth page.]
'   e.g.  Smith v. Jones (2000) 123 Cal.App.4th 456.
'
'   "See" or "Cf." may be italic and part of the italic block.
'   Ibid. and Id. at p. X. are also recognized.
'   Semicolon-separated chains are wrapped in ONE pair of parentheses.
'
'   If the citation is followed by ("quoted text.") or (" quoted text."):
'     ( -> [    ) -> ]    punct before closing quote is removed
'     citation ) goes after ]
'
' PASS 2 � TEXT-PATTERN CITATIONS (from WrapCitations / DoCheckAndWrap)
'   Handles citations that may lack an italic case name anchor:
'     a) Pilcrow / section-symbol (� / �) statutes
'     b) U.S. reporter citations
'     c) "at p." / "at pp." pin-cite citations
'     d) Exhibit references  (Ex. A.)
'     e) Plain Cal. citations not caught by Pass 1
'
'   Pass 2 iterates paragraph by paragraph and applies the same sentence-
'   boundary logic used by the live keystroke macro so that the two are
'   consistent.  Already-wrapped citations are skipped.
'
' PASS 3 � EDITORIAL / QUOTED PARENTHETICAL CLEANUP
'   Inside already-wrapped citations, converts trailing parentheticals to
'   California-style brackets.  Two kinds are converted:
'     a) Recognized editorial phrases, e.g.:
'          (citations omitted)  ?  [citations omitted]
'          (internal quotation marks omitted)  ?  [internal quotation marks omitted]
'          (emphasis added)  ?  [emphasis added]   etc.
'     b) Quoted-text parentheticals, e.g.:
'          ("some quoted language.")  ?  ["some quoted language."]
'   The outer citation parentheses are NOT moved; only the inner
'   parentheticals are converted to brackets in place.
'
' HOW TO INSTALL:
'   1. Press Alt+F11 to open the VBA editor.
'   2. Go to Insert > Module.
'   3. Paste this entire code into the module.
'   4. Run via Alt+F8 > WrapLegalCitation > Run.
'
' Always run on a copy of your document first.
' =============================================================================

Option Explicit

' ?????????????????????????????????????????????????????????????????????????????
'  SHARED HELPERS
' ?????????????????????????????????????????????????????????????????????????????

Private Function IsQualifyingPunct(c As String) As Boolean
    Dim n As Long
    If Len(c) = 0 Then IsQualifyingPunct = False: Exit Function
    n = AscW(c)
    Select Case n
        Case 46, 59, 34, 8221, 8217, 8220, 8216
            IsQualifyingPunct = True
        Case Else
            IsQualifyingPunct = False
    End Select
End Function

Private Function IsItalic(oDoc As Document, pos As Long) As Boolean
    IsItalic = (oDoc.Range(pos, pos + 1).Font.Italic = True)
End Function

Private Function CharAt(oDoc As Document, pos As Long) As String
    CharAt = oDoc.Range(pos, pos + 1).text
End Function

Private Function IsClosingQuote(c As String) As Boolean
    Dim n As Long
    If Len(c) = 0 Then IsClosingQuote = False: Exit Function
    n = AscW(c)
    IsClosingQuote = (n = 34 Or n = 8221)
End Function

Private Function IsOpeningQuote(c As String) As Boolean
    Dim n As Long
    If Len(c) = 0 Then IsOpeningQuote = False: Exit Function
    n = AscW(c)
    IsOpeningQuote = (n = 34 Or n = 8220)
End Function

Private Function IsStrippablePunct(c As String) As Boolean
    Dim n As Long
    If Len(c) = 0 Then IsStrippablePunct = False: Exit Function
    n = AscW(c)
    IsStrippablePunct = (n = 46 Or n = 44 Or n = 59 Or n = 33 Or n = 63)
End Function

' ?????????????????????????????????????????????????????????????????????????????
'  PASS 1 HELPERS  (italic-anchor citation boundary detection)
' ?????????????????????????????????????????????????????????????????????????????

' Scan forward from nStart through mixed italic/non-italic text.
' Collect non-italic characters. Stop when we detect a sentence boundary
' AFTER having seen "Cal." -- boundary = space followed by capital letter,
' or a paragraph break, or a semicolon.
' Returns position of last non-space character of the citation, or 0.
' Sets bHitSemicolon = True if stopped at ";".
Private Function FindCalTerminalPos(oDoc As Document, nStart As Long, _
                                    nDocEnd As Long, _
                                    ByRef bHitSemicolon As Boolean) As Long
    Dim nPos          As Long
    Dim sChar         As String
    Dim oChar         As Range
    Dim sAfter        As String
    Dim bFoundCal     As Boolean
    Dim nLastNonSpace As Long

    sAfter = ""
    nPos = nStart
    bFoundCal = False
    nLastNonSpace = 0
    bHitSemicolon = False
    FindCalTerminalPos = 0

    Do While nPos < nDocEnd

        Set oChar = oDoc.Range(nPos, nPos + 1)
        sChar = oChar.text
        Dim nAsc As Long
        nAsc = AscW(sChar)

        If oChar.Font.Italic = True Then
            ' Hit italic text (next case name). If we found Cal. already,
            ' the citation has ended.
            If bFoundCal Then
                FindCalTerminalPos = nLastNonSpace
                Exit Do
            End If
            nPos = nPos + 1

        ElseIf nAsc = 13 Or nAsc = 11 Or nAsc = 12 Then
            ' Paragraph / line break
            If bFoundCal Then
                FindCalTerminalPos = nLastNonSpace
                Exit Do
            End If
            nPos = nPos + 1

        ElseIf bFoundCal And sChar = ";" Then
            bHitSemicolon = True
            FindCalTerminalPos = nLastNonSpace
            Exit Do

        ElseIf bFoundCal And sChar = " " Then
            ' Space -- peek ahead to see if next non-space is uppercase (sentence end)
            ' BUT skip past ( and [ since those are part of citation parentheticals
            Dim nPeek As Long
            nPeek = nPos + 1
            Dim sPeek As String
            sPeek = ""
            Do While nPeek < nDocEnd And nPeek < nPos + 10
                sPeek = CharAt(oDoc, nPeek)
                Dim nPA As Long
                nPA = AscW(sPeek)
                If sPeek = " " Or sPeek = "(" Or sPeek = "[" Then
                    nPeek = nPeek + 1
                ElseIf nPA >= 65 And nPA <= 90 Then
                    ' Capital letter -- but only a boundary if the last
                    ' non-space char was a sentence-ending character:
                    ' period, ), ], digit
                    Dim nLAsc As Long
                    If nLastNonSpace > 0 Then
                        nLAsc = AscW(CharAt(oDoc, nLastNonSpace))
                        If nLAsc = 46 Or nLAsc = 41 Or nLAsc = 93 Or _
                           (nLAsc >= 48 And nLAsc <= 57) Then
                            FindCalTerminalPos = nLastNonSpace
                            nPos = nDocEnd  ' signal to exit outer loop
                            nPeek = nDocEnd
                        End If
                    End If
                    nPeek = nDocEnd
                Else
                    ' Non-capital, non-space -- not a sentence boundary
                    nPeek = nDocEnd
                End If
            Loop

            If nPos < nDocEnd Then
                sAfter = sAfter & sChar
                nPos = nPos + 1
            End If

        Else
            sAfter = sAfter & sChar
            If sChar <> " " Then nLastNonSpace = nPos
            nPos = nPos + 1

            If Not bFoundCal Then
                If InStr(sAfter, "Cal.") > 0 Then bFoundCal = True
            End If
        End If

        If Len(sAfter) > 300 Then Exit Do

    Loop

    If FindCalTerminalPos = 0 And bFoundCal And nLastNonSpace > 0 Then
        FindCalTerminalPos = nLastNonSpace
    End If

End Function

' ?????????????????????????????????????????????????????????????????????????????
'  PASS 2 HELPERS  (text-pattern citation detection, ported from WrapCitations)
' ?????????????????????????????????????????????????????????????????????????????

' Returns the 0-based character offset within paragraph string s where the
' citation starts (i.e. the position of the first character of the citation,
' relative to the paragraph start).  Uses ". " boundaries with a 3-word
' minimum on the preceding segment, and requires the next word to be
' capitalised.  "Ex." is never treated as a sentence boundary.
Private Function GetCiteStart(s As String) As Long
    Dim lastBoundary As Long: lastBoundary = 0
    Dim segStart As Long: segStart = 1
    Dim i As Long
    For i = 1 To Len(s) - 1
        Dim bBound As Boolean: bBound = False
        Dim nxtSeg As Long: nxtSeg = 0

        ' Case 1: standard period + space
        If Mid(s, i, 2) = ". " Then
            bBound = True
            nxtSeg = i + 2

        ' Case 2: period + closing quote/paren + space (e.g.  ." C  or  .' C )
        ElseIf Mid(s, i, 1) = "." And i + 2 <= Len(s) Then
            Dim cq As Long: cq = AscW(Mid(s, i + 1, 1))
            If (cq = 8221 Or cq = 8217 Or cq = 34 Or cq = 39 Or cq = 41) Then
                If Mid(s, i + 2, 1) = " " Then
                    bBound = True
                    nxtSeg = i + 3
                End If
            End If
        End If

        If bBound Then
            ' Exception: "Ex." is never a sentence boundary
            Dim bExAbbrev As Boolean: bExAbbrev = False
            If i >= 3 Then
                If Mid(s, i - 2, 2) = "Ex" Then
                    If i = 3 Then
                        bExAbbrev = True
                    Else
                        Dim chBefore As String: chBefore = Mid(s, i - 3, 1)
                        If chBefore = " " Or chBefore = "," Or chBefore = "(" Then bExAbbrev = True
                    End If
                End If
            End If

            ' Exception: " v." (case-name separator) is never a sentence boundary
            Dim bVAbbrev As Boolean: bVAbbrev = False
            If i >= 2 Then
                If Mid(s, i - 1, 1) = "v" Then
                    If i = 2 Then
                        bVAbbrev = True
                    Else
                        Dim chBeforeV As String: chBeforeV = Mid(s, i - 2, 1)
                        If chBeforeV = " " Then bVAbbrev = True
                    End If
                End If
            End If

            If Not bExAbbrev And Not bVAbbrev Then
                If GetWordCount(Mid(s, segStart, i - segStart)) >= 3 Then
                    Dim capCheck As Long: capCheck = nxtSeg
                    If capCheck <= Len(s) Then
                        If Mid(s, capCheck, 1) Like "[A-Z]" Then
                            lastBoundary = nxtSeg - 1
                        End If
                    End If
                End If
            End If
            segStart = nxtSeg
        End If
    Next i
    GetCiteStart = lastBoundary   ' 0-based offset into paragraph (same as WrapCitations.bas)
End Function

Private Function GetWordCount(ByVal t As String) As Long
    t = Trim(t)
    If Len(t) = 0 Then GetWordCount = 0: Exit Function
    Dim count As Long: count = 1
    Dim i As Long
    For i = 1 To Len(t) - 1
        If Mid(t, i, 1) = " " And Mid(t, i + 1, 1) <> " " Then count = count + 1
    Next i
    GetWordCount = count
End Function

' Returns True if the range lS..lE is already wrapped in parentheses.
Private Function IsAlreadyWrapped(oDoc As Document, lS As Long, lE As Long) As Boolean
    IsAlreadyWrapped = False
    On Error Resume Next
    If lS >= 1 Then
        If oDoc.Range(lS - 1, lS).text = "(" Then IsAlreadyWrapped = True
    End If
    If Not IsAlreadyWrapped And lE > lS Then
        If Left(oDoc.Range(lS, lE).text, 1) = "(" Then IsAlreadyWrapped = True
    End If
    On Error GoTo 0
End Function

' Insert ( at lS and ) at lE with correct formatting.
' Returns 2 (net characters inserted) so the caller can adjust counters.
Private Function WrapRangeP2(oDoc As Document, lS As Long, lE As Long, _
                              Optional bItalicAll As Boolean = False, _
                              Optional bItalicFirst3 As Boolean = False, _
                              Optional bSkipOpen As Boolean = False) As Long
    If bItalicAll Then oDoc.Range(lS, lE).Font.Italic = True
    If bItalicFirst3 Then oDoc.Range(lS, lS + 3).Font.Italic = True

    ' Insert ) first (higher offset) to avoid index shifting
    Dim oClose As Range: Set oClose = oDoc.Range(lE, lE)
    oClose.InsertAfter ")"
    Dim oCloseChar As Range: Set oCloseChar = oDoc.Range(lE, lE + 1)
    oCloseChar.Font.Name = "Times New Roman"
    oCloseChar.Font.Size = 12
    oCloseChar.Font.Italic = False
    oCloseChar.Font.Bold = False

    If Not bSkipOpen Then
        ' Insert ( at start
        Dim oOpen As Range: Set oOpen = oDoc.Range(lS, lS)
        oOpen.InsertBefore "("
        Dim oOpenChar As Range: Set oOpenChar = oDoc.Range(lS, lS + 1)
        oOpenChar.Font.Name = "Times New Roman"
        oOpenChar.Font.Size = 12
        oOpenChar.Font.Italic = False
        oOpenChar.Font.Bold = False
        WrapRangeP2 = 2   ' two characters were inserted
    Else
        WrapRangeP2 = 1   ' only ) inserted
    End If
End Function

' PASS 2 - paragraph-by-paragraph sweep that faithfully replays the live
' keystroke macro (WrapCitations.DoCheckAndWrap) across the whole document.
'
' For each paragraph we scan every position ending in ".".  At each such
' position we build s = RTrim(text from paragraph start up to that period,
' inclusive) and run the exact same guard/pattern logic used on space-press.
' After any wrap we re-fetch the paragraph and restart the scan so that the
' newly-inserted ".)" becomes a wrap-boundary for the next pass - mirroring
' what would happen if the user had typed the paragraph interactively.

' Returns the number of additional citations wrapped.
Private Function RunPass2(oDoc As Document) As Long
    Dim nWrapped As Long: nWrapped = 0
    Dim oPar     As Paragraph

    For Each oPar In oDoc.Paragraphs

        Dim lParStart As Long
        Dim sPar      As String
        Dim lc        As Long
        Dim iPos      As Long
        Dim nIter     As Long
        Dim nAdded    As Long
        Dim bWrapped  As Boolean

        ' Per-paragraph outer loop: keep scanning until no wrap happens in
        ' a full left-to-right pass (bounded by nIter to prevent runaway).
        nIter = 0
        Do
            bWrapped = False
            nIter = nIter + 1
            If nIter > 50 Then Exit Do

            lParStart = oPar.Range.start
            sPar = oPar.Range.text
            If Len(sPar) > 0 Then
                lc = AscW(Right(sPar, 1))
                If lc = 13 Or lc = 11 Or lc = 12 Or lc = 7 Then
                    sPar = Left(sPar, Len(sPar) - 1)
                End If
            End If
            If Len(sPar) = 0 Then Exit Do

            ' Trigger only at real "space-press moments": a "." that is the
            ' last character of the paragraph OR is followed by a space.  This
            ' mirrors the live macro (which runs on the spacebar or Enter) and
            ' avoids false fires on periods inside abbreviations, decimals,
            ' or already-wrapped citations (where the "." is followed by ")").
            For iPos = 1 To Len(sPar)
                If Mid(sPar, iPos, 1) = "." Then
                    Dim bTrigger As Boolean: bTrigger = False
                    If iPos = Len(sPar) Then
                        bTrigger = True
                    ElseIf Mid(sPar, iPos + 1, 1) = " " Then
                        bTrigger = True
                    End If
                    If bTrigger Then
                        nAdded = DoCheckAndWrapAt(oDoc, lParStart, Left(sPar, iPos))
                        If nAdded > 0 Then
                            nWrapped = nWrapped + 1
                            bWrapped = True
                            Exit For   ' positions have shifted; restart scan
                        End If
                    End If
                End If
            Next iPos
        Loop While bWrapped

    Next oPar

    RunPass2 = nWrapped
End Function

' Runs the WrapCitations.DoCheckAndWrap guards and patterns against a
' simulated space-press where the "cursor" is immediately after s.
' Returns the number of characters inserted (0, 1, or 2); nonzero means a
' citation was wrapped.
Private Function DoCheckAndWrapAt(oDoc As Document, lParStart As Long, sIn As String) As Long
    DoCheckAndWrapAt = 0

    Dim s As String: s = RTrim(sIn)
    If Len(s) = 0 Or Right(s, 1) <> "." Then Exit Function

    ' Hard stops for "Ex.", "Exs.", "p.", "pp." (abbreviation trailing periods)
    If Len(s) >= 3 And Right(s, 3) = "Ex." Then Exit Function
    If Len(s) >= 4 And Right(s, 4) = "Exs." Then Exit Function
    If Len(s) >= 2 And Right(s, 2) = "p." Then Exit Function
    If Len(s) >= 3 And Right(s, 3) = "pp." Then Exit Function

    ' Complete exhibit reference (used later for Pattern 3c guardrail)
    Dim bExhibit As Boolean: bExhibit = False
    If Len(s) >= 7 Then
        If Right(s, 7) Like " Ex. [A-Z]." Then bExhibit = True
    End If

    ' Quote depth: never wrap mid-quotation
    ' Straight double quotes (AscW 34) are direction-ambiguous -- the same
    ' character opens and closes -- so they are tracked by parity: the flag
    ' toggles on each one, and an odd count (flag still True) means a straight
    ' quote is open.  Curly singles get their own depth so a right single
    ' quote (8217) only closes a pending 8216 opener; with no opener pending
    ' it is an apostrophe (e.g. "plaintiff's") and must be ignored rather
    ' than counted as a closer.
    Dim qDepth As Long, qSingle As Long, qi As Long
    Dim bStraightOpen As Boolean: bStraightOpen = False
    For qi = 1 To Len(s)
        Dim qc As Long: qc = AscW(Mid(s, qi, 1))
        If qc = 34 Then
            bStraightOpen = Not bStraightOpen
        ElseIf qc = 8220 Then
            qDepth = qDepth + 1
        ElseIf qc = 8221 Then
            If qDepth > 0 Then qDepth = qDepth - 1
        ElseIf qc = 8216 Then
            qSingle = qSingle + 1
        ElseIf qc = 8217 Then
            If qSingle > 0 Then qSingle = qSingle - 1
        End If
    Next qi
    If bStraightOpen Then qDepth = qDepth + 1
    If qDepth + qSingle > 0 Then Exit Function

    ' Paren balance: if paragraph has unmatched "(", close with ")" only.
    Dim bSkipOpen As Boolean: bSkipOpen = False
    Dim nOpen As Long, nClose As Long, pi As Long
    For pi = 1 To Len(s)
        Select Case Mid(s, pi, 1)
            Case "(": nOpen = nOpen + 1
            Case ")": nClose = nClose + 1
        End Select
    Next pi
    If nOpen > nClose Then bSkipOpen = True

    Dim lS As Long, lE As Long, lOff As Long

    ' WRAPPED-CITATION BOUNDARY GUARD: ignore everything up to the last
    ' ".)" or right-double-quote that ends a previously wrapped citation.
    Dim lBoundary As Long: lBoundary = 0
    Dim iBound As Long
    For iBound = Len(s) - 1 To 2 Step -1
        If Mid(s, iBound, 2) = ".)" Then
            lBoundary = iBound + 2: Exit For
        ElseIf AscW(Mid(s, iBound, 1)) = 8221 Then
            lBoundary = iBound + 1: Exit For
        End If
    Next iBound
    Dim sSearch As String
    If lBoundary > 0 And lBoundary <= Len(s) Then
        sSearch = Mid(s, lBoundary)
    Else
        sSearch = s
    End If

    ' Pattern 1: Ibid.
    If Right(s, 5) = "Ibid." Then
        lS = lParStart + Len(s) - 5: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then
            DoCheckAndWrapAt = WrapRangeP2(oDoc, lS, lE, True, False, bSkipOpen)
        End If
        Exit Function
    End If

    ' Pattern 2: Id. at
    If InStr(sSearch, "Id. at") > 0 Then
        Dim lId As Long: lId = InStrRev(s, "Id. at")
        lS = lParStart + lId - 1: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then
            DoCheckAndWrapAt = WrapRangeP2(oDoc, lS, lE, False, True, bSkipOpen)
        End If
        Exit Function
    End If

    ' Pattern 3a: pilcrow
    If InStr(sSearch, ChrW(182)) > 0 Then
        lOff = GetCiteStart(s)
        lS = lParStart + lOff: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then
            DoCheckAndWrapAt = WrapRangeP2(oDoc, lS, lE, False, False, bSkipOpen)
        End If
        Exit Function
    End If

    ' Pattern 3b: " at p." / " at pp."
    If InStr(sSearch, " at p.") > 0 Or InStr(sSearch, " at pp.") > 0 Then
        Dim lAtP As Long
        lAtP = InStrRev(s, " at pp.")
        If lAtP = 0 Then lAtP = InStrRev(s, " at p.")
        If lAtP > 1 Then
            lOff = GetCiteStart(Left(s, lAtP - 1))
        Else
            lOff = 0
        End If
        lS = lParStart + lOff: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then
            DoCheckAndWrapAt = WrapRangeP2(oDoc, lS, lE, False, False, bSkipOpen)
        End If
        Exit Function
    End If

    ' Pattern 3c: Cal., U.S., sec, or a bare exhibit reference (header item
    ' (d)).  Reporter/section citations require a trailing digit/")"; an
    ' exhibit reference (bExhibit) reaches the wrap logic even when no
    ' reporter/section marker is present, and keeps its existing
    ' trailing-digit bypass.
    If InStr(sSearch, "Cal.") > 0 Or InStr(sSearch, "U.S.") > 0 Or InStr(sSearch, ChrW(167)) > 0 Or bExhibit Then
        Dim charBefore As String: charBefore = Mid(s, Len(s) - 1, 1)
        If (charBefore Like "[0-9)]") Or bExhibit Then
            lOff = GetCiteStart(s)
            lS = lParStart + lOff: lE = lParStart + Len(s)
            If Not IsAlreadyWrapped(oDoc, lS, lE) Then
                DoCheckAndWrapAt = WrapRangeP2(oDoc, lS, lE, False, False, bSkipOpen)
            End If
        End If
    End If
End Function


' ?????????????????????????????????????????????????????????????????????????????
'  PASS 3 HELPERS  (editorial / quoted parenthetical bracket conversion)
' ?????????????????????????????????????????????????????????????????????????????

Private Function IsEditorialPhrase(s As String) As Boolean
    ' Returns True if s (trimmed, case-insensitive) is a recognized editorial
    ' phrase used in California legal citations.
    Dim phrases(29) As String
    phrases(0) = "citations omitted"
    phrases(1) = "citation omitted"
    phrases(2) = "internal quotation marks omitted"
    phrases(3) = "internal citations omitted"
    phrases(4) = "internal citations and quotation marks omitted"
    phrases(5) = "citation and internal quotation marks omitted"
    phrases(6) = "citations and internal quotation marks omitted"
    phrases(7) = "footnotes omitted"
    phrases(8) = "footnote omitted"
    phrases(9) = "alterations omitted"
    phrases(10) = "alteration omitted"
    phrases(11) = "emphasis added"
    phrases(12) = "emphasis omitted"
    phrases(13) = "emphasis in original"
    phrases(14) = "brackets omitted"
    phrases(15) = "bracket omitted"
    phrases(16) = "original brackets omitted"
    phrases(17) = "ellipsis omitted"
    phrases(18) = "ellipses omitted"
    phrases(19) = "capitalization omitted"
    phrases(20) = "some capitalization omitted"
    phrases(21) = "punctuation omitted"
    phrases(22) = "quotation marks omitted"
    phrases(23) = "quotation marks and citations omitted"
    phrases(24) = "omission marks omitted"
    phrases(25) = "internal quotation marks and citations omitted"
    phrases(26) = "internal quotation marks, citations, and footnotes omitted"
    phrases(27) = "italics omitted"
    phrases(28) = "italics in original"
    phrases(29) = "cleaned up"

    Dim sLower As String: sLower = LCase(Trim(s))
    Dim i As Long
    For i = 0 To UBound(phrases)
        If sLower = phrases(i) Then
            IsEditorialPhrase = True
            Exit Function
        End If
    Next i
    IsEditorialPhrase = False
End Function

' Returns True if the content between ( and ) is a quoted-text parenthetical,
' i.e. it starts with an opening quote character (straight or curly).
Private Function IsQuotedParenthetical(sContent As String) As Boolean
    Dim sT As String: sT = LTrim(sContent)
    If Len(sT) = 0 Then IsQuotedParenthetical = False: Exit Function
    Dim n As Long: n = AscW(Left(sT, 1))
    IsQuotedParenthetical = (n = 34 Or n = 8220 Or n = 8216)
End Function

' Pass 3: walk every paragraph, find wrapped citations (outer parens that look
' like citations), then convert any trailing editorial or quoted-text
' parentheticals inside them to brackets.  Outer parens stay put.
' Returns the number of individual parentheticals converted.
Private Function RunPass3(oDoc As Document) As Long
    ' ?? All variables declared at function scope (VBA requirement) ??
    Dim nFixed      As Long
    Dim oPar        As Paragraph
    Dim lParStart   As Long
    Dim sText       As String
    Dim lc          As Long
    Dim iSearch     As Long
    Dim iClose      As Long
    Dim nDepth      As Long
    Dim iScan       As Long
    Dim ch          As String
    Dim sInner      As String
    Dim bCite       As Boolean
    Dim iY          As Long
    Dim sYr         As String
    Dim sWork       As String
    Dim bTrailDot   As Boolean
    Dim editStarts(20) As Long
    Dim editEnds(20)   As Long
    Dim nEdit       As Long
    Dim iRight      As Long
    Dim iEClose     As Long
    Dim iEOpen      As Long
    Dim iB          As Long
    Dim nID         As Long
    Dim sContent    As String
    Dim iEdit       As Long
    Dim docOpen     As Long
    Dim docClose    As Long
    Dim oRP3        As Range

    nFixed = 0

    For Each oPar In oDoc.Paragraphs

        lParStart = oPar.Range.start
        sText = oPar.Range.text

        ' Strip paragraph mark
        If Len(sText) > 0 Then
            lc = AscW(Right(sText, 1))
            If lc = 13 Or lc = 11 Or lc = 12 Or lc = 7 Then
                sText = Left(sText, Len(sText) - 1)
            End If
        End If
        If Len(sText) = 0 Then GoTo P3NextPar

        ' Scan every ( in the paragraph
        For iSearch = 1 To Len(sText)

            If Mid(sText, iSearch, 1) <> "(" Then GoTo P3NextChar

            ' Find matching outer closing )
            iClose = 0
            nDepth = 1
            For iScan = iSearch + 1 To Len(sText)
                ch = Mid(sText, iScan, 1)
                If ch = "(" Then
                    nDepth = nDepth + 1
                ElseIf ch = ")" Then
                    nDepth = nDepth - 1
                    If nDepth = 0 Then
                        iClose = iScan
                        Exit For
                    End If
                End If
            Next iScan
            If iClose = 0 Then GoTo P3NextChar

            ' Quick citation filter
            sInner = Mid(sText, iSearch + 1, iClose - iSearch - 1)
            bCite = False
            If InStr(sInner, "Cal.") > 0 Then bCite = True
            If InStr(sInner, "U.S.") > 0 Then bCite = True
            If InStr(sInner, "Ibid.") > 0 Then bCite = True
            If InStr(sInner, "Id.") > 0 Then bCite = True
            If Not bCite Then
                For iY = 1 To Len(sInner) - 5
                    If Mid(sInner, iY, 1) = "(" Then
                        sYr = Mid(sInner, iY + 1, 4)
                        If sYr Like "19##" Or sYr Like "20##" Then
                            bCite = True: Exit For
                        End If
                    End If
                Next iY
            End If
            If Not bCite Then GoTo P3NextChar

            ' ?? Collect trailing editorial/quoted parentheticals inside sInner ??
            sWork = sInner
            bTrailDot = False
            If Right(sWork, 1) = "." Then
                bTrailDot = True
                sWork = Left(sWork, Len(sWork) - 1)
            End If
            sWork = RTrim(sWork)

            nEdit = 0

            iRight = Len(sWork)
            Do
                Do While iRight >= 1 And Mid(sWork, iRight, 1) = " "
                    iRight = iRight - 1
                Loop
                If iRight < 1 Then Exit Do
                If Mid(sWork, iRight, 1) <> ")" Then Exit Do

                iEClose = iRight
                iEOpen = 0
                nID = 1
                For iB = iEClose - 1 To 1 Step -1
                    If Mid(sWork, iB, 1) = ")" Then
                        nID = nID + 1
                    ElseIf Mid(sWork, iB, 1) = "(" Then
                        nID = nID - 1
                        If nID = 0 Then
                            iEOpen = iB
                            Exit For
                        End If
                    End If
                Next iB
                If iEOpen = 0 Then Exit Do

                sContent = Mid(sWork, iEOpen + 1, iEClose - iEOpen - 1)

                If Not IsEditorialPhrase(sContent) And Not IsQuotedParenthetical(sContent) Then Exit Do

                editStarts(nEdit) = iEOpen
                editEnds(nEdit) = iEClose
                nEdit = nEdit + 1

                iRight = iEOpen - 1
            Loop

            If nEdit = 0 Then GoTo P3NextChar

            ' ?? Apply bracket conversions RIGHT TO LEFT ??
            ' Single-char replacements: no offset shift needed.
            For iEdit = 0 To nEdit - 1
                docOpen = lParStart + iSearch + editStarts(iEdit) - 1
                docClose = lParStart + iSearch + editEnds(iEdit) - 1

                Set oRP3 = oDoc.Range(docClose, docClose + 1)
                oRP3.text = "]"

                Set oRP3 = oDoc.Range(docOpen, docOpen + 1)
                oRP3.text = "["

                nFixed = nFixed + 1
            Next iEdit

            ' Re-read paragraph text to stay current for multi-cite paragraphs
            sText = oPar.Range.text
            If Len(sText) > 0 Then
                lc = AscW(Right(sText, 1))
                If lc = 13 Or lc = 11 Or lc = 12 Or lc = 7 Then
                    sText = Left(sText, Len(sText) - 1)
                End If
            End If

P3NextChar:
        Next iSearch

P3NextPar:
    Next oPar

    RunPass3 = nFixed
End Function

' ?????????????????????????????????????????????????????????????????????????????
'  MAIN ENTRY POINT
' ?????????????????????????????????????????????????????????????????????????????

Sub WrapLegalCitation()

    Dim oDoc     As Document
    Dim oSearch  As Range
    Dim nWrapped As Long
    Dim nDocEnd  As Long
    Dim oUndo    As UndoRecord

    Set oDoc = ActiveDocument

    ' Track Changes off for the run, restored in CleanUp: with revisions on,
    ' deletions stay in the position stream and every hard-coded offset
    ' adjustment in the passes lands wrong.
    Dim bPrevTrack As Boolean
    bPrevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False

    Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "Wrap Legal Citations"

    ' Any runtime error must still land on CleanUp: otherwise ScreenUpdating
    ' stays off and the custom undo record stays open -- and an open record
    ' makes the next spacebar press fail inside CheckAndWrap, silently killing
    ' citation wrapping for the rest of the session.
    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    nWrapped = 0

    ' =========================================================================
    '  PASS 1 � ITALIC-ANCHOR SWEEP  (original logic, unchanged)
    ' =========================================================================

    Set oSearch = oDoc.content
    oSearch.Find.ClearFormatting
    oSearch.Find.Font.Italic = True
    oSearch.Find.text = ""
    oSearch.Find.Forward = True
    oSearch.Find.Wrap = wdFindStop
    oSearch.Find.Format = True
    oSearch.Find.MatchCase = False
    oSearch.Find.MatchWildcards = False

    Do While oSearch.Find.Execute

        ' Refresh the document end on every iteration: each wrap below inserts
        ' characters, so a snapshot taken once before the loop goes stale and
        ' the character walks (and FindCalTerminalPos, which receives nDocEnd)
        ' would stop short of the real end, mangling or skipping a trailing
        ' citation.  No insertions happen between here and those uses, so the
        ' value stays current for the whole iteration.
        nDocEnd = oDoc.content.End - 1

        Dim nBlockStart  As Long
        Dim nBlockEnd    As Long
        Dim nPos         As Long
        Dim sChar        As String
        Dim bQualifies   As Boolean
        Dim bAlready     As Boolean
        Dim bAtParaStart As Boolean
        Dim oPar         As Paragraph
        Dim nParStart    As Long
        Dim nOpenPos     As Long
        Dim sBlockText   As String
        Dim sBlockCore   As String
        Dim bIsIbid      As Boolean
        Dim bIsId        As Boolean

        ' --- Expand to full contiguous italic block ---
        nBlockStart = oSearch.start
        nBlockEnd = oSearch.End

        nPos = nBlockStart - 1
        Do While nPos >= 0
            If IsItalic(oDoc, nPos) Then
                nBlockStart = nPos
                nPos = nPos - 1
            Else
                nPos = -1
            End If
        Loop

        nPos = nBlockEnd
        Do While nPos < nDocEnd
            If IsItalic(oDoc, nPos) Then
                nBlockEnd = nPos + 1
                nPos = nPos + 1
            Else
                nPos = nDocEnd + 1
            End If
        Loop

        sBlockText = oDoc.Range(nBlockStart, nBlockEnd).text

        ' Strip See/Cf for type detection
        sBlockCore = sBlockText
        If Left(sBlockCore, 4) = "See " Then sBlockCore = Mid(sBlockCore, 5)
        If Left(sBlockCore, 4) = "Cf. " Then sBlockCore = Mid(sBlockCore, 5)

        bIsIbid = (sBlockCore = "Ibid.")
        bIsId = (Left(sBlockCore, 3) = "Id." And Len(sBlockCore) = 3)

        ' --- Determine nOpenPos ---
        nOpenPos = nBlockStart
        If Left(sBlockText, 4) <> "See " And Left(sBlockText, 4) <> "Cf. " Then
            Dim s5Before As String
            s5Before = ""
            Dim nb As Long
            For nb = 5 To 1 Step -1
                If nBlockStart - nb >= 0 Then
                    s5Before = s5Before & CharAt(oDoc, nBlockStart - nb)
                End If
            Next nb
            If Right(s5Before, 4) = "See " Then
                nOpenPos = nBlockStart - 4
            ElseIf Right(s5Before, 4) = "Cf. " Then
                nOpenPos = nBlockStart - 4
            End If
        End If

        ' --- Check not already wrapped ---
        bAlready = False
        nPos = nOpenPos - 1
        Dim nFoundParen As Long
        nFoundParen = -1
        Do While nPos >= 0
            sChar = CharAt(oDoc, nPos)
            If sChar = " " Then
                nPos = nPos - 1
            ElseIf sChar = "(" Then
                nFoundParen = nPos
                nPos = -1
            Else
                nPos = -1
            End If
        Loop

        If nFoundParen >= 0 Then
            If nFoundParen = 0 Then
                bAlready = True
            Else
                Dim nBPAsc As Long
                nBPAsc = AscW(CharAt(oDoc, nFoundParen - 1))
                If nBPAsc = 32 Or nBPAsc = 13 Or nBPAsc = 11 Then
                    bAlready = True
                End If
            End If
        End If

        If Not bAlready Then

            ' --- Check preceding character qualifies ---
            bQualifies = False
            Set oPar = oDoc.Range(nOpenPos, nOpenPos).Paragraphs(1)
            nParStart = oPar.Range.start

            bAtParaStart = True
            nPos = nOpenPos - 1
            Do While nPos >= nParStart
                sChar = CharAt(oDoc, nPos)
                If sChar <> " " Then
                    bAtParaStart = False
                    nPos = nParStart - 1
                End If
                nPos = nPos - 1
            Loop

            If bAtParaStart Then
                bQualifies = True
            Else
                nPos = nOpenPos - 1
                Do While nPos >= nParStart
                    sChar = CharAt(oDoc, nPos)
                    If sChar <> " " Then
                        If IsQualifyingPunct(sChar) Or sChar = "(" Then
                            bQualifies = True
                        End If
                        nPos = nParStart - 1
                    End If
                    nPos = nPos - 1
                Loop
            End If

            If bQualifies Then

                Dim nTermAbs   As Long
                Dim bSemicolon As Boolean
                nTermAbs = 0
                bSemicolon = False

                If bIsIbid Then
                    Dim sAfterIbid As String
                    sAfterIbid = ""
                    If nBlockEnd < nDocEnd Then sAfterIbid = CharAt(oDoc, nBlockEnd)
                    If sAfterIbid = ";" Then
                        nTermAbs = FindCalTerminalPos(oDoc, nBlockEnd + 1, nDocEnd, bSemicolon)
                    Else
                        nTermAbs = nBlockEnd - 1
                    End If

                ElseIf bIsId Then
                    Dim sAfterId  As String
                    Dim bFoundAtP As Boolean
                    Dim oCharId   As Range
                    sAfterId = ""
                    bFoundAtP = False
                    nPos = nBlockEnd
                    Do While nPos < nDocEnd And nTermAbs = 0
                        Set oCharId = oDoc.Range(nPos, nPos + 1)
                        sChar = oCharId.text
                        If oCharId.Font.Italic = True Then
                            nPos = nPos + 1
                        Else
                            sAfterId = sAfterId & sChar
                            nPos = nPos + 1
                            If Not bFoundAtP Then
                                If InStr(sAfterId, "at p.") > 0 Then bFoundAtP = True
                            End If
                            If bFoundAtP Then
                                If sChar = "." Then
                                    Dim sPrev As String
                                    sPrev = ""
                                    If Len(sAfterId) >= 2 Then
                                        sPrev = Mid(sAfterId, Len(sAfterId) - 1, 1)
                                    End If
                                    If sPrev >= "0" And sPrev <= "9" Then
                                        nTermAbs = nPos - 1
                                    End If
                                ElseIf sChar = ";" Then
                                    nTermAbs = FindCalTerminalPos(oDoc, nPos, nDocEnd, bSemicolon)
                                End If
                            End If
                            If Len(sAfterId) > 100 Then Exit Do
                        End If
                    Loop

                Else
                    nTermAbs = FindCalTerminalPos(oDoc, nBlockEnd, nDocEnd, bSemicolon)
                End If

                If nTermAbs > 0 Then

                    ' --- Check for quoted parenthetical after nTermAbs ---
                    Dim bHasQuote        As Boolean
                    Dim nQuoteOpenParen  As Long
                    Dim nQuoteCloseQuote As Long
                    Dim nQuoteCloseParen As Long

                    bHasQuote = False
                    nQuoteOpenParen = 0
                    nQuoteCloseQuote = 0
                    nQuoteCloseParen = 0

                    nPos = nTermAbs + 1
                    Do While nPos < nDocEnd And nPos <= nTermAbs + 10
                        sChar = CharAt(oDoc, nPos)
                        If sChar = " " Then
                            nPos = nPos + 1
                        ElseIf sChar = "(" Then
                            Dim nAfterParen As Long
                            nAfterParen = nPos + 1
                            Do While nAfterParen < nDocEnd And nAfterParen < nPos + 5
                                Dim sAP As String
                                sAP = CharAt(oDoc, nAfterParen)
                                If sAP = " " Then
                                    nAfterParen = nAfterParen + 1
                                ElseIf IsOpeningQuote(sAP) Then
                                    nQuoteOpenParen = nPos
                                    Dim nQPos As Long
                                    nQPos = nAfterParen + 1
                                    Do While nQPos < nDocEnd And nQPos < nPos + 2000
                                        Dim sQC As String
                                        sQC = CharAt(oDoc, nQPos)
                                        If IsClosingQuote(sQC) Then
                                            nQuoteCloseQuote = nQPos
                                            Dim nCP As Long
                                            nCP = nQPos + 1
                                            Do While nCP < nDocEnd And nCP < nQPos + 5
                                                Dim sCP As String
                                                sCP = CharAt(oDoc, nCP)
                                                If sCP = ")" Then
                                                    nQuoteCloseParen = nCP
                                                    bHasQuote = True
                                                    nCP = nQPos + 5
                                                ElseIf sCP = " " Then
                                                    nCP = nCP + 1
                                                Else
                                                    nCP = nQPos + 5
                                                End If
                                            Loop
                                            nQPos = nPos + 2000
                                        Else
                                            nQPos = nQPos + 1
                                        End If
                                    Loop
                                    nAfterParen = nPos + 5
                                Else
                                    nAfterParen = nPos + 5
                                End If
                            Loop
                            nPos = nTermAbs + 11
                        Else
                            nPos = nTermAbs + 11
                        End If
                    Loop

                    ' --- Insert/replace RIGHT TO LEFT ---
                    Dim oInsert As Range

                    If bHasQuote Then

                        ' 1. Replace ) with ] at nQuoteCloseParen
                        Set oInsert = oDoc.Range(nQuoteCloseParen, nQuoteCloseParen + 1)
                        oInsert.text = "]"

                        ' 2. Remove strippable punct immediately before closing quote
                        If IsStrippablePunct(CharAt(oDoc, nQuoteCloseQuote - 1)) Then
                            Set oInsert = oDoc.Range(nQuoteCloseQuote - 1, nQuoteCloseQuote)
                            oInsert.text = ""
                            nQuoteCloseParen = nQuoteCloseParen - 1
                        End If

                        ' 3. Insert citation ) after ]
                        Set oInsert = oDoc.Range(nQuoteCloseParen + 1, nQuoteCloseParen + 1)
                        oInsert.InsertAfter ")"
                        Set oInsert = oDoc.Range(nQuoteCloseParen + 1, nQuoteCloseParen + 2)
                        oInsert.Font.Italic = False

                        ' 4. Replace ( with [
                        Set oInsert = oDoc.Range(nQuoteOpenParen, nQuoteOpenParen + 1)
                        oInsert.text = "["

                        ' 5. Insert citation ( at nOpenPos
                        Set oInsert = oDoc.Range(nOpenPos, nOpenPos)
                        oInsert.InsertBefore "("
                        Set oInsert = oDoc.Range(nOpenPos, nOpenPos + 1)
                        oInsert.Font.Italic = False

                    Else

                        ' Standard: ) after terminal pos, ( at open pos
                        Set oInsert = oDoc.Range(nTermAbs + 1, nTermAbs + 1)
                        oInsert.InsertAfter ")"
                        Set oInsert = oDoc.Range(nTermAbs + 1, nTermAbs + 2)
                        oInsert.Font.Italic = False

                        Set oInsert = oDoc.Range(nOpenPos, nOpenPos)
                        oInsert.InsertBefore "("
                        Set oInsert = oDoc.Range(nOpenPos, nOpenPos + 1)
                        oInsert.Font.Italic = False

                    End If

                    nWrapped = nWrapped + 1

                End If
            End If
        End If

        oSearch.Collapse wdCollapseEnd

    Loop

    ' =========================================================================
    '  PASS 2 � TEXT-PATTERN SWEEP  (ported from WrapCitations / DoCheckAndWrap)
    ' =========================================================================

    Dim nPass2 As Long
    nPass2 = RunPass2(oDoc)
    nWrapped = nWrapped + nPass2

    ' =========================================================================
    '  PASS 3 � EDITORIAL / QUOTED PARENTHETICAL BRACKET CONVERSION
    '  Runs after Pass 1 & 2 so all citations are already wrapped before we
    '  look for inner parentheticals to convert.
    ' =========================================================================

    Dim nPass3 As Long
    nPass3 = RunPass3(oDoc)

CleanUp:
    Dim lErrNum As Long, sErrDesc As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    On Error Resume Next
    oUndo.EndCustomRecord
    oDoc.TrackRevisions = bPrevTrack
    Application.ScreenUpdating = True
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Wrap Legal Citations hit an error and stopped:" & vbCrLf & vbCrLf & _
               "Error " & lErrNum & ": " & sErrDesc, _
               vbExclamation, "Wrap Legal Citations"
    Else
        MsgBox "Done." & vbCrLf & _
               "  Pass 1 (italic-anchor):      " & (nWrapped - nPass2) & " citation(s) wrapped" & vbCrLf & _
               "  Pass 2 (text-pattern):       " & nPass2 & " citation(s) wrapped" & vbCrLf & _
               "  Pass 3 (bracket conversion): " & nPass3 & " parenthetical(s) converted", _
               vbInformation, "Wrap Legal Citations"
    End If

End Sub




