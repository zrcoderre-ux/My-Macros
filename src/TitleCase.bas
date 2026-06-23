Attribute VB_Name = "TitleCase"
' ============================================================
' TitleCase Macro for Microsoft Word  (patched)
'
' Applies legal/title-case capitalization rules:
'   - Capitalizes all words by default
'   - Lowercases articles: a, an, the
'   - Lowercases coordinating conjunctions: for, and, nor, but, or, yet, so
'   - Lowercases prepositions of 4 letters or fewer
'   - Lowercases words preceded by a connective hyphen (e.g. "murder" in "Felony-murder")
'   - Always capitalizes the first and last word
'   - Always capitalizes the first word after a colon, em dash, en dash, or double hyphen
'
' Pre-processing (runs before title-casing):
'   - Replaces paragraph marks with spaces (in string-space, not on the live document)
'   - Converts numbered list items (e.g. "1. " or "15. ") to parenthetical form ("(1) ", "(15) ")
'     for single-digit (1-9) and two-digit (10-29) numbers, joined with "; " and "; and "
'
' Behavior:
'   - If text is highlighted, applies pre-processing + title-case to the selection in place.
'   - If no text is highlighted, pastes the clipboard contents at the cursor
'     and applies pre-processing + title-case to the pasted text.
'   - Output is always formatted as Times New Roman, 12pt.
'   - No external library references required.
'
' --- CHANGES FROM PRIOR VERSION (crash fixes) ---------------------------------
'   [#1] One error handler now covers BOTH branches, and EndCustomRecord is
'        wrapped so it always runs. A thrown error can no longer leave an open
'        UndoRecord that crashes a later run.
'   [#2] Paragraph-mark handling and list joins now happen entirely in a VBA
'        string. The document is read once and written once. A single trailing
'        paragraph/cell mark is excluded from the range before write-back, so the
'        macro no longer deletes a structural paragraph mark.
'        NOTE: the result now stays its OWN paragraph instead of merging into the
'        following paragraph (the old version merged it as a side effect).
'   [+]  Final selection comes from the Range object, not Len()-based position
'        arithmetic, which removes the off-by-N range risk.
'   [+]  IsWhitespaceChar / EndsWithBreak use AscW instead of Asc, so a character
'        outside the system code page can no longer raise a runtime error.
' ------------------------------------------------------------------------------

Option Explicit

Sub ApplyTitleCase()
    Dim oSel As Selection
    Set oSel = Selection

    ' --- Wrap everything in one named undo record so Ctrl+Z reverses it all at once ---
    Dim oUndo As UndoRecord
    Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "Apply Title Case"

    ' [#1] Single handler covering BOTH branches. Any error routes to CleanUp,
    ' which always closes the undo record.
    On Error GoTo CleanUp

    Dim mutRng As Range

    If oSel.Type = wdSelectionIP Then
        ' --- No selection: paste clipboard as plain text, then process it ---
        Dim pasteStart As Long
        pasteStart = oSel.start

        On Error GoTo NoClipboard
        oSel.PasteSpecial DataType:=wdPasteText
        On Error GoTo CleanUp

        Dim pasteEnd As Long
        pasteEnd = Selection.start

        If pasteEnd <= pasteStart Then
            MsgBox "The clipboard appears to be empty or contains no plain text. " & _
                   "Please copy some text and try again.", _
                   vbInformation, "Nothing to Process"
            GoTo CleanUp
        End If

        Set mutRng = ActiveDocument.Range(pasteStart, pasteEnd)
    Else
        ' --- Text is highlighted ---
        Set mutRng = oSel.Range.Duplicate
    End If

    ' [#2] Exclude a single trailing paragraph/cell mark so the .Text write-back
    ' below cannot delete a structural mark.
    TrimTrailingMark mutRng

    If mutRng.End <= mutRng.start Then GoTo CleanUp   ' nothing left to process

    ' --- Read once, transform in string-space, write back once ---
    Dim rawText As String
    rawText = mutRng.text

    Dim processed As String
    processed = PreProcessText(rawText)        ' paragraph joins + list normalization
    processed = TitleCaseString(processed)     ' casing

    mutRng.text = processed                    ' single write-back; mutRng auto-extends

    mutRng.Select
    ApplyFormatting Selection
    Selection.Collapse wdCollapseEnd

CleanUp:
    On Error Resume Next
    oUndo.EndCustomRecord
    Exit Sub

NoClipboard:
    MsgBox "Could not paste from the clipboard. Please try copying your text again.", _
           vbExclamation, "Clipboard Error"
    Resume CleanUp
End Sub


' ============================================================
' [#2] Excludes ONE trailing paragraph mark (Chr 13) or cell
' mark (Chr 7) from the range's span, so a later .Text
' assignment replaces content only and leaves the structural
' mark in place.
' ============================================================
Sub TrimTrailingMark(rng As Range)
    If rng.End <= rng.start Then Exit Sub
    Dim t As String
    t = rng.text
    If Len(t) = 0 Then Exit Sub
    Dim c As Long
    c = AscW(Right(t, 1))
    If c = 13 Or c = 7 Then
        rng.End = rng.End - 1
    End If
End Sub


' ============================================================
' [#2] Pure string pre-processing (no document interaction).
' Replaces paragraph/line/cell marks with spaces, then
' conditionally converts numbered list markers to parenthetical
' form with semicolons and "and" before the last item.
'
' Recognised input formats (N = 1-29, sequential from 1, no gaps):
'   Format A: "N. xyz"   — number, period, space
'   Format B: "N.xyz"    — number, period, no space (content immediately follows)
'   Format C: "N) xyz"   — number, closing paren, space
'   Format D: "(N) xyz"  — already wrapped (semicolon/and logic only)
'
' Rules:
'   - Formats are consistent within one run (not mixed)
'   - Sequence must start at 1 and be unbroken (1,2,3... no gaps)
'   - Only applies if the sequential run from 1 is 3 or more items
'   - After wrapping, "; " replaces the space before each "(N)"
'     from the 2nd through 2nd-to-last; "; and " before the last
'   - Result pattern: (1) foo; (2) bar; and (3) baz
' ============================================================
Function PreProcessText(ByVal s As String) As String

    ' Replace paragraph/line/cell marks with spaces (string-space only).
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(11), " ")   ' manual line break
    s = Replace(s, Chr(7), " ")    ' cell mark (defensive)

    ' -------------------------------------------------------
    ' Detect format and measure sequential run from 1.
    ' fmtCode: 0=none 1="N. " 2="N." 3="N) " 4="(N) "
    ' -------------------------------------------------------
    Dim fmtCode As Integer
    fmtCode = 0
    Dim seqCount As Integer
    seqCount = 0

    Dim fmt As Integer
    For fmt = 1 To 4
        Dim runLen As Integer
        runLen = 0
        Dim n As Integer
        For n = 1 To 29
            If ListItemExists(s, n, fmt) Then
                runLen = runLen + 1
            Else
                Exit For
            End If
        Next n
        If runLen >= 3 Then
            fmtCode = fmt
            seqCount = runLen
            Exit For
        End If
    Next fmt

    If fmtCode = 0 Then
        PreProcessText = s      ' no qualifying sequence
        Exit Function
    End If

    ' -------------------------------------------------------
    ' Normalise formats 1, 2, 3 to format 4 ("(N) ").
    ' Replace largest numbers first to avoid substring collisions
    ' (e.g. replacing "1." before "10." would corrupt "10.").
    ' -------------------------------------------------------
    Dim i As Integer

    If fmtCode <> 4 Then
        For n = seqCount To 1 Step -1
            Dim newMark As String
            newMark = "(" & CStr(n) & ") "
            Dim p As Long
            p = 1
            Do
                p = FindListItem(s, n, fmtCode, p)
                If p = 0 Then Exit Do
                Dim oldLen As Integer
                oldLen = ListItemLen(n, fmtCode)
                s = Left(s, p - 1) & newMark & Mid(s, p + oldLen)
                p = p + Len(newMark)
            Loop
        Next n
    End If

    ' -------------------------------------------------------
    ' Insert "; " before items (2)..(seqCount-1) and
    ' "; and " before item (seqCount).
    ' -------------------------------------------------------
    Dim sepPositions() As Long
    ReDim sepPositions(seqCount - 2)
    Dim sepCount As Integer
    sepCount = 0

    For n = 2 To seqCount
        Dim marker As String
        marker = " (" & CStr(n) & ") "
        Dim mp As Long
        mp = InStr(s, marker)
        If mp > 0 Then
            sepPositions(sepCount) = mp
            sepCount = sepCount + 1
        End If
    Next n

    ' Sort ascending (bubble sort — small array)
    Dim j As Integer
    Dim tmp As Long
    For i = 0 To sepCount - 2
        For j = i + 1 To sepCount - 1
            If sepPositions(j) < sepPositions(i) Then
                tmp = sepPositions(j)
                sepPositions(j) = sepPositions(i)
                sepPositions(i) = tmp
            End If
        Next j
    Next i

    ' Process right-to-left so earlier positions stay valid
    Dim sp As Long
    Dim sep As String
    Dim prefix As String
    For i = sepCount - 1 To 0 Step -1
        sp = sepPositions(i)
        If i = sepCount - 1 Then
            sep = "; and "
        Else
            sep = "; "
        End If
        ' Skip if already separated.
        prefix = Left(s, sp - 1)
        If Len(prefix) >= 1 Then
            If Right(prefix, 1) = ";" Then GoTo NextSep
        End If
        If Len(prefix) >= 2 Then
            If Right(prefix, 2) = "; " Then GoTo NextSep
        End If
        If Len(prefix) >= 5 Then
            If Right(prefix, 5) = "; and" Then GoTo NextSep
        End If
        s = Left(s, sp - 1) & sep & Mid(s, sp + 1)
NextSep:
    Next i

    PreProcessText = s
End Function


' ============================================================
' Returns True if list item number [n] in format [fmt] exists
' in string [s] as a whole-number token (not a substring of
' a longer number).
'
' fmt: 1="N. "  2="N." (non-space after)  3="N) "  4="(N) "
' ============================================================
Function ListItemExists(s As String, n As Integer, fmt As Integer) As Boolean
    ListItemExists = (FindListItem(s, n, fmt, 1) > 0)
End Function


' ============================================================
' Finds the first occurrence of list item [n] in format [fmt]
' at or after position [startPos] in [s].
' Returns the 1-based start position, or 0 if not found.
' ============================================================
Function FindListItem(s As String, n As Integer, fmt As Integer, startPos As Long) As Long
    FindListItem = 0
    Dim numStr As String
    numStr = CStr(n)
    Dim nLen As Integer
    nLen = Len(numStr)

    Dim srch As String
    Select Case fmt
        Case 1: srch = numStr & ". "
        Case 2: srch = numStr & "."
        Case 3: srch = numStr & ") "
        Case 4: srch = "(" & numStr & ") "
    End Select

    Dim p As Long
    p = InStr(startPos, s, srch)
    Do While p > 0
        Dim ok As Boolean
        ok = True

        ' For digit-led formats, check the char before the number is not a digit
        If fmt = 1 Or fmt = 2 Or fmt = 3 Then
            If p > 1 Then
                Dim chBefore As String
                chBefore = Mid(s, p - 1, 1)
                If chBefore >= "0" And chBefore <= "9" Then ok = False
                ' For format 3 ("N) "), also reject if preceded by "(" —
                ' that means the text is already in format 4 ("(N) ").
                If ok And fmt = 3 Then
                    If chBefore = "(" Then ok = False
                End If
            End If
        End If

        ' For format 2, check the char after the period is not a space or digit
        If ok And fmt = 2 Then
            Dim afterPeriod As Long
            afterPeriod = p + nLen + 1   ' position after the "."
            If afterPeriod <= Len(s) Then
                Dim afterChar As String
                afterChar = Mid(s, afterPeriod, 1)
                If afterChar = " " Then ok = False          ' would be format 1 instead
                If afterChar >= "0" And afterChar <= "9" Then ok = False  ' digit follows — e.g. "3." inside "1793.2"
            End If
        End If

        If ok Then
            FindListItem = p
            Exit Function
        End If

        p = InStr(p + 1, s, srch)
    Loop
End Function


' ============================================================
' Returns the character length of a list item marker for
' format [fmt] and number [n] (used during replacement).
' ============================================================
Function ListItemLen(n As Integer, fmt As Integer) As Integer
    Dim nLen As Integer
    nLen = Len(CStr(n))
    Select Case fmt
        Case 1: ListItemLen = nLen + 2   ' "N. "
        Case 2: ListItemLen = nLen + 1   ' "N."
        Case 3: ListItemLen = nLen + 2   ' "N) "
        Case 4: ListItemLen = nLen + 3   ' "(N) "
    End Select
End Function


' ============================================================
' Applies Times New Roman 12pt to the current Selection
' (also clears bold/italic across the result).
' ============================================================
Sub ApplyFormatting(oSel As Selection)
    With oSel.Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
        .Italic = False
    End With
End Sub


' ============================================================
' Core function: applies title-case rules to a string
' ============================================================
Function TitleCaseString(inputText As String) As String

    ' --- Define word lists ---

    ' Articles
    Dim articles(2) As String
    articles(0) = "a"
    articles(1) = "an"
    articles(2) = "the"

    ' Coordinating conjunctions (FANBOYS)
    Dim conjunctions(6) As String
    conjunctions(0) = "for"
    conjunctions(1) = "and"
    conjunctions(2) = "nor"
    conjunctions(3) = "but"
    conjunctions(4) = "or"
    conjunctions(5) = "yet"
    conjunctions(6) = "so"

    ' Prepositions of 4 letters or fewer
    Dim preps(31) As String
    preps(0) = "a"
    preps(1) = "as"
    preps(2) = "at"
    preps(3) = "by"
    preps(4) = "for"
    preps(5) = "from"
    preps(6) = "in"
    preps(7) = "into"
    preps(8) = "like"
    preps(9) = "near"
    preps(10) = "of"
    preps(11) = "off"
    preps(12) = "on"
    preps(13) = "onto"
    preps(14) = "out"
    preps(15) = "over"
    preps(16) = "past"
    preps(17) = "per"
    preps(18) = "plus"
    preps(19) = "than"
    preps(20) = "thru"
    preps(21) = "till"
    preps(22) = "to"
    preps(23) = "up"
    preps(24) = "upon"
    preps(25) = "via"
    preps(26) = "with"
    preps(27) = "amid"
    preps(28) = "anti"
    preps(29) = "re"
    preps(30) = "et"
    preps(31) = "seq"

    ' Protected acronyms — these are never altered regardless of position
    Dim acronyms(13) As String
    acronyms(0) = "FAC"
    acronyms(1) = "SAC"
    acronyms(2) = "TAC"
    acronyms(3) = "CEQA"
    acronyms(4) = "CD"
    acronyms(5) = "CEO"
    acronyms(6) = "IIED"
    acronyms(7) = "LLC"
    acronyms(8) = "LLP"
    acronyms(9) = "LP"
    acronyms(10) = "LLLP"
    acronyms(11) = "PC"
    acronyms(12) = "GP"
    acronyms(13) = "FEHA"

    ' --- Protect "et seq." (case-insensitive) by substituting a placeholder ---
    Dim etSeqPlaceholder As String
    etSeqPlaceholder = "et seq"
    Dim workText As String
    workText = inputText
    Dim esPos As Long
    esPos = 1
    Do
        Dim esIdx As Long
        esIdx = 0
        Dim candidate As String
        Dim ci As Long
        For ci = esPos To Len(workText) - 6
            candidate = LCase(Mid(workText, ci, 7))
            If candidate = "et seq." Then
                esIdx = ci
                Exit For
            End If
        Next ci
        If esIdx = 0 Then Exit Do
        workText = Left(workText, esIdx - 1) & etSeqPlaceholder & Mid(workText, esIdx + 7)
        esPos = esIdx + Len(etSeqPlaceholder)
    Loop

    ' --- Tokenize into words and whitespace runs ---
    Dim tokens() As String
    tokens = SplitPreserveAll(workText)

    Dim i As Integer
    Dim wordIndex As Integer
    Dim totalWords As Integer

    ' Count total word tokens to identify the last word
    totalWords = 0
    For i = 0 To UBound(tokens)
        If Not IsWhitespaceToken(tokens(i)) Then
            totalWords = totalWords + 1
        End If
    Next i

    ' --- Process each token ---
    Dim result As String
    result = ""
    wordIndex = 0
    Dim capitalizeNext As Boolean
    capitalizeNext = True  ' Always capitalize the first word
    Dim tok As String
    Dim isFirst As Boolean
    Dim isLast As Boolean
    Dim bareWord As String
    Dim leadPunct As String
    Dim trailPunct As String
    Dim afterHyphen As Boolean
    Dim nextIsAct As Boolean
    Dim ki As Integer
    Dim kBare As String, klp As String, ktp As String
    Dim bwParts() As String
    Dim bwIdx As Integer
    Dim hypPos As Integer

    For i = 0 To UBound(tokens)
        tok = tokens(i)

        If IsWhitespaceToken(tok) Then
            result = result & tok
        Else
            wordIndex = wordIndex + 1
            isFirst = (wordIndex = 1)
            isLast = (wordIndex = totalWords)

            leadPunct = ""
            trailPunct = ""
            bareWord = StripPunctuation(tok, leadPunct, trailPunct)

            afterHyphen = IsAfterHyphen(tok, bareWord)

            Dim acronymMatch As String
            acronymMatch = MatchAcronym(bareWord, acronyms)

            Dim shouldCap As Boolean
            Dim newBare As String

            If acronymMatch <> "" Then
                ' Protected acronym: preserve all-caps form, skip all other rules
                newBare = acronymMatch
            Else
                If isFirst Or isLast Then
                    shouldCap = True
                ElseIf capitalizeNext Then
                    shouldCap = True
                ElseIf afterHyphen Then
                    shouldCap = False
                ElseIf IsInList(LCase(bareWord), articles) Then
                    shouldCap = False
                ElseIf IsInList(LCase(bareWord), conjunctions) Then
                    shouldCap = False
                ElseIf IsInList(LCase(bareWord), preps) Then
                    shouldCap = False
                Else
                    shouldCap = True
                End If

                If shouldCap Then
                    nextIsAct = False
                    For ki = i + 1 To UBound(tokens)
                        If Not IsWhitespaceToken(tokens(ki)) Then
                            klp = "": ktp = ""
                            kBare = StripPunctuation(tokens(ki), klp, ktp)
                            If LCase(kBare) = "act" Then nextIsAct = True
                            Exit For
                        End If
                    Next ki
                    If InStr(bareWord, "-") = 0 Or nextIsAct Then
                        bwParts = Split(bareWord, "-")
                        newBare = ""
                        For bwIdx = 0 To UBound(bwParts)
                            If bwIdx > 0 Then newBare = newBare & "-"
                            If Len(bwParts(bwIdx)) > 0 Then
                                newBare = newBare & UCase(Left(bwParts(bwIdx), 1)) & LCase(Mid(bwParts(bwIdx), 2))
                            End If
                        Next bwIdx
                    Else
                        hypPos = InStr(bareWord, "-")
                        newBare = UCase(Left(bareWord, 1)) & LCase(Mid(bareWord, 2, hypPos - 2)) & "-" & LCase(Mid(bareWord, hypPos + 1))
                    End If
                Else
                    newBare = LCase(bareWord)
                End If
            End If

            result = result & leadPunct & newBare & trailPunct

            capitalizeNext = EndsWithBreak(tok)
        End If
    Next i

    ' --- Restore "et seq." placeholder ---
    Dim finalResult As String
    finalResult = result
    Dim rp As Long
    rp = InStr(finalResult, etSeqPlaceholder)
    Do While rp > 0
        finalResult = Left(finalResult, rp - 1) & "et seq." & Mid(finalResult, rp + Len(etSeqPlaceholder))
        rp = InStr(rp + 7, finalResult, etSeqPlaceholder)
    Loop

    TitleCaseString = finalResult
End Function


' ============================================================
' Split a string into alternating word/whitespace tokens
' ============================================================
Function SplitPreserveAll(s As String) As String()
    Dim result() As String
    ReDim result(0)
    Dim count As Integer
    count = 0

    Dim i As Integer
    Dim ch As String
    Dim current As String
    Dim inWhitespace As Boolean
    current = ""

    If Len(s) = 0 Then
        result(0) = ""
        SplitPreserveAll = result
        Exit Function
    End If

    inWhitespace = IsWhitespaceChar(Mid(s, 1, 1))

    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Dim chIsWS As Boolean
        chIsWS = IsWhitespaceChar(ch)

        If chIsWS = inWhitespace Then
            current = current & ch
        Else
            If Len(current) > 0 Then
                If count > 0 Then ReDim Preserve result(count)
                result(count) = current
                count = count + 1
            End If
            current = ch
            inWhitespace = chIsWS
        End If
    Next i

    If Len(current) > 0 Then
        If count > 0 Then ReDim Preserve result(count)
        result(count) = current
    End If

    SplitPreserveAll = result
End Function


' ============================================================
' Returns True if the character is whitespace.
' Uses AscW so characters outside the system code page cannot
' raise a runtime error.
' ============================================================
Function IsWhitespaceChar(ch As String) As Boolean
    Dim c As Long
    c = AscW(ch)
    IsWhitespaceChar = (c = 32 Or c = 9 Or c = 13 Or c = 10 Or c = 160)
End Function


' ============================================================
' Returns True if the entire token is whitespace
' ============================================================
Function IsWhitespaceToken(tok As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(tok)
        If Not IsWhitespaceChar(Mid(tok, i, 1)) Then
            IsWhitespaceToken = False
            Exit Function
        End If
    Next i
    IsWhitespaceToken = True
End Function


' ============================================================
' Strip leading/trailing punctuation; return bare word core.
' leadPunct and trailPunct are set by reference.
' ============================================================
Function StripPunctuation(tok As String, ByRef leadPunct As String, ByRef trailPunct As String) As String
    leadPunct = ""
    trailPunct = ""
    Dim s As String
    s = tok

    ' Dotted acronym like L.A.M.C. or L.A.M.C — return as-is.
    If IsDottedAcronym(s) Then
        StripPunctuation = s
        Exit Function
    End If

    Do While Len(s) > 0 And Not IsLetterOrDigit(Left(s, 1))
        leadPunct = leadPunct & Left(s, 1)
        s = Mid(s, 2)
    Loop

    Do While Len(s) > 0 And Not IsLetterOrDigit(Right(s, 1))
        trailPunct = Right(s, 1) & trailPunct
        s = Left(s, Len(s) - 1)
    Loop

    StripPunctuation = s
End Function


' ============================================================
' Returns True if the token looks like a dotted acronym,
' e.g. L.A.M.C. or U.S.A or F.E.H.A.
' ============================================================
Function IsDottedAcronym(tok As String) As Boolean
    IsDottedAcronym = False
    Dim n As Integer
    n = Len(tok)
    If n < 3 Then Exit Function  ' minimum: "A." or "A.B"

    Dim i As Integer
    For i = 1 To n
        Dim ch As String
        ch = Mid(tok, i, 1)
        If i Mod 2 = 1 Then
            If Not IsLetterOrDigit(ch) Then Exit Function   ' odd positions must be letters/digits
        Else
            If ch <> "." Then Exit Function                 ' even positions must be periods
        End If
    Next i

    IsDottedAcronym = True
End Function


' ============================================================
' Returns True if ch is a letter or digit
' ============================================================
Function IsLetterOrDigit(ch As String) As Boolean
    If Len(ch) = 0 Then
        IsLetterOrDigit = False
        Exit Function
    End If
    Dim c As Long
    c = AscW(ch)
    IsLetterOrDigit = (c >= 65 And c <= 90) Or _
                      (c >= 97 And c <= 122) Or _
                      (c >= 48 And c <= 57) Or _
                      (c > 127)
End Function


' ============================================================
' Returns True if bareWord is immediately preceded by a hyphen
' within the token (connective hyphen rule)
' ============================================================
Function IsAfterHyphen(tok As String, bareWord As String) As Boolean
    IsAfterHyphen = False
    If Len(bareWord) = 0 Then Exit Function

    Dim pos As Integer
    pos = InStr(LCase(tok), LCase(bareWord))
    If pos <= 1 Then Exit Function

    If Mid(tok, pos - 1, 1) = "-" Then
        IsAfterHyphen = True
    End If
End Function


' ============================================================
' Returns True if the token ends with a break character
' (colon, em dash, en dash, or double hyphen) that requires
' the following word to be capitalized.
' Uses AscW only, so no Asc() code-page error risk.
' ============================================================
Function EndsWithBreak(tok As String) As Boolean
    EndsWithBreak = False
    If Len(tok) = 0 Then Exit Function

    Dim lastChar As String
    lastChar = Right(tok, 1)
    Dim cw As Long
    cw = AscW(lastChar)

    If lastChar = ":" Then EndsWithBreak = True: Exit Function
    If cw = 8212 Or cw = 151 Then EndsWithBreak = True: Exit Function   ' em dash
    If cw = 8211 Or cw = 150 Then EndsWithBreak = True: Exit Function   ' en dash

    If Len(tok) >= 2 Then
        If Right(tok, 2) = "--" Then EndsWithBreak = True: Exit Function
    End If
End Function


' ============================================================
' If word matches a protected acronym (case-insensitive),
' returns the canonical all-caps form; otherwise returns "".
' ============================================================
Function MatchAcronym(word As String, lst() As String) As String
    Dim i As Integer
    For i = 0 To UBound(lst)
        If LCase(word) = LCase(lst(i)) Then
            MatchAcronym = lst(i)
            Exit Function
        End If
    Next i
    MatchAcronym = ""
End Function


' ============================================================
' Returns True if word (already lowercased) is in the array
' ============================================================
Function IsInList(word As String, lst() As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(lst)
        If word = lst(i) Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False
End Function


