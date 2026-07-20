Attribute VB_Name = "WrapCitations"
Option Explicit
' NOTE: do NOT add Option Private Module here. CheckAndWrap / CheckAndWrapEnter
' are assigned to the spacebar/return keys via KeyBindings.Add, and a macro in an
' Option Private Module module cannot be resolved as a key-binding target --
' RegisterWrapKeyBindings would fail with runtime error 5346.

Private bWrapBusy As Boolean

' Both key handlers run under On Error Resume Next for their whole body: they
' fire on EVERY space/Enter in EVERY document, so an error must never (a) throw
' a dialog at the user mid-keystroke -- in a read-only or protected document the
' TypeText itself errors where native typing would just refuse -- or (b) skip
' the bWrapBusy reset, which would silently kill wrapping for the session (e.g.
' StartCustomRecord raises if a crashed macro left a custom record open).
Public Sub CheckAndWrap()
    On Error Resume Next
    If bWrapBusy Then
        Selection.TypeText " "
        Exit Sub
    End If

    bWrapBusy = True
    Dim oUndo As UndoRecord: Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "Auto Wrap Citation"
    DoCheckAndWrap
    oUndo.EndCustomRecord
    bWrapBusy = False
    Selection.TypeText " "
    ApplyAutoCorrect
End Sub

Public Sub CheckAndWrapEnter()
    On Error Resume Next
    If bWrapBusy Then
        Selection.TypeParagraph
        Exit Sub
    End If

    bWrapBusy = True
    Dim oUndo As UndoRecord: Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "Auto Wrap Citation"
    DoCheckAndWrap
    oUndo.EndCustomRecord
    bWrapBusy = False
    Selection.TypeParagraph
End Sub

Private Sub ApplyAutoCorrect()
    On Error GoTo Done
    Dim oDoc As Document: Set oDoc = ActiveDocument
    Dim lCur As Long: lCur = Selection.start
    If lCur < 2 Then GoTo Done

    Dim lWordEnd As Long: lWordEnd = lCur - 1
    Dim lWordStart As Long: lWordStart = lWordEnd

    Do While lWordStart > 0
        Dim n As Long: n = AscW(oDoc.Range(lWordStart - 1, lWordStart).text)
        Select Case n
            Case 32, 13, 11, 9, 7, 160: Exit Do
            Case Else: lWordStart = lWordStart - 1
        End Select
    Loop

    If lWordStart >= lWordEnd Then GoTo Done

    Dim sWord As String: sWord = oDoc.Range(lWordStart, lWordEnd).text
    Dim sNew As String: sNew = Application.AutoCorrect.Entries(sWord).Value

    If Err.Number = 0 And sNew <> sWord Then
        oDoc.Range(lWordStart, lWordEnd).text = sNew
        Selection.start = lWordStart + Len(sNew) + 1
        Selection.End = Selection.start
    End If
Done:
    Err.Clear
End Sub

Private Sub DoCheckAndWrap()
    If Selection.Type <> wdSelectionIP Then Exit Sub

    ' Main body only: Selection offsets in a footnote/header/text box are
    ' story-relative, but every range below goes through ActiveDocument.Range
    ' (the main text story) -- wrapping there would insert parens into an
    ' unrelated spot in the body. The caller still types the space/paragraph.
    If Selection.StoryType <> wdMainTextStory Then Exit Sub

    ' Skip while Track Changes is on: revision-marked deletions still appear in
    ' Range.Text, so the paragraph-prefix analysis would run on text that is no
    ' longer really there. The keystroke itself still goes through.
    If ActiveDocument.TrackRevisions Then Exit Sub

    Dim oPar As Paragraph: Set oPar = Selection.Paragraphs(1)
    If oPar Is Nothing Then Exit Sub

    ' 1. STRICT PARAGRAPH BOUNDARY: The macro cannot see previous paragraphs
    Dim lParStart As Long: lParStart = oPar.Range.start
    Dim lCursor As Long: lCursor = Selection.Range.start
    If lCursor <= lParStart Then Exit Sub

    Dim s As String: s = RTrim(ActiveDocument.Range(lParStart, lCursor).text)
    If Len(s) = 0 Or Right(s, 1) <> "." Then Exit Sub

    ' 2. HARD STOP: never wrap on the period in "Ex.", "p.", or "pp."
    If Len(s) >= 3 And Right(s, 3) = "Ex." Then Exit Sub
    If Len(s) >= 4 And Right(s, 4) = "Exs." Then Exit Sub
    If Len(s) >= 2 And Right(s, 2) = "p." Then Exit Sub
    If Len(s) >= 3 And Right(s, 3) = "pp." Then Exit Sub

    ' 3. EXHIBIT REFERENCE CHECK (used later for Pattern 3c guardrail)
    Dim bExhibit As Boolean: bExhibit = False
    If Len(s) >= 7 Then
        If Right(s, 7) Like " Ex. [A-Z]." Then bExhibit = True
    End If

    ' 3. QUOTE DEPTH CHECK: Includes single curly quotes for nested legal cites
    Dim qDepth As Long, qi As Long
    For qi = 1 To Len(s)
        Dim qc As Long: qc = AscW(Mid(s, qi, 1))
        If qc = 34 Or qc = 8220 Then
            qDepth = qDepth + 1
        ElseIf qc = 8221 Or qc = 8217 Then
            If qDepth > 0 Then qDepth = qDepth - 1
        End If
    Next qi
    If qDepth > 0 Then Exit Sub

    ' PAREN BALANCE: if the paragraph has an unmatched open "(" before the
    ' cursor, the user opened the citation manually.  Close with ")" only;
    ' do not insert another "(".
    Dim bSkipOpen As Boolean: bSkipOpen = False
    Dim nOpen As Long, nClose As Long, pi As Long
    For pi = 1 To Len(s)
        Select Case Mid(s, pi, 1)
            Case "(": nOpen = nOpen + 1
            Case ")": nClose = nClose + 1
        End Select
    Next pi
    If nOpen > nClose Then bSkipOpen = True

    Dim oDoc As Document: Set oDoc = ActiveDocument
    Dim lS As Long, lE As Long, lOff As Long

    ' WRAPPED-CITATION BOUNDARY GUARD
    ' Previously wrapped citations end with ".)" or a right double-quote
    ' (U+201D, e.g. a quoted passage ending).  Find the last such
    ' occurrence and ignore everything before it; this prevents markers
    ' from an already-wrapped citation from triggering on the plain
    ' sentence that follows (e.g. one that merely ends in a number).
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

    ' TRIGGER STRING: a copy of sSearch with the contents of every BALANCED
    ' (...) pair removed.  All trigger-token detection below runs against this
    ' so that a citation marker sealed inside a finished parenthetical (e.g.
    ' "(Pen. Code, sec 187)" or "(Smith, supra, at p. 5)" embedded in prose)
    ' does NOT fire the wrapper.  Text inside an unmatched/unclosed "(" is
    ' preserved, so a manually-opened citation still triggers.  Wrap offsets
    ' are still computed from the original string (s).
    Dim sTrig As String: sTrig = StripClosedParens(sSearch)

    ' Pattern 1: Ibid.
    If Right(s, 5) = "Ibid." Then
        lS = lParStart + Len(s) - 5: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then WrapRange oDoc, lS, lE, True, False, bSkipOpen
        Exit Sub
    End If

    ' Pattern 2: Id. at  (must appear after the last wrapped citation,
    ' and outside any closed parenthetical)
    If InStr(sTrig, "Id. at") > 0 Then
        Dim lId As Long: lId = InStrRev(s, "Id. at")
        lS = lParStart + lId - 1: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then WrapRange oDoc, lS, lE, False, True, bSkipOpen
        Exit Sub
    End If

    ' Pattern 3a: Pilcrow (\xb6)  (must appear after the last wrapped citation,
    ' and outside any closed parenthetical -- a pilcrow inside a balanced
    ' "(Fallah Decl. para 9, 12)" embedded in prose is already a finished
    ' citation and must not cause the surrounding sentence to be wrapped).
    If InStr(sTrig, ChrW(182)) > 0 Then
        lOff = GetCiteStart(s)
        lS = lParStart + lOff: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then WrapRange oDoc, lS, lE, False, False, bSkipOpen
        Exit Sub
    End If

    ' Pattern 3b: " at p." / " at pp."  (must appear after the last wrapped
    ' citation, and outside any closed parenthetical)
    If InStr(sTrig, " at p.") > 0 Or InStr(sTrig, " at pp.") > 0 Then
        Dim lAtP As Long
        lAtP = InStrRev(s, " at pp.")
        If lAtP = 0 Then lAtP = InStrRev(s, " at p.")
        If lAtP > 1 Then
            lOff = GetCiteStart(Left(s, lAtP - 1))
        Else
            lOff = 0
        End If
        lS = lParStart + lOff: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then WrapRange oDoc, lS, lE, False, False, bSkipOpen
        Exit Sub
    End If

    ' Pattern 3c: Cal., U.S., or section sign  (must appear after the last
    ' wrapped citation, and outside any closed parenthetical).
    ' Also requires the sentence to end in a digit or ")" (page/year ref).
    If InStr(sTrig, "Cal.") > 0 Or InStr(sTrig, "U.S.") > 0 Or InStr(sTrig, ChrW(167)) > 0 Then
        Dim charBefore As String: charBefore = Mid(s, Len(s) - 1, 1)
        If (charBefore Like "[0-9)]") Or bExhibit Then
            lOff = GetCiteStart(s)
            lS = lParStart + lOff: lE = lParStart + Len(s)
            If Not IsAlreadyWrapped(oDoc, lS, lE) Then WrapRange oDoc, lS, lE, False, False, bSkipOpen
        End If
        Exit Sub
    End If

    ' Pattern 3d: UMF/AMF No. <n>. or UMF/AMF Nos. <list/range>.
    '   e.g.  UMF No. 6.        -> (UMF No. 6.)
    '         UMF Nos. 3-7.     -> (UMF Nos. 3-7.)
    '         AMF Nos. 3, 5, 7. -> (AMF Nos. 3, 5, 7.)
    If MatchesUmfAmfCitation(sTrig) Then
        lOff = GetCiteStart(s)
        lS = lParStart + lOff: lE = lParStart + Len(s)
        If Not IsAlreadyWrapped(oDoc, lS, lE) Then WrapRange oDoc, lS, lE, False, False, bSkipOpen
    End If
End Sub

' Returns True if the trimmed sentence ends with a UMF/AMF citation pattern:
'   "UMF No. <digits>."     or  "AMF No. <digits>."
'   "UMF Nos. <list>."      or  "AMF Nos. <list>."
' where <list> is digits, optionally followed by more digits separated by
' commas, spaces, or hyphens (e.g., "3", "3-7", "3, 5, 7", "3-7, 10").
Private Function MatchesUmfAmfCitation(s As String) As Boolean
    MatchesUmfAmfCitation = False
    If Len(s) < 8 Then Exit Function          ' shortest is "UMF No. 1."
    If Right(s, 1) <> "." Then Exit Function

    ' Walk backward over the trailing "<digits/commas/spaces/hyphens>."
    Dim i As Long: i = Len(s) - 1              ' position before final period
    Dim sawDigit As Boolean: sawDigit = False
    Do While i >= 1
        Dim ch As String: ch = Mid(s, i, 1)
        If ch Like "[0-9]" Then
            sawDigit = True
            i = i - 1
        ElseIf ch = "," Or ch = " " Or ch = "-" Then
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    If Not sawDigit Then Exit Function

    ' The walk above consumes spaces too, so after the loop i points at the
    ' period of "No." / "Nos." (the first char that wasn't digit/comma/space/-).
    '   ... U M F   N o .  <space>  3 - 7 .
    '               ^ i
    If i < 1 Then Exit Function
    If Mid(s, i, 1) <> "." Then Exit Function

    ' Now look for "No" or "Nos" ending just before that period
    Dim abbrevEnd As Long: abbrevEnd = i - 1   ' last letter of No/Nos
    If abbrevEnd < 2 Then Exit Function

    Dim foundAbbrev As Boolean: foundAbbrev = False
    Dim abbrevStart As Long

    ' Try "Nos" (3 chars) first
    If abbrevEnd >= 3 Then
        If Mid(s, abbrevEnd - 2, 3) = "Nos" Then
            abbrevStart = abbrevEnd - 2
            foundAbbrev = True
        End If
    End If
    ' Try "No" (2 chars) if Nos didn't match
    If Not foundAbbrev Then
        If Mid(s, abbrevEnd - 1, 2) = "No" Then
            abbrevStart = abbrevEnd - 1
            foundAbbrev = True
        End If
    End If
    If Not foundAbbrev Then Exit Function

    ' Need a space before "No"/"Nos"
    If abbrevStart - 1 < 1 Then Exit Function
    If Mid(s, abbrevStart - 1, 1) <> " " Then Exit Function

    ' Need "UMF" or "AMF" immediately before that space
    If abbrevStart - 4 < 1 Then Exit Function
    Dim prefix As String: prefix = Mid(s, abbrevStart - 4, 3)
    If prefix <> "UMF" And prefix <> "AMF" Then Exit Function

    ' Make sure UMF/AMF isn't part of a larger word (e.g., "XUMF")
    If abbrevStart - 4 > 1 Then
        Dim chBefore As String: chBefore = Mid(s, abbrevStart - 5, 1)
        If chBefore Like "[A-Za-z0-9]" Then Exit Function
    End If

    MatchesUmfAmfCitation = True
End Function

' Returns a copy of s with the contents of every BALANCED (...) pair removed,
' including the parentheses themselves and any nesting.  Characters inside an
' unmatched/unclosed "(" are preserved (so a manually-opened citation that the
' user expects the macro to close still triggers), as are any stray unmatched
' ")".  This is used only to decide whether a trigger token appears OUTSIDE a
' finished parenthetical; the wrap range itself is computed from the original
' string so offsets remain correct.
Private Function StripClosedParens(s As String) As String
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    Dim remove() As Boolean: ReDim remove(1 To n)
    Dim stack() As Long: ReDim stack(1 To n)
    Dim sp As Long: sp = 0
    Dim i As Long
    For i = 1 To n
        Dim ch As String: ch = Mid(s, i, 1)
        If ch = "(" Then
            sp = sp + 1
            stack(sp) = i
        ElseIf ch = ")" Then
            If sp > 0 Then
                Dim op As Long: op = stack(sp)
                sp = sp - 1
                Dim k As Long
                For k = op To i
                    remove(k) = True      ' mark the whole balanced pair for removal
                Next k
            End If
        End If
    Next i

    Dim out As String: out = ""
    For i = 1 To n
        If Not remove(i) Then out = out & Mid(s, i, 1)
    Next i
    StripClosedParens = out
End Function

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

        ' Case 2: period + closing quote/paren + space  (e.g.  ." C  or  .' C )
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
    GetCiteStart = lastBoundary
End Function
Private Function GetWordCount(ByVal t As String) As Long
    t = Trim(t)
    If Len(t) = 0 Then
        GetWordCount = 0
        Exit Function
    End If
    Dim count As Long: count = 1
    Dim i As Long
    For i = 1 To Len(t) - 1
        If Mid(t, i, 1) = " " And Mid(t, i + 1, 1) <> " " Then
            count = count + 1
        End If
    Next i
    GetWordCount = count
End Function

Private Function FindCodeAbbrevBefore(s As String, lBefore As Long) As Long
    Dim abbrevs(): abbrevs = Array("Pub. Resources Code", "Pub. Contract Code", "Bus. & Prof. Code", "Cal. U. Com. Code", "Health & Saf. Code", "Welf. & Inst. Code", "Food & Agr. Code", "Harb. & Nav. Code", "Mil. & Vet. Code", "Rev. & Tax. Code", "Unemp. Ins. Code", "Fish & G. Code", "Sts. & Hy. Code", "Code Civ. Proc.", "Cal. Code Regs.", "Pub. Util. Code", "U. Com. Code", "Corp. Code", "Elec. Code", "Evid. Code", "Civ. Code", "Fam. Code", "Fin. Code", "Gov. Code", "Ins. Code", "Lab. Code", "Pen. Code", "Prob. Code", "Veh. Code", "Wat. Code", "Ed. Code")
    Dim sSearch As String: sSearch = Left(s, lBefore - 1)
    Dim bestPos As Long: bestPos = 0
    Dim bestEnd As Long: bestEnd = 0
    Dim bestLen As Long: bestLen = 0
    Dim i As Long
    For i = 0 To UBound(abbrevs)
        Dim lPos As Long: lPos = InStrRev(sSearch, abbrevs(i))
        If lPos > 0 Then
            Dim lEnd As Long: lEnd = lPos + Len(abbrevs(i)) - 1
            If lEnd > bestEnd Or (lEnd = bestEnd And Len(abbrevs(i)) > bestLen) Then
                bestPos = lPos: bestEnd = lEnd: bestLen = Len(abbrevs(i))
            End If
        End If
    Next i
    FindCodeAbbrevBefore = bestPos
End Function

Private Function FindSignalBefore(s As String, lCodeStart As Long) As Long
    Dim sigs(): sigs = Array("See also ", "see also ", "But see ", "but see ", "Contra, ", "contra, ", "E.g., ", "e.g., ", "Cf. ", "cf. ", "See ", "see ")
    Dim i As Long
    For i = 0 To UBound(sigs)
        Dim lSigStart As Long: lSigStart = lCodeStart - Len(sigs(i))
        If lSigStart >= 1 Then
            If Mid(s, lSigStart, Len(sigs(i))) = sigs(i) Then
                FindSignalBefore = lSigStart
                Exit Function
            End If
        End If
    Next i
    FindSignalBefore = lCodeStart
End Function

Private Sub WrapRange(oDoc As Document, lS As Long, lE As Long, Optional bItalicAll As Boolean = False, Optional bItalicFirst3 As Boolean = False, Optional bSkipOpen As Boolean = False)
    If bItalicAll Then oDoc.Range(lS, lE).Font.Italic = True
    If bItalicFirst3 Then oDoc.Range(lS, lS + 3).Font.Italic = True
    oDoc.Range(lE, lE).InsertAfter ")"
    FormatRange oDoc.Range(lE, lE + 1)
    If Not bSkipOpen Then
        oDoc.Range(lS, lS).InsertBefore "("
        FormatRange oDoc.Range(lS, lS + 1)
        Selection.SetRange lE + 2, lE + 2
    Else
        Selection.SetRange lE + 1, lE + 1
    End If
End Sub

Private Sub FormatRange(r As Range)
    r.Font.Name = "Times New Roman": r.Font.Size = 12: r.Font.Italic = False: r.Font.Bold = False
End Sub

Private Function IsAlreadyWrapped(oDoc As Document, lS As Long, lE As Long) As Boolean
    IsAlreadyWrapped = False: On Error Resume Next
    If lS >= 1 Then
        If oDoc.Range(lS - 1, lS).text = "(" Then IsAlreadyWrapped = True
    End If
    If Not IsAlreadyWrapped And lE > lS Then
        If Left(oDoc.Range(lS, lE).text, 1) = "(" Then IsAlreadyWrapped = True
    End If
    On Error GoTo 0
End Function






