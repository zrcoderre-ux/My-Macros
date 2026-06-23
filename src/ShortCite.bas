Attribute VB_Name = "ShortCite"
'==============================================================================
' USER-DEFINED TYPES  (must appear before all Subs/Functions in VBA)
'==============================================================================

Type CitInfo
    caseName           As String
    shortName          As String
    reporter           As String
    initialPage        As String
    pincite            As String
    year               As String
    startChar          As Long
    length             As Long
    hasOpenParen       As Boolean
    hasCloseParen      As Boolean
    isMidSentence      As Boolean
    shortNameOverride  As String
    signal             As String
    isBare             As Boolean    ' True = no surrounding parentheses
    bracketNote        As String     ' optional trailing [text] inside cite, e.g. [italics added]
    isCompound         As Boolean    ' True = cite is one of several in a (; )-separated compound cite
    fnTail             As String     ' optional ", fn. omitted" / ", fns. omitted" / ", fn. N" tail
End Type

Type RepInfo
    citStartChar       As Long
    citLength          As Long
    newText            As String
    italicWord         As String
    italicOffset       As Long
    caseNameLen        As Long
    isSupra            As Boolean
    isFirstOccurrence  As Boolean
    signalLen          As Long
    parenNameOffset    As Long
    parenNameLen       As Long
    isBare             As Boolean    ' True = bare cite (no outer parens)
    isCompound         As Boolean    ' True = cite is one entry in a compound cite
End Type

Type DocCite
    citeType           As String
    normKey            As String
    caseName           As String
    shortName          As String
    reporter           As String
    initialPage        As String
    pincite            As String
    year               As String
    signal             As String
    shortNameOverride  As String
    absStart           As Long
    textLen            As Long
    isBare             As Boolean
    bracketNote        As String     ' optional [text] retained inside the cite
    isCompound         As Boolean    ' True = cite is one entry in a compound cite
    fnTail             As String     ' optional ", fn. omitted" / ", fns. omitted" / ", fn. N" tail
    paraIdx            As Long       ' paragraph index (1-based) where cite was found
End Type

Type ShortCiteInfo
    citeType           As String
    shortName          As String
    reporter           As String
    pincite            As String
    signal             As String
    startChar          As Long
    citLength          As Long
    isMidSentence      As Boolean
    inQuote            As Boolean
    isBare             As Boolean    ' True = bare supra (no outer parens)
    bracketNote        As String     ' optional trailing [text] inside cite
    isCompound         As Boolean    ' True = supra is one entry in a compound cite
End Type

Type HintLine
    normKey     As String
    caseName    As String
    shortName   As String
    reporter    As String
    initialPage As String
    year        As String
    bmName      As String    ' Word bookmark name used in Phase-4 deletion
End Type

'==============================================================================
' CALIFORNIA STYLE MANUAL CITATION CONVERTER  v5.1
'==============================================================================
Option Explicit

' === module-level phase tracker (used in ErrHandler) ===
Private gPhase As String


'==============================================================================
' MAIN ENTRY POINT
'==============================================================================

Public Sub ConvertToShortCitations()
On Error GoTo ErrHandler
    gPhase = "Init"

    Dim resp As Integer
    resp = MsgBox( _
        "CSM Citation Converter v5.1" & vbCrLf & vbCrLf & _
        "This macro will:" & vbCrLf & _
        "  1. Convert subsequent citations to short form" & vbCrLf & _
        "  2. Add short-name parentheticals to first occurrences" & vbCrLf & _
        "  3. Update existing Ibid./Id./supra cites on re-run" & vbCrLf & _
        "  4. Fix any supra that appears before its full citation" & vbCrLf & _
        "  5. Preserve introductory signals (See, Cf., etc.)" & vbCrLf & _
        "  6. Handle bare (inline) citations" & vbCrLf & _
        "  7. Complete in-quote citations using hint-line paragraphs" & vbCrLf & _
        "  8. Leave (Ibid.) / (Id.) alone after non-case parentheticals" & vbCrLf & _
        "     (declarations, motions, briefs, RJN, etc.)" & vbCrLf & vbCrLf & _
        "Work on a COPY of your document.  Continue?", _
        vbYesNo + vbQuestion, "CSM Citation Converter")
    If resp <> vbYes Then Exit Sub

    Dim Doc As Document: Set Doc = ActiveDocument
    Dim prevTrackRevisions As Boolean: prevTrackRevisions = Doc.TrackRevisions
    Doc.TrackRevisions = True
    Application.UndoRecord.StartCustomRecord ("CSM Short Citation Conversion")
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim multiDict         As Object: Set multiDict = CreateObject("Scripting.Dictionary"): multiDict.CompareMode = 1
    Dim snToKey           As Object: Set snToKey = CreateObject("Scripting.Dictionary"): snToKey.CompareMode = 1
    Dim rpToKey           As Object: Set rpToKey = CreateObject("Scripting.Dictionary"): rpToKey.CompareMode = 1
    Dim preScanInfo       As Object: Set preScanInfo = CreateObject("Scripting.Dictionary"): preScanInfo.CompareMode = 1
    Dim altPartyToKey     As Object: Set altPartyToKey = CreateObject("Scripting.Dictionary"): altPartyToKey.CompareMode = 1
    Dim shortNameOverrides As Object: Set shortNameOverrides = CreateObject("Scripting.Dictionary"): shortNameOverrides.CompareMode = 1
    Dim needsShortName    As Object: Set needsShortName = CreateObject("Scripting.Dictionary"): needsShortName.CompareMode = 1

    Dim docCites()  As DocCite: ReDim docCites(0 To 400): Dim dcc As Long: dcc = 0
    Dim hintLines() As HintLine: ReDim hintLines(0 To 100): Dim hlC As Long: hlC = 0
    Dim usedHints() As Boolean: ReDim usedHints(0 To 100)

    gPhase = "Phase 1: PreScan"
    PreScanDocument Doc, docCites, dcc, multiDict, snToKey, rpToKey, preScanInfo, _
                    altPartyToKey, shortNameOverrides, hintLines, hlC, needsShortName

    gPhase = "Phase 1.5: InQuote"
    Dim b15Dict  As Object: Set b15Dict = CreateObject("Scripting.Dictionary"): b15Dict.CompareMode = 1
    Dim b15Count As Long: b15Count = 0
    ProcessInQuoteCompletions Doc, hintLines, hlC, usedHints, docCites, dcc, snToKey, _
                               rpToKey, altPartyToKey, preScanInfo, b15Dict, b15Count

    gPhase = "Phase 2: Swaps"
    Dim swapCount As Long
    swapCount = PerformSwaps(Doc, docCites, dcc, snToKey, rpToKey, multiDict, shortNameOverrides)

    gPhase = "Phase 3: Main start"
    Dim caseDict As Object: Set caseDict = CreateObject("Scripting.Dictionary"): caseDict.CompareMode = 1

    Dim changeCount  As Long: changeCount = 0
    Dim orphanCount  As Long: orphanCount = 0
    Dim gPrevKey     As String: gPrevKey = ""
    Dim gPrevPincite As String: gPrevPincite = ""

    Dim PARA As Paragraph
    Dim paraIdx As Long: paraIdx = 0
    For Each PARA In Doc.Paragraphs
        paraIdx = paraIdx + 1
        gPhase = "Phase 3: Para#" & paraIdx & " start=" & PARA.Range.start & _
                 " '" & Left(Replace(PARA.Range.text, Chr(13), ""), 80) & "'"
        ProcessParagraph PARA, caseDict, snToKey, rpToKey, multiDict, preScanInfo, _
                         altPartyToKey, shortNameOverrides, b15Dict, _
                         changeCount, gPrevKey, gPrevPincite, orphanCount, needsShortName
    Next PARA

    gPhase = "Phase 3.5: Parens"
    Dim parenCount As Long: parenCount = 0
    ProcessParentheticals Doc, caseDict, preScanInfo, multiDict, snToKey, rpToKey, _
                          altPartyToKey, parenCount

    gPhase = "Phase 4: Delete hints"
    If hlC > 0 Then DeleteUsedHintLines Doc, hintLines, hlC, usedHints

    Application.ScreenUpdating = True
    Application.UndoRecord.EndCustomRecord
    Doc.TrackRevisions = prevTrackRevisions

    Dim msg As String
    msg = "Done! " & changeCount & " citation(s) converted."
    If swapCount > 0 Then msg = msg & vbCrLf & swapCount & " orphan supra(s) replaced with full citation."
    If b15Count > 0 Then msg = msg & vbCrLf & b15Count & " in-quote citation(s) completed."
    If parenCount > 0 Then msg = msg & vbCrLf & parenCount & " short-name parenthetical(s) removed."
    If orphanCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "WARNING: " & orphanCount & " short cite(s) could not be resolved " & _
              "and have been highlighted in yellow."
    End If
    msg = msg & vbCrLf & vbCrLf & "Ctrl+Z undoes all changes."
    MsgBox msg, vbInformation, "CSM Citation Converter"
    Exit Sub


ErrHandler:
    Application.ScreenUpdating = True
    Application.UndoRecord.EndCustomRecord
    On Error Resume Next
    Doc.TrackRevisions = prevTrackRevisions
    On Error GoTo 0
    MsgBox "Error " & Err.Number & " at line: " & Erl & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Phase: " & gPhase, _
           vbExclamation, "CSM Citation Converter"
End Sub

'==============================================================================
' PHASE 1 - PRE-SCAN
'==============================================================================
Private Sub PreScanDocument(Doc As Document, _
                             docCites() As DocCite, _
                             ByRef dcc As Long, _
                             multiDict As Object, _
                             snToKey As Object, _
                             rpToKey As Object, _
                             preScanInfo As Object, _
                             altPartyToKey As Object, _
                             shortNameOverrides As Object, _
                             hintLines() As HintLine, _
                             ByRef hlC As Long, _
                             needsShortName As Object)

    Dim reLong     As Object: Set reLong = CreateObject("VBScript.RegExp")
    Dim reSupra    As Object: Set reSupra = CreateObject("VBScript.RegExp")
    Dim reBareSup  As Object: Set reBareSup = CreateObject("VBScript.RegExp")
    reLong.Global = True: reLong.Multiline = False: reLong.Pattern = BuildLongCitePattern()
    reSupra.Global = True: reSupra.Multiline = False: reSupra.Pattern = BuildSupraPattern()
    reBareSup.Global = True: reBareSup.Multiline = False: reBareSup.Pattern = BuildBareSupraPattern()

    Dim allLong()  As DocCite: Dim alC As Long: ReDim allLong(0 To 400): alC = 0
    Dim allSupra() As DocCite: Dim asc As Long: ReDim allSupra(0 To 400): asc = 0

    Dim PARA As Paragraph
    Dim psParaIdx As Long: psParaIdx = 0
    For Each PARA In Doc.Paragraphs
        psParaIdx = psParaIdx + 1
        Dim pt As String
        pt = PARA.Range.text
        If Len(pt) > 0 And Right(pt, 1) = Chr(13) Then pt = Left(pt, Len(pt) - 1)
        pt = NormalizeSpaces(pt)
        If Len(pt) = 0 Then GoTo PSNextPara

        Dim qm() As Boolean: ReDim qm(1 To Len(pt)): BuildQuoteMask pt, qm

        ' ------ Long cites ------
        Dim msL As Object: Set msL = reLong.Execute(pt)
        Dim prevAlC As Long: prevAlC = alC
        Dim mL As Object
        For Each mL In msL
            Dim ldc As DocCite
            If ScanLongCite(mL, PARA, pt, qm, ldc) Then
                ldc.paraIdx = psParaIdx
                If alC > UBound(allLong) - 1 Then ReDim Preserve allLong(0 To UBound(allLong) + 100)
                allLong(alC) = ldc: alC = alC + 1
            End If
        Next mL

        ' Check if this paragraph is a hint line
        If msL.count = 1 And (alC - prevAlC) = 1 Then
            Dim ptTrim As String: ptTrim = Trim(pt)
            Dim addedLdc As DocCite: addedLdc = allLong(alC - 1)
            Dim citeRelStart As Long: citeRelStart = addedLdc.absStart - PARA.Range.start + 1
            If citeRelStart <= 3 And (citeRelStart - 1 + addedLdc.textLen) >= Len(ptTrim) - 3 Then
                If hlC > UBound(hintLines) - 1 Then ReDim Preserve hintLines(0 To UBound(hintLines) + 50)
                hintLines(hlC).normKey = addedLdc.normKey
                hintLines(hlC).caseName = addedLdc.caseName
                hintLines(hlC).shortName = addedLdc.shortName
                hintLines(hlC).reporter = addedLdc.reporter
                hintLines(hlC).initialPage = addedLdc.initialPage
                hintLines(hlC).year = addedLdc.year
                Dim bmn As String: bmn = "CSMHint" & hlC
                Doc.Bookmarks.Add Name:=bmn, Range:=PARA.Range
                hintLines(hlC).bmName = bmn
                hlC = hlC + 1
            End If
        End If

        ' ------ Supra cites (parenthetical) ------
        Dim msS As Object: Set msS = reSupra.Execute(pt)
        Dim mS As Object
        For Each mS In msS
            Dim sdc As DocCite
            If ScanSupraCite(mS, PARA, pt, qm, sdc) Then
                sdc.paraIdx = psParaIdx
                If asc > UBound(allSupra) - 1 Then ReDim Preserve allSupra(0 To UBound(allSupra) + 100)
                allSupra(asc) = sdc: asc = asc + 1
            End If
        Next mS

        ' ------ Bare supra cites ------
        Dim msBS2 As Object: Set msBS2 = reBareSup.Execute(pt)
        Dim mBS2  As Object
        For Each mBS2 In msBS2
            Dim bss2 As Long: bss2 = mBS2.FirstIndex + 1
            Dim bsl2 As Long: bsl2 = mBS2.length
            If bss2 > 1 Then
                If Mid(pt, bss2 - 1, 1) = "(" Then GoTo PSNextBS
                Dim bss2Back As Long: bss2Back = bss2 - 2
                Do While bss2Back >= 1
                    If Mid(pt, bss2Back, 1) <> " " Then Exit Do
                    bss2Back = bss2Back - 1
                Loop
                If bss2Back >= 1 Then
                    If Mid(pt, bss2Back, 1) = "," Then GoTo PSNextBS
                End If
            End If
            If IsInsideQuote(bss2, bsl2, qm) Then GoTo PSNextBS
            Dim bsdc As DocCite
            bsdc.citeType = "supra"
            bsdc.signal = Trim(mBS2.SubMatches(0))
            bsdc.shortName = Trim(mBS2.SubMatches(1))
            bsdc.reporter = Trim(mBS2.SubMatches(2))
            bsdc.pincite = CleanPincite(Trim(mBS2.SubMatches(3)))
            bsdc.absStart = PARA.Range.start + bss2 - 1
            bsdc.textLen = bsl2
            bsdc.normKey = ""
            bsdc.isBare = True
            bsdc.paraIdx = psParaIdx
            If asc > UBound(allSupra) - 1 Then ReDim Preserve allSupra(0 To UBound(allSupra) + 100)
            allSupra(asc) = bsdc: asc = asc + 1
PSNextBS:
        Next mBS2

PSNextPara:
    Next PARA

    ' Build lookup tables from long cites
    Dim occCount As Object: Set occCount = CreateObject("Scripting.Dictionary"): occCount.CompareMode = 1
    Dim i As Long
    For i = 0 To alC - 1
        Dim lk As String: lk = allLong(i).normKey
        Dim lSN As String: lSN = allLong(i).shortName
        If Not occCount.Exists(lk) Then
            occCount.Add lk, 1
            If Not snToKey.Exists(LCase(lSN)) Then snToKey.Add LCase(lSN), lk
            If Not rpToKey.Exists(LCase(allLong(i).reporter)) Then rpToKey.Add LCase(allLong(i).reporter), lk
        Else
            occCount(lk) = occCount(lk) + 1
            If Not multiDict.Exists(lk) Then multiDict.Add lk, True
        End If
        If dcc > UBound(docCites) - 1 Then ReDim Preserve docCites(0 To UBound(docCites) + 100)
        docCites(dcc) = allLong(i): dcc = dcc + 1
    Next i

    ' Build preScanInfo
    For i = 0 To alC - 1
        If Not preScanInfo.Exists(allLong(i).normKey) Then
            preScanInfo.Add allLong(i).normKey, _
                allLong(i).caseName & "|" & _
                allLong(i).year & "|" & _
                allLong(i).reporter & "|" & _
                allLong(i).initialPage & "|" & _
                allLong(i).shortName & "|" & _
                allLong(i).shortNameOverride
        End If
    Next i

    ' Build altPartyToKey
    For i = 0 To alC - 1
        Dim apVPos As Long: apVPos = InStr(allLong(i).caseName, " v. ")
        If apVPos > 1 Then
            Dim apPla As String: apPla = LCase(Trim(Left(allLong(i).caseName, apVPos - 1)))
            Dim apDef As String: apDef = LCase(Trim(Mid(allLong(i).caseName, apVPos + 4)))
            If Not altPartyToKey.Exists(apPla) Then altPartyToKey.Add apPla, allLong(i).normKey
            If Not altPartyToKey.Exists(apDef) Then altPartyToKey.Add apDef, allLong(i).normKey
        End If
    Next i

    ' Detect rival shortNameOverrides
    Dim snOvScan As Object: Set snOvScan = CreateObject("Scripting.Dictionary"): snOvScan.CompareMode = 1
    For i = 0 To alC - 1
        Dim lnkS  As String: lnkS = allLong(i).normKey
        Dim lSnOS As String: lSnOS = allLong(i).shortNameOverride
        If lSnOS = "" Then lSnOS = allLong(i).shortName
        Dim lDefS As String: lDefS = ExtractShortName(allLong(i).caseName)
        If LCase(Trim(lSnOS)) <> LCase(Trim(lDefS)) Then
            If Not snOvScan.Exists(lnkS) Then snOvScan.Add lnkS, lSnOS
        End If
    Next i
    Dim snOvKey As Variant
    For Each snOvKey In snOvScan.Keys
        Dim snovNK As String: snovNK = CStr(snOvKey)
        If Not shortNameOverrides.Exists(snovNK) Then
            shortNameOverrides.Add snovNK, snOvScan(snovNK)
        End If
        If preScanInfo.Exists(snovNK) Then
            Dim snovArr() As String: snovArr = Split(preScanInfo(snovNK), "|")
            If UBound(snovArr) >= 3 Then
                preScanInfo(snovNK) = snovArr(0) & "|" & snovArr(1) & "|" & snovArr(2) & "|" & _
                                       snovArr(3) & "|" & snOvScan(snovNK) & "|" & snOvScan(snovNK)
            End If
        End If
    Next snOvKey

    For i = 0 To asc - 1
        Dim ssnl As String: ssnl = LCase(allSupra(i).shortName)
        Dim srpl As String: srpl = LCase(allSupra(i).reporter)
        Dim sNK  As String: sNK = ""
        If snToKey.Exists(ssnl) Then sNK = snToKey(ssnl)
        If sNK = "" And rpToKey.Exists(srpl) Then sNK = rpToKey(srpl)
        If sNK = "" Then sNK = TryResolveByParty(ssnl, srpl, altPartyToKey, preScanInfo)

        If sNK <> "" And Not snToKey.Exists(ssnl) Then
            Dim bIsParty As Boolean: bIsParty = False
            If altPartyToKey.Exists(ssnl) Then
                bIsParty = (LCase(Trim(altPartyToKey(ssnl))) = LCase(Trim(sNK)))
            End If
            If bIsParty Then
                Dim newSN As String: newSN = allSupra(i).shortName
                If shortNameOverrides.Exists(sNK) Then
                    shortNameOverrides(sNK) = newSN
                Else
                    shortNameOverrides.Add sNK, newSN
                End If
                snToKey.Add ssnl, sNK
                If preScanInfo.Exists(sNK) Then
                    Dim psiArr() As String: psiArr = Split(preScanInfo(sNK), "|")
                    If UBound(psiArr) >= 3 Then
                        preScanInfo(sNK) = psiArr(0) & "|" & psiArr(1) & "|" & psiArr(2) & "|" & _
                                           psiArr(3) & "|" & newSN & "|" & newSN
                    End If
                End If
            End If
        End If

        allSupra(i).normKey = sNK
        If sNK <> "" And Not multiDict.Exists(sNK) Then multiDict.Add sNK, True
        If dcc > UBound(docCites) - 1 Then ReDim Preserve docCites(0 To UBound(docCites) + 100)
        docCites(dcc) = allSupra(i): dcc = dcc + 1
    Next i

    If dcc > 1 Then SortDocCitesByAbsStart docCites, dcc

    ' ------------------------------------------------------------------
    ' Compute needsShortName: for each normKey, decide whether ANY
    ' subsequent occurrence of that case will be rendered as supra
    ' (rather than Ibid./Id.). If yes, the short name parenthetical
    ' on the first occurrence is needed; otherwise it can be skipped.
    '
    ' Rules (mirroring Phase 3's decision at line ~1161):
    '   A subsequent long cite becomes Ibid./Id. iff:
    '     - not isBare, AND
    '     - not isCompound, AND
    '     - previous docCite in same paragraph has same normKey
    '   Otherwise it becomes supra.
    '   Any "supra" docCite (parenthetical or bare) explicitly uses
    '   the short name and therefore needs it.
    ' ------------------------------------------------------------------
    Dim seenFirst As Object: Set seenFirst = CreateObject("Scripting.Dictionary")
    seenFirst.CompareMode = 1
    Dim lastNkInPara As String: lastNkInPara = ""
    Dim curParaIdx As Long: curParaIdx = -1
    Dim nsi As Long
    For nsi = 0 To dcc - 1
        If docCites(nsi).paraIdx <> curParaIdx Then
            curParaIdx = docCites(nsi).paraIdx
            lastNkInPara = ""
        End If
        Dim nsiNK As String: nsiNK = docCites(nsi).normKey
        If nsiNK <> "" Then
            If docCites(nsi).citeType = "supra" Then
                ' Any supra occurrence uses the short name explicitly
                If Not needsShortName.Exists(nsiNK) Then needsShortName.Add nsiNK, True
            ElseIf docCites(nsi).citeType = "long" Then
                If Not seenFirst.Exists(nsiNK) Then
                    seenFirst.Add nsiNK, True
                Else
                    ' Subsequent long cite: would it become Ibid./Id. or supra?
                    Dim becomesSupra As Boolean: becomesSupra = False
                    If docCites(nsi).isBare Or docCites(nsi).isCompound Then
                        becomesSupra = True
                    ElseIf lastNkInPara <> nsiNK Then
                        becomesSupra = True
                    End If
                    If becomesSupra Then
                        If Not needsShortName.Exists(nsiNK) Then needsShortName.Add nsiNK, True
                    End If
                End If
            End If
            lastNkInPara = nsiNK
        End If
    Next nsi
End Sub

'------------------------------------------------------------------------------
' Extract a DocCite from a long-cite regex match.
'------------------------------------------------------------------------------
Private Function ScanLongCite(m As Object, PARA As Paragraph, _
                               pt As String, qm() As Boolean, _
                               dc As DocCite) As Boolean
    ScanLongCite = False
    Dim mStart As Long: mStart = m.FirstIndex + 1
    Dim mLen   As Long: mLen = m.length
    If IsInsideQuote(mStart, mLen, qm) Then Exit Function
    If InStr(1, m.Value, "supra", vbTextCompare) > 0 Then Exit Function
    If InStr(1, m.Value, "Ibid.", vbBinaryCompare) > 0 Then Exit Function

    Dim sb As Long: sb = mStart - 1
    Dim foundParen As Boolean: foundParen = False
    Do While sb >= 1
        If Mid(pt, sb, 1) <> " " Then Exit Do
        sb = sb - 1
    Loop
    If sb >= 1 Then
        If Mid(pt, sb, 1) = "(" Then
            mLen = mLen + (mStart - sb): mStart = sb
            foundParen = True
        End If
    End If

    Dim sig As String: sig = ""
    If foundParen Then
        sig = DetectSignal(pt, mStart + 1)
    Else
        sig = DetectSignalBefore(pt, mStart)
        If sig <> "" Then
            Dim sigExpand As Long: sigExpand = Len(sig) + 1
            mStart = mStart - sigExpand
            mLen = mLen + sigExpand
        End If
    End If

    Dim afterPos As Long: afterPos = mStart + mLen

    ' --- Detect ", fn. omitted" tail BEFORE closing-paren extension ---
    ' Handles e.g. "(Smith ... 644, fn. omitted.)"   regex pincite group ate "644, "
    ' so afterPos points at "f" of "fn".  Also handles no-pincite case "(Smith ... 635, fn. omitted.)"
    Dim fnTail As String: fnTail = ""
    If foundParen Then
        Dim fnt1 As String, fnc1 As Long
        If MatchFnTail(pt, afterPos, fnt1, fnc1) Then
            fnTail = fnt1
            mLen = mLen + fnc1
            afterPos = afterPos + fnc1
        End If
    End If

    If foundParen Then
        Dim cP  As Boolean: cP = (m.SubMatches(6) = ")")
        Dim cPd As Boolean: cPd = False
        Dim pn  As Integer
        For pn = 1 To 2
            If afterPos <= Len(pt) Then
                Dim nC As String: nC = Mid(pt, afterPos, 1)
                If nC = ")" And Not cP Then mLen = mLen + 1: afterPos = afterPos + 1: cP = True
                If nC = "." And Not cPd Then mLen = mLen + 1: afterPos = afterPos + 1: cPd = True
            End If
        Next pn
    End If

    Dim bracketNote As String: bracketNote = ""
    Dim bnPos As Long: bnPos = afterPos
    If bnPos <= Len(pt) Then
        If Mid(pt, bnPos, 1) = " " Then bnPos = bnPos + 1
    End If
    If bnPos <= Len(pt) Then
        If Mid(pt, bnPos, 1) = "[" Then
            Dim bnEnd As Long: bnEnd = FindClosingBracket(pt, bnPos)
            If bnEnd > bnPos Then
                bracketNote = Mid(pt, bnPos, bnEnd - bnPos + 1)
                mLen = bnEnd - mStart + 1: afterPos = bnEnd + 1

                ' --- Detect ", fn. omitted" tail AFTER a bracket note ---
                ' Handles e.g. "(Smith ... 644 [italics added], fn. omitted.)"
                If fnTail = "" Then
                    Dim fnt2 As String, fnc2 As Long
                    If MatchFnTail(pt, afterPos, fnt2, fnc2) Then
                        fnTail = fnt2
                        mLen = mLen + fnc2
                        afterPos = afterPos + fnc2
                    End If
                End If

                Dim bn2P As Boolean: bn2P = False: Dim bn2Pd As Boolean: bn2Pd = False
                Dim bn2n As Integer
                For bn2n = 1 To 2
                    If afterPos <= Len(pt) Then
                        Dim bnc As String: bnc = Mid(pt, afterPos, 1)
                        If bnc = ")" And Not bn2P Then mLen = mLen + 1: afterPos = afterPos + 1: bn2P = True
                        If bnc = "." And Not bn2Pd Then mLen = mLen + 1: afterPos = afterPos + 1: bn2Pd = True
                    End If
                Next bn2n
            End If
        End If
    End If

    Dim snO       As String: snO = ""
    Dim origAfter As Long: origAfter = m.FirstIndex + 1 + m.length
    Dim sp        As Long
    If bracketNote <> "" Then sp = afterPos Else sp = origAfter
    Do While sp <= Len(pt)
        If Mid(pt, sp, 1) <> " " Then Exit Do
        sp = sp + 1
    Loop
    If sp <= Len(pt) Then
        If Mid(pt, sp, 1) = "(" Then
            Dim cp2 As Long: cp2 = InStr(sp + 1, pt, ")")
            If cp2 > sp + 1 Then
                Dim pnTxt As String: pnTxt = Mid(pt, sp + 1, cp2 - sp - 1)
                If Len(Trim(pnTxt)) > 0 And Len(Trim(pnTxt)) <= 60 And InStr(pnTxt, Chr(13)) = 0 Then
                    Dim fc As String: fc = Left(Trim(pnTxt), 1)
                    If (fc >= "A" And fc <= "Z") Or (fc >= "a" And fc <= "z") Then
                        snO = Trim(pnTxt)
                        Dim nA As Long: nA = cp2 + 1
                        Dim snBN As String: snBN = ""
                        Dim snBNPos As Long: snBNPos = nA
                        If snBNPos <= Len(pt) Then
                            If Mid(pt, snBNPos, 1) = " " Then snBNPos = snBNPos + 1
                        End If
                        If snBNPos <= Len(pt) Then
                            If Mid(pt, snBNPos, 1) = "[" Then
                                Dim snBNEnd As Long: snBNEnd = FindClosingBracket(pt, snBNPos)
                                If snBNEnd > snBNPos Then
                                    snBN = Mid(pt, snBNPos, snBNEnd - snBNPos + 1)
                                    bracketNote = snBN
                                    nA = snBNEnd + 1
                                End If
                            End If
                        End If
                        Dim c3P As Boolean: c3P = False: Dim c3Pd As Boolean: c3Pd = False
                        Dim p3 As Integer
                        For p3 = 1 To 2
                            If nA <= Len(pt) Then
                                Dim nc3 As String: nc3 = Mid(pt, nA, 1)
                                If nc3 = ")" And Not c3P Then nA = nA + 1: c3P = True
                                If nc3 = "." And Not c3Pd Then nA = nA + 1: c3Pd = True
                            End If
                        Next p3
                        mLen = nA - mStart
                    End If
                End If
            End If
        End If
    End If

    ' Bare-cite italic walk-forward
    Dim cn As String
    If Not foundParen Then
        Dim origMatchEnd As Long: origMatchEnd = m.FirstIndex + 1 + Len(m.SubMatches(0))
        Dim italicStart As Long: italicStart = 0
        Dim walkPos As Long
        For walkPos = mStart To origMatchEnd - 1
            Dim walkRng As Range
            Set walkRng = PARA.Range.Document.Range( _
                PARA.Range.start + walkPos - 1, _
                PARA.Range.start + walkPos)
            If walkRng.Font.Italic = True Then
                italicStart = walkPos
                Exit For
            End If
        Next walkPos
        Dim cnRaw As String
        If italicStart > mStart Then
            mLen = mLen - (italicStart - mStart)
            mStart = italicStart
            cnRaw = Mid(pt, italicStart, origMatchEnd - italicStart)
        Else
            cnRaw = Trim(m.SubMatches(0))
        End If
        cn = CleanCaseName(Trim(cnRaw))
    Else
        cn = CleanCaseName(Trim(m.SubMatches(0)))
    End If
    If Len(cn) < 3 Then Exit Function
    If sig <> "" And Left(cn, Len(sig) + 1) = sig & " " Then
        cn = Trim(Mid(cn, Len(sig) + 2))
    End If
    Dim nk As String: nk = LCase(Trim(cn))
    Dim sn As String: If snO <> "" Then sn = snO Else sn = ExtractShortName(cn)

    dc.citeType = "long"
    dc.normKey = nk
    dc.caseName = cn
    dc.shortName = sn
    dc.reporter = Trim(m.SubMatches(2)) & " " & Trim(m.SubMatches(3))
    dc.initialPage = Trim(m.SubMatches(4))
    dc.pincite = CleanPincite(Trim(m.SubMatches(5)))
    dc.year = m.SubMatches(1)
    dc.signal = sig
    dc.shortNameOverride = snO
    dc.bracketNote = bracketNote
    dc.fnTail = fnTail
    dc.absStart = PARA.Range.start + mStart - 1
    dc.textLen = mLen
    dc.isBare = Not foundParen

    Dim cpChar As String
    Dim cpPos  As Long: cpPos = afterPos
    Do While cpPos <= Len(pt)
        If Mid(pt, cpPos, 1) <> " " Then Exit Do
        cpPos = cpPos + 1
    Loop
    If cpPos <= Len(pt) Then cpChar = Mid(pt, cpPos, 1) Else cpChar = ""
    dc.isCompound = foundParen And (cpChar = ";" Or cpChar = "]")

    ScanLongCite = True
End Function

'------------------------------------------------------------------------------
Private Function ScanSupraCite(m As Object, PARA As Paragraph, _
                                pt As String, qm() As Boolean, _
                                dc As DocCite) As Boolean
    ScanSupraCite = False
    Dim s As Long: s = m.FirstIndex + 1
    Dim l As Long: l = m.length
    If IsInsideQuote(s, l, qm) Then Exit Function

    dc.citeType = "supra"
    dc.signal = Trim(m.SubMatches(0))
    dc.shortName = Trim(m.SubMatches(1))
    dc.reporter = Trim(m.SubMatches(2))
    dc.pincite = CleanPincite(Trim(m.SubMatches(3)))
    dc.absStart = PARA.Range.start + s - 1
    dc.textLen = l
    dc.normKey = ""
    dc.isBare = False
    ScanSupraCite = True
End Function

'==============================================================================
' PHASE 1.5 - IN-QUOTE CITATION COMPLETIONS
'==============================================================================
Private Sub ProcessInQuoteCompletions(Doc As Document, _
                                       hintLines() As HintLine, _
                                       ByVal hlC As Long, _
                                       usedHints() As Boolean, _
                                       docCites() As DocCite, _
                                       ByVal dcc As Long, _
                                       snToKey As Object, _
                                       rpToKey As Object, _
                                       altPartyToKey As Object, _
                                       preScanInfo As Object, _
                                       b15Dict As Object, _
                                       ByRef b15Count As Long)

    Dim reSupra As Object: Set reSupra = CreateObject("VBScript.RegExp")
    reSupra.Global = True: reSupra.Multiline = False
    reSupra.Pattern = BuildSupraPattern()

    Dim repAbsStart(99)  As Long
    Dim repAbsEnd(99)    As Long
    Dim repNewText(99)   As String
    Dim repItalicOff(99) As Long
    Dim repItalicLen(99) As Long
    Dim repItalicOff2(99) As Long
    Dim repItalicLen2(99) As Long
    Dim repC As Long

    Dim PARA As Paragraph
    For Each PARA In Doc.Paragraphs
        Dim pt As String: pt = PARA.Range.text
        If Right(pt, 1) = Chr(13) Then pt = Left(pt, Len(pt) - 1)
        pt = NormalizeSpaces(pt)
        If Len(pt) = 0 Then GoTo PIQCNext

        Dim qm() As Boolean: ReDim qm(1 To Len(pt)): BuildQuoteMask pt, qm

        Dim hasQ As Boolean: hasQ = False
        Dim qi As Long
        For qi = 1 To Len(pt)
            If qm(qi) Then hasQ = True: Exit For
        Next qi
        If Not hasQ Then GoTo PIQCNext

        repC = 0

        '--- B1 ---
        Dim msS As Object: Set msS = reSupra.Execute(pt)
        Dim mS As Object
        For Each mS In msS
            Dim bss As Long: bss = mS.FirstIndex + 1
            Dim bsl As Long: bsl = mS.length
            If IsInsideQuote(bss, bsl, qm) Then
                Dim bSig  As String: bSig = Trim(mS.SubMatches(0))
                Dim bName As String: bName = Trim(mS.SubMatches(1))
                Dim bRep  As String: bRep = Trim(mS.SubMatches(2))
                Dim bPin  As String: bPin = CleanPincite(Trim(mS.SubMatches(3)))
                Dim bDot  As String
                If mS.SubMatches.count >= 6 Then
                    bDot = "" & mS.SubMatches(5)
                Else
                    bDot = ""
                End If

                Dim b1MT As String: b1MT = mS.Value
                Dim b1AtP As Long: b1AtP = InStr(1, b1MT, " at p", vbTextCompare)
                If b1AtP > 0 Then
                    Dim b1PS As Long: b1PS = bss + b1AtP - 1
                    Do While b1PS <= Len(pt)
                        If Mid(pt, b1PS, 1) >= "0" And Mid(pt, b1PS, 1) <= "9" Then Exit Do
                        b1PS = b1PS + 1
                    Loop
                    Do While b1PS <= Len(pt)
                        Dim b1PC As String: b1PC = Mid(pt, b1PS, 1)
                        If (b1PC >= "0" And b1PC <= "9") Or b1PC = "-" Or b1PC = "," Or b1PC = " " Then
                            b1PS = b1PS + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    Do While b1PS <= Len(pt)
                        If Mid(pt, b1PS, 1) <> " " Then Exit Do
                        b1PS = b1PS + 1
                    Loop
                    If b1PS <= Len(pt) Then
                        If Mid(pt, b1PS, 1) = "[" Then
                            Dim b1TE As Long: b1TE = FindClosingBracket(pt, b1PS)
                            If b1TE > b1PS Then
                                Dim b1NE As Long: b1NE = b1TE + 1
                                If b1NE <= Len(pt) Then
                                    If Mid(pt, b1NE, 1) = "." Then b1NE = b1NE + 1
                                End If
                                If b1NE <= Len(pt) Then
                                    If Mid(pt, b1NE, 1) = ")" Then b1NE = b1NE + 1
                                End If
                                Dim b1NL As Long: b1NL = b1NE - bss
                                If b1NL > bsl Then bsl = b1NL
                            End If
                        End If
                    End If
                End If

                Dim hIdx As Long: hIdx = MatchHintLine(bName, bRep, hintLines, hlC)
                Dim b1HL As HintLine
                Dim b1UsedHint As Boolean: b1UsedHint = False
                If hIdx >= 0 Then
                    b1HL = hintLines(hIdx): b1UsedHint = True
                Else
                    Dim b1NK As String: b1NK = ""
                    If rpToKey.Exists(LCase(Trim(bRep))) Then b1NK = rpToKey(LCase(Trim(bRep)))
                    If b1NK = "" Then
                        If snToKey.Exists(LCase(Trim(bName))) Then b1NK = snToKey(LCase(Trim(bName)))
                    End If
                    If b1NK <> "" And preScanInfo.Exists(b1NK) Then
                        b1HL = HintLineFromNormKey(b1NK, preScanInfo)
                    Else
                        GoTo B1NextS
                    End If
                End If

                If repC < 99 Then
                    Dim b1IO As Long, b1IL As Long, b1IO2 As Long, b1IL2 As Long
                    Dim b1NT As String
                    b1NT = BuildB1Transformed(bSig, bName, bRep, bPin, bDot, b1HL, b1IO, b1IL, b1IO2, b1IL2)
                    repAbsStart(repC) = PARA.Range.start + bss - 1
                    repAbsEnd(repC) = PARA.Range.start + bss - 1 + bsl
                    repNewText(repC) = b1NT
                    repItalicOff(repC) = b1IO
                    repItalicLen(repC) = b1IL
                    repItalicOff2(repC) = b1IO2
                    repItalicLen2(repC) = b1IL2
                    repC = repC + 1
                    If b1UsedHint Then
                        If hIdx >= 0 And hIdx <= UBound(usedHints) Then usedHints(hIdx) = True
                    End If
                    RegisterInB15Dict b1HL, b15Dict, snToKey
                    b15Count = b15Count + 1
                End If
B1NextS:
            End If
        Next mS

        '--- B2 ---
        Dim b2Rng As Range: Set b2Rng = PARA.Range.Duplicate
        Do
            b2Rng.Find.ClearFormatting
            b2Rng.Find.Font.Italic = True
            b2Rng.Find.text = ""
            b2Rng.Find.Forward = True
            b2Rng.Find.Wrap = wdFindStop
            b2Rng.Find.MatchWildcards = False
            If Not b2Rng.Find.Execute Then Exit Do

            Dim irAbsS As Long: irAbsS = b2Rng.start
            Dim irAbsE As Long: irAbsE = b2Rng.End
            Dim irPPos As Long: irPPos = irAbsS - PARA.Range.start + 1
            Dim irLen  As Long: irLen = irAbsE - irAbsS

            Dim isInLongCite As Boolean: isInLongCite = False
            Dim dci As Long
            For dci = 0 To dcc - 1
                If docCites(dci).citeType = "long" Then
                    If irAbsS >= docCites(dci).absStart And _
                       irAbsS < docCites(dci).absStart + docCites(dci).textLen Then
                        isInLongCite = True: Exit For
                    End If
                End If
            Next dci
            If isInLongCite Then GoTo B2NextRun

            If irPPos >= 1 And irPPos + irLen - 1 <= Len(pt) Then
                If IsInsideQuote(irPPos, irLen, qm) Then
                    Dim irTxt   As String: irTxt = b2Rng.text
                    Dim cleanIR As String: cleanIR = cleanItalicText(irTxt)
                    If Len(cleanIR) >= 2 Then

                        Dim b2Keys(49)   As String
                        Dim b2MatchCount As Long: b2MatchCount = 0
                        FindAllMatchingNormKeys cleanIR, preScanInfo, snToKey, altPartyToKey, _
                                                b2Keys, b2MatchCount

                        Dim b2NormKey As String: b2NormKey = ""
                        If b2MatchCount = 1 Then
                            b2NormKey = b2Keys(0)
                        ElseIf b2MatchCount > 1 Then
                            b2NormKey = DisambiguateByReporter(b2Keys, b2MatchCount, pt, qm, preScanInfo)
                            If b2NormKey = "" Then
                                Dim yRng As Range
                                Set yRng = Doc.Range(irAbsS, irAbsE)
                                yRng.HighlightColorIndex = wdYellow
                                GoTo B2NextRun
                            End If
                        End If

                        If b2NormKey <> "" Then
                            If Not HasPriorLongCiteInDocCites(b2NormKey, irAbsS, docCites, dcc) Then
                                Dim b2HIdx As Long: b2HIdx = FindHintLineByNormKey(b2NormKey, hintLines, hlC)
                                Dim b2HL   As HintLine
                                If b2HIdx >= 0 Then
                                    b2HL = hintLines(b2HIdx)
                                Else
                                    b2HL = HintLineFromNormKey(b2NormKey, preScanInfo)
                                End If

                                Dim b2IO As Long, b2IL As Long, b2IO2 As Long, b2IL2 As Long
                                Dim b2NT As String
                                b2NT = BuildB2Replacement(cleanIR, b2HL, b2IO, b2IL, b2IO2, b2IL2)
                                If b2NT <> cleanIR And repC < 99 Then
                                    Dim isDup As Boolean: isDup = False
                                    Dim dri As Long
                                    For dri = 0 To repC - 1
                                        If repAbsStart(dri) = irAbsS Then isDup = True: Exit For
                                    Next dri
                                    If Not isDup Then
                                        repAbsStart(repC) = irAbsS
                                        repAbsEnd(repC) = irAbsE
                                        repNewText(repC) = b2NT
                                        repItalicOff(repC) = b2IO
                                        repItalicLen(repC) = b2IL
                                        repItalicOff2(repC) = b2IO2
                                        repItalicLen2(repC) = b2IL2
                                        repC = repC + 1
                                        If b2HIdx >= 0 And b2HIdx <= UBound(usedHints) Then usedHints(b2HIdx) = True
                                        RegisterInB15Dict b2HL, b15Dict, snToKey
                                        b15Count = b15Count + 1
                                    End If
                                End If
                            End If
                        End If

                    End If
                End If
            End If

B2NextRun:
            Dim b2Next As Long: b2Next = b2Rng.End
            If b2Next >= PARA.Range.End Then Exit Do
            Set b2Rng = Doc.Range(b2Next, PARA.Range.End)
        Loop

        If repC > 0 Then
            Dim si2 As Long, sj2 As Long
            For si2 = 0 To repC - 2
                For sj2 = si2 + 1 To repC - 1
                    If repAbsStart(sj2) > repAbsStart(si2) Then
                        Dim tS As Long: tS = repAbsStart(si2): repAbsStart(si2) = repAbsStart(sj2): repAbsStart(sj2) = tS
                        Dim tE As Long: tE = repAbsEnd(si2): repAbsEnd(si2) = repAbsEnd(sj2): repAbsEnd(sj2) = tE
                        Dim tT As String: tT = repNewText(si2): repNewText(si2) = repNewText(sj2): repNewText(sj2) = tT
                        Dim tIO As Long: tIO = repItalicOff(si2): repItalicOff(si2) = repItalicOff(sj2): repItalicOff(sj2) = tIO
                        Dim tIL As Long: tIL = repItalicLen(si2): repItalicLen(si2) = repItalicLen(sj2): repItalicLen(sj2) = tIL
                        Dim tIO2 As Long: tIO2 = repItalicOff2(si2): repItalicOff2(si2) = repItalicOff2(sj2): repItalicOff2(sj2) = tIO2
                        Dim tIL2 As Long: tIL2 = repItalicLen2(si2): repItalicLen2(si2) = repItalicLen2(sj2): repItalicLen2(sj2) = tIL2
                    End If
                Next sj2
            Next si2
            Dim ri As Long
            For ri = 0 To repC - 1
                Dim repRng As Range
                Set repRng = Doc.Range(repAbsStart(ri), repAbsEnd(ri))
                repRng.text = repNewText(ri)
                TrimTrailingItalic Doc, repAbsStart(ri), Len(repNewText(ri))
                Dim flatRng As Range
                Set flatRng = Doc.Range(repAbsStart(ri), repAbsStart(ri) + Len(repNewText(ri)))
                flatRng.Font.Italic = False
                If repItalicLen(ri) > 0 Then
                    Dim italRng As Range
                    Set italRng = Doc.Range(repAbsStart(ri) + repItalicOff(ri), _
                                            repAbsStart(ri) + repItalicOff(ri) + repItalicLen(ri))
                    italRng.Font.Italic = True
                End If
                If repItalicLen2(ri) > 0 Then
                    Dim italRng2 As Range
                    Set italRng2 = Doc.Range(repAbsStart(ri) + repItalicOff2(ri), _
                                             repAbsStart(ri) + repItalicOff2(ri) + repItalicLen2(ri))
                    italRng2.Font.Italic = True
                End If
            Next ri
        End If

PIQCNext:
    Next PARA
End Sub

'==============================================================================
' PHASE 2 - FIX ORPHAN SUPRAS
'==============================================================================
Private Function PerformSwaps(Doc As Document, _
                               docCites() As DocCite, _
                               dcc As Long, _
                               snToKey As Object, _
                               rpToKey As Object, _
                               multiDict As Object, _
                               shortNameOverrides As Object) As Long
    PerformSwaps = 0
    If dcc = 0 Then Exit Function

    Dim firstLongIdx As Object: Set firstLongIdx = CreateObject("Scripting.Dictionary")
    firstLongIdx.CompareMode = 1
    Dim i As Long
    For i = 0 To dcc - 1
        If docCites(i).citeType = "long" And Not firstLongIdx.Exists(docCites(i).normKey) Then
            firstLongIdx.Add docCites(i).normKey, i
        End If
    Next i

    Dim seenLong  As Object: Set seenLong = CreateObject("Scripting.Dictionary"): seenLong.CompareMode = 1
    Dim swapIdxs() As Long: ReDim swapIdxs(0 To 100): Dim swapC As Long: swapC = 0

    For i = 0 To dcc - 1
        If docCites(i).citeType = "long" Then
            seenLong(docCites(i).normKey) = True
        ElseIf docCites(i).citeType = "supra" Then
            Dim sNK As String: sNK = docCites(i).normKey
            If sNK = "" Then
                Dim snl As String: snl = LCase(docCites(i).shortName)
                Dim rpl As String: rpl = LCase(docCites(i).reporter)
                If snToKey.Exists(snl) Then sNK = snToKey(snl)
                If sNK = "" And rpToKey.Exists(rpl) Then sNK = rpToKey(rpl)
                docCites(i).normKey = sNK
            End If
            If sNK <> "" And Not seenLong.Exists(sNK) Then
                If swapC > UBound(swapIdxs) - 1 Then ReDim Preserve swapIdxs(0 To UBound(swapIdxs) + 100)
                swapIdxs(swapC) = i: swapC = swapC + 1
            End If
        End If
    Next i

    If swapC = 0 Then Exit Function

    Dim j As Long, tmp As Long
    For i = 0 To swapC - 2
        For j = i + 1 To swapC - 1
            If docCites(swapIdxs(j)).absStart > docCites(swapIdxs(i)).absStart Then
                tmp = swapIdxs(i): swapIdxs(i) = swapIdxs(j): swapIdxs(j) = tmp
            End If
        Next j
    Next i

    For i = 0 To swapC - 1
        Dim si  As Long: si = swapIdxs(i)
        Dim lnk As String: lnk = docCites(si).normKey
        If Not firstLongIdx.Exists(lnk) Then GoTo NextSwap
        Dim li2 As Long: li2 = firstLongIdx(lnk)

        Dim newLong As String
        Dim swEffSN As String
        If shortNameOverrides.Exists(lnk) Then
            swEffSN = shortNameOverrides(lnk)
        ElseIf docCites(li2).shortNameOverride <> "" Then
            swEffSN = docCites(li2).shortNameOverride
        Else
            swEffSN = docCites(li2).shortName
        End If
        newLong = BuildFullCiteTextSN(docCites(li2), docCites(si).pincite, _
                                      docCites(si).signal, True, swEffSN, _
                                      docCites(si).fnTail)

        Dim swRng As Range
        Set swRng = Doc.Range(start:=docCites(si).absStart, _
                              End:=docCites(si).absStart + docCites(si).textLen)
        swRng.text = newLong

        TrimTrailingItalic Doc, docCites(si).absStart, Len(newLong)

        Dim cnAbs As Long: cnAbs = docCites(si).absStart + 1
        If docCites(si).signal <> "" Then cnAbs = cnAbs + Len(docCites(si).signal) + 1
        Dim cnLen As Long: cnLen = Len(docCites(li2).caseName)
        If cnLen > 0 Then
            Dim cnRng As Range
            Set cnRng = Doc.Range(start:=cnAbs, End:=cnAbs + cnLen)
            cnRng.Font.Italic = True
        End If

        Dim pnQ2 As Long: pnQ2 = InStrRev(newLong, "(" & swEffSN & ")")
        If pnQ2 > 0 Then
            Dim pnRng2 As Range
            Set pnRng2 = Doc.Range(start:=docCites(si).absStart + pnQ2, _
                                    End:=docCites(si).absStart + pnQ2 + Len(swEffSN))
            pnRng2.Font.Italic = True
        End If

        PerformSwaps = PerformSwaps + 1
NextSwap:
    Next i
End Function

'==============================================================================
' PHASE 3 - PROCESS ONE PARAGRAPH
'==============================================================================
Private Sub ProcessParagraph(PARA As Paragraph, _
                              caseDict As Object, _
                              snToKey As Object, _
                              rpToKey As Object, _
                              multiDict As Object, _
                              preScanInfo As Object, _
                              altPartyToKey As Object, _
                              shortNameOverrides As Object, _
                              b15Dict As Object, _
                              ByRef changeCount As Long, _
                              ByRef gPrevKey As String, _
                              ByRef gPrevPincite As String, _
                              ByRef orphanCount As Long, _
                              needsShortName As Object)

    Dim pt As String
    pt = PARA.Range.text
    If Len(pt) > 0 And Right(pt, 1) = Chr(13) Then pt = Left(pt, Len(pt) - 1)
    pt = NormalizeSpaces(pt)
    If Len(pt) = 0 Then Exit Sub

    Dim qm() As Boolean: ReDim qm(1 To Len(pt)): BuildQuoteMask pt, qm

    Dim cits() As CitInfo: ReDim cits(0 To 60): Dim citC As Long: citC = 0
    FindFullCitations PARA, pt, qm, cits, citC

    Dim scs() As ShortCiteInfo: ReDim scs(0 To 60): Dim scC As Long: scC = 0
    FindExistingShortCites pt, qm, scs, scC

    If citC = 0 And scC = 0 Then Exit Sub

    Dim reps() As RepInfo: ReDim reps(0 To citC + scC + 10): Dim repC As Long: repC = 0

    Dim prevKey     As String: prevKey = ""
    Dim prevPincite As String: prevPincite = ""
    Dim li As Long: li = 0
    Dim si As Long: si = 0

    Do While li < citC Or si < scC

        Dim doLong As Boolean
        If li < citC And si < scC Then
            doLong = (cits(li).startChar <= scs(si).startChar)
        ElseIf li < citC Then: doLong = True
        Else: doLong = False
        End If

        If doLong Then
            Dim cit As CitInfo: cit = cits(li)
            Dim nk  As String: nk = LCase(Trim(cit.caseName))

            If Not caseDict.Exists(nk) Then
                Dim regSN As String
                If cit.shortNameOverride <> "" Then
                    regSN = cit.shortNameOverride
                ElseIf shortNameOverrides.Exists(nk) Then
                    regSN = shortNameOverrides(nk)
                Else
                    regSN = cit.shortName
                End If
                caseDict.Add nk, regSN & "|" & cit.reporter & "|" & cit.initialPage & _
                                  "|" & cit.year & "|" & cit.caseName

                If Not snToKey.Exists(LCase(regSN)) Then snToKey.Add LCase(regSN), nk
                If Not rpToKey.Exists(LCase(cit.reporter)) Then rpToKey.Add LCase(cit.reporter), nk

                If multiDict.Exists(nk) And cit.shortNameOverride = "" And needsShortName.Exists(nk) Then
                    Dim sigLenF As Long: sigLenF = IIf(cit.signal <> "", Len(cit.signal) + 1, 0)
                    Dim fRep As RepInfo
                    fRep.citStartChar = cit.startChar
                    fRep.citLength = cit.length
                    fRep.newText = BuildFullCiteText2(cit, regSN)
                    fRep.italicWord = ""
                    fRep.italicOffset = -1
                    fRep.caseNameLen = Len(cit.caseName)
                    fRep.isSupra = False
                    fRep.isFirstOccurrence = True
                    fRep.signalLen = sigLenF
                    fRep.isBare = cit.isBare
                    Dim fpTxt As String: fpTxt = fRep.newText
                    Dim pnQ As Long: pnQ = InStrRev(fpTxt, "(" & regSN & ")")
                    fRep.parenNameOffset = pnQ
                    fRep.parenNameLen = Len(regSN)
                    reps(repC) = fRep: repC = repC + 1
                End If

                prevKey = nk: prevPincite = cit.pincite
                gPrevKey = nk: gPrevPincite = cit.pincite

            Else
                Dim rv  As String: rv = caseDict(nk)
                Dim rpa() As String: rpa = Split(rv, "|")
                Dim rSN  As String
                Dim rRep As String
                If UBound(rpa) >= 1 Then
                    rSN = rpa(0)
                    rRep = rpa(1)
                Else
                    rSN = cit.shortName
                    rRep = cit.reporter
                End If

                Dim newTxt  As String
                Dim itWd    As String
                Dim itOff   As Long
                Dim sigLen2 As Long: sigLen2 = IIf(cit.signal <> "", Len(cit.signal) + 1, 0)

                If Not cit.isBare And Not cit.isCompound And nk = prevKey And prevKey <> "" Then
                    If cit.pincite = prevPincite Then
                        newTxt = BuildIbidStr(cit.isMidSentence, cit.bracketNote)
                        itWd = "Ibid."
                    Else
                        newTxt = BuildIdStr(cit.pincite, cit.isMidSentence, cit.bracketNote)
                        itWd = "Id."
                    End If
                Else
                    If cit.isCompound Then
                        newTxt = "(" & BuildSupraStr(cit.signal, rSN, rRep, cit.pincite, True, True, cit.bracketNote)
                    Else
                        newTxt = BuildSupraStr(cit.signal, rSN, rRep, cit.pincite, cit.isMidSentence, cit.isBare, cit.bracketNote)
                    End If
                    itWd = "supra"
                End If

                itOff = InStr(1, newTxt, itWd, vbBinaryCompare) - 1

                Dim cRep As RepInfo
                cRep.citStartChar = cit.startChar
                cRep.citLength = cit.length
                cRep.newText = newTxt
                cRep.italicWord = itWd
                cRep.italicOffset = itOff
                cRep.caseNameLen = Len(rSN)
                cRep.isSupra = (itWd = "supra")
                cRep.isFirstOccurrence = False
                cRep.signalLen = sigLen2
                cRep.isBare = cit.isBare
                cRep.isCompound = cit.isCompound
                reps(repC) = cRep: repC = repC + 1

                prevKey = nk: prevPincite = cit.pincite
                gPrevKey = nk: gPrevPincite = cit.pincite
            End If
            li = li + 1

        Else
            Dim sC As ShortCiteInfo: sC = scs(si)
            ClearHighlightIfYellow PARA, sC.startChar, sC.citLength

            If sC.inQuote Then
                If sC.citeType = "supra" Then
                    Dim iqNK As String: iqNK = ""
                    If snToKey.Exists(LCase(sC.shortName)) Then iqNK = snToKey(LCase(sC.shortName))
                    If iqNK = "" And rpToKey.Exists(LCase(sC.reporter)) Then iqNK = rpToKey(LCase(sC.reporter))
                    If iqNK = "" Or Not caseDict.Exists(iqNK) Then
                        ApplyHighlight PARA, sC.startChar, sC.citLength
                        orphanCount = orphanCount + 1
                    Else
                        prevKey = iqNK: prevPincite = sC.pincite
                        gPrevKey = iqNK: gPrevPincite = sC.pincite
                    End If
                ElseIf sC.citeType = "ibid" Or sC.citeType = "id" Then
                    If Not HasPrecedingFullCiteInQuote(pt, qm, sC.startChar) Then
                        ApplyHighlight PARA, sC.startChar, sC.citLength
                        orphanCount = orphanCount + 1
                    Else
                        If sC.citeType = "id" Then
                            prevPincite = sC.pincite: gPrevPincite = sC.pincite
                        End If
                    End If
                End If
                GoTo NextSC
            End If

            Dim scNK      As String: scNK = ""
            Dim scSN      As String: scSN = ""
            Dim scRep     As String: scRep = ""
            Dim scPincite As String: scPincite = sC.pincite
            Dim scOK      As Boolean: scOK = False

            Select Case sC.citeType

                Case "supra"
                    If snToKey.Exists(LCase(sC.shortName)) Then scNK = snToKey(LCase(sC.shortName))
                    If scNK = "" And rpToKey.Exists(LCase(sC.reporter)) Then scNK = rpToKey(LCase(sC.reporter))

                    If scNK = "" Then
                        scNK = TryResolveByParty(LCase(sC.shortName), LCase(sC.reporter), _
                                                  altPartyToKey, preScanInfo)
                        If scNK <> "" Then
                            If Not shortNameOverrides.Exists(scNK) Then
                                shortNameOverrides(scNK) = sC.shortName
                            End If
                            If Not snToKey.Exists(LCase(sC.shortName)) Then
                                snToKey.Add LCase(sC.shortName), scNK
                            End If
                            If preScanInfo.Exists(scNK) Then
                                Dim ppArr() As String: ppArr = Split(preScanInfo(scNK), "|")
                                If UBound(ppArr) >= 3 Then
                                    preScanInfo(scNK) = ppArr(0) & "|" & ppArr(1) & "|" & ppArr(2) & _
                                                        "|" & ppArr(3) & "|" & sC.shortName & "|" & sC.shortName
                                End If
                            End If
                        End If
                    End If

                    If scNK = "" Then
                        ApplyHighlight PARA, sC.startChar, sC.citLength
                        orphanCount = orphanCount + 1
                        GoTo NextSC
                    End If

                    If Not caseDict.Exists(scNK) Then
                        If b15Dict.Exists(scNK) Then
                            caseDict.Add scNK, b15Dict(scNK)
                            Dim b15v() As String: b15v = Split(b15Dict(scNK), "|")
                            If UBound(b15v) >= 1 Then
                                If Not snToKey.Exists(LCase(b15v(0))) Then snToKey.Add LCase(b15v(0)), scNK
                                If Not rpToKey.Exists(LCase(b15v(1))) Then rpToKey.Add LCase(b15v(1)), scNK
                            End If
                            Dim scRVb() As String: scRVb = Split(caseDict(scNK), "|")
                            If UBound(scRVb) >= 1 Then
                                scSN = scRVb(0): scRep = scRVb(1): scOK = True
                            End If
                            GoTo SCSupraDone
                        End If
                        If Not preScanInfo.Exists(scNK) Then
                            ApplyHighlight PARA, sC.startChar, sC.citLength
                            orphanCount = orphanCount + 1
                            GoTo NextSC
                        End If
                        Dim psi()  As String: psi = Split(preScanInfo(scNK), "|")
                        If UBound(psi) < 5 Then
                            ApplyHighlight PARA, sC.startChar, sC.citLength
                            orphanCount = orphanCount + 1
                            GoTo NextSC
                        End If
                        Dim psiCN  As String: psiCN = psi(0)
                        Dim psiYr  As String: psiYr = psi(1)
                        Dim psiRp  As String: psiRp = psi(2)
                        Dim psiIP  As String: psiIP = psi(3)
                        Dim psiSN  As String: psiSN = psi(4)
                        Dim psiSnO As String: psiSnO = psi(5)
                        Dim psiEff As String: If psiSnO <> "" Then psiEff = psiSnO Else psiEff = psiSN
                        caseDict.Add scNK, psiEff & "|" & psiRp & "|" & psiIP & "|" & psiYr & "|" & psiCN
                        If Not snToKey.Exists(LCase(psiEff)) Then snToKey.Add LCase(psiEff), scNK
                        If Not rpToKey.Exists(LCase(psiRp)) Then rpToKey.Add LCase(psiRp), scNK
                        Dim psiPin As String: psiPin = sC.pincite
                        Dim psiBN  As String: psiBN = IIf(sC.bracketNote <> "", " " & sC.bracketNote, "")
                        Dim psiTxt As String: psiTxt = "("
                        If sC.signal <> "" Then psiTxt = psiTxt & sC.signal & " "
                        psiTxt = psiTxt & psiCN & " (" & psiYr & ") " & psiRp & " " & psiIP
                        If psiPin <> "" Then psiTxt = psiTxt & ", " & psiPin
                        psiTxt = psiTxt & " (" & psiEff & ")" & psiBN & ".)"
                        Dim psiSL As Long: psiSL = IIf(sC.signal <> "", Len(sC.signal) + 1, 0)
                        Dim psiPQ As Long: psiPQ = InStrRev(psiTxt, "(" & psiEff & ")")
                        Dim psiR  As RepInfo
                        psiR.citStartChar = sC.startChar
                        psiR.citLength = sC.citLength
                        psiR.newText = psiTxt
                        psiR.italicWord = ""
                        psiR.italicOffset = -1
                        psiR.caseNameLen = Len(psiCN)
                        psiR.isSupra = False
                        psiR.isFirstOccurrence = True
                        psiR.signalLen = psiSL
                        psiR.parenNameOffset = psiPQ
                        psiR.parenNameLen = Len(psiEff)
                        psiR.isBare = False
                        reps(repC) = psiR: repC = repC + 1
                        prevKey = scNK: prevPincite = psiPin
                        gPrevKey = scNK: gPrevPincite = psiPin
                        GoTo NextSC
                    End If

                    Dim scRVs() As String: scRVs = Split(caseDict(scNK), "|")
                    If UBound(scRVs) >= 1 Then
                        scSN = scRVs(0): scRep = scRVs(1): scOK = True
                    End If
SCSupraDone:

                Case "ibid"
                    ' If the most recent preceding parenthetical is a non-case
                    ' document reference (Declaration, Motion, Brief, etc.),
                    ' leave the (Ibid.) alone -- it refers to the non-case doc.
                    If RefersToNonCaseDocument(pt, sC.startChar, PARA) Then
                        GoTo NextSC
                    End If
                    If prevKey <> "" Then
                        scNK = prevKey: scPincite = prevPincite
                    ElseIf gPrevKey <> "" And caseDict.Exists(gPrevKey) Then
                        scNK = gPrevKey: scPincite = gPrevPincite
                    Else
                        ApplyHighlight PARA, sC.startChar, sC.citLength
                        orphanCount = orphanCount + 1
                        GoTo NextSC
                    End If
                    Dim scRVi() As String: scRVi = Split(caseDict(scNK), "|")
                    If UBound(scRVi) >= 1 Then
                        scSN = scRVi(0): scRep = scRVi(1): scOK = True
                    End If

                Case "id"
                    ' If the most recent preceding parenthetical is a non-case
                    ' document reference (Declaration, Motion, Brief, etc.),
                    ' leave the (Id. at ...) alone -- it refers to the non-case doc.
                    If RefersToNonCaseDocument(pt, sC.startChar, PARA) Then
                        GoTo NextSC
                    End If
                    If prevKey <> "" Then
                        scNK = prevKey
                    ElseIf gPrevKey <> "" And caseDict.Exists(gPrevKey) Then
                        scNK = gPrevKey
                    Else
                        ApplyHighlight PARA, sC.startChar, sC.citLength
                        orphanCount = orphanCount + 1
                        GoTo NextSC
                    End If
                    Dim scRVd() As String: scRVd = Split(caseDict(scNK), "|")
                    If UBound(scRVd) >= 1 Then
                        scSN = scRVd(0): scRep = scRVd(1): scOK = True
                    End If

            End Select

            If Not scOK Then GoTo NextSC

            Dim scNewTxt As String
            Dim scItWd   As String
            Dim scItOff  As Long
            Dim scSigLen As Long: scSigLen = IIf(sC.signal <> "", Len(sC.signal) + 1, 0)

            If Not sC.isBare And scNK = prevKey And prevKey <> "" Then
                If scPincite = prevPincite Then
                    scNewTxt = BuildIbidStr(sC.isMidSentence, sC.bracketNote)
                    scItWd = "Ibid."
                Else
                    scNewTxt = BuildIdStr(scPincite, sC.isMidSentence, sC.bracketNote)
                    scItWd = "Id."
                End If
            Else
                scNewTxt = BuildSupraStr(sC.signal, scSN, scRep, scPincite, sC.isMidSentence, sC.isBare, sC.bracketNote)
                scItWd = "supra"
            End If

            If scNewTxt = Mid(pt, sC.startChar, sC.citLength) Then
                prevKey = scNK: prevPincite = scPincite
                gPrevKey = scNK: gPrevPincite = scPincite
                GoTo NextSC
            End If

            scItOff = InStr(1, scNewTxt, scItWd, vbBinaryCompare) - 1

            Dim scRepObj As RepInfo
            scRepObj.citStartChar = sC.startChar
            scRepObj.citLength = sC.citLength
            scRepObj.newText = scNewTxt
            scRepObj.italicWord = scItWd
            scRepObj.italicOffset = scItOff
            scRepObj.caseNameLen = Len(scSN)
            scRepObj.isSupra = (scItWd = "supra")
            scRepObj.isFirstOccurrence = False
            scRepObj.signalLen = scSigLen
            scRepObj.isBare = sC.isBare
            reps(repC) = scRepObj: repC = repC + 1

            prevKey = scNK: prevPincite = scPincite
            gPrevKey = scNK: gPrevPincite = scPincite

NextSC:
            si = si + 1
        End If
    Loop

    If repC = 0 Then Exit Sub

    SortRepsDescending reps, repC

    Dim paraStart As Long: paraStart = PARA.Range.start
    Dim j As Long
    For j = 0 To repC - 1
        ApplyReplacement PARA.Range.Document, paraStart, reps(j)
        changeCount = changeCount + 1
    Next j
End Sub

'==============================================================================
' PHASE 3.5 - PARENTHETICAL KEEP / REMOVE / RENAME
'==============================================================================
Private Sub ProcessParentheticals(Doc As Document, _
                                   caseDict As Object, _
                                   preScanInfo As Object, _
                                   multiDict As Object, _
                                   snToKey As Object, _
                                   rpToKey As Object, _
                                   altPartyToKey As Object, _
                                   ByRef parenCount As Long)

    Dim reLong As Object: Set reLong = CreateObject("VBScript.RegExp")
    reLong.Global = True: reLong.Multiline = False
    reLong.Pattern = BuildLongCitePattern()

    Dim seenFirst     As Object: Set seenFirst = CreateObject("Scripting.Dictionary"): seenFirst.CompareMode = 1
    Dim firstCiteInfo As Object: Set firstCiteInfo = CreateObject("Scripting.Dictionary"): firstCiteInfo.CompareMode = 1

    Dim PARA As Paragraph
    For Each PARA In Doc.Paragraphs
        Dim pt As String: pt = PARA.Range.text
        If Len(pt) > 0 And Right(pt, 1) = Chr(13) Then pt = Left(pt, Len(pt) - 1)
        pt = NormalizeSpaces(pt)
        If Len(pt) = 0 Then GoTo P35NextPara1

        Dim qm() As Boolean: ReDim qm(1 To Len(pt)): BuildQuoteMask pt, qm

        Dim mS As Object: Set mS = reLong.Execute(pt)
        Dim m  As Object
        For Each m In mS
            Dim dc As DocCite
            If Not ScanLongCite(m, PARA, pt, qm, dc) Then GoTo P35NextM1
            If seenFirst.Exists(dc.normKey) Then GoTo P35NextM1

            Dim effSN As String
            If dc.shortNameOverride <> "" Then
                effSN = dc.shortNameOverride
            ElseIf caseDict.Exists(dc.normKey) Then
                Dim cdArr() As String: cdArr = Split(caseDict(dc.normKey), "|")
                If UBound(cdArr) >= 0 Then effSN = cdArr(0) Else effSN = dc.shortName
            Else
                effSN = dc.shortName
            End If

            Dim hasParen As Boolean: hasParen = (dc.shortNameOverride <> "")
            If Not hasParen And multiDict.Exists(dc.normKey) Then
                Dim fullCiteTxt As String
                Dim citRng2 As Range
                Set citRng2 = Doc.Range(dc.absStart, dc.absStart + dc.textLen)
                fullCiteTxt = citRng2.text
                If InStr(fullCiteTxt, "(" & effSN & ")") > 0 Then hasParen = True
            End If

            seenFirst.Add dc.normKey, True
            firstCiteInfo.Add dc.normKey, _
                dc.absStart & "|" & dc.textLen & "|" & effSN & "|" & _
                dc.caseName & "|" & dc.reporter & "|" & IIf(hasParen, "1", "0")
P35NextM1:
        Next m
P35NextPara1:
    Next PARA

    If firstCiteInfo.count = 0 Then Exit Sub

    Dim reSupra   As Object: Set reSupra = CreateObject("VBScript.RegExp")
    Dim reBareSup As Object: Set reBareSup = CreateObject("VBScript.RegExp")
    reSupra.Global = True: reSupra.Multiline = False: reSupra.Pattern = BuildSupraPattern()
    reBareSup.Global = True: reBareSup.Multiline = False: reBareSup.Pattern = BuildBareSupraPattern()

    Dim nkv As Variant
    For Each nkv In firstCiteInfo.Keys
        Dim nk As String: nk = CStr(nkv)
        Dim fci() As String: fci = Split(firstCiteInfo(nk), "|")
        If UBound(fci) < 5 Then GoTo P35NextNK
        Dim fciAbsStart As Long: fciAbsStart = CLng(fci(0))
        Dim fciEffSN    As String: fciEffSN = fci(2)
        Dim fciCaseName As String: fciCaseName = fci(3)
        Dim fciReporter As String: fciReporter = fci(4)
        Dim fciHasParen As Boolean: fciHasParen = (fci(5) = "1")

        If Not fciHasParen Then GoTo P35NextNK

        Dim vPos As Long: vPos = InStr(1, fciCaseName, " v. ", vbTextCompare)
        Dim fciPla As String: fciPla = ""
        Dim fciDef As String: fciDef = ""
        If vPos > 1 Then
            fciPla = Trim(Left(fciCaseName, vPos - 1))
            fciDef = Trim(Mid(fciCaseName, vPos + 4))
        End If

        Dim foundExact   As Boolean: foundExact = False
        Dim foundPartial As Boolean: foundPartial = False
        Dim bestPartial  As String:  bestPartial = ""

        Dim para2 As Paragraph
        For Each para2 In Doc.Paragraphs
            Dim paraEnd As Long: paraEnd = para2.Range.End - 1
            If paraEnd <= fciAbsStart Then GoTo P35NextPara2

            Dim paraAbsStart As Long: paraAbsStart = para2.Range.start

            Dim pt2Live As String: pt2Live = para2.Range.text
            If Len(pt2Live) > 0 And Right(pt2Live, 1) = Chr(13) Then _
                pt2Live = Left(pt2Live, Len(pt2Live) - 1)
            pt2Live = NormalizeSpaces(pt2Live)

            Dim searchFrom As Long
            If paraAbsStart > fciAbsStart Then
                searchFrom = 1
            Else
                searchFrom = fciAbsStart - paraAbsStart + 2
            End If

            If searchFrom <= Len(pt2Live) Then
                Dim fciEndInPara As Long
                If paraAbsStart <= fciAbsStart Then
                    fciEndInPara = fciAbsStart - paraAbsStart + CLng(fci(1))
                Else
                    fciEndInPara = 0
                End If
                Dim hitPos As Long: hitPos = searchFrom
                Do
                    hitPos = InStr(hitPos, pt2Live, fciEffSN, vbTextCompare)
                    If hitPos = 0 Then Exit Do
                    If hitPos > fciEndInPara Then
                        foundExact = True
                        Exit Do
                    End If
                    hitPos = hitPos + 1
                Loop
            End If

            If foundExact Then GoTo P35DecideNK

            Dim pt2 As String: pt2 = pt2Live
            Dim scanFromChar As Long
            If paraAbsStart >= fciAbsStart Then
                scanFromChar = 1
            Else
                scanFromChar = fciAbsStart - paraAbsStart + 1
            End If

            Dim msS2 As Object: Set msS2 = reSupra.Execute(pt2)
            Dim mS2 As Object
            For Each mS2 In msS2
                Dim sPos2 As Long: sPos2 = mS2.FirstIndex + 1
                If sPos2 < scanFromChar Then GoTo P35NextSup2
                Dim sSN2  As String: sSN2 = Trim(mS2.SubMatches(1))
                Dim sRep2 As String: sRep2 = Trim(mS2.SubMatches(2))
                If LCase(sRep2) = LCase(fciReporter) Then
                    If LCase(sSN2) <> LCase(fciEffSN) Then
                        If IsPartyPrefix(sSN2, fciPla, fciDef) Then
                            foundPartial = True
                            If bestPartial = "" Or Len(sSN2) < Len(bestPartial) Then
                                bestPartial = sSN2
                            End If
                        End If
                    End If
                End If
P35NextSup2:
            Next mS2

            Dim msBS2 As Object: Set msBS2 = reBareSup.Execute(pt2)
            Dim mBS2 As Object
            For Each mBS2 In msBS2
                Dim bsPos2 As Long: bsPos2 = mBS2.FirstIndex + 1
                If bsPos2 < scanFromChar Then GoTo P35NextBS2
                If bsPos2 > 1 Then
                    If Mid(pt2, bsPos2 - 1, 1) = "(" Then GoTo P35NextBS2
                End If
                Dim bsSN2  As String: bsSN2 = Trim(mBS2.SubMatches(1))
                Dim bsRep2 As String: bsRep2 = Trim(mBS2.SubMatches(2))
                If LCase(bsRep2) = LCase(fciReporter) Then
                    If LCase(bsSN2) <> LCase(fciEffSN) Then
                        If IsPartyPrefix(bsSN2, fciPla, fciDef) Then
                            foundPartial = True
                            If bestPartial = "" Or Len(bsSN2) < Len(bestPartial) Then
                                bestPartial = bsSN2
                            End If
                        End If
                    End If
                End If
P35NextBS2:
            Next mBS2
P35NextPara2:
        Next para2

P35DecideNK:
        If foundExact Then GoTo P35NextNK

        If foundPartial Then
            Dim para3 As Paragraph
            For Each para3 In Doc.Paragraphs
                Dim pt3 As String: pt3 = para3.Range.text
                If Len(pt3) > 0 And Right(pt3, 1) = Chr(13) Then pt3 = Left(pt3, Len(pt3) - 1)
                pt3 = NormalizeSpaces(pt3)
                If Len(pt3) = 0 Then GoTo P35NextPara3

                Dim msR As Object: Set msR = reSupra.Execute(pt3)
                Dim mR As Object
                Dim renStarts(49) As Long, renEnds(49) As Long
                Dim renNewTxts(49) As String: Dim renC As Long: renC = 0
                For Each mR In msR
                    If LCase(Trim(mR.SubMatches(1))) = LCase(bestPartial) And _
                       LCase(Trim(mR.SubMatches(2))) = LCase(fci(4)) Then
                        Dim rSig  As String: rSig = Trim(mR.SubMatches(0))
                        Dim rPin  As String: rPin = CleanPincite(Trim(mR.SubMatches(3)))
                        Dim rBN   As String: rBN = Trim(mR.SubMatches(4))
                        Dim rRep3 As String: rRep3 = Trim(mR.SubMatches(2))
                        Dim rML As Long: rML = mR.length
                        Dim rMS As Long: rMS = mR.FirstIndex + 1
                        Dim rMT As String: rMT = mR.Value
                        Dim rAtP As Long: rAtP = InStr(1, rMT, " at p", vbTextCompare)
                        If rAtP > 0 Then
                            Dim rPS As Long: rPS = rMS + rAtP - 1
                            Do While rPS <= Len(pt3)
                                If Mid(pt3, rPS, 1) >= "0" And Mid(pt3, rPS, 1) <= "9" Then Exit Do
                                rPS = rPS + 1
                            Loop
                            Do While rPS <= Len(pt3)
                                Dim rPC As String: rPC = Mid(pt3, rPS, 1)
                                If (rPC >= "0" And rPC <= "9") Or rPC = "-" Or rPC = "," Or rPC = " " Then
                                    rPS = rPS + 1
                                Else
                                    Exit Do
                                End If
                            Loop
                            Do While rPS <= Len(pt3)
                                If Mid(pt3, rPS, 1) <> " " Then Exit Do
                                rPS = rPS + 1
                            Loop
                            If rPS <= Len(pt3) Then
                                If Mid(pt3, rPS, 1) = "[" Then
                                    Dim rTE As Long: rTE = FindClosingBracket(pt3, rPS)
                                    If rTE > rPS Then
                                        rBN = Mid(pt3, rPS, rTE - rPS + 1)
                                        Dim rNE As Long: rNE = rTE + 1
                                        If rNE <= Len(pt3) Then
                                            If Mid(pt3, rNE, 1) = "." Then rNE = rNE + 1
                                        End If
                                        If rNE <= Len(pt3) Then
                                            If Mid(pt3, rNE, 1) = ")" Then rNE = rNE + 1
                                        End If
                                        Dim rNL As Long: rNL = rNE - rMS
                                        If rNL > rML Then rML = rNL
                                    End If
                                End If
                            End If
                        End If
                        Dim rNewT As String
                        rNewT = BuildSupraStr(rSig, fciEffSN, rRep3, rPin, False, False, rBN)
                        If renC < 49 Then
                            renStarts(renC) = para3.Range.start + mR.FirstIndex
                            renEnds(renC) = para3.Range.start + mR.FirstIndex + rML
                            renNewTxts(renC) = rNewT
                            renC = renC + 1
                        End If
                    End If
                Next mR
                Dim ri3 As Long
                For ri3 = renC - 1 To 0 Step -1
                    Dim renRng As Range
                    Set renRng = Doc.Range(renStarts(ri3), renEnds(ri3))
                    renRng.text = renNewTxts(ri3)
                    TrimTrailingItalic Doc, renStarts(ri3), Len(renNewTxts(ri3))
                    Dim supIdx As Long: supIdx = InStr(renNewTxts(ri3), "supra")
                    If supIdx > 0 Then
                        Dim supRng As Range
                        Set supRng = Doc.Range(renStarts(ri3) + supIdx - 1, _
                                               renStarts(ri3) + supIdx - 1 + 5)
                        supRng.Font.Italic = True
                    End If
                    Dim snOff As Long: snOff = 1 + IIf(rSig <> "", Len(rSig) + 1, 0)
                    Dim snRng As Range
                    Set snRng = Doc.Range(renStarts(ri3) + snOff, _
                                          renStarts(ri3) + snOff + Len(fciEffSN))
                    snRng.Font.Italic = True
                Next ri3

                Dim msRB As Object: Set msRB = reBareSup.Execute(pt3)
                Dim mRB As Object
                Dim renBStarts(49) As Long, renBEnds(49) As Long
                Dim renBNewTxts(49) As String: Dim renBC As Long: renBC = 0
                For Each mRB In msRB
                    Dim bsOff As Long: bsOff = mRB.FirstIndex + 1
                    If bsOff > 1 Then
                        If Mid(pt3, bsOff - 1, 1) = "(" Then GoTo P35NextRB
                    End If
                    If LCase(Trim(mRB.SubMatches(1))) = LCase(bestPartial) And _
                       LCase(Trim(mRB.SubMatches(2))) = LCase(fci(4)) Then
                        Dim rbSig  As String: rbSig = Trim(mRB.SubMatches(0))
                        Dim rbPin  As String: rbPin = CleanPincite(Trim(mRB.SubMatches(3)))
                        Dim rbRep3 As String: rbRep3 = Trim(mRB.SubMatches(2))
                        Dim rbBN   As String: rbBN = Trim(mRB.SubMatches(4))
                        Dim rbNewT As String
                        rbNewT = BuildSupraStr(rbSig, fciEffSN, rbRep3, rbPin, False, True, rbBN)
                        If renBC < 49 Then
                            renBStarts(renBC) = para3.Range.start + mRB.FirstIndex
                            renBEnds(renBC) = para3.Range.start + mRB.FirstIndex + mRB.length
                            renBNewTxts(renBC) = rbNewT
                            renBC = renBC + 1
                        End If
                    End If
P35NextRB:
                Next mRB
                Dim rbi3 As Long
                For rbi3 = renBC - 1 To 0 Step -1
                    Dim renBRng As Range
                    Set renBRng = Doc.Range(renBStarts(rbi3), renBEnds(rbi3))
                    renBRng.text = renBNewTxts(rbi3)
                    TrimTrailingItalic Doc, renBStarts(rbi3), Len(renBNewTxts(rbi3))
                    Dim bsupIdx As Long: bsupIdx = InStr(renBNewTxts(rbi3), "supra")
                    If bsupIdx > 0 Then
                        Dim bsupRng As Range
                        Set bsupRng = Doc.Range(renBStarts(rbi3) + bsupIdx - 1, _
                                                renBStarts(rbi3) + bsupIdx - 1 + 5)
                        bsupRng.Font.Italic = True
                    End If
                    Dim bsnOff As Long: bsnOff = IIf(rbSig <> "", Len(rbSig) + 1, 0)
                    Dim bsnRng As Range
                    Set bsnRng = Doc.Range(renBStarts(rbi3) + bsnOff, _
                                           renBStarts(rbi3) + bsnOff + Len(fciEffSN))
                    bsnRng.Font.Italic = True
                Next rbi3
P35NextPara3:
            Next para3
            GoTo P35NextNK
        End If

        ' Remove parenthetical
        Dim para4 As Paragraph
        For Each para4 In Doc.Paragraphs
            If para4.Range.End - 1 < fciAbsStart Then GoTo P35NextPara4
            Dim pt4 As String: pt4 = para4.Range.text
            If Len(pt4) > 0 And Right(pt4, 1) = Chr(13) Then pt4 = Left(pt4, Len(pt4) - 1)
            pt4 = NormalizeSpaces(pt4)
            If Len(pt4) = 0 Then GoTo P35NextPara4

            Dim qm4() As Boolean: ReDim qm4(1 To Len(pt4)): BuildQuoteMask pt4, qm4
            Dim ms4 As Object: Set ms4 = reLong.Execute(pt4)
            Dim m4 As Object
            For Each m4 In ms4
                Dim dc4 As DocCite
                If Not ScanLongCite(m4, para4, pt4, qm4, dc4) Then GoTo P35NextM4
                If LCase(Trim(dc4.caseName)) <> nk Then GoTo P35NextM4

                Dim rmSig As String: rmSig = dc4.signal
                Dim rmBN  As String: rmBN = IIf(dc4.bracketNote <> "", " " & dc4.bracketNote, "")
                Dim rmFn  As String: rmFn = IIf(dc4.fnTail <> "", ", " & dc4.fnTail, "")
                Dim rmTxt As String
                If dc4.isBare Then
                    rmTxt = IIf(rmSig <> "", rmSig & " ", "") & _
                            dc4.caseName & " (" & dc4.year & ") " & dc4.reporter & " " & dc4.initialPage
                    If dc4.pincite <> "" Then rmTxt = rmTxt & ", " & dc4.pincite
                    rmTxt = rmTxt & rmBN
                Else
                    rmTxt = "("
                    If rmSig <> "" Then rmTxt = rmTxt & rmSig & " "
                    rmTxt = rmTxt & dc4.caseName & " (" & dc4.year & ") " & _
                            dc4.reporter & " " & dc4.initialPage
                    If dc4.pincite <> "" Then rmTxt = rmTxt & ", " & dc4.pincite
                    rmTxt = rmTxt & rmBN & rmFn & ".)"
                End If

                Dim rmRng As Range
                Set rmRng = Doc.Range(dc4.absStart, dc4.absStart + dc4.textLen)
                rmRng.text = rmTxt

                Dim rmFlatRng As Range
                Set rmFlatRng = Doc.Range(dc4.absStart, dc4.absStart + Len(rmTxt))
                rmFlatRng.Font.Italic = False

                Dim rmParenSkip As Long: rmParenSkip = IIf(dc4.isBare, 0, 1)
                Dim rmSigLen    As Long: rmSigLen = IIf(rmSig <> "", Len(rmSig) + 1, 0)
                Dim rmNameAbs   As Long: rmNameAbs = dc4.absStart + rmParenSkip + rmSigLen
                Dim rmNameRng   As Range
                Set rmNameRng = Doc.Range(rmNameAbs, rmNameAbs + Len(dc4.caseName))
                rmNameRng.Font.Italic = True

                parenCount = parenCount + 1
                Exit For
P35NextM4:
            Next m4
            If para4.Range.start >= fciAbsStart And para4.Range.End > fciAbsStart Then Exit For
P35NextPara4:
        Next para4

P35NextNK:
    Next nkv
End Sub

'--- FindClosingBracket ---
Private Function FindClosingBracket(text As String, startPos As Long) As Long
    FindClosingBracket = startPos
    If startPos < 1 Or startPos > Len(text) Then Exit Function
    If Mid(text, startPos, 1) <> "[" Then Exit Function
    Dim depth As Long: depth = 0
    Dim i As Long
    For i = startPos To Len(text)
        Dim c As String: c = Mid(text, i, 1)
        If c = "[" Then
            depth = depth + 1
        ElseIf c = "]" Then
            depth = depth - 1
            If depth = 0 Then
                FindClosingBracket = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ExtractBracketNoteFromSupraMatch(m As Object, pt As String) As String
    ExtractBracketNoteFromSupraMatch = ""
    Dim endParen As Long: endParen = m.FirstIndex + m.length
    Dim scanPos  As Long: scanPos = endParen - 1
    scanPos = scanPos - 1
    If scanPos >= 1 Then
        If Mid(pt, scanPos, 1) = "." Then scanPos = scanPos - 1
    End If
    Do While scanPos >= 1
        If Mid(pt, scanPos, 1) <> " " Then Exit Do
        scanPos = scanPos - 1
    Loop
    If scanPos < 1 Then Exit Function
    If Mid(pt, scanPos, 1) <> "]" Then Exit Function
    Dim bDepth As Long: bDepth = 0
    Dim bScan  As Long: bScan = scanPos
    Dim bOpen  As Long: bOpen = 0
    Do While bScan >= 1
        Dim bChr As String: bChr = Mid(pt, bScan, 1)
        If bChr = "]" Then bDepth = bDepth + 1
        If bChr = "[" Then
            bDepth = bDepth - 1
            If bDepth = 0 Then bOpen = bScan: Exit Do
        End If
        bScan = bScan - 1
    Loop
    If bOpen > 0 Then
        ExtractBracketNoteFromSupraMatch = Mid(pt, bOpen, scanPos - bOpen + 1)
    End If
End Function

Private Function IsPartyPrefix(shortName As String, plaintiff As String, defendant As String) As Boolean
    Dim sn As String: sn = LCase(Trim(shortName))
    If Len(sn) = 0 Then IsPartyPrefix = False: Exit Function
    If plaintiff <> "" And Left(LCase(Trim(plaintiff)), Len(sn)) = sn Then
        IsPartyPrefix = True: Exit Function
    End If
    If defendant <> "" And Left(LCase(Trim(defendant)), Len(sn)) = sn Then
        IsPartyPrefix = True: Exit Function
    End If
    IsPartyPrefix = False
End Function

'==============================================================================
' PHASE 4 - DELETE USED HINT-LINE PARAGRAPHS
'==============================================================================
Private Sub DeleteUsedHintLines(Doc As Document, _
                                 hintLines() As HintLine, _
                                 hlC As Long, _
                                 usedHints() As Boolean)
    Dim i As Long
    For i = hlC - 1 To 0 Step -1
        If i <= UBound(usedHints) Then
            If usedHints(i) Then
                Dim bmn As String: bmn = hintLines(i).bmName
                If Doc.Bookmarks.Exists(bmn) Then
                    Dim delRng As Range: Set delRng = Doc.Bookmarks(bmn).Range
                    If delRng.End < Doc.content.End Then delRng.End = delRng.End + 1
                    delRng.Delete
                    If Doc.Bookmarks.Exists(bmn) Then Doc.Bookmarks(bmn).Delete
                End If
            End If
        End If
    Next i
End Sub

'==============================================================================
' CITATION DETECTION - FULL CITATIONS
'==============================================================================
Private Sub FindFullCitations(PARA As Paragraph, _
                               pt As String, _
                               qm() As Boolean, _
                               citations() As CitInfo, _
                               ByRef citCount As Long)

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.Multiline = False: re.Pattern = BuildLongCitePattern()

    Dim mS As Object: Set mS = re.Execute(pt)
    Dim m  As Object
    For Each m In mS
        Dim dc As DocCite
        If Not ScanLongCite(m, PARA, pt, qm, dc) Then GoTo NextFFC

        If citCount > UBound(citations) - 1 Then ReDim Preserve citations(0 To UBound(citations) + 20)

        Dim sC As Long: sC = CLng(dc.absStart - PARA.Range.start + 1)

        citations(citCount).caseName = dc.caseName
        citations(citCount).shortName = dc.shortName
        citations(citCount).reporter = dc.reporter
        citations(citCount).initialPage = dc.initialPage
        citations(citCount).pincite = dc.pincite
        citations(citCount).year = dc.year
        citations(citCount).shortNameOverride = dc.shortNameOverride
        citations(citCount).signal = dc.signal
        citations(citCount).startChar = sC
        citations(citCount).length = dc.textLen
        citations(citCount).hasOpenParen = Not dc.isBare
        citations(citCount).hasCloseParen = Not dc.isBare
        citations(citCount).isMidSentence = DetectMidSentence(pt, sC)
        citations(citCount).isBare = dc.isBare
        citations(citCount).bracketNote = dc.bracketNote
        citations(citCount).isCompound = dc.isCompound
        citations(citCount).fnTail = dc.fnTail
        citCount = citCount + 1
NextFFC:
    Next m
End Sub

'==============================================================================
' CITATION DETECTION - EXISTING SHORT CITES
'==============================================================================
Private Sub FindExistingShortCites(pt As String, _
                                    qm() As Boolean, _
                                    scs() As ShortCiteInfo, _
                                    ByRef scC As Long)

    Dim reSupra As Object: Set reSupra = CreateObject("VBScript.RegExp")
    reSupra.Global = True: reSupra.Multiline = False
    reSupra.Pattern = BuildSupraPattern()

    Dim msS As Object: Set msS = reSupra.Execute(pt)
    Dim mS  As Object
    For Each mS In msS
        Dim ss As Long: ss = mS.FirstIndex + 1
        Dim sl As Long: sl = mS.length
        Dim sBN As String
        If mS.SubMatches.count >= 5 Then
            sBN = Trim("" & mS.SubMatches(4))
        Else
            sBN = ""
        End If

        Dim matchText As String: matchText = mS.Value
        Dim pinEnd As Long
        Dim supraCom As Long: supraCom = InStr(1, matchText, "supra,", vbTextCompare)
        Dim atPPos As Long
        If supraCom > 0 Then
            atPPos = InStr(supraCom, matchText, " at p", vbTextCompare)
        Else
            atPPos = InStr(1, matchText, " at p", vbTextCompare)
        End If
        If atPPos > 0 Then
            Dim pScan As Long: pScan = ss + atPPos - 1
            Do While pScan <= Len(pt)
                Dim pCh As String: pCh = Mid(pt, pScan, 1)
                If pCh >= "0" And pCh <= "9" Then Exit Do
                pScan = pScan + 1
            Loop
            Do While pScan <= Len(pt)
                Dim pCh2 As String: pCh2 = Mid(pt, pScan, 1)
                If (pCh2 >= "0" And pCh2 <= "9") Or pCh2 = "-" Or pCh2 = "," Or pCh2 = " " Then
                    pScan = pScan + 1
                Else
                    Exit Do
                End If
            Loop
            If pScan > ss + sl Then pScan = ss + sl
            pinEnd = pScan
        Else
            pinEnd = ss + sl
        End If

        Dim bnScan As Long: bnScan = pinEnd
        Do While bnScan <= Len(pt)
            If Mid(pt, bnScan, 1) <> " " Then Exit Do
            bnScan = bnScan + 1
        Loop

        If bnScan <= Len(pt) Then
            If Mid(pt, bnScan, 1) = "[" Then
                Dim trueEnd As Long: trueEnd = FindClosingBracket(pt, bnScan)
                If trueEnd > bnScan Then
                    sBN = Mid(pt, bnScan, trueEnd - bnScan + 1)
                    Dim nse As Long: nse = trueEnd + 1
                    If nse <= Len(pt) Then
                        If Mid(pt, nse, 1) = "." Then nse = nse + 1
                    End If
                    If nse <= Len(pt) Then
                        If Mid(pt, nse, 1) = ")" Then nse = nse + 1
                    End If
                    Dim newSl As Long: newSl = nse - ss
                    If newSl > sl Then sl = newSl
                End If
            ElseIf sBN <> "" Then
                Dim bnStart As Long: bnStart = 0
                Dim bnSrch  As Long: bnSrch = ss + sl - 1
                Do While bnSrch >= ss
                    If Mid(pt, bnSrch, 1) = "[" Then bnStart = bnSrch: Exit Do
                    bnSrch = bnSrch - 1
                Loop
                If bnStart > 0 Then
                    Dim trueEnd2 As Long: trueEnd2 = FindClosingBracket(pt, bnStart)
                    If trueEnd2 > bnStart Then
                        sBN = Mid(pt, bnStart, trueEnd2 - bnStart + 1)
                        Dim nse2 As Long: nse2 = trueEnd2 + 1
                        If nse2 <= Len(pt) Then
                            If Mid(pt, nse2, 1) = "." Then nse2 = nse2 + 1
                        End If
                        If nse2 <= Len(pt) Then
                            If Mid(pt, nse2, 1) = ")" Then nse2 = nse2 + 1
                        End If
                        Dim newSl2 As Long: newSl2 = nse2 - ss
                        If newSl2 > sl Then sl = newSl2
                    End If
                End If
            End If
        End If

        Dim bnEndPos As Long
        If sBN <> "" Then
            Dim bnSearchPos As Long: bnSearchPos = InStr(ss, pt, sBN, vbBinaryCompare)
            If bnSearchPos > 0 Then
                bnEndPos = bnSearchPos + Len(sBN)
                Do While bnEndPos <= Len(pt)
                    If Mid(pt, bnEndPos, 1) <> " " Then Exit Do
                    bnEndPos = bnEndPos + 1
                Loop
                If bnEndPos <= Len(pt) Then
                    If Mid(pt, bnEndPos, 1) = ";" Then
                        Dim bnOpenPos As Long
                        bnOpenPos = InStr(ss, pt, "[")
                        Dim cmpTrueEnd As Long
                        If bnOpenPos > 0 Then
                            cmpTrueEnd = FindClosingBracket(pt, bnOpenPos)
                            If cmpTrueEnd > bnOpenPos Then
                                sBN = Mid(pt, bnOpenPos, cmpTrueEnd - bnOpenPos + 1)
                                sl = cmpTrueEnd - ss + 1
                            Else
                                GoTo NextSupra
                            End If
                        Else
                            GoTo NextSupra
                        End If
                        GoTo AddCompoundSupra
                    End If
                End If
            End If
        Else
            bnEndPos = ss + sl
            Do While bnEndPos <= Len(pt)
                If Mid(pt, bnEndPos, 1) <> " " Then Exit Do
                bnEndPos = bnEndPos + 1
            Loop
            If bnEndPos <= Len(pt) Then
                If Mid(pt, bnEndPos, 1) = ";" Then GoTo NextSupra
            End If
        End If

        If scC > UBound(scs) - 1 Then ReDim Preserve scs(0 To UBound(scs) + 20)
        scs(scC).citeType = "supra"
        scs(scC).signal = Trim(mS.SubMatches(0))
        scs(scC).shortName = Trim(mS.SubMatches(1))
        scs(scC).reporter = Trim(mS.SubMatches(2))
        scs(scC).pincite = CleanPincite(Trim(mS.SubMatches(3)))
        scs(scC).bracketNote = sBN
        scs(scC).startChar = ss
        scs(scC).citLength = sl
        scs(scC).isMidSentence = DetectMidSentence(pt, ss)
        scs(scC).inQuote = IsInsideQuote(ss, sl, qm)
        scs(scC).isBare = False
        scs(scC).isCompound = False
        scC = scC + 1
        GoTo NextSupra
AddCompoundSupra:
        Dim csSS As Long: csSS = ss + 1
        Dim csSL As Long: csSL = sl - 1
        If scC > UBound(scs) - 1 Then ReDim Preserve scs(0 To UBound(scs) + 20)
        scs(scC).citeType = "supra"
        scs(scC).signal = Trim(mS.SubMatches(0))
        scs(scC).shortName = Trim(mS.SubMatches(1))
        scs(scC).reporter = Trim(mS.SubMatches(2))
        scs(scC).pincite = CleanPincite(Trim(mS.SubMatches(3)))
        scs(scC).bracketNote = sBN
        scs(scC).startChar = csSS
        scs(scC).citLength = csSL
        scs(scC).isMidSentence = True
        scs(scC).inQuote = IsInsideQuote(csSS, csSL, qm)
        scs(scC).isBare = True
        scs(scC).isCompound = True
        scC = scC + 1
NextSupra:
    Next mS

    '--- Bare supra ---
    Dim reBareSupra As Object: Set reBareSupra = CreateObject("VBScript.RegExp")
    reBareSupra.Global = True: reBareSupra.Multiline = False
    reBareSupra.Pattern = BuildBareSupraPattern()

    Dim msBS As Object: Set msBS = reBareSupra.Execute(pt)
    Dim mBS  As Object
    For Each mBS In msBS
        Dim bss As Long: bss = mBS.FirstIndex + 1
        Dim bsl As Long: bsl = mBS.length
        If bss > 1 Then
            If Mid(pt, bss - 1, 1) = "(" Then GoTo NextBS
            Dim bssBack As Long: bssBack = bss - 2
            Do While bssBack >= 1
                If Mid(pt, bssBack, 1) <> " " Then Exit Do
                bssBack = bssBack - 1
            Loop
            If bssBack >= 1 Then
                If Mid(pt, bssBack, 1) = "," Then GoTo NextBS
            End If
        End If
        If IsInsideQuote(bss, bsl, qm) Then GoTo NextBS
        If scC > UBound(scs) - 1 Then ReDim Preserve scs(0 To UBound(scs) + 20)
        scs(scC).citeType = "supra"
        scs(scC).signal = Trim(mBS.SubMatches(0))
        scs(scC).shortName = Trim(mBS.SubMatches(1))
        scs(scC).reporter = Trim(mBS.SubMatches(2))
        scs(scC).pincite = CleanPincite(Trim(mBS.SubMatches(3)))
        Dim bsBN As String
        If mBS.SubMatches.count >= 5 Then
            bsBN = Trim("" & mBS.SubMatches(4))
        Else
            bsBN = ""
        End If
        scs(scC).bracketNote = bsBN
        scs(scC).startChar = bss
        scs(scC).citLength = bsl
        scs(scC).isMidSentence = DetectMidSentence(pt, bss)
        scs(scC).inQuote = False
        scs(scC).isBare = True
        scC = scC + 1
NextBS:
    Next mBS

    '--- Ibid. ---
    Dim reIbid As Object: Set reIbid = CreateObject("VBScript.RegExp")
    reIbid.Global = True: reIbid.Multiline = False
    reIbid.Pattern = "\([Ii]bid\.(\s*\[[\s\S]*?\])?\)"

    Dim msI As Object: Set msI = reIbid.Execute(pt)
    Dim mI  As Object
    For Each mI In msI
        Dim is_ As Long: is_ = mI.FirstIndex + 1
        Dim il  As Long: il = mI.length
        If scC > UBound(scs) - 1 Then ReDim Preserve scs(0 To UBound(scs) + 20)
        scs(scC).citeType = "ibid"
        scs(scC).pincite = ""
        Dim ibBN As String
        If mI.SubMatches.count >= 1 Then
            ibBN = Trim("" & mI.SubMatches(0))
        Else
            ibBN = ""
        End If
        scs(scC).bracketNote = ibBN
        scs(scC).startChar = is_
        scs(scC).citLength = il
        scs(scC).isMidSentence = DetectMidSentence(pt, is_)
        scs(scC).inQuote = IsInsideQuote(is_, il, qm)
        scs(scC).isBare = False
        scC = scC + 1
    Next mI

    '--- Id. ---
    Dim reId As Object: Set reId = CreateObject("VBScript.RegExp")
    reId.Global = True: reId.Multiline = False
    reId.Pattern = "\([Ii]d\. at pp?\. (\d[\d\-,\s]*?)(\s*\[[\s\S]*?\])?(\.?)\)"

    Dim msD As Object: Set msD = reId.Execute(pt)
    Dim mD  As Object
    For Each mD In msD
        Dim ds As Long: ds = mD.FirstIndex + 1
        Dim dl As Long: dl = mD.length
        If scC > UBound(scs) - 1 Then ReDim Preserve scs(0 To UBound(scs) + 20)
        scs(scC).citeType = "id"
        scs(scC).pincite = CleanPincite(Trim(mD.SubMatches(0)))
        Dim idBN As String
        If mD.SubMatches.count >= 2 Then
            idBN = Trim("" & mD.SubMatches(1))
        Else
            idBN = ""
        End If
        scs(scC).bracketNote = idBN
        scs(scC).startChar = ds
        scs(scC).citLength = dl
        scs(scC).isMidSentence = DetectMidSentence(pt, ds)
        scs(scC).inQuote = IsInsideQuote(ds, dl, qm)
        scs(scC).isBare = False
        scC = scC + 1
    Next mD

    If scC > 1 Then SortShortCitesAsc scs, scC
End Sub

'==============================================================================
' TEXT BUILDERS
'==============================================================================

Private Function BuildFullCiteText(longInfo As DocCite, _
                                    usePincite As String, _
                                    useSignal As String, _
                                    addParen As Boolean) As String
    Dim sn As String
    If longInfo.shortNameOverride <> "" Then sn = longInfo.shortNameOverride Else sn = longInfo.shortName
    BuildFullCiteText = BuildFullCiteTextSN(longInfo, usePincite, useSignal, addParen, sn, "")
End Function

Private Function BuildFullCiteTextSN(longInfo As DocCite, _
                                      usePincite As String, _
                                      useSignal As String, _
                                      addParen As Boolean, _
                                      shortName As String, _
                                      Optional useFnTail As String = "") As String
    Dim s As String
    s = "("
    If useSignal <> "" Then s = s & useSignal & " "
    s = s & longInfo.caseName & " (" & longInfo.year & ") " & _
            longInfo.reporter & " " & longInfo.initialPage
    If usePincite <> "" Then s = s & ", " & usePincite
    If useFnTail <> "" Then s = s & ", " & useFnTail
    If addParen Then s = s & " (" & shortName & ")"
    BuildFullCiteTextSN = s & ".)"
End Function

Private Function BuildFullCiteText2(cit As CitInfo, shortName As String) As String
    Dim bn As String: bn = IIf(cit.bracketNote <> "", " " & cit.bracketNote, "")
    Dim fnSfx As String: fnSfx = IIf(cit.fnTail <> "", ", " & cit.fnTail, "")
    Dim s  As String
    If cit.isBare Then
        If cit.signal <> "" Then s = cit.signal & " " Else s = ""
        s = s & cit.caseName & " (" & cit.year & ") " & cit.reporter & " " & cit.initialPage
        If cit.pincite <> "" Then s = s & ", " & cit.pincite
        BuildFullCiteText2 = s & fnSfx & " (" & shortName & ") " & bn
    Else
        s = "("
        If cit.signal <> "" Then s = s & cit.signal & " "
        s = s & cit.caseName & " (" & cit.year & ") " & cit.reporter & " " & cit.initialPage
        If cit.pincite <> "" Then s = s & ", " & cit.pincite
        If cit.bracketNote <> "" Then
            BuildFullCiteText2 = s & " " & cit.bracketNote & fnSfx & " (" & shortName & ").)"
        Else
            BuildFullCiteText2 = s & fnSfx & " (" & shortName & ").)"
        End If
    End If
End Function

Private Function BuildIbidStr(isMidSentence As Boolean, Optional bracketNote As String = "") As String
    If bracketNote <> "" Then
        BuildIbidStr = "(Ibid. " & bracketNote & ".)"
    Else
        BuildIbidStr = "(Ibid.)"
    End If
End Function

Private Function BuildIdStr(pincite As String, isMidSentence As Boolean, Optional bracketNote As String = "") As String
    Dim pp As String: pp = PageOrPages(pincite)
    Dim core As String
    If pincite <> "" Then
        If bracketNote <> "" Then
            core = "Id. at " & pp & " " & pincite & " " & bracketNote & "."
        Else
            core = "Id. at " & pp & " " & pincite & "."
        End If
    Else
        If bracketNote <> "" Then
            core = "Id. " & bracketNote & "."
        Else
            core = "Id."
        End If
    End If
    BuildIdStr = "(" & core & ")"
End Function

Private Function BuildSupraStr(signal As String, _
                                shortName As String, _
                                reporter As String, _
                                pincite As String, _
                                isMidSentence As Boolean, _
                                Optional isBare As Boolean = False, _
                                Optional bracketNote As String = "") As String
    Dim addPeriod As Boolean: addPeriod = (Not isMidSentence)
    Dim bn As String: bn = IIf(bracketNote <> "", " " & bracketNote, "")
    Dim core As String
    If pincite <> "" Then
        Dim pp As String: pp = PageOrPages(pincite)
        If addPeriod Then core = shortName & ", supra, " & reporter & " at " & pp & " " & pincite & bn & "." _
                     Else core = shortName & ", supra, " & reporter & " at " & pp & " " & pincite & bn
    Else
        If addPeriod Then core = shortName & ", supra, " & reporter & bn & "." _
                     Else core = shortName & ", supra, " & reporter & bn
    End If
    If isBare Then
        If signal <> "" Then BuildSupraStr = signal & " " & core Else BuildSupraStr = core
    Else
        If signal <> "" Then BuildSupraStr = "(" & signal & " " & core & ")" _
                        Else BuildSupraStr = "(" & core & ")"
    End If
End Function

Private Function BuildIbidText(cit As CitInfo) As String
    BuildIbidText = BuildIbidStr(cit.isMidSentence)
End Function
Private Function BuildIdText(cit As CitInfo) As String
    BuildIdText = BuildIdStr(cit.pincite, cit.isMidSentence)
End Function
Private Function BuildSupraText(cit As CitInfo, shortName As String, reporter As String) As String
    BuildSupraText = BuildSupraStr(cit.signal, shortName, reporter, cit.pincite, cit.isMidSentence, cit.isBare)
End Function

'==============================================================================
' APPLYING REPLACEMENTS WITH FORMATTING
'==============================================================================
Private Sub ApplyReplacement(Doc As Document, paraStart As Long, rep As RepInfo)

    Dim citRng As Range
    Set citRng = Doc.Range(start:=paraStart + rep.citStartChar - 1, _
                           End:=paraStart + rep.citStartChar - 1 + rep.citLength)
    citRng.text = rep.newText

    Dim newStart As Long: newStart = paraStart + rep.citStartChar - 1

    TrimTrailingItalic Doc, newStart, Len(rep.newText)

    Dim flatRepRng As Range
    Set flatRepRng = Doc.Range(newStart, newStart + Len(rep.newText))
    flatRepRng.Font.Italic = False

    If Not rep.isSupra And Not rep.isFirstOccurrence Then
        Dim afterEnd As Long: afterEnd = newStart + Len(rep.newText)
        If afterEnd + 1 <= Doc.content.End Then
            Dim peekRng  As Range
            Set peekRng = Doc.Range(start:=afterEnd, End:=afterEnd + 1)
            Dim peekCh As String: peekCh = peekRng.text
            If peekCh <> " " And peekCh <> Chr(13) And peekCh <> Chr(7) And Len(peekCh) = 1 Then
                Dim spRng As Range: Set spRng = Doc.Range(start:=afterEnd, End:=afterEnd)
                spRng.text = " "
            End If
        End If
    End If

    If rep.italicOffset >= 0 And Len(rep.italicWord) > 0 Then
        Dim itRng As Range
        Set itRng = Doc.Range(start:=newStart + rep.italicOffset, _
                              End:=newStart + rep.italicOffset + Len(rep.italicWord))
        itRng.Font.Italic = True
    End If

    If rep.caseNameLen > 0 And (rep.isSupra Or rep.isFirstOccurrence) Then
        Dim parenSkip As Long: parenSkip = IIf(rep.isBare, 0, 1)
        Dim nameAbs As Long: nameAbs = newStart + parenSkip + rep.signalLen
        Dim nameRng As Range
        Set nameRng = Doc.Range(start:=nameAbs, End:=nameAbs + rep.caseNameLen)
        nameRng.Font.Italic = True
    End If

    If rep.isFirstOccurrence And rep.parenNameLen > 0 And rep.parenNameOffset > 0 Then
        Dim pnAbs As Long: pnAbs = newStart + rep.parenNameOffset
        Dim pnRng As Range
        Set pnRng = Doc.Range(start:=pnAbs, End:=pnAbs + rep.parenNameLen)
        pnRng.Font.Italic = True
    End If
End Sub

Private Sub TrimTrailingItalic(Doc As Document, pasteStart As Long, pasteLen As Long)
    Dim pos As Long: pos = pasteStart + pasteLen
    Dim deleted As Long: deleted = 0
    Do
        If pos >= Doc.content.End Then Exit Do
        Dim chk As Range
        Set chk = Doc.Range(pos, pos + 1)
        If chk.Font.Italic = False Then Exit Do
        If deleted >= 10 Then Exit Do
        chk.Delete
        deleted = deleted + 1
    Loop

    Dim pasteEnd As Long: pasteEnd = pasteStart + pasteLen
    If pasteEnd >= 3 And pasteEnd + 1 < Doc.content.End Then
        Dim tailRng As Range
        Set tailRng = Doc.Range(pasteEnd - 2, pasteEnd + 2)
        If tailRng.text = ".)" & ".)" Then
            Dim dupRng As Range
            Set dupRng = Doc.Range(pasteEnd, pasteEnd + 2)
            dupRng.Delete
        End If
    End If
End Sub

Private Sub ClearHighlightIfYellow(PARA As Paragraph, startChar As Long, length As Long)
    Dim rng As Range
    Set rng = PARA.Range.Document.Range( _
        start:=PARA.Range.start + startChar - 1, _
        End:=PARA.Range.start + startChar - 1 + length)
    If rng.HighlightColorIndex = wdYellow Then
        rng.HighlightColorIndex = wdNoHighlight
    End If
End Sub

Private Sub ApplyHighlight(PARA As Paragraph, startChar As Long, length As Long)
    Dim hlRng As Range
    Set hlRng = PARA.Range.Document.Range( _
        start:=PARA.Range.start + startChar - 1, _
        End:=PARA.Range.start + startChar - 1 + length)
    hlRng.HighlightColorIndex = wdYellow
End Sub

'==============================================================================
' NON-CASE DOCUMENT DETECTION
'------------------------------------------------------------------------------
' Returns True iff the parenthetical immediately preceding `posIbid` looks like
' a reference to a non-case document (Declaration, Pleading, Brief, RJN, etc.)
' rather than a case citation. Used to prevent (Ibid.) / (Id. at ...) cites that
' refer to non-case documents from being "fixed" to parenthetical case cites.
'
' First tries to find the most recent preceding parenthetical within the current
' paragraph text `pt`. If none exists in the current paragraph, falls back to
' scanning previous paragraphs via the Document (so cross-paragraph (Ibid.)
' references still get classified correctly).
'
' Algorithm (per parenthetical scanned, walking backward):
'   1. Locate the most recent ')' and walk back to its matching '('.
'   2. Extract the content between '(' and ')'.
'   3. If the content has a case reporter pattern (Cal.4th, F.3d, etc.) OR a
'      supra token, treat as a case cite -> return False.
'   4. If the content is itself an Ibid./Id., keep walking back.
'   5. If the content matches a non-case document marker, return True.
'   6. Otherwise (ambiguous), keep walking back to the next parenthetical.
'   7. If we run out without classifying, return False (conservative).
'------------------------------------------------------------------------------
Private Function RefersToNonCaseDocument(pt As String, posIbid As Long, _
                                          Optional PARA As Paragraph) As Boolean
    RefersToNonCaseDocument = False

    ' First scan within the current paragraph
    Dim result As Integer  ' 0 = inconclusive, 1 = case, 2 = non-case
    result = ScanForPrecedingCiteType(pt, posIbid)
    If result = 2 Then RefersToNonCaseDocument = True: Exit Function
    If result = 1 Then Exit Function   ' case cite found -> definitely not non-case

    ' Inconclusive in current paragraph -- fall back to previous paragraphs.
    ' Limit the lookback window to a reasonable size to avoid scanning
    ' arbitrarily large prior text on big documents.
    If PARA Is Nothing Then Exit Function
    Dim Doc As Document
    On Error Resume Next
    Set Doc = PARA.Range.Document
    On Error GoTo 0
    If Doc Is Nothing Then Exit Function

    Dim paraStart As Long: paraStart = PARA.Range.start
    If paraStart <= 0 Then Exit Function

    Const LOOKBACK_LIMIT As Long = 4000
    Dim lookbackStart As Long: lookbackStart = paraStart - LOOKBACK_LIMIT
    If lookbackStart < 0 Then lookbackStart = 0

    ' Pull the recent text of the document up to (but not including) the
    ' current paragraph, capped at LOOKBACK_LIMIT characters.
    Dim prevText As String
    On Error Resume Next
    prevText = Doc.Range(lookbackStart, paraStart).text
    On Error GoTo 0
    If Len(prevText) = 0 Then Exit Function
    prevText = NormalizeSpaces(prevText)

    result = ScanForPrecedingCiteType(prevText, Len(prevText) + 1)
    If result = 2 Then RefersToNonCaseDocument = True
End Function

'------------------------------------------------------------------------------
' Walks backward in `text` from `endPos` looking for parenthetical citations.
' Returns:
'   0 = inconclusive (no recognizable cite found, or ran off the start)
'   1 = preceding cite is a case citation (has reporter or supra)
'   2 = preceding cite is a non-case document (Decl., Mot., Br., etc.)
'------------------------------------------------------------------------------
Private Function ScanForPrecedingCiteType(text As String, endPos As Long) As Integer
    ScanForPrecedingCiteType = 0
    If endPos < 2 Then Exit Function

    Dim scanStart As Long: scanStart = endPos - 1
    If scanStart > Len(text) Then scanStart = Len(text)
    Dim safety As Long: safety = 0

    Do While scanStart >= 1 And safety < 50
        safety = safety + 1

        ' Skip whitespace
        Do While scanStart >= 1
            Dim sC As String: sC = Mid(text, scanStart, 1)
            If sC <> " " And sC <> Chr(13) And sC <> Chr(9) And sC <> Chr(10) Then Exit Do
            scanStart = scanStart - 1
        Loop
        If scanStart < 1 Then Exit Function

        ' Find the most recent ')' at or before scanStart
        Dim closePos As Long: closePos = 0
        Dim sp As Long: sp = scanStart
        Do While sp >= 1
            If Mid(text, sp, 1) = ")" Then closePos = sp: Exit Do
            sp = sp - 1
        Loop
        If closePos = 0 Then Exit Function

        ' Walk back to matching '('
        Dim depth As Long: depth = 1
        Dim openPos As Long: openPos = 0
        Dim p As Long: p = closePos - 1
        Do While p >= 1
            Dim c As String: c = Mid(text, p, 1)
            If c = ")" Then
                depth = depth + 1
            ElseIf c = "(" Then
                depth = depth - 1
                If depth = 0 Then
                    openPos = p
                    Exit Do
                End If
            End If
            p = p - 1
        Loop
        If openPos = 0 Then Exit Function

        Dim content As String: content = Mid(text, openPos + 1, closePos - openPos - 1)
        Dim cT As String: cT = Trim(content)
        Dim ctL As String: ctL = LCase(cT)

        ' If this is itself an Ibid./Id., walk past it
        If Left(ctL, 5) = "ibid." Or ctL = "ibid" Or _
           Left(ctL, 3) = "id." Or Left(ctL, 4) = "id. " Or Left(ctL, 4) = "id, " Then
            scanStart = openPos - 1
            GoTo NextParen
        End If

        ' If contains "supra", it's a case short cite
        If InStr(1, content, "supra", vbTextCompare) > 0 Then
            ScanForPrecedingCiteType = 1
            Exit Function
        End If

        ' If contains a case reporter pattern, it's a case cite
        Dim reRep As Object: Set reRep = CreateObject("VBScript.RegExp")
        reRep.Pattern = ReporterPattern()
        reRep.Global = False
        reRep.IgnoreCase = False
        If reRep.Test(content) Then
            ScanForPrecedingCiteType = 1
            Exit Function
        End If

        ' If matches a non-case document marker, it's non-case
        If HasNonCaseDocumentMarker(content) Then
            ScanForPrecedingCiteType = 2
            Exit Function
        End If

        ' Ambiguous -- keep walking back
        scanStart = openPos - 1
NextParen:
    Loop
End Function

'------------------------------------------------------------------------------
' Returns True if the given parenthetical content contains a recognizable
' non-case-document marker: declaration, affidavit, motion, opposition, reply,
' brief, complaint, answer, RJN, petition, response, memorandum, order,
' transcript, deposition, exhibit, etc. Also recognizes paragraph references
' (the paragraph symbol or "para.") in the absence of a reporter, which strongly
' indicates a pleading/declaration pin-cite.
'------------------------------------------------------------------------------
Private Function HasNonCaseDocumentMarker(content As String) As Boolean
    HasNonCaseDocumentMarker = False
    If Len(content) = 0 Then Exit Function

    ' Paragraph symbol (¶ = ChrW(182)) without a reporter strongly indicates
    ' a pleading/declaration pin-cite.
    If InStr(1, content, ChrW(182), vbBinaryCompare) > 0 Then
        HasNonCaseDocumentMarker = True
        Exit Function
    End If

    ' Section symbol (§ = ChrW(167)) without a reporter indicates a statute
    ' or code section reference -- not a case citation.
    If InStr(1, content, ChrW(167), vbBinaryCompare) > 0 Then
        HasNonCaseDocumentMarker = True
        Exit Function
    End If

    ' Pattern 1: full-word markers (require word boundaries on both sides)
    Dim reW As Object: Set reW = CreateObject("VBScript.RegExp")
    reW.IgnoreCase = True
    reW.Global = False
    reW.Pattern = "\b(declaration|affidavit|motion|opposition|" & _
                  "reply|complaint|answer|brief|petition|response|" & _
                  "memorandum|order|judgment|judgement|" & _
                  "transcript|hearing|deposition|exhibit|" & _
                  "stipulation|interrogatory|interrogatories|" & _
                  "subpoena|ruling|" & _
                  "plaintiffs?|defendants?|respondents?|appellants?|" & _
                  "petitioners?|movants?|" & _
                  "appendix|minute order)\b"
    If reW.Test(content) Then
        HasNonCaseDocumentMarker = True
        Exit Function
    End If

    ' Pattern 2: abbreviations ending in period (require leading word boundary
    ' only; trailing context is the literal period itself).
    Dim reA As Object: Set reA = CreateObject("VBScript.RegExp")
    reA.IgnoreCase = True
    reA.Global = False
    reA.Pattern = "\b(decl\.|aff\.|mot\.|opp\.|oppo\.|opp'?n|" & _
                  "compl\.|ans\.|br\.|" & _
                  "pet\.|resp\.|memo\.|" & _
                  "tr\.|rt\.|hrg\.|dep\.|depo\.|" & _
                  "exh?\.|stip\.|rog\.|" & _
                  "para\.|paras\.|" & _
                  "app\.|appx\.|" & _
                  "pls?\.|defs?\.|" & _
                  "m\.o\.|" & _
                  "repl\.)"
    If reA.Test(content) Then
        HasNonCaseDocumentMarker = True
        Exit Function
    End If

    ' Pattern 3: distinctive acronyms (require both word boundaries, no period)
    Dim reAcr As Object: Set reAcr = CreateObject("VBScript.RegExp")
    reAcr.IgnoreCase = False  ' acronyms are case-sensitive
    reAcr.Global = False
    reAcr.Pattern = "\b(RJN|FAC|SAC|TAC|MIL)\b"
    If reAcr.Test(content) Then
        HasNonCaseDocumentMarker = True
        Exit Function
    End If

    ' "Request for judicial notice" (case-insensitive substring, distinctive)
    If InStr(1, content, "request for judicial notice", vbTextCompare) > 0 Then
        HasNonCaseDocumentMarker = True
        Exit Function
    End If
End Function

'==============================================================================
' HELPER FUNCTIONS
'==============================================================================

Private Function DetectMidSentence(pt As String, startChar As Long) As Boolean
    Dim scanPos As Long: scanPos = startChar - 1
    Do While scanPos >= 1
        Dim c As String: c = Mid(pt, scanPos, 1)
        If c = " " Or c = "(" Or c = ")" Or _
           c = Chr(34) Or c = ChrW(8220) Or c = ChrW(8221) Or _
           c = ChrW(8216) Or c = ChrW(8217) Then
            scanPos = scanPos - 1
        Else
            Exit Do
        End If
    Loop
    If scanPos < 1 Then
        DetectMidSentence = False
    Else
        Dim pc As String: pc = Mid(pt, scanPos, 1)
        DetectMidSentence = Not (pc = "." Or pc = "!" Or pc = "?")
    End If
End Function

Private Function DetectSignal(pt As String, pos As Long) As String
    Dim sigs(6) As String
    sigs(0) = "See generally "
    sigs(1) = "See also "
    sigs(2) = "But see "
    sigs(3) = "See "
    sigs(4) = "Cf. "
    sigs(5) = "Accord "
    sigs(6) = "Contra "
    Dim i As Long
    For i = 0 To 6
        Dim sl As Long: sl = Len(sigs(i))
        If pos + sl - 1 <= Len(pt) Then
            If Mid(pt, pos, sl) = sigs(i) Then
                DetectSignal = Trim(sigs(i))
                Exit Function
            End If
        End If
    Next i
    DetectSignal = ""
End Function

Private Function DetectSignalBefore(pt As String, pos As Long) As String
    Dim sigs(6) As String
    sigs(0) = "See generally "
    sigs(1) = "See also "
    sigs(2) = "But see "
    sigs(3) = "See "
    sigs(4) = "Cf. "
    sigs(5) = "Accord "
    sigs(6) = "Contra "
    Dim i As Long
    For i = 0 To 6
        Dim sl As Long: sl = Len(sigs(i))
        If pos - sl >= 1 Then
            If Mid(pt, pos - sl, sl) = sigs(i) Then
                DetectSignalBefore = Trim(sigs(i))
                Exit Function
            End If
        End If
    Next i
    DetectSignalBefore = ""
End Function

Private Function IsGovernmentalParty(party As String) As Boolean
    Dim p As String: p = Trim(party)
    Select Case LCase(p)
        Case "people", "united states", "state"
            IsGovernmentalParty = True: Exit Function
    End Select
    If LCase(Left(p, 8)) = "city of " Then IsGovernmentalParty = True: Exit Function
    If LCase(Left(p, 10)) = "county of " Then IsGovernmentalParty = True: Exit Function
    IsGovernmentalParty = False
End Function

Private Function ExtractShortName(caseName As String) As String
    Dim vPos As Long: vPos = InStr(1, caseName, " v. ", vbTextCompare)
    If vPos > 1 Then
        Dim plaintiff As String: plaintiff = Trim(Left(caseName, vPos - 1))
        Dim defendant As String: defendant = Trim(Mid(caseName, vPos + 4))
        If IsGovernmentalParty(plaintiff) Then
            ExtractShortName = defendant
        Else
            ExtractShortName = plaintiff
        End If
    Else
        ExtractShortName = Trim(caseName)
    End If
End Function

Private Function CleanCaseName(s As String) As String
    Dim r As String: r = Trim(s)
    Do While Len(r) > 0
        Select Case Right(r, 1)
            Case ".", ",", " ", "*": r = Left(r, Len(r) - 1)
            Case Else: Exit Do
        End Select
    Loop
    CleanCaseName = Trim(r)
End Function

'------------------------------------------------------------------------------
' Match a ", fn. omitted" / ", fns. omitted" / ", fn. N" / ", fns. N, M" tail
' starting at or just after position `startPos` in `pt`. Returns "" if not found.
'
' On success:
'   - fnTail       = the matched text WITHOUT the leading ", " (e.g. "fn. omitted")
'   - consumedLen  = total chars consumed in pt (including leading ", " if present)
'
' Handles two cases:
'   (a) startPos points directly at "fn" or "fns"   the comma+space was already
'       absorbed by the regex pincite group.
'   (b) startPos points at ","   there was no pinpoint so the regex didn't eat
'       the ", " before "fn".
'------------------------------------------------------------------------------
Private Function MatchFnTail(pt As String, ByVal startPos As Long, _
                              ByRef fnTail As String, _
                              ByRef consumedLen As Long) As Boolean
    MatchFnTail = False
    fnTail = ""
    consumedLen = 0
    If startPos < 1 Or startPos > Len(pt) Then Exit Function

    Dim p As Long: p = startPos
    Dim leadLen As Long: leadLen = 0

    ' If we're sitting on a comma, skip past ", " (case b)
    If Mid(pt, p, 1) = "," Then
        p = p + 1
        Do While p <= Len(pt)
            If Mid(pt, p, 1) <> " " Then Exit Do
            p = p + 1
        Loop
        leadLen = p - startPos
    End If

    If p > Len(pt) - 3 Then Exit Function

    ' Probe for "fn." or "fns."
    Dim probe As String: probe = LCase(Mid(pt, p, 4))
    If Left(probe, 3) <> "fn." And probe <> "fns." Then Exit Function

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = "^fns?\.\s+(?:omitted|\d+(?:[\-,\s]+\d+)*)"
    Dim mm As Object: Set mm = re.Execute(Mid(pt, p))
    If mm.count = 0 Then Exit Function

    fnTail = mm(0).Value
    consumedLen = leadLen + Len(fnTail)
    MatchFnTail = True
End Function

Private Function CleanPincite(s As String) As String
    Dim r As String: r = Trim(s)
    If Left(r, 1) = "," Then r = Trim(Mid(r, 2))
    Do While Len(r) > 0
        Select Case Right(r, 1)
            Case ".", ",", " ", ")", "]": r = Left(r, Len(r) - 1)
            Case Else: Exit Do
        End Select
    Loop
    CleanPincite = Trim(r)
End Function

Private Function cleanItalicText(s As String) As String
    Dim r As String: r = Trim(s)
    Do While Len(r) > 0
        Select Case Right(r, 1)
            Case ".", ",", ";", ":", " ": r = Left(r, Len(r) - 1)
            Case Else: Exit Do
        End Select
    Loop
    cleanItalicText = Trim(r)
End Function

Private Function NormalizeSpaces(s As String) As String
    Dim r As String: r = s
    r = Replace(r, Chr(160), Chr(32))
    r = Replace(r, ChrW(8239), Chr(32))
    r = Replace(r, ChrW(8201), Chr(32))
    r = Replace(r, ChrW(8202), Chr(32))
    NormalizeSpaces = r
End Function

Private Sub BuildQuoteMask(text As String, mask() As Boolean)
    Dim i As Long
    Dim inQ      As Boolean: inQ = False
    Dim inBrack  As Long: inBrack = 0
    For i = 1 To Len(text)
        Dim c As String: c = Mid(text, i, 1)
        If c = "[" Then
            inBrack = inBrack + 1
        ElseIf c = "]" And inBrack > 0 Then
            inBrack = inBrack - 1
        End If
        If Not inQ Then
            If inBrack = 0 And (c = Chr(34) Or c = ChrW(8220)) Then
                inQ = True: mask(i) = True
            Else
                mask(i) = False
            End If
        Else
            mask(i) = True
            If inBrack = 0 And (c = Chr(34) Or c = ChrW(8221)) Then inQ = False
        End If
    Next i
End Sub

Private Function IsInsideQuote(startPos As Long, length As Long, mask() As Boolean) As Boolean
    Dim i As Long: Dim mx As Long: mx = UBound(mask)
    For i = startPos To startPos + length - 1
        If i >= 1 And i <= mx Then
            If mask(i) Then IsInsideQuote = True: Exit Function
        End If
    Next i
    IsInsideQuote = False
End Function

Private Function PageOrPages(pincite As String) As String
    If InStr(pincite, "-") > 0 Or InStr(pincite, ",") > 0 Then
        PageOrPages = "pp."
    Else
        PageOrPages = "p."
    End If
End Function

Private Sub SortRepsDescending(reps() As RepInfo, count As Long)
    Dim i As Long, j As Long, tmp As RepInfo
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If reps(j).citStartChar > reps(i).citStartChar Then
                tmp = reps(i): reps(i) = reps(j): reps(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub SortDocCitesByAbsStart(arr() As DocCite, count As Long)
    Dim i As Long, j As Long, tmp As DocCite
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If arr(j).absStart < arr(i).absStart Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub SortShortCitesAsc(arr() As ShortCiteInfo, count As Long)
    Dim i As Long, j As Long, tmp As ShortCiteInfo
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If arr(j).startChar < arr(i).startChar Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

'==============================================================================
' FEATURE-B HELPERS
'==============================================================================

Private Function MatchHintLine(capturedName As String, capturedRep As String, _
                                hintLines() As HintLine, hlC As Long) As Long
    MatchHintLine = -1
    Dim i As Long
    Dim cRep As String: cRep = LCase(Trim(capturedRep))
    Dim cNm  As String: cNm = LCase(Trim(capturedName))
    For i = 0 To hlC - 1
        If LCase(Trim(hintLines(i).reporter)) = cRep Then
            MatchHintLine = i: Exit Function
        End If
    Next i
    For i = 0 To hlC - 1
        If LCase(Trim(hintLines(i).caseName)) = cNm Or _
           LCase(Trim(hintLines(i).shortName)) = cNm Then
            MatchHintLine = i: Exit Function
        End If
    Next i
End Function

Private Function MatchHintLineByName(cleanText As String, _
                                      hintLines() As HintLine, hlC As Long) As Long
    MatchHintLineByName = -1
    If Len(cleanText) < 2 Then Exit Function
    Dim cT As String: cT = LCase(cleanText)
    Dim i As Long
    For i = 0 To hlC - 1
        If LCase(Trim(hintLines(i).caseName)) = cT Or _
           LCase(Trim(hintLines(i).shortName)) = cT Then
            MatchHintLineByName = i: Exit Function
        End If
    Next i
End Function

Private Function HasPriorLongCiteInDocCites(normKey As String, absPos As Long, _
                                             docCites() As DocCite, dcc As Long) As Boolean
    HasPriorLongCiteInDocCites = False
    Dim i As Long
    Dim nk As String: nk = LCase(Trim(normKey))
    For i = 0 To dcc - 1
        If docCites(i).citeType = "long" And _
           LCase(Trim(docCites(i).normKey)) = nk And _
           docCites(i).absStart < absPos Then
            HasPriorLongCiteInDocCites = True
            Exit Function
        End If
    Next i
End Function

Private Sub RegisterInB15Dict(hl As HintLine, b15Dict As Object, snToKey As Object)
    Dim bKey As String: bKey = hl.normKey
    If bKey = "" Then Exit Sub
    If Not b15Dict.Exists(bKey) Then
        b15Dict.Add bKey, hl.shortName & "|" & hl.reporter & "|" & _
                          hl.initialPage & "|" & hl.year & "|" & hl.caseName
    End If
    If hl.shortName <> "" Then
        If Not snToKey.Exists(LCase(hl.shortName)) Then snToKey.Add LCase(hl.shortName), bKey
    End If
End Sub

Private Function BuildB1Transformed(sig As String, capturedName As String, _
                                     capturedRep As String, capturedPin As String, _
                                     trailDot As String, hl As HintLine, _
                                     ByRef italicOff As Long, _
                                     ByRef italicLen As Long, _
                                     ByRef italicOff2 As Long, _
                                     ByRef italicLen2 As Long) As String
    italicOff2 = 0: italicLen2 = 0

    Dim s As String: s = "("
    Dim prefixLen As Long: prefixLen = 1 + IIf(sig <> "", Len(sig) + 1, 0)
    Dim nameStartInS As Long

    If sig <> "" Then s = s & sig & " "

    If LCase(Trim(capturedName)) = LCase(Trim(hl.caseName)) Then
        nameStartInS = Len(s)
        italicLen = Len(capturedName)
        s = s & capturedName
    Else
        Dim restored As String, nameInResOff As Long
        BuildRestoredNameB1 capturedName, hl.caseName, restored, nameInResOff
        nameStartInS = Len(s) + nameInResOff
        italicLen = Len(capturedName)
        s = s & restored

        Dim vPos As Long: vPos = InStr(1, hl.caseName, " v. ", vbTextCompare)
        If vPos > 1 Then
            Dim bPla As String: bPla = Trim(Left(hl.caseName, vPos - 1))
            Dim bDef As String: bDef = Trim(Mid(hl.caseName, vPos + 4))
            If LCase(Trim(capturedName)) = LCase(bDef) Then
                italicOff2 = prefixLen + 1
                italicLen2 = Len(bPla) + 3
            Else
                italicOff2 = prefixLen + Len(capturedName) + Len(" [v. ")
                italicLen2 = Len(bDef)
            End If
        End If
    End If

    italicOff = nameStartInS

    s = s & " [(" & hl.year & ")] " & capturedRep & _
            " [" & hl.initialPage & ",] " & capturedPin & trailDot & ")"

    BuildB1Transformed = s
End Function

Private Sub BuildRestoredNameB1(shortName As String, caseName As String, _
                                  ByRef result As String, _
                                  ByRef shortNameOffset As Long)
    Dim vPos As Long: vPos = InStr(1, caseName, " v. ", vbTextCompare)
    If vPos < 1 Then
        result = shortName: shortNameOffset = 0: Exit Sub
    End If
    Dim plaintiff As String: plaintiff = Trim(Left(caseName, vPos - 1))
    Dim defendant As String: defendant = Trim(Mid(caseName, vPos + 4))

    If LCase(Trim(shortName)) = LCase(defendant) Then
        Dim pfx As String: pfx = "[" & plaintiff & " v.] "
        result = pfx & shortName
        shortNameOffset = Len(pfx)
    ElseIf LCase(Trim(shortName)) = LCase(plaintiff) Then
        result = shortName & " [v. " & defendant & "]"
        shortNameOffset = 0
    Else
        result = shortName: shortNameOffset = 0
    End If
End Sub

Private Function BuildB2Replacement(cleanItalicText As String, hl As HintLine, _
                                     ByRef italicOff As Long, _
                                     ByRef italicLen As Long, _
                                     ByRef italicOff2 As Long, _
                                     ByRef italicLen2 As Long) As String
    italicOff2 = 0: italicLen2 = 0

    Dim vPos As Long: vPos = InStr(1, hl.caseName, " v. ", vbTextCompare)

    If LCase(Trim(cleanItalicText)) = LCase(Trim(hl.caseName)) Then
        italicOff = 0
        italicLen = Len(cleanItalicText)
        BuildB2Replacement = cleanItalicText & _
            " [(" & hl.year & ") " & hl.reporter & " " & hl.initialPage & "]"

    ElseIf LCase(Trim(cleanItalicText)) = LCase(Trim(hl.shortName)) And vPos > 0 Then
        Dim plaintiff As String: plaintiff = Trim(Left(hl.caseName, vPos - 1))
        Dim defendant As String: defendant = Trim(Mid(hl.caseName, vPos + 4))

        If LCase(Trim(cleanItalicText)) = LCase(defendant) Then
            Dim pfx2 As String: pfx2 = "[" & plaintiff & " v.] "
            italicOff = Len(pfx2)
            italicLen = Len(cleanItalicText)
            italicOff2 = 1
            italicLen2 = Len(plaintiff) + 3
            BuildB2Replacement = pfx2 & cleanItalicText & _
                " [(" & hl.year & ") " & hl.reporter & " " & hl.initialPage & "]"
        Else
            italicOff = 0
            italicLen = Len(cleanItalicText)
            italicOff2 = Len(cleanItalicText) + Len(" [v. ")
            italicLen2 = Len(defendant)
            BuildB2Replacement = cleanItalicText & _
                " [v. " & defendant & " (" & hl.year & ") " & hl.reporter & " " & hl.initialPage & "]"
        End If
    Else
        italicOff = 0
        italicLen = Len(cleanItalicText)
        BuildB2Replacement = cleanItalicText
    End If
End Function

Private Function TryResolveByParty(lcShortName As String, lcReporter As String, _
                                    altPartyToKey As Object, _
                                    preScanInfo As Object) As String
    TryResolveByParty = ""
    If Not altPartyToKey.Exists(lcShortName) Then Exit Function
    Dim nk As String: nk = altPartyToKey(lcShortName)
    If Not preScanInfo.Exists(nk) Then Exit Function
    Dim parts() As String: parts = Split(preScanInfo(nk), "|")
    If UBound(parts) < 2 Then Exit Function
    If LCase(Trim(parts(2))) = LCase(Trim(lcReporter)) Then TryResolveByParty = nk
End Function

Private Sub FindAllMatchingNormKeys(cleanText As String, _
                                     preScanInfo As Object, _
                                     snToKey As Object, _
                                     altPartyToKey As Object, _
                                     matchKeys() As String, _
                                     ByRef matchCount As Long)
    matchCount = 0
    Dim cT As String: cT = LCase(Trim(cleanText))
    If Len(cT) < 2 Then Exit Sub

    Dim nk As Variant
    For Each nk In preScanInfo.Keys
        Dim parts() As String: parts = Split(preScanInfo(nk), "|")
        If UBound(parts) < 5 Then GoTo FAMNKNext
        Dim pCN  As String: pCN = LCase(Trim(parts(0)))
        Dim pSN  As String: pSN = LCase(Trim(parts(4)))
        Dim pSnO As String: pSnO = LCase(Trim(parts(5)))

        Dim matched As Boolean: matched = False
        If cT = pCN Then matched = True
        If Not matched And cT = pSN Then matched = True
        If Not matched And pSnO <> "" And cT = pSnO Then matched = True
        If Not matched Then
            Dim vp As Long: vp = InStr(1, parts(0), " v. ", vbTextCompare)
            If vp > 1 Then
                If cT = LCase(Trim(Left(parts(0), vp - 1))) Then matched = True
                If Not matched Then
                    If cT = LCase(Trim(Mid(parts(0), vp + 4))) Then matched = True
                End If
            End If
        End If

        If matched Then
            Dim isDupNK As Boolean: isDupNK = False
            Dim ki As Long
            For ki = 0 To matchCount - 1
                If LCase(matchKeys(ki)) = LCase(CStr(nk)) Then isDupNK = True: Exit For
            Next ki
            If Not isDupNK Then
                If matchCount <= UBound(matchKeys) Then
                    matchKeys(matchCount) = CStr(nk)
                    matchCount = matchCount + 1
                End If
            End If
        End If
FAMNKNext:
    Next nk
End Sub

Private Function HintLineFromNormKey(normKey As String, preScanInfo As Object) As HintLine
    Dim hl As HintLine
    If Not preScanInfo.Exists(normKey) Then HintLineFromNormKey = hl: Exit Function
    Dim parts() As String: parts = Split(preScanInfo(normKey), "|")
    If UBound(parts) < 5 Then HintLineFromNormKey = hl: Exit Function
    hl.normKey = normKey
    hl.caseName = parts(0)
    hl.year = parts(1)
    hl.reporter = parts(2)
    hl.initialPage = parts(3)
    hl.shortName = IIf(Trim(parts(5)) <> "", Trim(parts(5)), Trim(parts(4)))
    hl.bmName = ""
    HintLineFromNormKey = hl
End Function

Private Function FindHintLineByNormKey(normKey As String, _
                                        hintLines() As HintLine, hlC As Long) As Long
    FindHintLineByNormKey = -1
    Dim i As Long
    For i = 0 To hlC - 1
        If LCase(Trim(hintLines(i).normKey)) = LCase(Trim(normKey)) Then
            FindHintLineByNormKey = i: Exit Function
        End If
    Next i
End Function

Private Function DisambiguateByReporter(matchKeys() As String, matchCount As Long, _
                                         paraText As String, qm() As Boolean, _
                                         preScanInfo As Object) As String
    DisambiguateByReporter = ""
    Dim foundKey   As String: foundKey = ""
    Dim foundCount As Long: foundCount = 0
    Dim i As Long
    For i = 0 To matchCount - 1
        Dim nk As String: nk = matchKeys(i)
        If preScanInfo.Exists(nk) Then
            Dim parts() As String: parts = Split(preScanInfo(nk), "|")
            If UBound(parts) >= 2 Then
                Dim rep As String: rep = Trim(parts(2))
                If Len(rep) > 0 Then
                    Dim pos As Long: pos = 1
                    Do
                        pos = InStr(pos, paraText, rep, vbTextCompare)
                        If pos = 0 Then Exit Do
                        If IsInsideQuote(pos, Len(rep), qm) Then
                            foundKey = nk
                            foundCount = foundCount + 1
                            Exit Do
                        End If
                        pos = pos + 1
                    Loop
                End If
            End If
        End If
    Next i
    If foundCount = 1 Then DisambiguateByReporter = foundKey
End Function

'==============================================================================
' REGEX PATTERN BUILDERS
'==============================================================================

Private Function ReporterPattern() As String
    ReporterPattern = "(?:Cal\.App\.[2-5]th|Cal\.App\.[23]d|Cal\.App\.|" & _
                      "Cal\.[2-5]th|Cal\.[23]d|Cal\.|" & _
                      "Cal\.Rptr\.[23]d|Cal\.Rptr\.|" & _
                      "U\.S\.|F\.[234]th|F\.[234]d|" & _
                      "F\.Supp\.[23]d|F\.Supp\.|" & _
                      "P\.[23]d|A\.[23]d|S\.W\.[23]d|N\.E\.[23]d)"
End Function

Private Function BuildLongCitePattern() As String
    BuildLongCitePattern = _
        "((?:[A-Z][^(]*?(?:\s+v\.\s+|\s+In re\s+|\s+))[^(]+?)" & _
        "\s*\((\d{4})\)\s+" & _
        "(\d+)\s+" & _
        "(" & ReporterPattern() & ")\s*" & _
        "(\d+)" & _
        "((?:,\s*\d[\d\s,\-]*)?)?" & _
        "(\))?"
End Function

Private Function BuildSupraPattern() As String
    BuildSupraPattern = _
        "\((?:(See generally |See also |But see |See |Cf\. |Accord |Contra ))?" & _
        "([A-Z][^)]*?),\s+supra,\s+" & _
        "(\d+\s+" & ReporterPattern() & ")" & _
        "\s+at\s+pp?\.\s+" & _
        "(\d[\d\-,\s]*?)" & _
        "(\s*\[[\s\S]*?\])?" & _
        "(\.?)\)"
End Function

Private Function BuildBareSupraPattern() As String
    BuildBareSupraPattern = _
        "((?:See generally |See also |But see |See |Cf\. |Accord |Contra ))?" & _
        "([A-Z][^\n,;()\[\]]*?)" & _
        ",\s+supra,\s+" & _
        "(\d+\s+" & ReporterPattern() & ")" & _
        "\s+at\s+pp?\.\s+" & _
        "(\d+(?:[-,\s]\d+)*)" & _
        "(\s*\[[\s\S]*?\])?"
End Function

Private Function HasPrecedingFullCiteInQuote(pt As String, _
                                              qm() As Boolean, _
                                              pos As Long) As Boolean
    HasPrecedingFullCiteInQuote = False
    If pos < 2 Then Exit Function

    Dim qStart As Long: qStart = pos
    Do While qStart > 1
        If Not qm(qStart - 1) Then Exit Do
        qStart = qStart - 1
    Loop

    If pos <= qStart Then Exit Function
    Dim textBefore As String: textBefore = Mid(pt, qStart, pos - qStart)
    If Len(Trim(textBefore)) = 0 Then Exit Function

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.Multiline = False
    re.Pattern = BuildLongCitePattern()
    HasPrecedingFullCiteInQuote = (re.Execute(textBefore).count > 0)
End Function






