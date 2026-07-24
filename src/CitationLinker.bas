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

' Provider whose search URLs the links point to: "lexis" or "westlaw".
' Persisted in the registry (SaveSetting/GetSetting) so it survives Word
' restarts. Flip it with the ToggleCitationProvider macro -- no code edit
' needed. New installs default to Westlaw.
Private Const PROVIDER_APP     As String = "MyMacros"
Private Const PROVIDER_SECTION As String = "CitationLinker"
Private Const PROVIDER_KEY     As String = "Provider"
Private Const PROVIDER_DEFAULT As String = "westlaw"

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

' Private: driven through ToggleCitationLinks (Ctrl+Shift+H), so it is kept off
' the Alt+F8 list. Still callable within this module.
Private Sub AddCitationLinks()
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

    ' Arm the handler for the whole shell/IO/parse phase, not just the linking
    ' loop: Python missing (Shell raises), an unwritable TEMP, a bridge that
    ' exits 0 without writing tmpOut, or a malformed TSV row all previously
    ' surfaced as raw unhandled runtime-error dialogs.
    On Error GoTo CleanUp

    ' Delete last run's output BEFORE running the bridge. If the bridge fails
    ' without writing, a stale TSV from a DIFFERENT document would otherwise be
    ' read back and its offsets applied to this one.
    On Error Resume Next
    Kill tmpOut
    On Error GoTo CleanUp

    WriteUtf8File tmpIn, docHtml

    ' Run the bridge and wait.
    Dim cmd As String
    cmd = Q(PYTHON_EXE) & " " & Q(SCRIPT_DIR & "\word_cite_bridge.py") & _
          " " & Q(tmpIn) & " " & Q(tmpOut)
    If Len(REPO_JSON) > 0 Then cmd = cmd & " " & Q(REPO_JSON)
    ' Provider is validated to "lexis"/"westlaw", so it needs no quoting.
    cmd = cmd & " --provider " & CitationProvider()

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

    ' Parse rows. Normalize CRLF first: a bridge writing in Windows text mode
    ' would otherwise leave a trailing CR on every row's last field, breaking
    ' the Find fallback that searches for that text verbatim.
    tsv = Replace(tsv, vbCrLf, vbLf)
    tsv = Replace(tsv, vbCr, vbLf)
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

    ' Catch subsequent "..., supra, <vol reporter> at p. <pages>" cites the
    ' bridge left unlinked. This happens when the full cite's short name is set
    ' by a parenthetical override -- e.g. "... 1251, 1261 (Grand Terrace)" -- that
    ' the extractor never ties back to "Grand Terrace, supra". We match on the
    ' reporter volume, which the supra shares verbatim with the full cite, and
    ' reuse that full cite's URL.
    LinkOrphanSupraCites doc, keep, added

CleanUp:
    ' Capture the error before any On Error statement clears it.
    Dim lErrN As Long, sErrD As String
    lErrN = Err.Number
    sErrD = Err.Description
    On Error Resume Next
    Kill tmpIn
    Kill tmpOut
    Application.ScreenUpdating = True
    On Error GoTo 0
    If lErrN <> 0 Then
        MsgBox "Citation Linker stopped after an error:" & vbCrLf & vbCrLf & _
               "Error " & lErrN & ": " & sErrD & vbCrLf & vbCrLf & _
               "If this mentions a missing file or path, check PYTHON_EXE and " & _
               "SCRIPT_DIR at the top of the module.", _
               vbExclamation, "Citation Linker"
    Else
        MsgBox "Linked " & added & " citation" & IIf(added = 1, "", "s") & _
               " (" & ProviderDisplay(CitationProvider()) & ").", _
               vbInformation, "Citation Linker"
    End If
End Sub


' Private: driven through ToggleCitationLinks (Ctrl+Shift+H), so it is kept off
' the Alt+F8 list. Still callable within this module.
Private Sub RemoveCitationLinks()
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
'
' Must stay a no-argument Public Sub: it is bound to Ctrl+Shift+H via
' KeyBindings.Add, and a macro that takes an argument cannot be a key-binding
' target (the binding fails with runtime error 5346). It therefore stays in the
' Alt+F8 list -- that is the price of being key-bindable.
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


' Flip the citation-link provider between Westlaw and Lexis+ and remember the
' choice across Word sessions. Run it again to switch back. Bind it to a
' shortcut if you switch often. The next AddCitationLinks uses the new provider.
Public Sub ToggleCitationProvider()
    Dim cur As String: cur = CitationProvider()
    Dim nxt As String
    If cur = "westlaw" Then nxt = "lexis" Else nxt = "westlaw"
    SaveSetting PROVIDER_APP, PROVIDER_SECTION, PROVIDER_KEY, nxt
    MsgBox "Citation links now point to " & ProviderDisplay(nxt) & "." & vbCrLf & vbCrLf & _
           "Run ToggleCitationProvider again to switch back to " & ProviderDisplay(cur) & ".", _
           vbInformation, "Citation Linker"
End Sub


' The provider whose search URLs the linker builds: "lexis" or "westlaw".
' Read from the registry each time so a toggle takes effect on the next run;
' defaults to Westlaw until changed.
Private Function CitationProvider() As String
    Dim p As String
    p = LCase$(Trim$(GetSetting(PROVIDER_APP, PROVIDER_SECTION, PROVIDER_KEY, PROVIDER_DEFAULT)))
    If p <> "lexis" And p <> "westlaw" Then p = PROVIDER_DEFAULT
    CitationProvider = p
End Function


Private Function ProviderDisplay(ByVal p As String) As String
    If LCase$(p) = "lexis" Then ProviderDisplay = "Lexis+" Else ProviderDisplay = "Westlaw"
End Function


' Private so it stays off the Alt+F8 list (rarely needed; use Ctrl+Shift+H for
' this tool's own links). Run it from the VBE if you ever need the "remove EVERY
' hyperlink" behavior.
Private Sub RemoveAllHyperlinks()
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

    ' Never start a citation link at the citation sentence's outer "(" (or a
    ' leading "[", quote, or space). The literal-text fallback in particular can
    ' hand us a range that begins with "(" -- e.g. "(Commodore Home Systems,
    ' Inc. v. Superior Court ...". Trim any such leading characters off the
    ' anchor so the hyperlink begins at the case name.
    On Error Resume Next
    Do While rng.Characters.count > 1
        Dim fch As String: fch = rng.Characters(1).text
        If fch = "(" Or fch = "[" Or fch = " " Or fch = Chr$(160) _
           Or fch = ChrW$(8220) Or fch = ChrW$(8216) Or fch = Chr$(34) Then
            rng.MoveStart wdCharacter, 1
        Else
            Exit Do
        End If
    Loop
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
        ' No case-name tail (no "(year)" / ", supra" / " v. ") INSIDE the link.
        ' Two supra shapes leave the short name outside that logic:
        '  (a) the display IS the short name and ", supra, <reporter>" follows
        '      OUTSIDE the link ("<link>Galleria Plus, Inc.</link>, supra,
        '      179 Cal.App.4th at p. 538"). The Hyperlink style stripped the
        '      short name's italic, so treat the whole display as the case name.
        '  (b) the short name sits BEFORE the link (the linker anchored on the
        '      reporter) -- italicize that preceding run.
        If LinkFollowedBySupra(disp) Then
            tailStart = Len(s) + 1        ' whole display is the case short name
        Else
            ' The link may begin with the "supra" signal (the orphan-supra
            ' linker pulls "supra" into the link so there's no gap). Italicize
            ' that "supra" inside the link, and the case short name before it.
            ItalicizeLeadingSupra disp
            ItalicizeSupraShortNameBefore disp
            Exit Sub
        End If
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

    ' Clean slate first: clear italic across the WHOLE display, extending one
    ' position back into the hidden field separator so the boundary's first
    ' character is reached too. This removes any stray italic -- e.g. a leading
    ' "(" left italic by an earlier build or a previous link/unlink cycle -- so
    ' only the case name ends up italic no matter the document's prior state.
    ActiveDocument.Range(disp.Characters(1).start - 1, disp.Characters(m).End).Font.Italic = False

    ' Now italicize the case-name run as one range. Only when it starts at the
    ' very first display character do we extend the start one position back into
    ' the field separator, so the boundary doesn't absorb the italic on that
    ' first letter. (Characters(1).Start is the true text position; the Range's
    ' own .Start points into the field code and must not be used here.)
    Dim startPos As Long
    startPos = disp.Characters(nameStart).start
    Dim extendedBack As Boolean: extendedBack = (nameStart = 1)
    If extendedBack Then startPos = startPos - 1

    ActiveDocument.Range(startPos, disp.Characters(nameEnd).End).Font.Italic = True

    ' Undo the leak from the back-extension. Extending the italic start one
    ' position before the first display character also italicizes whatever plain
    ' character sits immediately before the hyperlink field -- for a citation
    ' SENTENCE that is the outer "(" (e.g. "(Gutierrez v. Tostado (2025) ...").
    ' The first display letter's italic is stored on its own character run, so
    ' clearing italic on just that one preceding character removes the stray
    ' italic on the "(" without disturbing the case name. Harmless when the
    ' preceding character is a space (in-text cites): clearing invisible italic
    ' on a space changes nothing visible.
    If extendedBack Then
        ActiveDocument.Range(startPos, disp.Characters(1).start).Font.Italic = False
    End If

    ' When the whole supra cite is one link ("Galleria Plus, Inc., supra, 179
    ' Cal.App.4th at p. 538"), the case-name run above ends at ", supra"; also
    ' italicize the "supra" word so only the reporter stays roman. No-op when
    ' there is no ", supra" in the display.
    ItalicizeSupraWordInDisplay disp
End Sub

' Italicize a "supra" signal that appears inside the display AFTER the case name
' ("<name>, supra, <reporter>"). The clean-slate above left it roman; this adds
' the italic so the reporter alone stays roman. No-op when the display has no
' ", supra".
Private Sub ItalicizeSupraWordInDisplay(ByVal disp As Range)
    On Error Resume Next
    Dim s As String: s = disp.text
    Dim cp As Long: cp = InStr(1, s, ", supra", vbTextCompare)
    If cp < 1 Then Exit Sub
    Dim sp As Long: sp = cp + 2                  ' 1-based index of "supra" (after ", ")
    Dim m As Long: m = disp.Characters.count
    If sp < 1 Or sp > m Then Exit Sub
    Dim endIdx As Long: endIdx = sp + 4          ' "supra" is 5 characters
    If endIdx > m Then endIdx = m
    ActiveDocument.Range(disp.Characters(sp).start, disp.Characters(endIdx).End).Font.Italic = True
End Sub

' True when the text immediately AFTER the link begins with ", supra" -- i.e.
' the linked display is the case short name of a supra cite and the "supra,
' <reporter>" tail follows outside the link. The comma may be inside or outside
' the link, so leading whitespace, commas, non-breaking spaces, and the hidden
' field-end control marks are skipped before the "supra" test.
Private Function LinkFollowedBySupra(ByVal disp As Range) As Boolean
    On Error Resume Next
    Dim m As Long: m = disp.Characters.count
    If m < 1 Then Exit Function
    Dim aStart As Long: aStart = disp.Characters(m).End
    Dim after As String
    after = ActiveDocument.Range(aStart, aStart + 16).text
    Do While Len(after) > 0
        Dim c As String: c = Left$(after, 1)
        If c = " " Or c = "," Or c = Chr$(160) Or AscW(c) <= 31 Then
            after = Mid$(after, 2)
        Else
            Exit Do
        End If
    Loop
    LinkFollowedBySupra = (LCase$(Left$(after, 5)) = "supra")
End Function

' When a supra cite's link display begins with the "supra" signal (the orphan-
' supra linker pulls "supra" into the link so the hyperlink is continuous),
' italicize just that "supra" word inside the link -- the reporter that follows
' stays roman, matching legal style. No-op unless the display leads with "supra".
Private Sub ItalicizeLeadingSupra(ByVal disp As Range)
    On Error Resume Next
    Dim s As String: s = disp.text
    Dim m As Long: m = disp.Characters.count
    If m < 1 Then Exit Sub

    Dim p As Long: p = InStr(1, s, "supra", vbTextCompare)
    If p < 1 Or p > m Then Exit Sub
    ' Only the LEADING signal counts -- nothing but a comma/space may precede it.
    If Len(Trim$(Replace(Left$(s, p - 1), ",", ""))) > 0 Then Exit Sub

    Dim endIdx As Long: endIdx = p + 4               ' "supra" is 5 characters
    If endIdx > m Then endIdx = m

    ' Clean slate across the display (extend one back into the field separator),
    ' then italicize only the "supra" run.
    ActiveDocument.Range(disp.Characters(1).start - 1, disp.Characters(m).End).Font.Italic = False
    Dim startPos As Long: startPos = disp.Characters(p).start
    Dim extendedBack As Boolean: extendedBack = (p = 1)
    If extendedBack Then startPos = startPos - 1
    ActiveDocument.Range(startPos, disp.Characters(endIdx).End).Font.Italic = True
    If extendedBack Then
        ActiveDocument.Range(startPos, disp.Characters(1).start).Font.Italic = False
    End If
End Sub

' Italicize the short name of a supra cite that sits just BEFORE the link, e.g.
' the document reads "Rappleyea, supra, " and then the link (which now includes
' the "supra" signal, "supra, 8 Cal.4th at p. 982"). The short name is outside
' the hyperlink, so it is a plain document range (no field-boundary quirk). Only
' called when the in-link logic found nothing, so it never disturbs cites handled
' inside the link.
Private Sub ItalicizeSupraShortNameBefore(ByVal disp As Range)
    On Error Resume Next
    Dim linkStart As Long: linkStart = disp.start
    If linkStart < 8 Then Exit Sub

    Dim lookLen As Long: lookLen = 70
    If lookLen > linkStart Then lookLen = linkStart
    Dim base As Long: base = linkStart - lookLen
    Dim b As String: b = ActiveDocument.Range(base, linkStart).text
    If Len(b) = 0 Then Exit Sub

    ' This must be a supra context. The document reads "<short name>, supra,
    ' <reporter>"; depending on where the link starts, the text before it ends
    ' either with "..., supra" (link anchored on the reporter) or with "..., "
    ' (link now includes "supra"). Confirm via one of those two signals so this
    ' never italicizes the tail of an ordinary preceding sentence.
    Dim leadSupra As Boolean
    Dim dLead As String: dLead = disp.text
    Do While Len(dLead) > 0
        Dim dc As String: dc = Left$(dLead, 1)
        If dc = " " Or dc = "," Then dLead = Mid$(dLead, 2) Else Exit Do
    Loop
    leadSupra = (LCase$(Left$(dLead, 5)) = "supra")

    Dim t As String: t = b
    Do While Len(t) > 0
        Dim last As String: last = Right$(t, 1)
        If last = " " Or last = "," Then t = Left$(t, Len(t) - 1) Else Exit Do
    Loop
    Dim beforeSupra As Boolean: beforeSupra = (Len(t) >= 5 And LCase$(Right$(t, 5)) = "supra")

    If Not leadSupra And Not beforeSupra Then Exit Sub

    ' If "supra" trails the before-text (it's outside the link), drop it and its
    ' comma so we land on the short name -- same landing as the in-link case.
    If beforeSupra Then
        t = Left$(t, Len(t) - 5)
        Do While Len(t) > 0
            Dim l2 As String: l2 = Right$(t, 1)
            If l2 = " " Or l2 = "," Then t = Left$(t, Len(t) - 1) Else Exit Do
        Loop
    End If
    If Len(t) = 0 Then Exit Sub
    Dim nameEnd As Long: nameEnd = Len(t)            ' last char of the short name (index in b)

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


' After the bridge links are placed, hyperlink any "supra" cite it left behind.
' A subsequent cite such as "Grand Terrace, supra, 192 Cal.App.3d at pp.
' 1266-1267" shares its reporter volume ("192 Cal.App.3d") verbatim with the
' full cite that established the case, so we link the reporter-through-pincite
' span to that full cite's URL. This is the common miss when the full cite's
' short name comes from a parenthetical override ("... 1251, 1261 (Grand
' Terrace)") the extractor never associates with the short form. Best-effort:
' any failure is swallowed so it can never disturb the links already placed.
Private Sub LinkOrphanSupraCites(ByVal doc As Document, ByRef keep() As CiteRow, _
                                 ByRef added As Long)
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    ' ", supra, <volume reporter> at p. <pages>"  /  "at pp. <pages>".
    ' The reporter group is lazy so it stops at " at p". The page tail accepts
    ' one page or one hyphen/en-dash range and then STOPS: the old open class
    ' [\d,\s-]* ran through commas and spaces, so "at p. 982, 30 days later"
    ' linked through ", 30" and "pp. 1266-1267, and" carried a trailing comma
    ' into the hyperlink.
    re.Pattern = ",\s+supra,\s+(\d{1,4}\s+[A-Za-z][A-Za-z.\d ]*?)\s+at\s+pp?\.\s*" _
                 & "\d+(?:\s*[" & ChrW(8211) & "\-]\s*\d+)?"

    Dim p As Paragraph
    For Each p In doc.Paragraphs
        Dim raw As String
        raw = ParagraphRawText(p.Range)
        If Len(raw) = 0 Then GoTo NextPara

        Dim ms As Object
        Set ms = re.Execute(raw)
        If ms.Count = 0 Then GoTo NextPara

        Dim mm As Object
        For Each mm In ms
            Dim repVol As String
            repVol = Trim$(mm.SubMatches(0))
            If Len(repVol) = 0 Then GoTo NextMatch

            Dim url As String
            url = UrlForReporterVol(repVol, keep)
            If Len(url) = 0 Then GoTo NextMatch

            ' Link the ENTIRE supra cite as one hyperlink -- the case short name,
            ' the ", supra" connective, and the reporter through the pincite. The
            ' regex match starts at the ", supra" comma; walk back through the raw
            ' paragraph text to the start of the short name and link from there.
            ' Only the APPLIED span is widened -- the URL is still resolved from
            ' the reporter volume alone (UrlForReporterVol above). Falls back to
            ' the match alone when no short name is found.
            Dim matchStart As Long: matchStart = mm.FirstIndex + 1   ' 1-based, at the "," of ", supra"
            Dim nameStart As Long: nameStart = SupraShortNameStart(raw, matchStart)
            Dim linkText As String
            If nameStart > 0 And nameStart < matchStart Then
                linkText = Mid$(raw, nameStart, matchStart + Len(mm.Value) - nameStart)
            Else
                linkText = mm.Value
            End If

            LinkTextIfUnlinked p.Range, linkText, url, added
NextMatch:
        Next mm
NextPara:
    Next p
End Sub


' Given the raw paragraph text and the 1-based index of the ", supra" comma,
' return the 1-based index where the case SHORT NAME begins, so the orphan-supra
' link can start there -- or 0 to NOT extend (link from ", supra" only).
'
' The short name is delimited only when it sits behind a STRUCTURAL boundary:
' an opening "(" / "[" (the citation sentence's outer paren, never linked), or a
' ")" / "]" / ";" that closes a prior clause or citation -- exactly the
' parenthetical and string-cite shapes where supra cites live. A ". " boundary
' is deliberately NOT used: it can't be told apart from "v." or "Inc." inside a
' case name, and walking through open prose would swallow preceding words ("The
' court in Grand Terrace, supra..." -> the whole phrase). When no structural
' boundary is found before the paragraph start, return 0 so the link falls back
' to starting at ", supra" -- a smaller span is far safer than a wrong one.
Private Function SupraShortNameStart(ByVal raw As String, ByVal commaPos As Long) As Long
    Dim nameEnd As Long: nameEnd = commaPos - 1
    Do While nameEnd >= 1 And Mid$(raw, nameEnd, 1) = " ": nameEnd = nameEnd - 1
    Loop
    If nameEnd < 1 Then Exit Function

    Dim k As Long: k = nameEnd
    Dim foundBoundary As Boolean: foundBoundary = False
    Do While k >= 1
        Dim ch As String: ch = Mid$(raw, k, 1)
        If ch = "(" Or ch = ")" Or ch = "[" Or ch = "]" Or ch = ";" Then
            foundBoundary = True
            Exit Do
        End If
        k = k - 1
    Loop
    If Not foundBoundary Then Exit Function          ' no structural delimiter -> don't extend
    Dim nameStart As Long: nameStart = k + 1

    ' Skip leading spaces and citation signal words -- lowercase in running text
    ' ("see Galleria...") or Capitalized in a parenthetical ("(See Galleria...")
    ' -- so the link starts at the capitalized case name, not the signal.
    Do
        Do While nameStart <= nameEnd And Mid$(raw, nameStart, 1) = " ": nameStart = nameStart + 1
        Loop
        If nameStart > nameEnd Then Exit Function
        Dim wordEnd As Long: wordEnd = nameStart
        Do While wordEnd <= nameEnd And Mid$(raw, wordEnd, 1) <> " ": wordEnd = wordEnd + 1
        Loop
        Dim tok As String: tok = LCase$(Mid$(raw, nameStart, wordEnd - nameStart))
        Do While Len(tok) > 0 And (Right$(tok, 1) = "," Or Right$(tok, 1) = ".")
            tok = Left$(tok, Len(tok) - 1)
        Loop
        If IsCiteSignalWord(tok) Then
            nameStart = wordEnd         ' skip the signal word, keep scanning
        Else
            Exit Do
        End If
    Loop
    If nameStart > nameEnd Then Exit Function
    SupraShortNameStart = nameStart
End Function

' A leading citation signal word (case-insensitive, trailing punctuation already
' stripped) that precedes a case name and should stay OUT of the hyperlink.
Private Function IsCiteSignalWord(ByVal w As String) As Boolean
    Select Case w
        Case "see", "also", "generally", "cf", "accord", "contra", "but", _
             "compare", "e.g", "eg"
            IsCiteSignalWord = True
    End Select
End Function


' Return the URL of the linked full cite whose text contains reporter volume
' repVol (e.g. "192 Cal.App.3d"). Returns "" when none match or when the volume
' is claimed by two different URLs (ambiguous -- safer to leave it unlinked).
Private Function UrlForReporterVol(ByVal repVol As String, ByRef keep() As CiteRow) As String
    On Error GoTo Fail
    Dim wantUrl As String: wantUrl = ""

    Dim i As Long
    For i = LBound(keep) To UBound(keep)
        Dim t As String: t = keep(i).txt
        If Len(t) = 0 Or Len(keep(i).url) = 0 Then GoTo NextRow

        Dim pos As Long: pos = InStr(1, t, repVol, vbTextCompare)
        Do While pos > 0
            ' Require a non-digit (or string start) just before the volume so
            ' "192 Cal.App.3d" is not matched inside "1192 Cal.App.3d".
            Dim okLeft As Boolean: okLeft = (pos = 1)
            If Not okLeft Then okLeft = Not (Mid$(t, pos - 1, 1) Like "#")
            If okLeft Then
                If wantUrl = "" Then
                    wantUrl = keep(i).url
                ElseIf StrComp(wantUrl, keep(i).url, vbTextCompare) <> 0 Then
                    UrlForReporterVol = ""       ' ambiguous volume
                    Exit Function
                End If
                Exit Do
            End If
            pos = InStr(pos + 1, t, repVol, vbTextCompare)
        Loop
NextRow:
    Next i

    UrlForReporterVol = wantUrl
    Exit Function
Fail:
    UrlForReporterVol = ""
End Function


' Find needle inside scope with Word Find (so field/footnote positions are
' handled) and hyperlink the first occurrence that is not already linked.
Private Sub LinkTextIfUnlinked(ByVal scope As Range, ByVal needle As String, _
                               ByVal url As String, ByRef added As Long)
    On Error GoTo Done
    If Len(needle) = 0 Or Len(needle) > 250 Then Exit Sub

    Dim searchStart As Long: searchStart = scope.Start
    Dim guard As Long: guard = 0
    Do
        guard = guard + 1
        If guard > 50 Then Exit Do

        Dim fr As Range
        Set fr = ActiveDocument.Range(searchStart, scope.End)
        With fr.Find
            .ClearFormatting
            .text = needle
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            .MatchCase = True
            .Execute
        End With
        If Not fr.Find.Found Then Exit Do

        If fr.Hyperlinks.Count = 0 Then
            If AddLink(fr, url, "case") Then added = added + 1
            Exit Do
        End If

        ' This occurrence is already linked -- resume past it.
        searchStart = fr.End
        If searchStart >= scope.End Then Exit Do
    Loop
Done:
End Sub


Private Sub ResetLinkFormatting(ByVal rng As Range)
    On Error Resume Next

    ' Code-section heading: the whole line is a code section on its own (e.g.
    ' "Civil Code Section 1942.4"), optionally after a roman numeral + period
    ' ("I. Civil Code Section 1942.4"). Such a heading is underlined, so APPLY
    ' the underline to the (former) link and stop. This is the case the adjacency
    ' check below can't catch -- a line that is ENTIRELY one linked code section
    ' has no adjacent underlined character to key off, so the underline was lost.
    If ParaIsCodeSectionHeading(rng) Then
        rng.Font.Underline = wdUnderlineSingle
        rng.Font.ColorIndex = wdAuto
        Exit Sub
    End If

    ' Clear the hyperlink style's underline -- but NOT when the link sat inside
    ' text that is itself underlined (an underlined section heading containing a
    ' code section, e.g. "Retaliation Under Labor Code Sections 98.6 and
    ' 1102.5"). Look past the word-separating spaces to the nearest VISIBLE
    ' character on each side: probing only the single adjacent character read the
    ' separating space, and a heading whose underline does not paint those spaces
    ' (word-style underline, or spaces simply left un-underlined) then looked
    ' un-underlined on both sides, so the citation's underline was wrongly
    ' stripped while the surrounding words stayed underlined. When an underlined
    ' neighbor is found, restore that exact underline style across the range so
    ' it matches the rest of the heading.
    Dim nbr As WdUnderline
    nbr = NeighborUnderline(rng)
    If nbr <> wdUnderlineNone And nbr <> wdUndefined Then
        rng.Font.Underline = nbr
    Else
        rng.Font.Underline = wdUnderlineNone
    End If
    rng.Font.ColorIndex = wdAuto
End Sub


' Underline style of the nearest visible (non-whitespace) character next to rng,
' scanning left first and then right, but never past rng's own paragraph. Returns
' wdUnderlineNone when the nearest neighbor on each side is un-underlined (or the
' paragraph has no other visible character). Used by ResetLinkFormatting to tell
' a code section embedded in an underlined heading -- where the underline must be
' kept -- from an ordinary body citation, where it must be cleared. Word-
' separating spaces are skipped so a heading that underlines words but not the
' spaces between them is still recognized as underlined.
Private Function NeighborUnderline(ByVal rng As Range) As WdUnderline
    On Error Resume Next
    NeighborUnderline = wdUnderlineNone

    Dim para As Range
    Set para = rng.Paragraphs(1).Range
    Dim pStart As Long, pEnd As Long
    pStart = para.start
    pEnd = para.End

    Dim pos As Long, u As WdUnderline

    ' Left: first non-whitespace character before the link, within the paragraph.
    For pos = rng.start - 1 To pStart Step -1
        If Not IsSkippableChar(ActiveDocument.Range(pos, pos + 1).text) Then
            u = ActiveDocument.Range(pos, pos + 1).Font.Underline
            If u <> wdUnderlineNone And u <> wdUndefined Then
                NeighborUnderline = u
                Exit Function
            End If
            Exit For
        End If
    Next pos

    ' Right: first non-whitespace character after the link, within the paragraph.
    For pos = rng.End To pEnd - 1
        If Not IsSkippableChar(ActiveDocument.Range(pos, pos + 1).text) Then
            u = ActiveDocument.Range(pos, pos + 1).Font.Underline
            If u <> wdUnderlineNone And u <> wdUndefined Then NeighborUnderline = u
            Exit For
        End If
    Next pos
End Function


' True when ch is empty or a single whitespace character (space, tab, non-
' breaking space, paragraph/line marks, and the Unicode spaces recognized
' elsewhere in this module). Used to skip word separators when hunting for the
' nearest visible neighbor of a former link.
Private Function IsSkippableChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then
        IsSkippableChar = True
    Else
        IsSkippableChar = IsWhitespaceCode(AscW(Left$(ch, 1)))
    End If
End Function


' True when rng's paragraph is a standalone code-section heading: the whole
' line, after an optional roman-numeral prefix, is a single code-section
' citation and nothing else. Used to keep/apply the underline on such a heading
' when its hyperlink is removed.
Private Function ParaIsCodeSectionHeading(ByVal rng As Range) As Boolean
    On Error Resume Next
    Dim s As String
    s = rng.Paragraphs(1).Range.text

    ' Strip trailing paragraph/line/cell marks and spaces.
    Do While Len(s) > 0
        Dim c As String: c = Right$(s, 1)
        If c = vbCr Or c = vbLf Or c = Chr$(11) Or c = Chr$(12) Or c = Chr$(7) Or c = " " Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    s = Trim$(s)
    ' Headings are short; a length cap keeps a prose sentence that merely names
    ' a section from ever qualifying.
    If Len(s) = 0 Or Len(s) > 90 Then Exit Function

    ParaIsCodeSectionHeading = IsCodeSectionHeadingText(s)
End Function

' Regex test: an entire line that is a code-section citation, optionally led by
' a roman numeral + period. Accepts an optional code-name run ("Civil Code ",
' "Code Civ. Proc., ") before the section marker (Section / Sec. / section
' sign), then a section number with optional dotted parts and (a)(1)-style
' subdivisions -- and NOTHING after it, so "Section 5 of the lease" (a sentence)
' does not match. Case-insensitive.
Private Function IsCodeSectionHeadingText(ByVal s As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Global = False
        re.Pattern = "^(?:[IVXLCDM]{1,7}\.\s+)?" & _
                     "(?:[A-Za-z][A-Za-z.,'&/ ]*\s)?" & _
                     "(?:" & ChrW(167) & "|Section|Sec\.)\s*" & _
                     "\d[\d.]*(?:\s*\([A-Za-z0-9]+\))*\.?$"
    End If
    IsCodeSectionHeadingText = re.Test(s)
End Function


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
