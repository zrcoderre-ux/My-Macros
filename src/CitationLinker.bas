Attribute VB_Name = "CitationLinker"
'==============================================================================
' CitationLinker.bas
'------------------------------------------------------------------------------
' Hyperlinks every legal authority in the active Word document, and removes
' those links again on demand. Detection is delegated to citation_extractor.py
' (your existing tool) through word_cite_bridge.py, so there is one source of
' truth for citation parsing.
'
' Two bodies of authority the extractor does not read are read here instead --
' a bare "rule 3.1350(f)" (LinkRulesOfCourt) and a California Constitution
' reference (LinkCalConstitution) -- and both take their address from the
' extractor's own rows, so every link in the document lands on the same
' service.
'
' MACROS YOU RUN:
'   AddCitationLinks       - detect + hyperlink every authority (idempotent)
'   RemoveCitationLinks    - remove only the links this tool added (recommended)
'   RemoveAllHyperlinks    - remove EVERY hyperlink in the body (asks first)
'   ToggleCitationLinks    - Ctrl+Shift+H: remove this tool's links if any are
'                            present, otherwise apply them. Also runs the heading
'                            pass (HeadingFormat.ApplyHeadingFormat) on every
'                            press, one way -- see the note at that call.
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
' What the citation-italics pass decided about one character. 0 -- the default
' for every character it did not read -- means leave that character exactly as
' the document has it.
Private Const MARK_ITALIC As Byte = 1
Private Const MARK_ROMAN  As Byte = 2

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

    ' Declared up here because the early exits below now link too: a document
    ' the bridge finds nothing in can still carry Rules of Court references,
    ' and CleanUp reports this count on every path. The row array goes with it:
    ' those exits hand it to the rule pass before the bridge's rows have been
    ' filtered into it, so its declaration has to come first. Unallocated is the
    ' right state there -- no bridge rows means no case cite to borrow a URL
    ' from, and RuleUrl (like ConstUrl) falls through to a public address.
    Dim added As Long
    Dim keep() As CiteRow

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
        ' Nothing from the bridge is not nothing to link. The Rules of Court
        ' and Constitution passes read the document itself, so a ruling that
        ' cites a rule or article and no cases still has links to place;
        ' CleanUp reports what it did.
        Application.ScreenUpdating = False
        LinkRulesOfCourt doc, keep, added
        LinkCalConstitution doc, keep, added
        GoTo CleanUp
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
        ' As above: no rows from the bridge still leaves those passes their work.
        Application.ScreenUpdating = False
        LinkRulesOfCourt doc, keep, added
        LinkCalConstitution doc, keep, added
        GoTo CleanUp
    End If
    ReDim Preserve rows(0 To cnt - 1)

    SortRows rows
    keep = FilterOverlaps(rows)

    ' Apply links in reverse document order so any positional shift from a
    ' hyperlink field only affects text to the right of spans not yet linked.
    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    Dim curBlk As Long, hasN As NormResult
    curBlk = -1

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

    ' And the California Rules of Court references the bridge left bare. It
    ' hands over its own rows so those links can point where its do.
    LinkRulesOfCourt doc, keep, added

    ' And the California Constitution, which the bridge does not read at all.
    ' Same borrow, for the same reason: its links land on the provider the rest
    ' of the document's links land on.
    LinkCalConstitution doc, keep, added

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
    ElseIf added = 0 Then
        MsgBox "No legal authorities were detected.", vbInformation, "Citation Linker"
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
    Dim removed As Long, strayLeft As Long
    removed = RemoveCitationLinks_Quiet(doc, strayLeft)

    ' Say so when a URL was left in the text. Silence would read as a clean
    ' removal while a search URL sits in the middle of a sentence.
    Dim leftover As String
    If strayLeft > 0 Then
        leftover = vbCrLf & vbCrLf & strayLeft & " of them displayed a search URL " & _
                   "instead of a citation -- damage from an older build of this " & _
                   "macro -- and Word would not delete that text (it refuses some " & _
                   "deletions in a document carrying tracked changes). The links " & _
                   "are gone; delete the leftover URL text by hand."
    End If

    MsgBox "Removed " & removed & " citation link" & IIf(removed = 1, "", "s") & "." & _
           leftover, vbInformation, "Citation Linker"
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

    ' Heading housekeeping rides along on every press: "keep with next" on every
    ' section heading and the underline on roman-numeral titles. This shortcut is
    ' the one in constant use, so the headings come with it rather than costing a
    ' second keystroke. FormatHeadings (Ctrl+Shift+K) still runs the same pass on
    ' its own, and reports what it did; here it is silent, because the dialog
    ' above has already reported the press.
    '
    ' ONE WAY. Removing the citation links does NOT take the heading formatting
    ' back out: that formatting is how the document is meant to read, not
    ' something this macro owns and lends.
    '
    ' AFTER the link work, deliberately. Removing a link re-derives the underline
    ' around it (ResetLinkFormatting), which would clear an underline applied
    ' before it -- so the heading pass has to be the last word, not the first.
    Dim nKept As Long, nLined As Long
    HeadingFormat.ApplyHeadingFormat doc, nKept, nLined

    ' "supra" is always italicized, linked or not. The in-link italic logic only
    ' reaches a "supra" that sits INSIDE a link's display text, which leaves the
    ' common shape untouched: when the link covers the case short name alone, the
    ' ", supra, 179 Cal.App.4th at p. 538" that follows is outside it. Sweeping
    ' the body catches those, and the supra cites that never resolved to a link
    ' at all. One way, like the heading pass: removing the links does not take
    ' the italic back out -- it is the citation's own style, not link decoration.
    ' Runs last, after the link work, because adding a link re-derives the italic
    ' across its display and would otherwise overwrite this.
    ItalicizeSupraEverywhere doc

    ' Straight quotes to curly, both families -- ' and " alike -- direction-aware,
    ' so an opening mark curls open and a closing one (and every apostrophe)
    ' curls closed. Silent, and one way like the passes above: it is typographic
    ' cleanup the document always wants, not link decoration to be taken back
    ' out on the next press. The close review runs the same helper on
    ' apostrophes only; here the doubles come along, because a press of this
    ' shortcut is the user asking for the cleanup.
    modMain.SmartenStraightQuotes doc, True

    ' Citation italics, re-derived from the paragraph text and applied through
    ' Find. Runs after every pass above that touches italics, so where it can
    ' read a citation it has the last word on how that citation is set.
    NormalizeCitationItalics doc

    ' A blank line ahead of each ALL-CAPS label, so the top-level sections read
    ' apart. LAST of everything: it decides by where the text falls on the page,
    ' so every pass that can move text -- the links, the headings, the supra
    ' italics, the quotes -- has to be finished before it looks.
    HeadingFormat.SpaceCapsHeadings doc
End Sub

' Italicize every whole-word "supra" in the body. Returns how many were changed.
' Case-insensitive on the search but nothing is recased -- "Supra" opening a
' sentence stays capitalized, it just becomes italic.
Private Function ItalicizeSupraEverywhere(ByVal doc As Document) As Long
    On Error Resume Next

    Dim n As Long, guard As Long
    Dim r As Range
    Set r = doc.content.Duplicate

    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "supra"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .Format = False
    End With

    Do While r.Find.Execute
        guard = guard + 1
        If guard > 2000 Then Exit Do      ' no document has this many; backstop
        If r.Font.Italic <> True Then
            r.Font.Italic = True
            n = n + 1
        End If
        r.Collapse wdCollapseEnd
    Loop

    ItalicizeSupraEverywhere = n
End Function


'==============================================================================
' CITATION ITALICS -- FINAL PASS, READ FROM THE TEXT
'==============================================================================
' Every other italic routine in this module derives a citation's italics from ONE
' HYPERLINK'S DISPLAY TEXT and positions the result through that field's
' Characters collection. When a citation is covered by more than one link, or by
' a link that begins partway into it, that arithmetic has nothing to correct it,
' and the italics land on words the citation never puts in italics:
'
'   (Center for Self-Improvement & Community Development v. Lennar Corp. (2009)
'    173 Cal.App.4th 1543, 1554 (Center for Self-Improvement).)
'
' came back with "Center for" and the reporter tail italic, the rest of the case
' name roman, and the short name roman -- and the supra cite that followed it the
' same way.
'
' This pass ignores the links completely. It reads the paragraph's VISIBLE TEXT,
' works out the runs a California citation has, records a decision for each
' CHARACTER, and then writes those decisions one character at a time:
'
'   (Center ... v. Lennar Corp.   (2009) 173 Cal.App.4th 1543, 1554   (Center for
'    -------- italic ---------    ------------ roman ------------      Self-
'                                                                      Improvement)
'                                                                      -- italic --
'
' ONE CHARACTER AT A TIME is the point, and it is the third mechanism tried here.
' Arithmetic on Range.Start misses because .Start counts positions that .Text
' never returns -- a hyperlink's field code among them -- and a citation is
' exactly where the fields are. Word's Find misses because it will not match a
' run that crosses into or out of a field, which is where the reporter half of a
' citation always sits: the case name matched and went italic, the tail did not
' match and kept whatever italic it already had, and the whole citation read as
' italic, numbers and all. A single character is inside the field or outside it,
' never both, and the Characters collection IS the text the marks were worked out
' from, in order. See ApplyMarks.
'
' A citation it cannot read confidently is left exactly as it is -- the left edge
' of a case name is a guess in running prose, and a wrong guess italicizes the
' judge's own sentence. Body only, like the supra sweep above.
Private Function NormalizeCitationItalics(ByVal doc As Document) As Long
    On Error Resume Next
    Dim p As Paragraph, n As Long
    For Each p In doc.content.Paragraphs
        n = n + NormalizeParaItalics(p)
    Next p
    NormalizeCitationItalics = n
End Function

' One paragraph. Full cites are found by their " v. ", short cites by their
' ", supra" -- a full cite carrying both is simply seen twice, and the second
' look asks for the formatting the first one already applied.
Private Function NormalizeParaItalics(ByVal p As Paragraph) As Long
    On Error Resume Next
    Dim s As String
    s = p.Range.text
    If Len(s) = 0 Then Exit Function
    If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
    If Len(s) < 12 Then Exit Function

    ' One byte per character of the paragraph: 0 = not ours, leave it exactly as
    ' it is; MARK_ITALIC; MARK_ROMAN. Nothing is written to the document while
    ' the citations are being read -- the marks are collected first and applied
    ' in a single pass, so a citation read later cannot fight one read earlier.
    Dim marks() As Byte
    ReDim marks(1 To Len(s))

    Dim found As Long
    Dim i As Long

    ' ---- Full cites: "<name> v. <name> (year) <reporter>" ----
    i = 1
    Do
        Dim vp As Long: vp = InStr(i, s, " v. ", vbTextCompare)
        If vp = 0 Then Exit Do
        Dim nextI As Long: nextI = vp + 4
        found = found + MarkCitationAt(s, vp, nextI, marks)
        i = nextI
    Loop

    ' ---- Short cites: "<short name>, supra, <reporter>" ----
    i = 1
    Do
        Dim cp As Long: cp = InStr(i, s, ", supra", vbTextCompare)
        If cp = 0 Then Exit Do
        Dim nextC As Long: nextC = cp + 7
        found = found + MarkCitationAt(s, cp, nextC, marks)
        i = nextC
    Loop

    If found = 0 Then Exit Function
    NormalizeParaItalics = ApplyMarks(p, marks)
End Function

' Write the marks onto the paragraph, ONE CHARACTER AT A TIME.
'
' Character by character on purpose. The Characters collection is the paragraph's
' text, in order, one range per character -- so character i is exactly character
' i of the string the marks were worked out from, whatever fields the paragraph
' contains. Every mechanism tried before this one could miss by a whole run:
' arithmetic on Range.Start counts positions .Text never returns, and Find will
' not match a run that crosses into or out of a hyperlink field, which is where
' the reporter half of a citation always sits. A single character is inside the
' field or outside it, never both.
'
' Returns how many characters were actually changed.
Private Function ApplyMarks(ByVal p As Paragraph, ByRef marks() As Byte) As Long
    On Error Resume Next
    Dim top As Long: top = UBound(marks)

    ' Where each hyperlink's display text begins. Formatting assigned to one of
    ' those characters ALONE is absorbed by the field boundary and never reaches
    ' the letter -- and the linker anchors on the citation, so that character is
    ' the first letter of the case name. Those are written the long way round.
    Dim linkStarts() As Long
    Dim nLinks As Long
    nLinks = CollectLinkStarts(p, linkStarts)

    Dim i As Long, n As Long
    Dim ch As Range

    For Each ch In p.Range.Characters
        i = i + 1
        If i > top Then Exit For
        Select Case marks(i)
            Case MARK_ITALIC
                If SetCharItalic(ch, True, linkStarts, nLinks) Then n = n + 1
            Case MARK_ROMAN
                If SetCharItalic(ch, False, linkStarts, nLinks) Then n = n + 1
        End Select
    Next ch

    ApplyMarks = n
End Function

' Set one character's italic, and mean it.
'
' A plain assignment is enough almost everywhere, and where it is enough that is
' all this does -- a character already set the right way is not written at all,
' which keeps the pass out of the revision marks in a document edited with track
' changes on.
'
' The exception is the first character of a hyperlink's display text. Word keeps
' that character's formatting on the field boundary, so an assignment to the
' character by itself is swallowed and the letter comes back unchanged: the first
' letter of a case name stays roman while the rest of the name goes italic. The
' cure is the one ItalicizeCaseName already uses -- extend the range one position
' back into the boundary and format that -- with the character before it read
' first and put back after, since extending drags it along and it is usually the
' "(" that opens the citation.
'
' The same repair runs whenever a plain assignment does not take, so a boundary
' this code did not predict is still handled.
Private Function SetCharItalic(ByVal ch As Range, ByVal wantItalic As Boolean, _
                                ByRef linkStarts() As Long, ByVal nLinks As Long) As Boolean
    On Error Resume Next
    If ch.Font.Italic = wantItalic Then Exit Function        ' already right

    If Not StartsALink(ch.start, linkStarts, nLinks) Then
        ch.Font.Italic = wantItalic
        If ch.Font.Italic = wantItalic Then
            SetCharItalic = True
            Exit Function
        End If
    End If

    If ch.start < 1 Then Exit Function

    Dim before As Range: Set before = ch.Duplicate
    before.SetRange ch.start - 1, ch.start
    Dim prevState As Long: prevState = before.Font.Italic

    Dim ext As Range: Set ext = ch.Duplicate
    ext.SetRange ch.start - 1, ch.End
    ext.Font.Italic = wantItalic

    ' Put the dragged-along character back the way it was. wdUndefined is not a
    ' state to restore, so a mixed reading is left alone.
    If prevState = True Then
        before.Font.Italic = True
    ElseIf prevState = False Then
        before.Font.Italic = False
    End If

    SetCharItalic = True
End Function

' The display-text start of every hyperlink in the paragraph. Characters(1).Start
' is the first character the reader sees; the field's own .Start points into the
' field code and is no use here.
Private Function CollectLinkStarts(ByVal p As Paragraph, ByRef starts() As Long) As Long
    On Error Resume Next
    ReDim starts(0 To 0)
    Dim total As Long: total = p.Range.Hyperlinks.count
    If total < 1 Then Exit Function

    ReDim starts(1 To total)
    Dim k As Long, got As Long
    For k = 1 To total
        Dim hr As Range
        Set hr = p.Range.Hyperlinks(k).Range
        If hr.Characters.count > 0 Then
            got = got + 1
            starts(got) = hr.Characters(1).start
        End If
    Next k

    CollectLinkStarts = got
End Function

Private Function StartsALink(ByVal pos As Long, ByRef starts() As Long, _
                              ByVal nLinks As Long) As Boolean
    Dim k As Long
    For k = 1 To nLinks
        If starts(k) = pos Then
            StartsALink = True
            Exit Function
        End If
    Next k
End Function

' Mark characters a..b, clamped to the paragraph.
Private Sub MarkRun(ByRef marks() As Byte, ByVal a As Long, ByVal b As Long, _
                     ByVal what As Byte)
    If a < 1 Then a = 1
    If b > UBound(marks) Then b = UBound(marks)
    Dim i As Long
    For i = a To b
        marks(i) = what
    Next i
End Sub

' Format the one citation whose parties end at anchor -- the " v. " of a full
' cite, or the comma of a short cite's ", supra". nextScan comes back at the end
' of what was handled so the caller does not read the same citation twice.
' Returns 1 when the citation was read and formatted, 0 when it was left alone.
Private Function MarkCitationAt(ByVal s As String, ByVal anchor As Long, _
                                 ByRef nextScan As Long, ByRef marks() As Byte) As Long
    ' Where the roman tail starts: the "(year)" or the ", supra".
    Dim ts As Long
    If IsSupraTail(s, anchor) Then ts = anchor Else ts = CiteTailStartAfter(s, anchor)
    If ts <= 0 Then Exit Function

    ' Where the case name starts. 0 means the left edge could not be trusted.
    Dim boundary As Long
    Dim ns As Long: ns = CaseNameLeftEdge(s, anchor, boundary)
    If ns <= 0 Then Exit Function

    Dim ne As Long: ne = ts - 1
    Do While ne >= ns
        If Mid$(s, ne, 1) = " " Then ne = ne - 1 Else Exit Do
    Loop
    If ne < ns Then Exit Function

    Dim caseName As String: caseName = Mid$(s, ns, ne - ns + 1)
    If InStr(1, caseName, " v. ", vbTextCompare) = 0 And Not IsSupraTail(s, ts) Then Exit Function

    MarkRun marks, ns, ne, MARK_ITALIC

    ' The tail is roman: the date through the reporter and pincite. On a short
    ' cite the run starts AFTER ", supra" -- the supra signal is italic, and
    ' ItalicizeSupraEverywhere has just made it so.
    Dim rf As Long: rf = ts
    If IsSupraTail(s, ts) Then rf = ts + 7
    Dim te As Long: te = CiteTailEnd(s, ts)
    If te >= rf Then MarkRun marks, rf, te, MARK_ROMAN

    ' The parenthetical that introduces the case's short name is part of the case
    ' name and carries its italics.
    Dim a As Long, b As Long
    If ShortNameAfter(s, te + 1, caseName, a, b) Then
        MarkRun marks, a, b, MARK_ITALIC
    End If

    If te + 1 > nextScan Then nextScan = te + 1
    MarkCitationAt = 1
End Function

' True when ", supra" starts at position k.
Private Function IsSupraTail(ByVal s As String, ByVal k As Long) As Boolean
    If k < 1 Then Exit Function
    IsSupraTail = (LCase$(Mid$(s, k, 7)) = ", supra")
End Function

' The start of the roman tail after the parties: the ", supra" or the "(year)",
' whichever comes first. 0 when neither is there -- and then the citation is left
' alone, because without a tail there is nothing to say where the name stops.
Private Function CiteTailStartAfter(ByVal s As String, ByVal anchor As Long) As Long
    Dim rest As String: rest = Mid$(s, anchor)
    Dim pSup As Long: pSup = InStr(1, rest, ", supra", vbTextCompare)
    Dim pYr As Long: pYr = FindYearParen(rest)
    Dim p As Long
    If pSup > 0 And (pYr = 0 Or pSup < pYr) Then p = pSup Else p = pYr
    If p = 0 Then Exit Function
    CiteTailStartAfter = anchor + p - 1
End Function

' The last character of the roman tail: the date through the reporter and
' pincite, stopping before whatever comes next -- a short-name or explanatory
' parenthetical, a parallel cite after ";", the ")" closing the citation
' sentence, or (for an in-text cite) the sentence's own prose. Those are
' somebody else's to decide; only the span this is sure of gets rewritten.
'
' The tail is walked TOKEN BY TOKEN, and only tokens a citation tail can
' contain are taken (IsCiteTailToken): volume/page numbers, reporter
' abbreviations, and the pincite connectives. The bracket characters alone
' used to be the stop, and an IN-TEXT cite has none before the next
' parenthetical -- so its roman run swept through the prose that followed, and
' when a LATER citation sat in that prose ("In Doe v. Roe (2000) 20 Cal.4th
' 100, 110, and later in Bruesewitz v. Wyeth LLC (2011) 562 U.S. 223, 242,
' ..."), Doe's tail marked Bruesewitz's case name ROMAN and pushed nextScan
' past its " v. " anchor, so the name never got its italics back. A supra
' cite's tail did the same one loop later, overwriting italics the full-cite
' pass had just applied. The first prose word fails every token test, which
' stops the tail at the citation's true end.
Private Function CiteTailEnd(ByVal s As String, ByVal ts As Long) As Long
    Dim i As Long
    Dim closePos As Long: closePos = ts

    If Mid$(s, ts, 1) = "(" Then
        closePos = 0
        For i = ts + 1 To Len(s)
            If Mid$(s, i, 1) = ")" Then
                closePos = i
                Exit For
            End If
        Next i
        If closePos = 0 Then Exit Function
    End If

    Dim e As Long: e = closePos
    Dim c As String
    Dim tokEnd As Long
    i = closePos + 1
    Do While i <= Len(s)
        Do While i <= Len(s)
            If Mid$(s, i, 1) = " " Then i = i + 1 Else Exit Do
        Loop
        If i > Len(s) Then Exit Do
        c = Mid$(s, i, 1)
        If c = "(" Or c = "[" Or c = ")" Or c = "]" Or c = ";" Then Exit Do
        tokEnd = i
        Do While tokEnd <= Len(s)
            c = Mid$(s, tokEnd, 1)
            If c = " " Or c = "(" Or c = "[" Or c = ")" Or c = "]" Or c = ";" Then Exit Do
            tokEnd = tokEnd + 1
        Loop
        If Not IsCiteTailToken(Mid$(s, i, tokEnd - i)) Then Exit Do
        e = tokEnd - 1
        i = tokEnd
    Loop
    CiteTailEnd = e
End Function

' True when tok can sit inside a citation's roman tail: a volume, page, or pin
' range ("562", "223,", "1266-1267,"), a reporter abbreviation ("U.S.",
' "Cal.App.4th", "Rptr.", "S.Ct."), an ordinal reporter part ("2d", "4th"), or
' a pincite connective ("supra", "at", "p.", "pp.", "fn.", "fns."). The first
' word of following prose -- lowercase, or capitalized without a period --
' fails every test, and that is what ends the tail.
Private Function IsCiteTailToken(ByVal tok As String) As Boolean
    If Len(tok) = 0 Then Exit Function

    ' Bare punctuation rides along (the "," between "supra" and the reporter).
    Dim core As String: core = StripWordPunct(tok)
    If Len(core) = 0 Then
        IsCiteTailToken = True
        Exit Function
    End If

    Select Case LCase$(core)
        Case "supra", "at", "p", "pp", "fn", "fns"
            IsCiteTailToken = True
            Exit Function
    End Select

    ' Number run: digits with any of . , - and the en-dash between or after
    ' them ("223,", "1266-1267,", "242.").
    Dim i As Long, c As String
    Dim allNum As Boolean: allNum = True
    Dim hasDigit As Boolean
    For i = 1 To Len(tok)
        c = Mid$(tok, i, 1)
        If c Like "#" Then
            hasDigit = True
        ElseIf c <> "." And c <> "," And c <> "-" And c <> ChrW$(8211) Then
            allNum = False
            Exit For
        End If
    Next i
    If allNum And hasDigit Then
        IsCiteTailToken = True
        Exit Function
    End If

    ' Ordinal reporter part: a digit followed by letters ("2d", "4th").
    If Mid$(core, 1, 1) Like "#" Then
        Dim alnum As Boolean: alnum = True
        For i = 1 To Len(core)
            If Not Mid$(core, i, 1) Like "[A-Za-z0-9]" Then
                alnum = False
                Exit For
            End If
        Next i
        If alnum Then
            IsCiteTailToken = True
            Exit Function
        End If
    End If

    ' Reporter abbreviation: starts with a capital and carries a period
    ' ("U.S.", "Cal.App.4th", "L.Ed.2d"). A case name's first word has no
    ' period, so a following in-text citation stops the tail here too.
    If Mid$(tok, 1, 1) Like "[A-Z]" And InStr(1, tok, ".") > 0 Then
        IsCiteTailToken = True
    End If
End Function

' Where the case name that ends at anchor begins, or 0 when the left edge cannot
' be trusted -- and an untrusted edge means the citation is left alone, because
' guessing wide italicizes the judge's own prose.
'
' A CITATION SENTENCE answers this outright, and that is the shape nearly every
' cite in a ruling has: the "(" opens the citation, so the name starts right
' after it (past a "See" or "Accord,") and runs to the "(year)". Nothing is
' inspected word by word, which matters -- party names contain the very words a
' word filter has to distrust ("The Sports Authority", "Court Reporters, Inc.",
' "In-N-Out"), and filtering them is what leaves half a case name roman.
'
' Only an IN-TEXT cite ("... the court in Smith v. Jones (2020) ...") has no
' bracket to say where it begins, and only there does the walk below guess: a
' capitalized word joins the name, a lowercase one joins only as a CONNECTOR
' ("of", "for", "and", "ex rel.") and never becomes the start itself, so the
' start is always the leftmost capitalized word accepted. "The court in Smith v.
' Jones" starts at "Smith" -- "in" is a connector, "court" is not a name word,
' and the walk stops there.
'
' boundary comes back as the "(" or "[" the citation sits inside, or 0.
Private Function CaseNameLeftEdge(ByVal s As String, ByVal anchor As Long, _
                                   ByRef boundary As Long) As Long
    boundary = 0

    ' --- Citation sentence: the bracket IS the answer. ---
    Dim j As Long
    For j = anchor - 1 To 1 Step -1
        Dim c0 As String: c0 = Mid$(s, j, 1)
        ' A closed parenthetical, or a parallel-cite semicolon, means the bracket
        ' further left is not the one this citation opens with.
        If c0 = ")" Or c0 = "]" Or c0 = ";" Then Exit For
        If c0 = "(" Or c0 = "[" Then
            boundary = j
            CaseNameLeftEdge = SkipCiteSignals(s, j + 1, anchor)
            Exit Function
        End If
    Next j

    ' --- In-text cite: no bracket, so the edge has to be worked out. ---
    Dim best As Long: best = 0
    Dim k As Long: k = anchor - 1

    Do While k >= 1
        Do While k >= 1
            If Mid$(s, k, 1) = " " Then k = k - 1 Else Exit Do
        Loop
        If k < 1 Then Exit Do

        Dim c As String: c = Mid$(s, k, 1)
        If c = "(" Or c = "[" Or c = ";" Or IsQuoteChar(c) Then
            boundary = k
            Exit Do
        End If

        Dim ws As Long: ws = k
        Do While ws >= 1
            Dim wc As String: wc = Mid$(s, ws, 1)
            If wc = " " Or wc = "(" Or wc = "[" Then Exit Do
            ws = ws - 1
        Loop
        ws = ws + 1

        Dim w As String: w = Mid$(s, ws, k - ws + 1)
        Dim f As String: f = Left$(w, 1)

        If f >= "A" And f <= "Z" Then
            If IsSentenceLeadWord(w) And Not OpensInRe(s, ws) Then Exit Do
            best = ws
        ElseIf f >= "a" And f <= "z" Then
            If Not IsNameConnector(w) Then Exit Do
        ElseIf f = "&" Then
            ' joins the name, never starts it
        Else
            Exit Do
        End If

        k = ws - 1
    Loop

    CaseNameLeftEdge = best
End Function

' The first character of the case name inside a citation sentence: past the
' spaces, and past a leading signal -- "See", "See also", "Accord,", "But see",
' "Cf.", "e.g.," -- which is roman and no part of the name. Reuses the module's
' own signal-word list. "In" is deliberately not one: "In re Marriage of ..." is
' a case name. Falls back to fromPos if the whole span reads as signals, so a
' misread can only leave the name where it was.
Private Function SkipCiteSignals(ByVal s As String, ByVal fromPos As Long, _
                                  ByVal limit As Long) As Long
    Dim i As Long: i = fromPos
    Dim guard As Long

    Do While i < limit And guard < 6
        guard = guard + 1
        Do While i < limit
            If Mid$(s, i, 1) = " " Then i = i + 1 Else Exit Do
        Loop
        Dim e As Long: e = i
        Do While e < limit
            If Mid$(s, e, 1) = " " Then Exit Do
            e = e + 1
        Loop
        If e <= i Then Exit Do
        If Not IsCiteSignalWord(LCase$(StripWordPunct(Mid$(s, i, e - i)))) Then Exit Do
        i = e
    Loop

    Do While i < limit
        If Mid$(s, i, 1) = " " Then i = i + 1 Else Exit Do
    Loop
    If i >= limit Then i = fromPos
    SkipCiteSignals = i
End Function

' "In re ..." -- the one capitalized sentence word that does open a case name.
Private Function OpensInRe(ByVal s As String, ByVal ws As Long) As Boolean
    OpensInRe = (LCase$(Mid$(s, ws, 6)) = "in re ")
End Function

' Capitalized words that open a sentence or a signal, not a case name.
Private Function IsSentenceLeadWord(ByVal w As String) As Boolean
    Select Case LCase$(StripWordPunct(w))
        Case "see", "accord", "compare", "citing", "quoting", "but", "and", "also", _
             "cf", "eg", "ie", "the", "in", "under", "however", "here", "although", _
             "because", "while", "when", "where", "whether", "if", "as", "at", _
             "since", "thus", "nor", "following", "italics", "emphasis", "internal", _
             "original", "id", "ibid", "court", "plaintiff", "defendant", "he", _
             "she", "they", "it", "this", "that", "these", "those", "there"
            IsSentenceLeadWord = True
    End Select
End Function

' Lowercase words that sit INSIDE a case name.
Private Function IsNameConnector(ByVal w As String) As Boolean
    Select Case LCase$(StripWordPunct(w))
        Case "of", "for", "and", "the", "de", "del", "la", "las", "los", "van", _
             "von", "der", "den", "da", "du", "ex", "rel", "et", "al", "in", "re", _
             "dba", "aka", "fka", "nka", "v"
            IsNameConnector = True
    End Select
End Function

' A word without its surrounding punctuation, so "Corp.," compares as "corp".
Private Function StripWordPunct(ByVal w As String) As String
    Dim a As Long: a = 1
    Dim b As Long: b = Len(w)
    Do While a <= b
        If Mid$(w, a, 1) Like "[A-Za-z0-9]" Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Mid$(w, b, 1) Like "[A-Za-z0-9]" Then Exit Do
        b = b - 1
    Loop
    If b < a Then Exit Function
    StripWordPunct = Mid$(w, a, b - a + 1)
End Function

' The short-name parenthetical starting at or just after fromPos, as indices into
' s. False when what is there is any other kind of parenthetical.
Private Function ShortNameAfter(ByVal s As String, ByVal fromPos As Long, _
                                 ByVal caseName As String, _
                                 ByRef a As Long, ByRef b As Long) As Boolean
    Dim i As Long: i = fromPos
    If i < 1 Then i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) = " " Then i = i + 1 Else Exit Do
    Loop
    If i > Len(s) Then Exit Function
    If Mid$(s, i, 1) <> "(" Then Exit Function

    Dim cl As Long: cl = InStr(i + 1, s, ")")
    If cl = 0 Then Exit Function

    ShortNameAfter = ShortNameSpan(s, i, cl, caseName, a, b)
End Function


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
    On Error Resume Next          ' the toggle asks this first: never fail here
    Dim i As Long
    For i = 1 To doc.Hyperlinks.Count
        If Left$(doc.Hyperlinks(i).ScreenTip, Len(SCREENTIP_PREFIX)) = SCREENTIP_PREFIX Then
            HasCitationLinks = True
            Exit Function
        End If
    Next i
End Function


' strayLeft reports links whose display text was a URL that Word would not let us
' delete -- see the repair below. The field is gone either way; the text is not.
Private Function RemoveCitationLinks_Quiet(ByVal doc As Document, _
                                           Optional ByRef strayLeft As Long) As Long
    Dim removed As Long
    Dim i As Long, rng As Range
    Application.ScreenUpdating = False

    ' Per-link isolation, and never a raw error out of here. This runs from the
    ' toggle and from the top of AddCitationLinks, in BOTH cases before the caller
    ' has armed a handler -- so anything raised here reached the user as a bare
    ' runtime-error dialog and left ScreenUpdating off with it. One link Word
    ' won't let us touch must not cost the user the other forty.
    On Error Resume Next
    For i = doc.Hyperlinks.Count To 1 Step -1
        Dim tip As String
        tip = ""
        tip = doc.Hyperlinks(i).ScreenTip
        If Left$(tip, Len(SCREENTIP_PREFIX)) = SCREENTIP_PREFIX Then
            ' A link of ours whose display text IS its own address is not a
            ' citation link -- it is the damage AddLink now refuses to do (Word
            ' writing the search URL into the prose when handed an empty anchor).
            ' A real citation link displays the citation, never a URL, so that
            ' text goes with the field, which repairs a document linked before
            ' that guard existed. Otherwise the toggle would unlink the URL and
            ' leave it sitting in the sentence.
            Dim addr As String
            addr = ""
            addr = doc.Hyperlinks(i).Address
            Dim stray As Boolean
            stray = (Len(addr) > 0)
            If stray Then
                stray = (StrComp(doc.Hyperlinks(i).Range.text, addr, vbTextCompare) = 0)
            End If

            Set rng = Nothing
            Set rng = doc.Hyperlinks(i).Range
            Err.Clear
            doc.Hyperlinks(i).Delete

            If Err.Number = 0 Then
                removed = removed + 1

                Dim wiped As Boolean
                wiped = False
                If stray Then
                    ' Re-read the text before deleting any of it. Unlinking moves
                    ' text around, and a range that no longer holds the URL is a
                    ' range that now holds the user's prose -- deleting on faith
                    ' would eat a sentence. Word also refuses some deletions
                    ' outright (4198 "Command failed" -- a document carrying
                    ' tracked changes is the case seen in practice), so this is a
                    ' try, not an assumption: on refusal the text stays and the
                    ' link formatting is cleared like any other removal.
                    If StrComp(rng.text, addr, vbTextCompare) = 0 Then
                        Err.Clear
                        rng.Delete
                        wiped = (Err.Number = 0)
                        Err.Clear
                    End If
                    If Not wiped Then strayLeft = strayLeft + 1
                End If
                If Not wiped Then ResetLinkFormatting rng
            End If
            Err.Clear
        End If
    Next i
    Err.Clear
    On Error GoTo 0

    Application.ScreenUpdating = True
    RemoveCitationLinks_Quiet = removed
End Function


Private Function AddLink(ByVal rng As Range, ByVal url As String, ByVal typ As String, _
                         Optional ByVal noItalics As Boolean = False) As Boolean
    On Error GoTo Fail

    ' First, cut off any tail of the PREVIOUS SENTENCE the span opened with; then
    ' trim the punctuation left in front of the case name.
    On Error Resume Next
    TrimLeadingProse rng

    ' Never start a citation link at the citation sentence's outer "(" (or a
    ' leading "[", quote, or space). The literal-text fallback in particular can
    ' hand us a range that begins with "(" -- e.g. "(Commodore Home Systems,
    ' Inc. v. Superior Court ...". Trim any such leading characters off the
    ' anchor so the hyperlink begins at the case name.
    '
    ' CLOSING quotes belong in this set as much as opening ones. A citation that
    ' follows a quotation is preceded by "..." and a span that starts one
    ' character early opens on that closing quote, which used to stop the trim
    ' dead and leave the quote mark inside the link.
    Do While rng.Characters.count > 1
        Dim fch As String: fch = rng.Characters(1).text
        If fch = "(" Or fch = "[" Or fch = " " Or fch = Chr$(160) _
           Or fch = ChrW$(8220) Or fch = ChrW$(8216) Or fch = Chr$(34) _
           Or fch = ChrW$(8221) Or fch = ChrW$(8217) Or fch = "'" Then
            rng.MoveStart wdCharacter, 1
        Else
            Exit Do
        End If
    Loop

    ' Then push the anchor back OUT to the left, over any of the case name the
    ' extractor's span opened inside of -- see ExtendAnchorToCaseName. A rule
    ' reference carries no case name, so the flag that skips its italics skips
    ' this too.
    If Not noItalics Then ExtendAnchorToCaseName rng
    On Error GoTo Fail

    ' An anchor with no text is the one input Word answers by writing the URL
    ' into the document as literal text. Hyperlinks.Add reads a collapsed range
    ' as "put a link HERE", and with no anchor text to display it uses the
    ' address as the display text -- which is how a paragraph came back reading
    ' "https://plus.lexis.com/search/?...%20450Saxon reserved whether ...", the
    ' whole search URL pasted in front of the word it was supposed to link.
    ' A caller holding an empty range has already failed to find its citation, so
    ' there is nothing here to link: report the miss instead of writing to the
    ' document. This is the only place in the module that adds a link, so the
    ' check covers the offset path and both literal-text fallbacks at once.
    If rng Is Nothing Then Exit Function
    If rng.start >= rng.End Then Exit Function
    Dim anchorText As String
    anchorText = rng.text
    If Len(anchorText) = 0 Then Exit Function

    ' Anchoring inside an existing hyperlink is the other way the address ends up
    ' as text: nesting a field in a field is not something Word does, and what it
    ' does instead is unhelpful. It is also simply wrong -- the span is already
    ' linked. LinkTextIfUnlinked makes this test before calling; the literal-text
    ' fallback (FindAndLink) did not, and could land on a span an overlapping row
    ' had just linked.
    If rng.Hyperlinks.count > 0 Then Exit Function

    Dim h As Hyperlink
    Set h = ActiveDocument.Hyperlinks.Add(Anchor:=rng, Address:=url, _
        ScreenTip:=Left$(SCREENTIP_PREFIX & typ & " | " & url, 255))

    ' Post-condition, because the two guards above are a diagnosis and the damage
    ' this prevents is a URL sitting in the middle of a judge's prose: a citation
    ' link must display the words it was anchored to, never the address. If Word
    ' wrote the address anyway, take it straight back out -- drop the field first
    ' so the delete removes text and not a half-dismantled field -- and report the
    ' miss. An unlinked citation is a nuisance; a pasted search URL is a defect
    ' the user has to find and clean up by hand.
    If StrComp(h.Range.text, url, vbTextCompare) = 0 And _
       StrComp(anchorText, url, vbTextCompare) <> 0 Then
        Dim stray As Range
        Set stray = h.Range.Duplicate
        h.Delete                     ' unlink, leaving the text Word inserted
        stray.Delete                 ' and remove that text
        Exit Function                ' AddLink stays False
    End If

    ' Word's Hyperlink style drops the case-name italic. Rather than try to
    ' preserve the prior formatting through the field boundary (fragile --
    ' anything applied to the first display character gets absorbed), re-derive
    ' the italic from citation structure: in a case cite the case name is
    ' everything to the left of the "(year)" date, or of ", supra".
    ' Skipped for a rule reference: "rule 3.1350(f)" carries no case name, and
    ' the logic below is written to read a case cite, not a rule number.
    If Not noItalics Then ItalicizeCaseName h.Range

    AddLink = True
    Exit Function
Fail:
    AddLink = False
End Function

' Cut a leading run of the PREVIOUS SENTENCE off a citation anchor.
'
' Observed: a link that covered "English." (" as well as the case name --
'
'     ... nor explained to respondents who did not understand written
'     English." (Penilla v. Westmont Corp. (2016) 3 Cal.App.5th 205, 209.)
'
' -- so the judge's own quoted sentence ended up inside a Lexis link. The span
' arrives that way; the character trim in AddLink only strips punctuation, so a
' span opening on a WORD defeated it entirely.
'
' The whole span is examined, not a window of it. It used to be the first 60
' characters, which is enough for the sentence fragment above and nothing like
' enough for a BLOCK QUOTATION -- the shape a ruling uses constantly:
'
'     "It is the general rule that in actions for the conversion of personal
'     property ... it is sufficient to declare generally, that the property was
'     wrongfully converted." (Wendling Lumber Co. v. Glenwood Lumber Co. (1908)
'     153 Cal. 411, 414.)
'
' There the boundary sits some 250 characters into the span, the trim never saw
' it, and the whole quotation went into the link -- and then into the italics,
' since ItalicizeCaseName reads everything ahead of the "(year)" as the case
' name. The quotation came back italic from its first word to the case name's
' last.
'
' See ProseRunEnd for what counts as the boundary and which one wins.
'
' The prose that gets cut is also taken back OUT of italic when it is entirely
' italic -- which is the signature of a document linked by the build that had
' the window, and nothing a judge writes. (A quotation carrying emphasis on some
' of its words reads as mixed, not italic, and is left exactly as it is.) The
' case-name italics inside that prose are not lost: NormalizeCitationItalics
' runs after every link pass and re-derives them from the paragraph text.
Private Sub TrimLeadingProse(ByVal rng As Range)
    Dim s As String
    s = rng.text
    If Len(s) < 2 Then Exit Sub

    Dim cut As Long
    cut = ProseRunEnd(s)
    If cut = 0 Or cut >= Len(s) Then Exit Sub

    ' Repair pass, before the range gives the prose up. SubRangeByChars is the
    ' only safe way to turn a text index back into a range here -- the span can
    ' already contain a link from an earlier row, and .Start counts positions
    ' .Text never returns.
    Dim dropped As Range
    Set dropped = SubRangeByChars(rng, 1, cut)
    If Not dropped Is Nothing Then
        If dropped.Font.Italic = True Then dropped.Font.Italic = False
    End If

    rng.MoveStart wdCharacter, cut
End Sub

' The last character of the PROSE that a citation span opened inside of, or 0
' when the span opens on the citation itself.
'
' The boundary is sentence-ending punctuation, a CLOSING QUOTE, and then the "("
' or "[" that opens the citation sentence: `English." (Penilla`. All three are
' required. The quote alone would misfire on a citation carrying a quoted
' parenthetical -- `Smith v. Jones (2020) 1 Cal.5th 1 ["no."]` ends in the same
' two characters -- and it is the opening bracket after it that says a new
' citation sentence starts here rather than a parenthetical closing. The period
' alone would be worse still: every "Lumber Co. (1908)" in the language ends a
' case name with one.
'
' The LAST boundary wins, so a span that swallowed two sentences still lands on
' the citation -- but only among the boundaries a CITATION still follows
' (CiteFollows). That is what keeps the walk from cutting into the citation
' itself on the one shape where a boundary appears inside it, a quotation closed
' inside a parenthetical: `... 209 ["a quote." (Italics added.)]`. With no such
' boundary anywhere the last one found is used regardless, which is how a
' statute or rule span -- no case name to test for -- still gets its prose cut.
Private Function ProseRunEnd(ByVal s As String) As Long
    If Len(s) < 3 Then Exit Function

    Dim last As Long, lastCite As Long
    Dim i As Long
    For i = 1 To Len(s) - 2
        Dim c As String: c = Mid$(s, i, 1)
        If c = "." Or c = "!" Or c = "?" Then
            Dim nx As String: nx = Mid$(s, i + 1, 1)
            If nx = ChrW$(8221) Or nx = ChrW$(8217) Or nx = Chr$(34) Or nx = "'" Then
                If OpensCitation(s, i + 2) Then
                    last = i + 1                       ' through the quote
                    If CiteFollows(s, i + 2) Then lastCite = last
                End If
            End If
        End If
    Next i

    If lastCite > 0 Then ProseRunEnd = lastCite Else ProseRunEnd = last
End Function

' True when what follows position k still reads as an authority -- parties, a
' "(year)", a ", supra", or a section sign. Used to tell the boundary that opens
' the citation from one that merely sits inside it.
Private Function CiteFollows(ByVal s As String, ByVal k As Long) As Boolean
    Dim rest As String
    rest = Mid$(s, k)
    If Len(rest) = 0 Then Exit Function

    If InStr(1, rest, " v. ", vbTextCompare) > 0 Then
        CiteFollows = True
    ElseIf InStr(1, rest, ", supra", vbTextCompare) > 0 Then
        CiteFollows = True
    ElseIf FindYearParen(rest) > 0 Then
        CiteFollows = True
    ElseIf InStr(rest, ChrW$(167)) > 0 Then
        CiteFollows = True
    End If
End Function

' True when, from position k, the text opens a citation sentence: any run of
' spaces and then "(" or "[". Anything else -- more prose, a closing bracket, the
' end of the span -- means the quote just ended a parenthetical, not a sentence.
Private Function OpensCitation(ByVal s As String, ByVal k As Long) As Boolean
    Dim i As Long: i = k
    Do While i <= Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        If c = " " Or c = Chr$(160) Then
            i = i + 1
        Else
            OpensCitation = (c = "(" Or c = "[")
            Exit Function
        End If
    Loop
End Function

' Push a citation anchor back to the START of the case name.
'
' Where a full cite's span begins is the extractor's call, and two shapes defeat
' it -- both leaving the link, and with it the italics, opening in the middle of
' the case name:
'
'   NBC Subsidiary (KNBC-TV), Inc. v. Superior Court (1999) 20 Cal.4th 1178
'                             ^ the span opened here: the parenthetical ended
'                               the walk back through the parties
'
'   Board of Trustees v. Superior Court (2007) 149 Cal.App.4th 1154
'            ^ and here: a lowercase word ended it
'
' So the anchor is walked left again from wherever it arrived, a word at a time,
' and the link is widened to the leftmost CAPITALIZED word the walk accepts.
' What it accepts is deliberately narrow, because guessing wide pulls the
' judge's own prose into a Lexis link. See CaseNameStartBefore.
'
' Only a cite with PARTIES is touched -- " v. " somewhere in the anchor. A
' statute has no case name to widen to, and its span sits in running prose where
' the same walk would happily eat "the California Constitution and" out of
' "under the California Constitution and Code of Civil Procedure section 437c".
Private Sub ExtendAnchorToCaseName(ByVal rng As Range)
    Const MAX_BACK As Long = 140

    On Error Resume Next
    If rng Is Nothing Then Exit Sub
    If InStr(1, rng.text, " v. ", vbTextCompare) = 0 Then Exit Sub

    Dim para As Range
    Set para = rng.Paragraphs(1).Range
    If para Is Nothing Then Exit Sub
    If rng.start <= para.start Then Exit Sub

    ' The text ahead of the anchor, inside its own paragraph. Character
    ' positions have to run 1:1 with that string for MoveStart to land where the
    ' walk says they will; a field or a footnote reference in front of the anchor
    ' breaks that, and the length test is what catches it.
    Dim before As Range
    Set before = rng.Duplicate
    before.SetRange para.start, rng.start
    Dim s As String
    s = before.text
    If Len(s) = 0 Then Exit Sub
    If Len(s) <> rng.start - para.start Then Exit Sub

    Dim best As Long
    best = CaseNameStartBefore(s)
    If best = 0 Then Exit Sub

    Dim back As Long
    back = Len(s) - best + 1
    If back <= 0 Or back > MAX_BACK Then Exit Sub

    rng.MoveStart wdCharacter, -back

    ' Never widen INTO an existing link. Word does not nest one field inside
    ' another, and AddLink refuses an anchor that overlaps one -- so a widening
    ' that reaches a link already there would cost the whole citation its own.
    If rng.Hyperlinks.count > 0 Then rng.MoveStart wdCharacter, back
End Sub


' The 1-based index, in the text running up to a citation anchor, where the case
' name that anchor opens inside of begins -- or 0 when the walk accepted no word
' and the anchor should be left where the extractor put it.
'
' What the walk crosses, right to left:
'
'   - a capitalized word, unless it is one of the words that open a sentence or
'     a signal rather than a case name ("See", "The", "Here", "Court"). "In re"
'     is the exception, since that does open one;
'   - a lowercase word ONLY as a joiner -- "of", "for", "the", "de", "ex rel."
'     -- which joins a name and never starts it. See IsCaseNameJoiner for the
'     two words a case name does use that this walk still will not cross;
'   - a short parenthetical carrying no sentence punctuation, which is what
'     "(KNBC-TV)" is and what the ")" closing a PREVIOUS citation is not.
'
' Anything else -- a bracket, a semicolon, a quote, a lowercase word that is no
' joiner -- ends the walk where it stands. Since only a capitalized word ever
' becomes the answer, a walk that ends on a joiner gives back the capitalized
' word to its right, not the joiner.
Private Function CaseNameStartBefore(ByVal s As String) As Long
    Dim best As Long
    Dim k As Long: k = Len(s)
    Dim guard As Long

    Do While k >= 1 And guard < 16
        guard = guard + 1

        ' Spaces, and the comma that separates the parts of a party's name.
        Do While k >= 1
            Dim sc As String: sc = Mid$(s, k, 1)
            If sc = " " Or sc = Chr$(160) Or sc = "," Then k = k - 1 Else Exit Do
        Loop
        If k < 1 Then Exit Do

        Dim c As String: c = Mid$(s, k, 1)

        If c = ")" Then
            Dim op As Long: op = MatchingOpenParen(s, k)
            If op = 0 Then Exit Do
            If Not IsNamePartParen(Mid$(s, op + 1, k - op - 1)) Then Exit Do
            k = op - 1
        ElseIf c = "(" Or c = "[" Or c = "]" Or c = ";" Or c = ":" Or IsQuoteChar(c) Then
            Exit Do
        Else
            Dim ws As Long: ws = k
            Do While ws >= 1
                Dim wc As String: wc = Mid$(s, ws, 1)
                If wc = " " Or wc = Chr$(160) Or wc = "(" Or wc = "[" Or wc = ")" Then Exit Do
                ws = ws - 1
            Loop
            ws = ws + 1

            Dim w As String: w = Mid$(s, ws, k - ws + 1)
            Dim f As String: f = Left$(w, 1)

            If f >= "A" And f <= "Z" Then
                If IsSentenceLeadWord(w) And Not OpensInRe(s, ws) Then Exit Do
                best = ws
            ElseIf f >= "a" And f <= "z" Then
                If Not IsCaseNameJoiner(w) Then Exit Do
            ElseIf f = "&" Then
                ' Joins two parties ("Smith & Jones"); never opens the name.
            Else
                Exit Do
            End If
            k = ws - 1
        End If
    Loop

    CaseNameStartBefore = best
End Function


' The lowercase words this walk will cross to reach the rest of a case name.
'
' Deliberately NARROWER than IsNameConnector, which the italics pass uses. Two
' of its words cannot be crossed from a standing start in the middle of a
' sentence:
'
'   "and" -- "under the California Constitution and Smith v. Jones (2020)"
'   would give up "California Constitution and" to the link, and a case name
'   that really does carry an "and" ("Fair Employment and Housing v. Lucent")
'   is the rarer of the two by a wide margin;
'
'   "in"  -- "the Court of Appeal in Ochoa v. Superior Court" would give up
'   "Appeal in". "In re" is unaffected: its "In" is capitalized, and the
'   capitalized branch of the walk reads it through OpensInRe.
'
' The cost is a link that stops short on a case name built around one of those
' two words, which is the state the reference arrived in anyway.
Private Function IsCaseNameJoiner(ByVal w As String) As Boolean
    Select Case LCase$(StripWordPunct(w))
        Case "of", "for", "the", "de", "del", "la", "las", "los", "van", _
             "von", "der", "den", "da", "du", "ex", "rel", "et", "al", "re", _
             "dba", "aka", "fka", "nka"
            IsCaseNameJoiner = True
    End Select
End Function


' The position of the "(" opening the group that closes at posClose, or 0 when
' the group does not open inside this text.
Private Function MatchingOpenParen(ByVal s As String, ByVal posClose As Long) As Long
    Dim depth As Long
    Dim i As Long
    For i = posClose To 1 Step -1
        Dim c As String: c = Mid$(s, i, 1)
        If c = ")" Then
            depth = depth + 1
        ElseIf c = "(" Then
            depth = depth - 1
            If depth = 0 Then
                MatchingOpenParen = i
                Exit Function
            End If
        End If
    Next i
End Function


' True when a parenthetical sits INSIDE a case name -- "(KNBC-TV)", "(USA)" --
' rather than closing something the case name has nothing to do with. It has to
' be short, open on a letter or a digit, and carry none of the punctuation that
' ends a sentence or a citation. That last test does the work: it is what stops
' the walk at the ")" closing the citation before this one -- "... 20 Cal.4th
' 1178, 1212.)" -- instead of swallowing that citation whole.
Private Function IsNamePartParen(ByVal inner As String) As Boolean
    If Len(inner) = 0 Or Len(inner) > 40 Then Exit Function
    If Not Left$(inner, 1) Like "[A-Za-z0-9]" Then Exit Function

    Dim i As Long
    For i = 1 To Len(inner)
        Dim c As String: c = Mid$(inner, i, 1)
        If c = "." Or c = "!" Or c = "?" Or c = ";" Or c = ":" Then Exit Function
        If IsQuoteChar(c) Then Exit Function
    Next i

    IsNamePartParen = True
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

    ' A display that still carries the tail of the sentence BEFORE the citation
    ' -- the span opened there and TrimLeadingProse could not cut it -- must not
    ' hand that prose to the case name. Everything ahead of the "(year)" reads as
    ' the name otherwise, so a block quotation the span swallowed came back
    ' italic from its first word ("It is the general rule that ... the property
    ' was wrongfully converted." (Wendling Lumber Co. v. Glenwood Lumber Co.
    ' (1908) 153 Cal. 411, 414.)). The same boundary the trim looks for is looked
    ' for here, inside the part of the display that reads as the name, and the
    ' name starts after it. Belt and braces: the trim is what keeps that prose
    ' out of the LINK; this is what keeps it out of the ITALICS whatever else
    ' put it there.
    Dim proseEnd As Long
    proseEnd = ProseRunEnd(Left$(s, tailStart - 1))

    ' First letter of the case name: skip the prose above if there was any, then
    ' a leading outer "(", quote, or space, then any lowercase signal words
    ' ("see", "cf.", "see also"). A case short name always starts with a capital.
    Dim nameStart As Long: nameStart = 1
    If proseEnd > 0 And proseEnd < tailStart Then nameStart = proseEnd + 1
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
    ' Where the prose above was found, the slate starts at the citation instead:
    ' clearing across prose the link should never have covered would flatten the
    ' judge's own emphasis along with everything else.
    Dim clearFrom As Long: clearFrom = 1
    If proseEnd > 0 And proseEnd < m Then clearFrom = proseEnd + 1
    If clearFrom <= 1 Then
        ActiveDocument.Range(disp.Characters(1).start - 1, disp.Characters(m).End).Font.Italic = False
    Else
        ActiveDocument.Range(disp.Characters(clearFrom).start, disp.Characters(m).End).Font.Italic = False
    End If

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
    ' Clearing italic on just that one preceding character takes the stray
    ' italic off the "(" without disturbing the case name.
    '
    ' ONLY for a bracket, though. That undo ends AT the field boundary, and a
    ' range ending there can reach the first display character as surely as the
    ' back-extension reached it -- which is how a link opening on the case name
    ' itself, "<link>Inc. v. Superior Court (1999) ...", came back with its
    ' first letter roman in a document where the name around it was italic. A
    ' bracket left italic is visible and worth that risk. The space or comma an
    ' in-text cite follows is not: italic on a space shows nothing, so there is
    ' nothing there to trade the case name's first letter for.
    If extendedBack Then
        Dim leadCh As String
        leadCh = ActiveDocument.Range(startPos, disp.Characters(1).start).text
        If leadCh = "(" Or leadCh = "[" Then
            ActiveDocument.Range(startPos, disp.Characters(1).start).Font.Italic = False
        End If
    End If

    ' When the whole supra cite is one link ("Galleria Plus, Inc., supra, 179
    ' Cal.App.4th at p. 538"), the case-name run above ends at ", supra"; also
    ' italicize the "supra" word so only the reporter stays roman. No-op when
    ' there is no ", supra" in the display.
    ItalicizeSupraWordInDisplay disp

    ' And the parenthetical that introduces the case's short name, which the
    ' clean slate above would otherwise leave roman.
    ItalicizeShortNameParen disp, Mid$(s, nameStart, nameEnd - nameStart + 1), tailStart
End Sub

' Italicize the parenthetical that introduces a case's SHORT NAME. The short name
' is part of the case name and carries its italics:
'
'   (Center ... v. Lennar Corp. (2009) 173 Cal.App.4th 1543, 1554
'                                        (Center for Self-Improvement).)
'                                         ^^^^^^^^^^^^^^^^^^^^^^^^^^^ this
'
' Needed because the clean slate in ItalicizeCaseName clears italic across the
' whole display: without this the short name comes back roman on every press,
' even when it was typed correctly to begin with.
'
' A parenthetical qualifies only when it opens with a capital AND its first word
' is one of the case name's own words. That is what separates a short name from
' the parentheticals that sit in exactly the same place and are NOT italic --
' "(conc. opn. of Werdegar, J.)", "(en banc)", "(italics added)", and the
' "(2009)" date itself.
'
' Looked for inside the link first, then through the rest of the paragraph after
' it: the linker usually ends the display at the pincite, which leaves the short
' name in the plain text beyond the field.
Private Sub ItalicizeShortNameParen(ByVal disp As Range, ByVal caseName As String, _
                                     ByVal tailStart As Long)
    On Error Resume Next
    If Len(Trim$(caseName)) = 0 Then Exit Sub

    If MarkShortNameParen(disp, caseName, tailStart) Then Exit Sub

    Dim m As Long: m = disp.Characters.count
    If m < 1 Then Exit Sub
    Dim after As Range
    Set after = disp.Duplicate
    after.SetRange disp.Characters(m).End, disp.Paragraphs(1).Range.End
    If after Is Nothing Then Exit Sub
    MarkShortNameParen after, caseName, 1
End Sub

' Scan one range's text from startAt for the short-name parenthetical, and
' italicize it. True when one was found. Only the first few parentheticals are
' examined -- the short name follows its own citation immediately, so anything
' further along belongs to the next sentence.
Private Function MarkShortNameParen(ByVal scope As Range, ByVal caseName As String, _
                                     ByVal startAt As Long) As Boolean
    On Error Resume Next
    Dim s As String: s = scope.text
    If Len(s) = 0 Then Exit Function
    If startAt < 1 Then startAt = 1

    Dim tries As Long
    Dim op As Long: op = InStr(startAt, s, "(")
    Do While op > 0 And tries < 4
        tries = tries + 1
        Dim cl As Long: cl = InStr(op + 1, s, ")")
        If cl = 0 Then Exit Function
        Dim a As Long, b As Long
        If ShortNameSpan(s, op, cl, caseName, a, b) Then
            Dim r As Range
            Set r = SubRangeByChars(scope, a, b)
            If r Is Nothing Then Exit Function
            If r.Font.Italic <> True Then r.Font.Italic = True
            MarkShortNameParen = True
            Exit Function
        End If
        op = InStr(cl + 1, s, "(")
    Loop
End Function

' The run inside "(...)" that is a case short name, as indices into s. False when
' the parenthetical is anything else.
Private Function ShortNameSpan(ByVal s As String, ByVal op As Long, ByVal cl As Long, _
                                ByVal caseName As String, _
                                ByRef a As Long, ByRef b As Long) As Boolean
    Dim ns As Long: ns = op + 1
    Dim ne As Long: ne = cl - 1
    If ne < ns Then Exit Function

    ' "(hereafter Lennar)" / "(hereinafter, Lennar)": the signal word stays roman.
    Dim lead As Variant, v As Variant
    lead = Array("hereafter", "hereinafter")
    For Each v In lead
        If LCase$(Mid$(s, ns, Len(v))) = v Then
            ns = ns + Len(v)
            Do While ns <= ne
                If Mid$(s, ns, 1) = " " Or Mid$(s, ns, 1) = "," Then ns = ns + 1 Else Exit Do
            Loop
            Exit For
        End If
    Next v

    ' Quotation marks around the short name are not part of it.
    Do While ns <= ne
        If IsQuoteChar(Mid$(s, ns, 1)) Then ns = ns + 1 Else Exit Do
    Loop
    Do While ne >= ns
        If IsQuoteChar(Mid$(s, ne, 1)) Or Mid$(s, ne, 1) = " " Then ne = ne - 1 Else Exit Do
    Loop
    If ne < ns Then Exit Function

    Dim f As String: f = Mid$(s, ns, 1)
    If Not (f >= "A" And f <= "Z") Then Exit Function
    If Not WordAppearsIn(FirstWordOf(Mid$(s, ns, ne - ns + 1)), caseName) Then Exit Function

    a = ns
    b = ne
    ShortNameSpan = True
End Function

Private Function IsQuoteChar(ByVal c As String) As Boolean
    Select Case c
        Case Chr$(34), "'", ChrW$(8216), ChrW$(8217), ChrW$(8220), ChrW$(8221)
            IsQuoteChar = True
    End Select
End Function

' The leading run of letters, which is all the comparison below needs.
Private Function FirstWordOf(ByVal s As String) As String
    Dim i As Long
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "[A-Za-z]") Then Exit For
    Next i
    FirstWordOf = Left$(s, i - 1)
End Function

' True when w is one of text's words, compared on letters alone so a hyphen,
' ampersand or period between them changes nothing ("Self-Improvement" is two
' words either way, and "Corp." matches "Corp").
Private Function WordAppearsIn(ByVal w As String, ByVal text As String) As Boolean
    If Len(w) = 0 Then Exit Function
    WordAppearsIn = (InStr(1, " " & LettersOnlyLower(text) & " ", _
                           " " & LCase$(w) & " ", vbBinaryCompare) > 0)
End Function

' Lowercase the text with every non-letter reduced to a single space.
Private Function LettersOnlyLower(ByVal s As String) As String
    Dim i As Long, c As String, out As String, gap As Boolean
    gap = True
    For i = 1 To Len(s)
        c = LCase$(Mid$(s, i, 1))
        If c >= "a" And c <= "z" Then
            out = out & c
            gap = False
        ElseIf Not gap Then
            out = out & " "
            gap = True
        End If
    Next i
    LettersOnlyLower = Trim$(out)
End Function

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
    ' The link's TRUE text start. disp.Start points into the field code, so using
    ' it as the window's right edge put the window in the wrong place to begin
    ' with -- see SubRangeByChars for why the two coordinate systems differ.
    If disp.Characters.count < 1 Then Exit Sub
    Dim linkStart As Long: linkStart = disp.Characters(1).start
    If linkStart < 8 Then Exit Sub

    Dim lookLen As Long: lookLen = 70
    If lookLen > linkStart Then lookLen = linkStart
    Dim base As Long: base = linkStart - lookLen
    ' Hold the RANGE, not just its text: every index into b below has to be
    ' converted back through this range's own Characters collection, never by
    ' adding it to base.
    Dim bRange As Range
    Set bRange = ActiveDocument.Range(base, linkStart)
    If bRange Is Nothing Then Exit Sub
    Dim b As String: b = bRange.text
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

    ' Convert the two indices through the window's OWN characters. The old
    ' "base + nameStart - 1" read a text index as a document position: it is
    ' short by every field-code character between base and the name, so with any
    ' link inside the window -- and the first full cite of a case is always
    ' linked -- the italic landed on a span shifted off the short name. That is
    ' what italicized "Center for" of one cite and the reporter tail of the one
    ' before it, and left the rest of the name roman.
    Dim r As Range
    Set r = SubRangeByChars(bRange, nameStart, nameEnd)
    If Not r Is Nothing Then r.Font.Italic = True
End Sub

' The range covering characters a..b of rng's TEXT, or Nothing when that run is
' not there.
'
' This is the only safe way to turn an index into rng.Text back into a range. A
' range's .Text and its .Characters agree character for character, but .Start
' and .End count positions that .Text never returns -- a hyperlink's field code
' among them. Index arithmetic on .Start is therefore wrong wherever a link
' sits, which in a citation paragraph is everywhere.
Private Function SubRangeByChars(ByVal rng As Range, ByVal a As Long, _
                                  ByVal b As Long) As Range
    On Error Resume Next
    If rng Is Nothing Then Exit Function
    If a < 1 Or b < a Then Exit Function
    Dim m As Long: m = rng.Characters.count
    If b > m Then Exit Function
    ' Built off rng itself rather than ActiveDocument.Range: SetRange keeps the
    ' result in rng's OWN story, so a caller working in a footnote doesn't get a
    ' range pointing at those positions in the body.
    Dim out As Range
    Set out = rng.Duplicate
    out.SetRange rng.Characters(a).start, rng.Characters(b).End
    Set SubRangeByChars = out
End Function

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


'==============================================================================
' CALIFORNIA RULES OF COURT
'==============================================================================

' Hyperlink every California Rules of Court reference in the body.
'
' The bridge already reads a rule cited in FULL -- "Cal. Rules of Court, rule
' 3.1350(f)" -- and links it to Lexis or Westlaw along with everything else.
' What it does not read is the bare form the rest of a ruling uses once that
' full cite has been given: "rule 3.1350(f)", "rules 3.1350, 3.1354, and
' 8.204". In a California civil ruling a rule number written that way is the
' Rules of Court, so this pass links it as one.
'
' THE PROVIDER COMES FIRST. A bare reference takes the URL the bridge gave the
' full cite of the SAME rule elsewhere in the document, so it lands where the
' spelled-out cite lands -- the same borrow LinkOrphanSupraCites makes from a
' case's full cite, and the reason a rule link follows the Lexis/Westlaw toggle
' without knowing anything about either service. Only when the document never
' cites that rule in full is there nothing to borrow, and only then does the
' link fall back to the free official page on courts.ca.gov.
'
' What counts as a rule number is the Rules of Court's own numbering: a title
' number (1 through 10), a period, the rule number, and any run of subdivision
' parentheticals -- "rule 8.204(a)(1)(B)". A rule cited without a title number
' belongs to some other body of rules ("rule 12(b)(6)") and is left alone, as
' is a title outside 1-10.
'
' THE ONE EXCEPTION IS A LOCAL RULE. "local rule 3.57" and "Local Rules, rule
' 3.57" name a superior court's own rules, which are neither the Rules of Court
' nor on that site; a "local" in either of the two words before the reference
' takes it out of this pass entirely.
'
' Runs after the bridge links are placed and skips any span already sitting
' inside a hyperlink, so a rule the bridge did catch keeps its own link.
' Best-effort throughout: any failure is swallowed rather than allowed to
' disturb the links already placed.
Private Sub LinkRulesOfCourt(ByVal doc As Document, ByRef keep() As CiteRow, _
                             ByRef added As Long)
    On Error Resume Next

    ' One rule number: title, period, number, and any subdivision tail. The
    ' subdivision class excludes whitespace and nested parens so it cannot run
    ' away into a parenthetical of prose -- "rule 3.1350 (which the parties
    ' ignored)" links the rule number alone.
    Const NUM As String = "\d{1,2}\.\d{1,4}(?:\([^()\s]{1,8}\))*"

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    ' The word, the first number, and the rest of an enumeration after it.
    ' Which of those trailing numbers are really rules is decided in
    ' LinkRuleReference, from the word and the connectors; the pattern only has
    ' to reach far enough to offer them.
    re.Pattern = "\b(rules?)\s+" & NUM & _
                 "(?:(?:,\s*and\s+|,\s*or\s+|,\s*|\s+and\s+|\s+or\s+)" & NUM & ")*"

    Dim reNum As Object
    Set reNum = CreateObject("VBScript.RegExp")
    reNum.Global = True
    reNum.Pattern = NUM

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
            ' FirstIndex is 0-based; PrecededByLocal reads 1-based positions.
            If Not PrecededByLocal(raw, mm.FirstIndex + 1) Then
                LinkRuleReference p.Range, mm.Value, _
                    (StrComp(mm.SubMatches(0), "rules", vbTextCompare) = 0), _
                    reNum, keep, added
            End If
        Next mm
NextPara:
    Next p
End Sub


' Link one rule reference inside one paragraph: "rule 3.1350(f)", or a whole
' enumeration such as "rules 3.1350, 3.1354, and 8.204".
'
' The first number is linked together with the word that introduces it, so the
' display reads "rule 3.1350"; each later number is linked on its own, since
' the connectives between them are prose and do not belong in a link. Each gets
' its own address -- an enumeration can cross titles.
'
' Applied RIGHT TO LEFT. A hyperlink field moves the text after it, so linking
' the last number first leaves the offsets of the ones still to be linked
' untouched -- the same reason the main loop applies the bridge's rows in
' reverse document order.
'
' A trailing number is only taken when the text says the enumeration is one.
' "rules" plural says it outright; after a singular "rule" the connector has to
' be "and" or "or". A bare comma after "rule 3.1350" is far more often the end
' of the clause than another rule, and "rule 3.1350, 4.2 percent of ..." must
' not link "4.2". The first number that fails this, or that no address can be
' built for, ends the enumeration -- what follows it is no longer a list of
' rules.
Private Sub LinkRuleReference(ByVal para As Range, ByVal refText As String, _
                              ByVal isPlural As Boolean, ByVal reNum As Object, _
                              ByRef keep() As CiteRow, ByRef added As Long)
    On Error Resume Next

    Dim nums As Object
    Set nums = reNum.Execute(refText)
    If nums.Count = 0 Then Exit Sub

    ' Segment bounds, 1-based and inclusive, within refText. Segment 0 opens at
    ' the word "rule"; the rest open at their own number.
    Dim segA() As Long, segB() As Long, segUrl() As String
    ReDim segA(0 To nums.Count - 1)
    ReDim segB(0 To nums.Count - 1)
    ReDim segUrl(0 To nums.Count - 1)

    Dim n As Long
    Dim k As Long
    For k = 0 To nums.Count - 1
        If k > 0 Then
            If Not isPlural Then
                ' Everything between the previous number and this one.
                Dim gap As String
                gap = Mid$(refText, segB(k - 1) + 1, nums.Item(k).FirstIndex - segB(k - 1))
                If InStr(1, gap, "and", vbTextCompare) = 0 And _
                   InStr(1, gap, "or", vbTextCompare) = 0 Then Exit For
            End If
        End If

        Dim url As String
        url = RuleUrl(nums.Item(k).Value, keep)
        If Len(url) = 0 Then Exit For

        If k = 0 Then segA(k) = 1 Else segA(k) = nums.Item(k).FirstIndex + 1
        segB(k) = nums.Item(k).FirstIndex + Len(nums.Item(k).Value)
        segUrl(k) = url
        n = n + 1
    Next k
    If n = 0 Then Exit Sub

    Dim fr As Range
    Set fr = FindUnlinked(para, refText)
    If fr Is Nothing Then Exit Sub

    For k = n - 1 To 0 Step -1
        Dim seg As Range
        Set seg = SubRangeByChars(fr, segA(k), segB(k))
        If Not seg Is Nothing Then
            If AddLink(seg, segUrl(k), "rule", True) Then added = added + 1
        End If
    Next k
End Sub


' The address for one rule number, provider first.
'
' The bridge's own link for a full cite of this rule wins, so a bare "rule
' 3.1350" opens the same Lexis or Westlaw page "Cal. Rules of Court, rule
' 3.1350(f)" opens two paragraphs above it, and the Ctrl+Shift+H provider
' toggle governs both. The free official page is the fallback: a rule the
' document never cites in full has nothing to borrow, and so does every rule in
' a document the bridge read nothing in.
'
' Subdivisions are dropped from the lookup and from the fallback address alike.
' The rule is one document either way; "(f)(2)" says where to read in it.
Private Function RuleUrl(ByVal numText As String, ByRef keep() As CiteRow) As String
    Dim bare As String
    bare = numText
    Dim paren As Long
    paren = InStr(bare, "(")
    If paren > 0 Then bare = Left$(bare, paren - 1)

    RuleUrl = ProviderUrlForRule(bare, keep)
    If Len(RuleUrl) > 0 Then Exit Function

    Dim dot As Long
    dot = InStr(bare, ".")
    If dot < 2 Or dot >= Len(bare) Then Exit Function
    RuleUrl = RuleOfCourtUrl(Left$(bare, dot - 1), Mid$(bare, dot + 1))
End Function


' The URL the bridge gave a full cite of this rule, or "" when the document
' carries none. A row qualifies when its text names a rule at all and carries
' this rule number with no digit against either end -- so "3.135" is not read
' out of "3.1350", and "3.10" not out of "13.10". Two rows disagreeing on the
' address means the number is doing more than one job in this document, and
' there is no way to choose: the fallback takes it from there.
Private Function ProviderUrlForRule(ByVal bare As String, ByRef keep() As CiteRow) As String
    On Error GoTo Fail
    Dim wantUrl As String: wantUrl = ""

    Dim i As Long
    For i = LBound(keep) To UBound(keep)
        Dim t As String: t = keep(i).txt
        If Len(t) = 0 Or Len(keep(i).url) = 0 Then GoTo NextRow
        If InStr(1, t, "rule", vbTextCompare) = 0 Then GoTo NextRow

        Dim pos As Long: pos = InStr(1, t, bare, vbTextCompare)
        Do While pos > 0
            Dim okEdges As Boolean: okEdges = True
            If pos > 1 Then okEdges = Not (Mid$(t, pos - 1, 1) Like "[0-9.]")
            If okEdges Then
                Dim after As Long: after = pos + Len(bare)
                If after <= Len(t) Then okEdges = Not (Mid$(t, after, 1) Like "[0-9]")
            End If
            If okEdges Then
                If wantUrl = "" Then
                    wantUrl = keep(i).url
                ElseIf StrComp(wantUrl, keep(i).url, vbTextCompare) <> 0 Then
                    ProviderUrlForRule = ""      ' ambiguous rule number
                    Exit Function
                End If
                Exit Do
            End If
            pos = InStr(pos + 1, t, bare, vbTextCompare)
        Loop
NextRow:
    Next i

    ProviderUrlForRule = wantUrl
    Exit Function
Fail:
    ProviderUrlForRule = ""
End Function


' The courts.ca.gov address for one rule: title "3" and number "1350" in,
' "https://courts.ca.gov/cms/rules/index/three/rule3_1350" out. The site spells
' the title out in the path and writes the rule with an underscore for its
' period. Returns "" when the title number is not a Rules of Court title, which
' is the caller's signal to leave the reference unlinked.
Private Function RuleOfCourtUrl(ByVal titleNum As String, ByVal ruleNum As String) As String
    Dim titleWord As String
    titleWord = RuleTitleWord(titleNum)
    If Len(titleWord) = 0 Then Exit Function

    RuleOfCourtUrl = "https://courts.ca.gov/cms/rules/index/" & titleWord & _
                     "/rule" & CStr(CLng(titleNum)) & "_" & ruleNum
End Function


' Title number to the word courts.ca.gov puts in the path. The Rules of Court
' run from title one to title ten; anything else is not one of them.
Private Function RuleTitleWord(ByVal titleNum As String) As String
    Select Case CLng(titleNum)
        Case 1:  RuleTitleWord = "one"
        Case 2:  RuleTitleWord = "two"
        Case 3:  RuleTitleWord = "three"
        Case 4:  RuleTitleWord = "four"
        Case 5:  RuleTitleWord = "five"
        Case 6:  RuleTitleWord = "six"
        Case 7:  RuleTitleWord = "seven"
        Case 8:  RuleTitleWord = "eight"
        Case 9:  RuleTitleWord = "nine"
        Case 10: RuleTitleWord = "ten"
    End Select
End Function


' Hyperlink every California Constitution reference in the body: "Cal. Const.,
' art. I, S 1", "California Constitution, article VI, section 10", and the
' enumerated "Cal. Const., art. I, SS 7, 15". (S stands in for the section sign
' in these comments; the code writes it as ChrW$(167) so this file stays ASCII.)
'
' The extractor does not read constitutional cites, so the reading is done here.
' The ADDRESS, though, is the extractor's own: a constitutional cite goes to the
' provider as a SEARCH for the citation itself, exactly as a statute does --
' "Cal. Const., art. I, S 1" typed into Lexis lands on that section.
'
' What a provider's search URL looks like is not written down anywhere in this
' module; the extractor builds those. So one is borrowed from a row the
' extractor produced for THIS document and this citation is swapped in for the
' row's own search terms -- see ProviderSearchUrl. What comes back is the same
' search on the same service, which is how these links follow the Ctrl+Shift+H
' provider toggle without this pass knowing anything about either service. The
' rule pass borrows from the same rows for the same reason.
'
' The fallback, for a document the extractor read nothing in, is the free
' official text on leginfo.legislature.ca.gov -- as courts.ca.gov is the rule
' pass's fallback.
'
' Runs after the extractor's links are placed and skips any span already inside
' a hyperlink, so a cite something else caught keeps the link it has.
' Best-effort throughout: a failure here must not disturb the links already down.
Private Sub LinkCalConstitution(ByVal doc As Document, ByRef keep() As CiteRow, _
                                ByRef added As Long)
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = ConstPattern()

    Dim reNum As Object
    Set reNum = CreateObject("VBScript.RegExp")
    reNum.Global = True
    reNum.Pattern = "\d+(?:\.\d+)?"

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
            LinkConstReference p.Range, mm, reNum, keep, added
        Next mm
NextPara:
    Next p
End Sub


' The one reference this pass reads, in the forms a ruling writes it:
'
'   Cal. Const., art. I, S 1          California Constitution, article VI,
'   Cal. Const. art. XIII A, S 2        section 10
'   Cal. Const., art. I, SS 7, 15     Cal. Const., art. I, S 28, subd. (b)
'
' The article number is roman or arabic, with the letter that some of them carry
' ("art. XIII A" -- Proposition 13); the section marker is the sign, singular or
' doubled, or the word spelled out or abbreviated. What follows the first
' section number is picked up only so an enumeration can be offered to
' LinkConstReference, which decides for itself which of those numbers are really
' sections. A subdivision tail ("subd. (b)(4)") is left outside the reference:
' the section is one provision either way, and the subdivision says where in it
' to read.
Private Function ConstPattern() As String
    Dim sec As String: sec = ChrW$(167)
    Dim num As String: num = "\d+(?:\.\d+)?"

    ConstPattern = "\bCal(?:if(?:ornia)?)?\.?\s*,?\s*Const(?:itution)?\.?\s*,?\s*" & _
                   "art(?:icle)?\.?\s*((?:[IVXL]+|\d{1,2})(?:\s+[A-D])?)\s*,?\s*" & _
                   "(" & sec & sec & "|" & sec & "|sections|section|secs\.|sec\.)\s*" & _
                   "(" & num & ")" & _
                   "(?:\s*,\s*(?:and\s+|or\s+)?" & num & ")*"
End Function


' Link one constitutional reference inside one paragraph.
'
' The first section number is linked together with everything that introduces it,
' so the link reads "Cal. Const., art. I, S 1"; each later number in an
' enumeration is linked on its own, the connectives between them being prose.
' Each gets its own address, since each is its own section.
'
' A later number is only taken when the marker said SECTIONS -- "SS", "sections",
' "secs.". After a singular marker a comma and a number is far more often the
' rest of the sentence than another section.
'
' Applied right to left: a hyperlink field moves the text after it, so linking
' the last number first leaves the ones still to be linked where they were.
Private Sub LinkConstReference(ByVal para As Range, ByVal mm As Object, _
                               ByVal reNum As Object, ByRef keep() As CiteRow, _
                               ByRef added As Long)
    On Error Resume Next

    Dim refText As String: refText = mm.Value
    Dim art As String: art = ConstArticleRoman(mm.SubMatches(0))
    If Len(art) = 0 Then Exit Sub

    Dim marker As String: marker = mm.SubMatches(1)
    Dim firstNum As String: firstNum = mm.SubMatches(2)

    ' Where that first section number sits inside the reference. Found through
    ' the marker rather than by searching for the number outright, so an arabic
    ' ARTICLE number is never mistaken for it ("art. 1, S 1").
    Dim mpos As Long: mpos = InStr(1, refText, marker, vbTextCompare)
    If mpos = 0 Then Exit Sub
    Dim numStart As Long
    numStart = InStr(mpos + Len(marker), refText, firstNum, vbTextCompare)
    If numStart = 0 Then Exit Sub

    ' Segment bounds, 1-based and inclusive, within refText.
    Const MAXSEG As Long = 8
    Dim segA() As Long, segB() As Long, segUrl() As String
    ReDim segA(0 To MAXSEG - 1)
    ReDim segB(0 To MAXSEG - 1)
    ReDim segUrl(0 To MAXSEG - 1)

    Dim n As Long
    Dim k As Long

    segA(0) = 1
    segB(0) = numStart + Len(firstNum) - 1
    segUrl(0) = ConstUrl(art, firstNum, keep)
    If Len(segUrl(0)) = 0 Then Exit Sub
    n = 1

    If ConstMarkerIsPlural(marker) Then
        Dim tail As String
        tail = Mid$(refText, segB(0) + 1)
        Dim ms2 As Object
        Set ms2 = reNum.Execute(tail)
        For k = 0 To ms2.Count - 1
            If n >= MAXSEG Then Exit For
            Dim u2 As String
            u2 = ConstUrl(art, ms2.Item(k).Value, keep)
            If Len(u2) = 0 Then Exit For
            segA(n) = segB(0) + ms2.Item(k).FirstIndex + 1
            segB(n) = segA(n) + Len(ms2.Item(k).Value) - 1
            segUrl(n) = u2
            n = n + 1
        Next k
    End If

    Dim fr As Range
    Set fr = FindUnlinked(para, refText)
    If fr Is Nothing Then Exit Sub

    For k = n - 1 To 0 Step -1
        Dim seg As Range
        Set seg = SubRangeByChars(fr, segA(k), segB(k))
        If Not seg Is Nothing Then
            ' noItalics: a constitutional cite carries no case name, and nothing
            ' in it is italic.
            If AddLink(seg, segUrl(k), "constitution", True) Then added = added + 1
        End If
    Next k
End Sub


' The address for one article and section, provider first.
Private Function ConstUrl(ByVal artRoman As String, ByVal secNum As String, _
                          ByRef keep() As CiteRow) As String
    Dim cite As String
    cite = "Cal. Const., art. " & artRoman & ", " & ChrW$(167) & " " & secNum

    ConstUrl = ProviderSearchUrl(cite, keep)
    If Len(ConstUrl) > 0 Then Exit Function

    ' leginfo writes the section with its trailing period and the article
    ' letter with a "+" for the space: ".../?lawCode=CONS&sectionNum=SEC.+2.1.
    ' &article=XIII+A".
    ConstUrl = "https://leginfo.legislature.ca.gov/faces/codes_displaySection.xhtml" & _
               "?lawCode=CONS&sectionNum=SEC.+" & secNum & ".&article=" & _
               Replace(artRoman, " ", "+")
End Function


' A provider search URL for arbitrary citation text, built from one the
' extractor already produced for this document.
'
' Two ways in, tried in that order against each row's URL:
'
'   1. the query string names a parameter a search puts its terms in
'      ("pdsearchterms", "query", ...) -- replace that parameter's value;
'   2. the row's OWN citation text is findable in its URL -- replace it there.
'
' The first is the reliable one, because it does not care how the extractor
' encoded the terms; the second is the safety net for a service that names its
' parameter something this does not know. Either way the rest of the URL is left
' exactly as it was, so the service, the other parameters, and the Ctrl+Shift+H
' provider toggle all carry over untouched.
'
' "" when the document has no row to copy -- which includes the paths where the
' extractor found nothing at all and keep() was never allocated, hence the trap
' around a bare LBound.
Private Function ProviderSearchUrl(ByVal term As String, ByRef keep() As CiteRow) As String
    On Error GoTo Fail

    Dim i As Long
    For i = LBound(keep) To UBound(keep)
        Dim u As String: u = keep(i).url
        If Len(u) > 0 Then
            If LooksLikeSearchUrl(u) Then
                Dim out As String
                out = SwapSearchTerms(u, term)
                If Len(out) = 0 Then out = SwapCiteText(u, keep(i).txt, term)
                If Len(out) > 0 Then
                    ProviderSearchUrl = out
                    Exit Function
                End If
            End If
        End If
    Next i

Fail:
End Function


' The URL with the value of its search-terms parameter replaced by term, or ""
' when it carries no parameter this recognises. The replacement is written in
' whatever encoding the value it replaces was written in.
Private Function SwapSearchTerms(ByVal u As String, ByVal term As String) As String
    Dim names As Variant
    names = Array("pdsearchterms", "searchterms", "searchterm", "query", "terms", "q")

    Dim lu As String: lu = LCase$(u)
    Dim qs As Long: qs = InStr(1, lu, "?")
    If qs = 0 Then Exit Function

    Dim j As Long
    For j = LBound(names) To UBound(names)
        Dim key As String: key = names(j) & "="

        Dim pos As Long: pos = InStr(qs, lu, key)
        Do While pos > 0
            ' The parameter has to open the query string or follow an "&", so a
            ' short name is never read out of the tail of a longer one.
            Dim prev As String: prev = Mid$(u, pos - 1, 1)
            If prev = "?" Or prev = "&" Then
                Dim vs As Long: vs = pos + Len(key)
                Dim ve As Long: ve = InStr(vs, u, "&")
                If ve = 0 Then ve = Len(u) + 1
                SwapSearchTerms = Left$(u, vs - 1) & _
                                  EncodeSearchTerm(term, EncodingOf(Mid$(u, vs, ve - vs))) & _
                                  Mid$(u, ve)
                Exit Function
            End If
            pos = InStr(pos + 1, lu, key)
        Loop
    Next j
End Function


' The URL with the row's own citation text -- wherever inside it that text sits,
' in whichever of the three encodings it was written in -- replaced by term. ""
' when the citation is not findable in its own URL, which is the case for a URL
' that is not a search after all.
Private Function SwapCiteText(ByVal u As String, ByVal cite As String, _
                              ByVal term As String) As String
    If Len(cite) < 6 Then Exit Function

    Dim mode As Long
    For mode = 0 To 2
        Dim enc As String: enc = EncodeSearchTerm(cite, mode)
        If Len(enc) > 0 Then
            Dim pos As Long: pos = InStr(1, u, enc, vbTextCompare)
            If pos > 0 Then
                SwapCiteText = Left$(u, pos - 1) & EncodeSearchTerm(term, mode) & _
                               Mid$(u, pos + Len(enc))
                Exit Function
            End If
        End If
    Next mode
End Function


' Which of EncodeSearchTerm's spellings an existing parameter value is written
' in: "+" for a space, or percent-encoding. Percent-encoding is the answer when
' the value settles nothing, since a URL cannot carry a raw space anyway.
Private Function EncodingOf(ByVal val As String) As Long
    If InStr(1, val, "%20") > 0 Then
        EncodingOf = 1
    ElseIf InStr(1, val, "+") > 0 Then
        EncodingOf = 2
    Else
        EncodingOf = 1
    End If
End Function


' True when a URL reads as a search rather than a link to one document: it has a
' query string, and that query string names the parameter a search puts its
' terms in. Substituting into anything else would build an address for a
' document this citation has nothing to do with.
Private Function LooksLikeSearchUrl(ByVal u As String) As Boolean
    If InStr(1, u, "?") = 0 Then Exit Function
    Dim lu As String: lu = LCase$(u)
    LooksLikeSearchUrl = (InStr(1, lu, "search") > 0) Or _
                         (InStr(1, lu, "query") > 0) Or _
                         (InStr(1, lu, "terms") > 0)
End Function


' Citation text as a search URL writes it. Mode 0 leaves it alone, mode 1
' percent-encodes it with "%20" for a space, mode 2 with "+" -- the three
' spellings a search URL uses. Non-ASCII characters (the section sign) go out as
' percent-encoded UTF-8, which is what a browser sends and what the extractor's
' own URLs carry.
Private Function EncodeSearchTerm(ByVal s As String, ByVal mode As Long) As String
    If mode = 0 Then
        EncodeSearchTerm = s
        Exit Function
    End If

    Dim out As String
    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        Dim code As Long: code = AscW(ch)
        If code < 0 Then code = code + 65536

        If ch Like "[A-Za-z0-9]" Or ch = "-" Or ch = "_" Or ch = "." Or ch = "~" Then
            out = out & ch
        ElseIf ch = " " Then
            If mode = 2 Then out = out & "+" Else out = out & "%20"
        ElseIf code < 128 Then
            out = out & "%" & Right$("0" & Hex$(code), 2)
        ElseIf code < 2048 Then
            out = out & "%" & Hex$(&HC0 Or (code \ 64)) & _
                  "%" & Hex$(&H80 Or (code And 63))
        Else
            out = out & "%" & Hex$(&HE0 Or (code \ 4096)) & _
                  "%" & Hex$(&H80 Or ((code \ 64) And 63)) & _
                  "%" & Hex$(&H80 Or (code And 63))
        End If
    Next i

    EncodeSearchTerm = out
End Function


' True when the section marker is the plural one: "SS", "sections", "secs.".
Private Function ConstMarkerIsPlural(ByVal marker As String) As Boolean
    Dim t As String: t = LCase$(Trim$(marker))
    ConstMarkerIsPlural = (t = ChrW$(167) & ChrW$(167)) Or _
                          (t = "sections") Or (t = "secs.")
End Function


' The article as leginfo and the reporters both write it -- a roman numeral in
' capitals, and the letter after it when the article carries one ("XIII A"). A
' ruling that wrote the number in arabic gets it converted; anything that is
' neither roman nor arabic returns "", which leaves the reference unlinked.
Private Function ConstArticleRoman(ByVal art As String) As String
    Dim t As String: t = UCase$(Trim$(art))
    If Len(t) = 0 Then Exit Function

    ' Split off the article letter, whatever run of whitespace introduced it.
    Dim letter As String
    Dim sp As Long: sp = InStr(1, t, " ")
    If sp > 0 Then
        letter = Trim$(Mid$(t, sp + 1))
        t = Trim$(Left$(t, sp - 1))
        If Len(letter) <> 1 Then Exit Function
        If Not letter Like "[A-D]" Then Exit Function
    End If

    Dim num As String
    If t Like "*[!IVXL]*" Then
        If Not IsNumeric(t) Then Exit Function
        num = ArabicToRoman(CLng(t))
    Else
        num = t
    End If
    If Len(num) = 0 Then Exit Function

    If Len(letter) > 0 Then num = num & " " & letter
    ConstArticleRoman = num
End Function


' 1 to 39 as a roman numeral, which covers every article the California
' Constitution has. "" for anything outside that.
Private Function ArabicToRoman(ByVal n As Long) As String
    If n < 1 Or n > 39 Then Exit Function

    Dim out As String
    Do While n >= 10
        out = out & "X"
        n = n - 10
    Loop
    If n = 9 Then
        out = out & "IX"
        n = 0
    ElseIf n >= 5 Then
        out = out & "V"
        n = n - 5
    End If
    If n = 4 Then
        out = out & "IV"
        n = 0
    End If
    Do While n >= 1
        out = out & "I"
        n = n - 1
    Loop

    ArabicToRoman = out
End Function


' True when one of the two words before a rule reference is "local".
'
' Two words, not one, because the two shapes put a different number of them in
' the way: "local rule 3.57" sits right against the reference, while "Local
' Rules, rule 3.57" -- how a superior court's own rules are cited in full --
' puts "Rules," between them.
'
' The walk back over the punctuation between words is capped, so this cannot
' step across a long run of it into a preceding clause and find a "local" that
' has nothing to do with this rule.
Private Function PrecededByLocal(ByVal raw As String, ByVal matchStart As Long) As Boolean
    Const MAXGAP As Long = 4

    Dim i As Long
    i = matchStart - 1

    Dim w As Long
    For w = 1 To 2
        ' Back over the punctuation and spaces to the end of the previous word.
        Dim gap As Long
        gap = 0
        Do While i >= 1
            If IsRuleWordChar(Mid$(raw, i, 1)) Then Exit Do
            gap = gap + 1
            If gap > MAXGAP Then Exit Function
            i = i - 1
        Loop
        If i < 1 Then Exit Function

        ' And then across the word itself.
        Dim wordEnd As Long
        wordEnd = i
        Do While i >= 1
            If Not IsRuleWordChar(Mid$(raw, i, 1)) Then Exit Do
            i = i - 1
        Loop

        If StrComp(Mid$(raw, i + 1, wordEnd - i), "local", vbTextCompare) = 0 Then
            PrecededByLocal = True
            Exit Function
        End If
    Next w
End Function


Private Function IsRuleWordChar(ByVal ch As String) As Boolean
    IsRuleWordChar = (ch Like "[A-Za-z0-9]")
End Function


' Hyperlink the first occurrence of needle inside scope that is not already
' linked. The looking is FindUnlinked's; this puts the link on what it found.
Private Sub LinkTextIfUnlinked(ByVal scope As Range, ByVal needle As String, _
                               ByVal url As String, ByRef added As Long)
    Dim fr As Range
    Set fr = FindUnlinked(scope, needle)
    If fr Is Nothing Then Exit Sub
    If AddLink(fr, url, "case") Then added = added + 1
End Sub


' The first occurrence of needle inside scope that is not already hyperlinked,
' or Nothing. Word Find does the looking, so field and footnote positions are
' handled; occurrences already linked are walked past rather than returned,
' because a span inside a hyperlink is one some earlier pass has already
' claimed and Word will not nest a field inside a field.
Private Function FindUnlinked(ByVal scope As Range, ByVal needle As String) As Range
    On Error GoTo Done
    If Len(needle) = 0 Or Len(needle) > 250 Then Exit Function

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
            Set FindUnlinked = fr
            Exit Function
        End If

        ' This occurrence is already linked -- resume past it.
        searchStart = fr.End
        If searchStart >= scope.End Then Exit Do
    Loop
Done:
End Function


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

' Regex test plus a code-name check: an entire line that is a code-section
' citation, optionally led by a roman numeral + period. The pattern accepts an
' optional run of words before the section marker (Section / Sec. / section
' sign), then a section number with optional dotted parts and (a)(1)-style
' subdivisions -- and NOTHING after it, so "Section 5 of the lease" (a sentence)
' does not match. Case-insensitive.
'
' The pattern alone was not enough. Its leading run allows spaces, so it read a
' whole PROSE SENTENCE that happens to end on a code section as the code's name,
' whenever the sentence was short enough to clear the length cap above. "The sole
' cause of action is for Violation of Health and Safety Code Section 25249.6." is
' 84 characters, and the run swallowed everything in front of "Section" -- so
' unlinking that sentence's citation underlined it instead of clearing it. The
' run is now captured and has to read as the name of a code before the line
' counts as a heading.
Private Function IsCodeSectionHeadingText(ByVal s As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Global = False
        re.Pattern = "^(?:[IVXLCDM]{1,7}\.\s+)?" & _
                     "([A-Za-z][A-Za-z.,'&/ ]*\s)?" & _
                     "(?:" & ChrW(167) & "|Section|Sec\.)\s*" & _
                     "\d[\d.]*(?:\s*\([A-Za-z0-9]+\))*\.?$"
    End If

    Dim ms As Object
    Set ms = re.Execute(s)
    If ms.count = 0 Then Exit Function

    ' The run is optional, and a group that did not participate comes back Empty
    ' rather than "" -- hence the concatenation.
    IsCodeSectionHeadingText = IsCodeNameRun(Trim$("" & ms.Item(0).SubMatches(0)))
End Function


' True when s -- the run of words in front of the section marker -- reads as the
' NAME OF A CODE and nothing else. An empty run is true: "Section 1942.4" alone
' on a line needs no code name in front of it.
'
' A code name is short, it says "Code", and it is written in title case: every
' word is capitalized or abbreviated apart from the few lowercase connectors a
' code name carries ("Health AND Safety Code", "Code OF Civil Procedure"). A
' sentence fails all three. "The sole cause of action is for Violation of Health
' and Safety Code" runs twelve words and carries lowercase verbs and articles no
' code name has, so the line it opens is prose, and its citation is unlinked and
' un-underlined like every other citation in the body.
Private Function IsCodeNameRun(ByVal s As String) As Boolean
    ' "Welfare and Institutions Code" is four words, "Code of Civil Procedure"
    ' four; six leaves room for a "Cal." or a "Former" in front.
    Const MAX_WORDS As Long = 6

    Dim parts() As String
    Dim i As Long, n As Long
    Dim w As String, bare As String
    Dim firstBare As String, lastBare As String

    s = Trim$(s)
    If Len(s) = 0 Then
        IsCodeNameRun = True
        Exit Function
    End If

    parts = Split(s, " ")
    For i = LBound(parts) To UBound(parts)
        w = Trim$(parts(i))
        If Len(w) > 0 Then
            ' The ampersand of "Health & Saf. Code" is a word of the name but
            ' carries no letters, so it is counted and otherwise passed over.
            n = n + 1
            If n > MAX_WORDS Then Exit Function

            If w <> "&" Then
                bare = LettersOnlyLower(w)
                If Len(bare) = 0 Then Exit Function      ' punctuation on its own
                If Not IsCodeNameWord(w, bare) Then Exit Function
                ' A signal opens a citation sentence, never a code name.
                If n = 1 Then
                    If IsCiteSignalWord(bare) Then Exit Function
                    firstBare = bare
                End If
                lastBare = bare
            End If
        End If
    Next i

    ' The name has to actually name a code, at one end or the other: "Civil Code"
    ' and "Health & Saf. Code," end on it, "Code of Civil Procedure" opens on it.
    IsCodeNameRun = IsCodeWord(firstBare) Or IsCodeWord(lastBare)
End Function


' One word of a code name: capitalized or abbreviated ("Health", "Saf."), or one
' of the lowercase connectors a code name is allowed to carry. Binary compare, so
' "[A-Z]" is the capital letters and nothing else.
Private Function IsCodeNameWord(ByVal w As String, ByVal bare As String) As Boolean
    Select Case bare
        Case "and", "of", "the"
            IsCodeNameWord = True
        Case Else
            IsCodeNameWord = (Left$(w, 1) Like "[A-Z]")
    End Select
End Function


Private Function IsCodeWord(ByVal bare As String) As Boolean
    IsCodeWord = (bare = "code" Or bare = "codes")
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
