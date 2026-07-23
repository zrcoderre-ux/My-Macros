Attribute VB_Name = "DeAnonymize"
'==============================================================================
' DeAnonymize.bas
'------------------------------------------------------------------------------
' Reverse of the PDF-Linker pseudonym tool. PDF-Linker replaces every real
' party name, case number, and piece of PII in an exported .txt with a stable
' FAKE, and writes a key spreadsheet mapping real <-> fake. Claude drafts the
' tentative from that anonymized text, so the draft contains the fakes. This
' macro reads the key and swaps every fake back to its real value.
'
' The key file PDF-Linker writes is "pseudonym_key.xlsx", a worksheet with the
' columns:
'     Category | Real Value | Replacement | Source | Occurrences
' where "Replacement" is the fake that appears in the anonymized draft.
'
' MACROS YOU RUN:
'   DeAnonymizeTentative - locate the key, then replace every fake with its
'                          real value throughout the document (in place).
'   ReAnonymizeTentative - the reverse: replace every real value with its fake,
'                          then export the anonymized text as a NEW Markdown
'                          (.md) file so it is safe to share. Hyperlinks are
'                          stripped (keeping their display text) before the
'                          replacement pass, and the export's default filename
'                          is the faked version of the document's own title.
'                          Nothing is ever written back to the Word document:
'                          the real->fake scrub runs in memory only (so italic
'                          cited authorities can be detected), the body is read
'                          out as Markdown, and the window is then reloaded from
'                          the untouched original.
'
' NOTES:
'   - The draft is a regular document (no live mail-merge fields): every value
'     -- caption, party block, and body prose -- is plain text. Replacement
'     covers the main body, each section's headers and footers, and
'     footnotes/endnotes.
'   - De-anonymize replaces longest fakes first, re-anonymize replaces longest
'     real values first, so a bare-surname token never rewrites part of a
'     longer full name.
'   - Court identity (Department 515, Judge Honorable Alison Mackenzie, Judicial
'     Assistant Steve Temblador, Courtroom Assistant Nancy Quintanilla) lives in
'     the header. De-anonymize fills it in; re-anonymize blanks it (keeping the
'     labels). See ApplyCourtIdentity.
'   - Re-anonymize leaves names inside italic text alone: cited case names in a
'     brief are italicized, so a party surname that also names a published case
'     (e.g. "Nash v. Superior Court") is preserved rather than rewritten. This
'     mirrors PDF-Linker's rule -- renaming a cited decision is worse than
'     leaving a party name in -- and its caption exemption (the own caption/prose
'     aren't italic, so the current parties are still replaced).
'   - Reads .xlsx via Excel automation. The rare JSON fallback that PDF-Linker
'     writes only when openpyxl is missing is not supported.
'   - AUTOMATIC ON CLOSE: RunDeAnonymizeOnClose (called from the close-review in
'     clsAppEvents) restores real names when a dated OneDrive tentative is
'     closed -- once per document, and never for re-anonymize output. It keys
'     off a pseudonym_key.xlsx in the document's folder and does nothing
'     silently if there isn't one. Two document variables track state:
'     MM_DeAnonymizeDone and MM_ReAnonymizeCreated.
'==============================================================================
Option Explicit

' PDF-Linker writes "pseudonym_key.xlsx"; match that plus any de-duplicated
' copies Windows may create (e.g. "pseudonym_key (1).xlsx"). Newest wins.
Private Const KEY_PATTERN As String = "pseudonym_key*.xlsx"

' Leftover pseudonym-pool words are flagged in pink (wdPink) after de-anonymize.
' Pink is distinct from the close-review's green/turquoise (which get auto-
' cleared) and from the user's own yellow, so these leak flags stand out. Used
' directly at the highlight site; a wd* enum member isn't a valid Const value.

' Document variables (persisted inside the .docx) that gate the automatic
' de-anonymize-on-close: DEANON_DONE marks a document already de-anonymized;
' REANON_CREATED marks a document produced by the re-anonymize macro, which must
' never be de-anonymized.
Private Const DEANON_DONE_VAR    As String = "MM_DeAnonymizeDone"
Private Const REANON_CREATED_VAR As String = "MM_ReAnonymizeCreated"

' Session latch: once re-anonymize runs, the automatic de-anonymize-on-close is
' disabled for the REST OF THE WORD SESSION -- every document, no heuristics.
' Per-document flags can in principle be stripped by metadata cleanup or a
' non-Word round trip; this latch cannot, so the shared clean copy can never be
' un-anonymized by a close in the same session that produced it.
Private g_ReAnonThisSession As Boolean

Private Type Mapping
    real As String
    fake As String
End Type

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub DeAnonymizeTentative()
    On Error GoTo ErrH

    Dim oDoc As Document
    Set oDoc = ActiveDocument
    If oDoc Is Nothing Then Exit Sub

    Dim keyPath As String
    keyPath = ResolveKeyPath(oDoc)
    If Len(keyPath) = 0 Then Exit Sub          ' user cancelled the picker

    Dim maps() As Mapping
    Dim nMaps As Long
    If Not ReadPseudonymKey(keyPath, maps, nMaps) Then
        MsgBox "Could not read any real/fake mappings from:" & vbCrLf & vbCrLf & _
               keyPath & vbCrLf & vbCrLf & _
               "Make sure this is the pseudonym_key.xlsx PDF-Linker wrote " & _
               "(with 'Real Value' and 'Replacement' columns).", _
               vbExclamation, "De-Anonymize"
        Exit Sub
    End If

    ' Longest fake first: a bare token like "Thorne" must not rewrite part of a
    ' longer fake like "Barry Thorne" before that longer one is handled.
    SortMappingsByLenDesc maps, nMaps, True

    ' Accidental-target check: de-anonymizing RE-anonymize output would put
    ' the real names back into the shared clean copy. The automatic close
    ' hook refuses such documents outright; the manual macro warns loudly
    ' and defaults to No.
    If LooksReAnonymized(oDoc) Then
        If MsgBox("This document looks like RE-ANONYMIZE OUTPUT (the shared " & _
                  "clean copy). De-anonymizing it will put the real names " & _
                  "back into it." & vbCrLf & vbCrLf & _
                  "Are you sure you want to continue?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, _
                  "De-Anonymize") <> vbYes Then Exit Sub
    End If

    If MsgBox("Restore real names using " & nMaps & " mapping(s) from:" & vbCrLf & vbCrLf & _
              keyPath & vbCrLf & vbCrLf & _
              "This replaces every pseudonym throughout the document with its " & _
              "real value. Work on a copy if you want an easy way back.", _
              vbYesNo + vbQuestion, "De-Anonymize") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Dim prevTrack As Boolean: prevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False

    ' Turn AutoSave off for the run (cloud docs re-sync after every edit).
    Dim prevAutoSave As Boolean: prevAutoSave = False
    On Error Resume Next
    prevAutoSave = oDoc.AutoSaveOn
    oDoc.AutoSaveOn = False
    On Error GoTo ErrH
    Dim bStateSaved As Boolean: bStateSaved = True   ' ErrH may now restore

    ' Strip hyperlinks (keeping display text) before replacing: a fake name
    ' inside a link's display text has survived a first replacement pass in
    ' practice (e.g. a linked "(Surname Decl.)" record cite) and was only
    ' caught on a re-run after the close review had removed the links. The
    ' close review strips every link anyway, so do it up front here.
    StripHyperlinksEverywhere oDoc

    ' Deliberately NO custom UndoRecord: wrapping every replacement across a large
    ' document (dozens of terms, each many hits) into one custom undo record
    ' overflows and crashes Word. Word still records normal (multi-step) undo.
    Dim distinctHits As Long, i As Long
    For i = 1 To nMaps
        If ReplaceEverywhere(oDoc, maps(i).fake, maps(i).real) > 0 Then
            distinctHits = distinctHits + 1
        End If
        If i Mod 5 = 0 Then DoEvents      ' let Word service its queue; avoids overflow
    Next i

    On Error Resume Next
    oDoc.AutoSaveOn = prevAutoSave
    On Error GoTo ErrH
    oDoc.TrackRevisions = prevTrack

    ' Restore the court-identity header (Department 515, judge, courtroom staff).
    ApplyCourtIdentity oDoc, True

    ' Safety net: flag any pseudonym-pool word still present (even inside a
    ' larger word) in pink, so a fake the key missed doesn't slip through.
    Dim nFlags As Long
    nFlags = HighlightResidualPseudonyms(oDoc)

    Application.ScreenUpdating = True

    SetDocFlag oDoc, DEANON_DONE_VAR      ' don't auto-run again on close

    Dim sFlagLine As String
    If nFlags > 0 Then
        sFlagLine = vbCrLf & vbCrLf & "Highlighted " & nFlags & " leftover " & _
                    "pseudonym word(s) in pink -- review each in case a fake " & _
                    "slipped through."
    End If
    MsgBox "De-anonymized: restored " & distinctHits & " of " & nMaps & _
           " pseudonym(s), and filled in the court-identity header " & _
           "(department, judge, staff)." & sFlagLine & vbCrLf & vbCrLf & _
           "Review the result before finalizing.", vbInformation, "De-Anonymize"
    Exit Sub

ErrH:
    Dim eN As Long: eN = Err.Number
    Dim eD As String: eD = Err.Description
    On Error Resume Next
    ' Restore what the run changed (only if it got far enough to save state):
    ' errors used to leave TrackRevisions and AutoSave silently off.
    If bStateSaved Then
        oDoc.TrackRevisions = prevTrack
        oDoc.AutoSaveOn = prevAutoSave
    End If
    Application.ScreenUpdating = True
    MsgBox "De-Anonymize hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & eN & ": " & eD, vbExclamation, "De-Anonymize"
End Sub

'==============================================================================
' RE-ANONYMIZE  (reverse: real -> fake, exported as a clean Markdown file)
'==============================================================================
Public Sub ReAnonymizeTentative()
    On Error GoTo ErrH

    Dim oDoc As Document
    Set oDoc = ActiveDocument
    If oDoc Is Nothing Then Exit Sub

    Dim keyPath As String
    keyPath = ResolveKeyPath(oDoc)
    If Len(keyPath) = 0 Then Exit Sub          ' user cancelled the picker

    Dim maps() As Mapping
    Dim nMaps As Long
    If Not ReadPseudonymKey(keyPath, maps, nMaps) Then
        MsgBox "Could not read any real/fake mappings from:" & vbCrLf & vbCrLf & _
               keyPath & vbCrLf & vbCrLf & _
               "Make sure this is the pseudonym_key.xlsx PDF-Linker wrote " & _
               "(with 'Real Value' and 'Replacement' columns).", _
               vbExclamation, "Re-Anonymize"
        Exit Sub
    End If

    ' Longest real value first so a bare surname doesn't rewrite part of a
    ' longer full name before that longer one is handled.
    SortMappingsByLenDesc maps, nMaps, False

    ' Word's Find can only search for terms up to 255 characters, so a longer
    ' real value (a quoted block, a long address) can never be auto-replaced
    ' and would survive into the shared copy. Warn BEFORE doing anything.
    Dim nTooLong As Long, i As Long
    For i = 1 To nMaps
        If Len(maps(i).real) > 255 Then nTooLong = nTooLong + 1
    Next i
    If nTooLong > 0 Then
        If MsgBox(nTooLong & " mapping(s) in the key have a real value longer " & _
                  "than 255 characters, which Word's search cannot handle. " & _
                  "Those values will NOT be replaced and would remain in the " & _
                  "anonymized copy." & vbCrLf & vbCrLf & _
                  "Continue anyway (and review the output for them manually)?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Re-Anonymize") <> vbYes Then Exit Sub
    End If

    ' Choose where to write the Markdown file BEFORE changing anything, so the
    ' run can be cancelled with nothing touched. Default the filename to the
    ' FAKED version of the document's own title, so the export is recognizable
    ' but carries pseudonyms, not real party names.
    Dim savePath As String
    savePath = PickReAnonSavePath(oDoc, maps, nMaps)
    If Len(savePath) = 0 Then Exit Sub

    If MsgBox("Re-anonymize using " & nMaps & " mapping(s) and save an " & _
              "anonymized Markdown file to:" & vbCrLf & vbCrLf & savePath & vbCrLf & vbCrLf & _
              "The Word document is left unchanged -- only the .md file is written.", _
              vbYesNo + vbQuestion, "Re-Anonymize") <> vbYes Then Exit Sub

    ' From this point on, no automatic de-anonymize for the rest of the Word
    ' session (set even if the run errors out partway -- fail safe). This also
    ' keeps the close hook from firing when we discard the scratch edits below.
    g_ReAnonThisSession = True

    Application.ScreenUpdating = False

    ' The real->fake scrub runs IN MEMORY on the open document ONLY -- so italic
    ' cited authorities can be detected and preserved -- and the result is then
    ' read out as Markdown. The Word file itself is NEVER written: AutoSave is
    ' disabled before the first edit, no Save/SaveAs is ever issued, and the
    ' in-memory edits are discarded at the end (the window is reloaded from the
    ' untouched original on disk).
    Dim prevTrack As Boolean: prevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False
    On Error Resume Next
    oDoc.AutoSaveOn = False              ' must precede edits: AutoSave would
    On Error GoTo ErrH                   ' push real->fake edits to the ORIGINAL
    Dim bStateSaved As Boolean: bStateSaved = True   ' ErrH may now restore

    ' Strip hyperlinks (keeping display text) before replacing and exporting:
    ' link targets can carry real names/paths the Markdown must not contain,
    ' and a real name inside a link's display text is replaced more reliably
    ' once the link is gone. Runs on the in-memory scratch copy only -- the
    ' original file (reloaded below) keeps its links.
    StripHyperlinksEverywhere oDoc

    ' Reverse direction: replace each real value with its fake. protectCitations
    ' leaves names inside italic cited authorities alone, so a party surname that
    ' also names a published case isn't rewritten in the shared copy. No custom
    ' undo record (it overflows and crashes Word on large documents).
    Dim distinctHits As Long
    For i = 1 To nMaps
        If ReplaceEverywhere(oDoc, maps(i).real, maps(i).fake, True) > 0 Then
            distinctHits = distinctHits + 1
        End If
        If i Mod 5 = 0 Then DoEvents
    Next i

    ' Blank the court-identity header (Department 515, judge, courtroom staff) so
    ' the shared copy doesn't reveal them. ApplyCourtIdentity also scrubs the
    ' body, which is what the Markdown export reads.
    ApplyCourtIdentity oDoc, False

    ' Generic court-personnel scrub, ported from PDF-Linker's
    ' register_court_names: role labels ("Judicial Assistant: <name>", "Deputy
    ' Clerk: ...", "Court Reporter: ..."), judicial titles ("Hon. / Judge /
    ' Justice <name>"), and label-anchored department numbers ("Department 515",
    ' "Dept. 72") are blanked WHOEVER is named -- nothing is hard-coded, so a
    ' new assistant, a different judge, or a body-text mention in a form the
    ' exact-string pass doesn't know is still caught.
    ScrubCourtIdentityGeneric oDoc

    ' Leak gate, ported from PDF-Linker's quarantine rule (a leaked real value
    ' is worse than an aborted run): if any real value from the key survived
    ' outside a protected italic citation, warn BEFORE anything is written and
    ' let the run be aborted with nothing on disk.
    Dim nLeaks As Long, sLeakList As String
    nLeaks = CountRealLeaks(oDoc, maps, nMaps, sLeakList)
    Dim bWriteExport As Boolean: bWriteExport = True
    If nLeaks > 0 Then
        If MsgBox("WARNING: " & nLeaks & " occurrence(s) of real value(s) from " & _
                  "the key are STILL PRESENT after replacement:" & vbCrLf & vbCrLf & _
                  sLeakList & vbCrLf & _
                  "(Real names inside italic cited case names are exempt and " & _
                  "not counted.)" & vbCrLf & vbCrLf & _
                  "Write the anonymized export anyway?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Re-Anonymize") <> vbYes Then
            bWriteExport = False
        End If
    End If

    ' Read the now-anonymized body out as Markdown and write it to disk (UTF-8,
    ' no BOM). This is the only file the macro writes.
    Dim md As String
    If bWriteExport Then
        md = DocToMarkdown(oDoc)
        WriteUtf8NoBom savePath, md
    End If

    ' Discard the in-memory fake edits: reload the window from the untouched
    ' original so the user is back on the real-names document and a stray Ctrl+S
    ' can never push fakes into it. If the document was never saved to disk
    ' (no path to reopen), leave the scratch window in place with AutoSave off
    ' and warn instead.
    Dim origPath As String: origPath = ""
    On Error Resume Next
    If Len(oDoc.path) > 0 Then origPath = oDoc.FullName
    On Error GoTo ErrH

    Application.ScreenUpdating = True

    Dim closedOK As Boolean: closedOK = False
    Dim reloaded As Boolean: reloaded = False
    If Len(origPath) > 0 Then
        ' The scratch copy is (typically) a dated OneDrive tentative, so this
        ' Close would otherwise fire the close review in clsAppEvents -- a
        ' surprise "Review document before closing?" prompt on the scrubbed
        ' copy, and a "stay open" answer would cancel the close while we
        ' report success. Suppress the review with the same flag the mail
        ' merge uses. (g_ReAnonThisSession already suppresses the
        ' de-anonymize half of that hook.)
        modMain.gSkipCloseChecks = True
        On Error Resume Next
        oDoc.Close SaveChanges:=wdDoNotSaveChanges
        Err.Clear
        ' Probe whether the close actually happened: touching a closed
        ' Document object raises, so an error here means success.
        Dim sProbe As String
        sProbe = oDoc.name
        closedOK = (Err.Number <> 0)
        Err.Clear
        modMain.gSkipCloseChecks = False
        If closedOK Then
            Documents.Open FileName:=origPath, AddToRecentFiles:=False
            reloaded = (Err.Number = 0)
        End If
        On Error GoTo 0
    End If

    If Len(origPath) > 0 And Not closedOK Then
        On Error Resume Next
        oDoc.TrackRevisions = prevTrack     ' keep AutoSave OFF: window holds fakes
        On Error GoTo 0
    ElseIf Len(origPath) = 0 Then
        oDoc.TrackRevisions = prevTrack     ' keep AutoSave OFF: window holds fakes
    End If

    Dim tail As String
    If reloaded Then
        tail = "Your Word window has been reloaded from the original file " & _
               "(real names), which was never modified."
    ElseIf closedOK Then
        tail = "The scratch window (fake names) was discarded, but the original " & _
               "could not be reopened automatically -- open it yourself from:" & _
               vbCrLf & origPath & vbCrLf & "It was never modified."
    Else
        tail = "This window still holds the re-anonymized (fake) content and was " & _
               "NOT saved -- close it WITHOUT saving to discard those edits and " & _
               "get back to the untouched original."
    End If

    If bWriteExport Then
        MsgBox "Re-anonymized: replaced " & distinctHits & " of " & nMaps & _
               " value(s) and blanked the court-identity header." & vbCrLf & vbCrLf & _
               "Names inside italic cited case names were left as-is so a party " & _
               "surname that also names a published case wasn't rewritten -- check " & _
               "any italicized cites if a real party name should have been replaced." & _
               vbCrLf & vbCrLf & _
               "Saved an anonymized Markdown file to:" & vbCrLf & savePath & vbCrLf & vbCrLf & _
               tail, vbInformation, "Re-Anonymize"
    Else
        MsgBox "Re-anonymize ABORTED at the leak check: nothing was written to " & _
               "disk." & vbCrLf & vbCrLf & _
               "Fix the key (add the missing variant rows) and run again." & _
               vbCrLf & vbCrLf & tail, vbExclamation, "Re-Anonymize"
    End If
    Exit Sub

ErrH:
    Dim reN As Long: reN = Err.Number
    Dim reD As String: reD = Err.Description
    On Error Resume Next
    ' Restore TrackRevisions (errors used to leave it silently off). AutoSave
    ' is deliberately NOT re-enabled here: the window may hold partial
    ' real->fake edits, and re-enabling AutoSave would push them to the
    ' original cloud file before the user can close without saving.
    If bStateSaved Then oDoc.TrackRevisions = prevTrack
    Application.ScreenUpdating = True
    MsgBox "Re-Anonymize hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & reN & ": " & reD & vbCrLf & vbCrLf & _
           "If the error happened before the .md file was written, this window " & _
           "may hold partial re-anonymize edits that were NOT saved anywhere -- " & _
           "close it WITHOUT saving to get back to the untouched original. " & _
           "(AutoSave was left off for the same reason.)", _
           vbExclamation, "Re-Anonymize"
End Sub

' Ask where to write the anonymized Markdown file. Defaults to the document's
' folder and to the FAKED version of the document's own title (real values in
' the filename replaced with their pseudonyms via the key), so the export is
' recognizable without carrying real party names. Returns "" if cancelled.
' Always normalizes the result to a .md extension -- the SaveAs dialog can
' otherwise append a Word extension.
Private Function PickReAnonSavePath(ByVal oDoc As Document, _
                                     ByRef maps() As Mapping, _
                                     ByVal nMaps As Long) As String
    Dim folder As String
    folder = ""
    On Error Resume Next
    folder = oDoc.path
    On Error GoTo 0
    If Len(folder) = 0 Then folder = Environ$("USERPROFILE") & "\Documents"

    Dim fd As FileDialog
    Dim p As String
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Save the anonymized Markdown file as"
        .InitialFileName = folder & "\" & FakedDocTitle(oDoc, maps, nMaps) & ".md"
        If .Show <> -1 Then
            PickReAnonSavePath = ""
            Exit Function
        End If
        p = .SelectedItems(1)
    End With

    ' Drop any extension the dialog tacked on (it defaults to a Word type), then
    ' force .md, so the file is always written as Markdown.
    Dim dotPos As Long: dotPos = InStrRev(p, ".")
    Dim slashPos As Long: slashPos = InStrRev(p, "\")
    If dotPos > slashPos And dotPos > 0 Then
        Select Case LCase$(Mid$(p, dotPos + 1))
            Case "md", "markdown", "docx", "doc", "dot", "dotx", "dotm", "txt", "rtf", "xml"
                p = Left$(p, dotPos - 1)
        End Select
    End If
    If LCase$(Right$(p, 3)) <> ".md" Then p = p & ".md"
    PickReAnonSavePath = p
End Function

' The document's title (filename without extension) with every real value from
' the key replaced by its fake, longest real first (the maps are already sorted
' that way when this is called). Matching is case-insensitive and the fake is
' recased to mirror the casing found, same as the body replacement. Characters
' Windows forbids in filenames are folded to "-" as a safety net.
Private Function FakedDocTitle(ByVal oDoc As Document, _
                                ByRef maps() As Mapping, _
                                ByVal nMaps As Long) As String
    Dim t As String
    t = "Anonymized Draft"            ' fallback for an unnamed document
    On Error Resume Next
    t = oDoc.name
    On Error GoTo 0

    Dim dotPos As Long: dotPos = InStrRev(t, ".")
    If dotPos > 1 Then t = Left$(t, dotPos - 1)

    Dim i As Long
    For i = 1 To nMaps
        t = ReplaceCIString(t, maps(i).real, maps(i).fake)
    Next i

    Dim k As Long, ch As String
    For k = 1 To Len(t)
        ch = Mid$(t, k, 1)
        If InStr(1, "\/:*?""<>|", ch) > 0 Then Mid$(t, k, 1) = "-"
    Next k

    FakedDocTitle = Trim$(t)
    If Len(FakedDocTitle) = 0 Then FakedDocTitle = "Anonymized Draft"
End Function

' Case-insensitive replace of every occurrence of findText in s, recasing the
' replacement to mirror each occurrence's casing (via MatchCasing). Restarts
' the scan after each inserted replacement so an inserted fake is never itself
' rescanned.
Private Function ReplaceCIString(ByVal s As String, _
                                  ByVal findText As String, _
                                  ByVal replaceText As String) As String
    Dim res As String, pos As Long, hit As Long
    res = "": pos = 1
    If Len(findText) = 0 Then ReplaceCIString = s: Exit Function
    Do
        hit = InStr(pos, s, findText, vbTextCompare)
        If hit = 0 Then
            res = res & Mid$(s, pos)
            Exit Do
        End If
        res = res & Mid$(s, pos, hit - pos) & _
              MatchCasing(Mid$(s, hit, Len(findText)), replaceText)
        pos = hit + Len(findText)
    Loop
    ReplaceCIString = res
End Function

'==============================================================================
' HYPERLINK STRIPPING
'==============================================================================
' Remove every hyperlink, keeping its display text, from all the stories the
' replacement pass touches. The body goes through Citation Linker's quiet
' remover (which also resets the blue/underline link formatting); headers,
' footers, notes, and text boxes are handled directly here. Motivation: a fake
' name inside a link's display text has survived a replacement pass in practice
' (a linked "(Surname Decl.)" record cite), and for the Markdown export the
' link targets themselves can leak real names or file paths.
Private Sub StripHyperlinksEverywhere(ByVal oDoc As Document)
    On Error Resume Next

    CitationLinker.RemoveAllHyperlinks_Quiet oDoc
    Application.ScreenUpdating = False   ' the helper re-enables it on exit

    Dim sec As Section, hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then StripHyperlinksInRange hf.Range
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then StripHyperlinksInRange hf.Range
        Next hf
    Next sec

    If oDoc.Footnotes.count > 0 Then StripHyperlinksInRange oDoc.StoryRanges(wdFootnotesStory)
    If oDoc.Endnotes.count > 0 Then StripHyperlinksInRange oDoc.StoryRanges(wdEndnotesStory)

    Dim shp As Shape
    For Each shp In oDoc.Shapes
        If shp.TextFrame.HasText Then StripHyperlinksInRange shp.TextFrame.TextRange
    Next shp
End Sub

' Delete the hyperlinks in one range, newest-index first (the collection
' reindexes as links are deleted). Display text is retained.
Private Sub StripHyperlinksInRange(ByVal rng As Range)
    On Error Resume Next
    Dim i As Long
    For i = rng.Hyperlinks.count To 1 Step -1
        rng.Hyperlinks(i).Delete
    Next i
End Sub

'==============================================================================
' MARKDOWN EXPORT  (read the in-memory, already-anonymized body out as Markdown)
'==============================================================================
' Convert the document's main body to Markdown:
'   - paragraph styles Heading 1..6 / Title  ->  # .. ###### / #
'   - list paragraphs                        ->  "- " (bullet) or the number label
'   - bold / italic runs                     ->  **bold**, *italic*, ***both***
'   - footnote/endnote reference marks        ->  [^n], with the note texts
'                                                 collected into a trailing block
' Only the body is exported: headers/footers carry the court identity (already
' blanked) and have no place in Markdown. Formatting Markdown can't express
' (alignment, tab leaders in the caption, tables) is dropped but the text is
' kept. Text boxes/shapes are not exported.
Private Function DocToMarkdown(ByVal oDoc As Document) As String
    Dim sb As String
    Dim fnList As String            ' accumulates the "[^n]: ..." footnote block
    Dim fnCount As Long: fnCount = 0
    Dim firstBlock As Boolean: firstBlock = True

    Dim p As Paragraph
    For Each p In oDoc.content.Paragraphs
        Dim line As String
        line = ParagraphToMarkdown(oDoc, p, fnCount, fnList)
        If Len(line) > 0 Then
            If Not firstBlock Then sb = sb & vbCrLf & vbCrLf
            sb = sb & line
            firstBlock = False
        End If
    Next p

    If Len(fnList) > 0 Then sb = sb & vbCrLf & vbCrLf & fnList

    DocToMarkdown = sb & vbCrLf
End Function

' One body paragraph -> one Markdown block (or "" for an empty paragraph, which
' just becomes block separation). List prefix wins over heading prefix.
Private Function ParagraphToMarkdown(ByVal oDoc As Document, ByVal p As Paragraph, _
                                     ByRef fnCount As Long, ByRef fnList As String) As String
    ' Paragraph content without the trailing paragraph mark.
    Dim wr As Range: Set wr = p.Range.Duplicate
    If wr.Characters.count >= 1 Then wr.MoveEnd wdCharacter, -1

    Dim inner As String
    inner = InlineMarkdown(oDoc, wr, fnCount, fnList)
    If Len(Trim$(inner)) = 0 Then Exit Function

    ' List item?
    Dim listPrefix As String: listPrefix = ""
    On Error Resume Next
    If p.Range.ListFormat.ListType <> wdListNoNumbering Then
        If p.Range.ListFormat.ListType = wdListBullet Then
            listPrefix = "- "
        Else
            Dim ls As String: ls = Trim$(p.Range.ListFormat.ListString)
            If Len(ls) = 0 Then
                listPrefix = "- "
            ElseIf Right$(ls, 1) = "." Then
                listPrefix = ls & " "
            Else
                listPrefix = ls & ". "
            End If
        End If
    End If
    On Error GoTo 0
    If Len(listPrefix) > 0 Then
        ParagraphToMarkdown = listPrefix & inner
        Exit Function
    End If

    ParagraphToMarkdown = HeadingPrefix(p) & inner
End Function

' Map a Heading 1..6 / Title paragraph style to its Markdown "#" prefix (with a
' trailing space); returns "" for body styles.
Private Function HeadingPrefix(ByVal p As Paragraph) As String
    Dim sName As String: sName = ""
    On Error Resume Next
    sName = p.Style                 ' a Style object's default property is NameLocal
    On Error GoTo 0
    sName = LCase$(Trim$(sName))

    Dim level As Long: level = 0
    If Left$(sName, 8) = "heading " Then
        level = Val(Mid$(sName, 9))
    ElseIf sName = "title" Then
        level = 1
    End If
    If level < 1 Then Exit Function
    If level > 6 Then level = 6
    HeadingPrefix = String$(level, "#") & " "
End Function

' Inline formatting for one paragraph's content range. Uniform paragraphs (the
' common case -- plain body prose) are wrapped at most once without walking; only
' mixed-format paragraphs (an italic cited case name in a sentence) are walked
' character by character.
Private Function InlineMarkdown(ByVal oDoc As Document, ByVal wr As Range, _
                                ByRef fnCount As Long, ByRef fnList As String) As String
    Dim txt As String: txt = wr.text
    If Len(txt) = 0 Then Exit Function

    Dim boldUniform As Boolean, italicUniform As Boolean
    boldUniform = (wr.Font.Bold <> wdUndefined)
    italicUniform = (wr.Font.Italic <> wdUndefined)

    If boldUniform And italicUniform Then
        InlineMarkdown = Emph(MapText(oDoc, txt, fnCount, fnList), _
                              (wr.Font.Bold = True), (wr.Font.Italic = True))
    Else
        InlineMarkdown = WalkRuns(oDoc, wr, fnCount, fnList)
    End If
End Function

' Walk a mixed-format range character by character, grouping consecutive
' same-format characters into runs and wrapping each run in its emphasis markers.
Private Function WalkRuns(ByVal oDoc As Document, ByVal wr As Range, _
                          ByRef fnCount As Long, ByRef fnList As String) As String
    Dim result As String, runText As String
    Dim curB As Long, curI As Long
    curB = -1: curI = -1                    ' -1 = no run started yet
    Dim n As Long: n = wr.Characters.count
    Dim i As Long
    For i = 1 To n
        Dim ch As Range: Set ch = wr.Characters(i)
        Dim c As String: c = ch.text
        If c = Chr$(2) Then                 ' footnote/endnote reference mark
            If Len(runText) > 0 Then
                result = result & Emph(runText, curB = 1, curI = 1)
                runText = ""
            End If
            result = result & EmitNote(oDoc, fnCount, fnList)
            curB = -1: curI = -1
        Else
            Dim b As Long, it As Long
            b = IIf(ch.Font.Bold = True, 1, 0)
            it = IIf(ch.Font.Italic = True, 1, 0)
            If b <> curB Or it <> curI Then
                If Len(runText) > 0 Then
                    result = result & Emph(runText, curB = 1, curI = 1)
                    runText = ""
                End If
                curB = b: curI = it
            End If
            runText = runText & MapChar(c)
        End If
    Next i
    If Len(runText) > 0 Then result = result & Emph(runText, curB = 1, curI = 1)
    WalkRuns = result
End Function

' Map a plain (single-format) text run to Markdown, translating footnote
' reference marks and per-character specials.
Private Function MapText(ByVal oDoc As Document, ByVal s As String, _
                         ByRef fnCount As Long, ByRef fnList As String) As String
    Dim res As String, i As Long
    For i = 1 To Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        If c = Chr$(2) Then
            res = res & EmitNote(oDoc, fnCount, fnList)
        Else
            res = res & MapChar(c)
        End If
    Next i
    MapText = res
End Function

' Consume the next footnote/endnote reference (in document order): append its
' text to the trailing block and return the "[^n]" inline marker.
Private Function EmitNote(ByVal oDoc As Document, ByRef fnCount As Long, _
                          ByRef fnList As String) As String
    fnCount = fnCount + 1
    Dim body As String: body = FlattenNote(NoteText(oDoc, fnCount))
    If Len(fnList) > 0 Then fnList = fnList & vbCrLf
    fnList = fnList & "[^" & fnCount & "]: " & body
    EmitNote = "[^" & fnCount & "]"
End Function

' The text of the idx-th note, in document order. Footnotes are used when the
' document has any; endnotes otherwise. (Mixed foot/endnotes -- rare here -- fall
' back to the footnotes.)
Private Function NoteText(ByVal oDoc As Document, ByVal idx As Long) As String
    On Error Resume Next
    If oDoc.Footnotes.count > 0 Then
        If idx <= oDoc.Footnotes.count Then NoteText = oDoc.Footnotes(idx).Range.text
    ElseIf oDoc.Endnotes.count > 0 Then
        If idx <= oDoc.Endnotes.count Then NoteText = oDoc.Endnotes(idx).Range.text
    End If
End Function

' Collapse a note's text to a single Markdown line (footnote defs are one line):
' newlines/breaks become spaces, per-character specials are mapped, and nested
' reference marks are dropped.
Private Function FlattenNote(ByVal s As String) As String
    Dim res As String, i As Long
    For i = 1 To Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        Select Case c
            Case vbCr, vbLf, Chr$(11), Chr$(12): res = res & " "
            Case Chr$(2):                        ' nested reference mark: drop
            Case Else:                           res = res & MapChar(c)
        End Select
    Next i
    FlattenNote = Trim$(res)
End Function

' Translate one Word character to its Markdown equivalent: escape the characters
' Markdown treats as markup, and normalize Word's control characters.
Private Function MapChar(ByVal c As String) As String
    Select Case c
        Case vbTab:     MapChar = " "                 ' avoid a stray code block
        Case Chr$(11):  MapChar = "  " & vbCrLf       ' manual line break -> hard break
        Case Chr$(12):  MapChar = vbCrLf & vbCrLf     ' page break -> blank line
        Case Chr$(160): MapChar = " "                 ' non-breaking space
        Case Chr$(31):  MapChar = ""                  ' optional hyphen
        Case Chr$(30):  MapChar = "-"                 ' non-breaking hyphen
        Case Chr$(1), Chr$(5), Chr$(19), Chr$(20), Chr$(21)
            MapChar = ""    ' inline object / annotation / field control marks
        Case "\":       MapChar = "\\"
        Case "`":       MapChar = "\`"
        Case "*":       MapChar = "\*"
        Case "_":       MapChar = "\_"
        Case Else:      MapChar = c
    End Select
End Function

' Wrap text in bold/italic markers, moving any leading/trailing whitespace
' outside the markers so the emphasis parses. An all-whitespace or unformatted
' run is returned unchanged.
Private Function Emph(ByVal s As String, ByVal bold As Boolean, ByVal italic As Boolean) As String
    If Len(s) = 0 Then Exit Function
    Dim marker As String
    If bold Then marker = marker & "**"
    If italic Then marker = marker & "*"
    If Len(marker) = 0 Then
        Emph = s
        Exit Function
    End If

    ' Markdown emphasis cannot span a blank line (MapChar turns a page break
    ' into one): wrap each blank-line-separated piece separately.
    If InStr(s, vbCrLf & vbCrLf) > 0 Then
        Dim parts() As String, pi As Long
        parts = Split(s, vbCrLf & vbCrLf)
        For pi = LBound(parts) To UBound(parts)
            parts(pi) = Emph(parts(pi), bold, italic)
        Next pi
        Emph = Join(parts, vbCrLf & vbCrLf)
        Exit Function
    End If

    Dim iStart As Long, iEnd As Long
    iStart = 1
    Do While iStart <= Len(s)
        If Not IsWs(Mid$(s, iStart, 1)) Then Exit Do
        iStart = iStart + 1
    Loop
    If iStart > Len(s) Then                  ' all whitespace: nothing to emphasize
        Emph = s
        Exit Function
    End If
    iEnd = Len(s)
    Do While iEnd >= 1
        If Not IsWs(Mid$(s, iEnd, 1)) Then Exit Do
        iEnd = iEnd - 1
    Loop

    Emph = Left$(s, iStart - 1) & marker & Mid$(s, iStart, iEnd - iStart + 1) & _
           marker & Mid$(s, iEnd + 1)
End Function

Private Function IsWs(ByVal c As String) As Boolean
    IsWs = (c = " " Or c = vbCr Or c = vbLf Or c = vbTab)
End Function

' Write text to disk as UTF-8 without a byte-order mark (plain Markdown tools can
' choke on a BOM). ADODB.Stream writes a BOM, so we re-read the bytes past it and
' save those to the file.
Private Sub WriteUtf8NoBom(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                     ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText text

    st.Position = 0
    st.Type = 1                     ' adTypeBinary
    st.Position = 3                 ' skip the 3-byte UTF-8 BOM
    Dim bytes As Variant: bytes = st.Read
    st.Close

    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1                    ' adTypeBinary
    bin.Open
    bin.Write bytes
    bin.SaveToFile path, 2          ' adSaveCreateOverWrite
    bin.Close
End Sub

'==============================================================================
' AUTOMATIC DE-ANONYMIZE ON CLOSE
'==============================================================================
' Called from clsAppEvents.App_DocumentBeforeClose. Restores real names when a
' dated OneDrive tentative is closed, but ONLY if de-anonymize hasn't already
' run on it and it wasn't produced by the re-anonymize macro. Silent: with no
' pseudonym key in the document's folder it does nothing (the document isn't an
' anonymized draft, or the key is unavailable). Never sets Cancel, so it can't
' block the close.
Public Sub RunDeAnonymizeOnClose(ByVal Doc As Document)
    On Error Resume Next
    If Doc Is Nothing Then Exit Sub
    If HasDocFlag(Doc, DEANON_DONE_VAR) Then Exit Sub

    ' Re-anonymize output must NEVER be auto-restored: un-anonymizing the
    ' shared clean copy on close would put real names back into the one file
    ' that exists to not have them. Two gates:
    '   1. The session latch -- once re-anonymize has run in this Word
    '      session, auto-restore is off for EVERY close until Word restarts.
    '   2. LooksReAnonymized -- the per-document flag plus the filename, for
    '      re-anonymize output opened in a LATER session.
    If g_ReAnonThisSession Then Exit Sub
    If LooksReAnonymized(Doc) Then Exit Sub

    ' DocFolderLocal, not Doc.Path: for a synced OneDrive/SharePoint document
    ' Doc.Path is an https URL that FolderExists can't read, which silently
    ' disabled this hook for exactly the dated-OneDrive tentatives it targets.
    Dim folder As String: folder = DocFolderLocal(Doc)
    If Len(folder) = 0 Then Exit Sub
    Dim keyPath As String: keyPath = MostRecentKeyInFolder(folder)
    If Len(keyPath) = 0 Then Exit Sub

    Dim maps() As Mapping, nMaps As Long
    If Not ReadPseudonymKey(keyPath, maps, nMaps) Then Exit Sub

    SortMappingsByLenDesc maps, nMaps, True

    Application.ScreenUpdating = False
    Dim prevTrack As Boolean: prevTrack = Doc.TrackRevisions
    Doc.TrackRevisions = False
    Dim prevAutoSave As Boolean: prevAutoSave = False
    prevAutoSave = Doc.AutoSaveOn
    Doc.AutoSaveOn = False

    ' Strip hyperlinks first for the same reason as the manual macro: this
    ' hook runs BEFORE the close review's own link removal, and a fake name
    ' inside a link's display text has been missed on exactly this first pass.
    StripHyperlinksEverywhere Doc

    Dim i As Long
    For i = 1 To nMaps
        ReplaceEverywhere Doc, maps(i).fake, maps(i).real
        If i Mod 5 = 0 Then DoEvents
    Next i

    ' Restore the court-identity header (Department 515, judge, courtroom staff).
    ApplyCourtIdentity Doc, True

    Doc.AutoSaveOn = prevAutoSave
    Doc.TrackRevisions = prevTrack
    Application.ScreenUpdating = True

    SetDocFlag Doc, DEANON_DONE_VAR
End Sub

' True when the document is re-anonymize output. Three EXACT signals only:
'   1. The MM_ReAnonymizeCreated document variable -- the primary marker
'      (only present on legacy .docx output; the Markdown export carries no
'      document variables).
'   2. The filename contains "anonym" (legacy output, or a fallback title).
'   3. The filename ends in .md/.markdown -- the export is now a Markdown file
'      whose default name is the FAKED document title (which ends in the same
'      date as the original), so a re-anonymized .md opened in Word would pass
'      the close hook's dated-OneDrive gates with neither signal 1 nor 2 to
'      protect it. The only dated .md in those folders is re-anonymize output.
' A blanked-header content heuristic used to be a third signal, but it false-
' positived on ordinary documents (header text layout varies) and tripped the
' manual-macro warning on every run, so it was removed. Same-session safety no
' longer depends on this function at all -- g_ReAnonThisSession switches the
' close hook off for the whole session the moment re-anonymize runs.
Private Function LooksReAnonymized(ByVal Doc As Document) As Boolean
    On Error GoTo Assume                        ' fail CLOSED: unsure = re-anon
    LooksReAnonymized = False

    If HasDocFlag(Doc, REANON_CREATED_VAR) Then
        LooksReAnonymized = True
        Exit Function
    End If

    If InStr(1, Doc.name, "anonym", vbTextCompare) > 0 Then
        LooksReAnonymized = True
        Exit Function
    End If

    Dim nm As String: nm = LCase$(Doc.name)
    If Right$(nm, 3) = ".md" Or Right$(nm, 9) = ".markdown" Then
        LooksReAnonymized = True
    End If
    Exit Function

Assume:
    LooksReAnonymized = True
End Function

' --- Document flags, persisted as document variables inside the .docx --------
Private Function HasDocFlag(ByVal Doc As Document, ByVal name As String) As Boolean
    On Error Resume Next
    HasDocFlag = (Doc.Variables(name).Value = "1")
End Function

Private Sub SetDocFlag(ByVal Doc As Document, ByVal name As String)
    On Error Resume Next
    Doc.Variables(name).Value = "1"
End Sub

'==============================================================================
' KEY-FILE LOCATION
'==============================================================================
' Look in the active document's own folder for the newest pseudonym_key*.xlsx
' (the key travels with the document -- Downloads, a case folder, wherever).
' Fall back to a file picker, starting in that folder, if none is found.
Private Function ResolveKeyPath(ByVal oDoc As Document) As String
    Dim docFolder As String
    docFolder = DocFolderLocal(oDoc)   ' local path even for OneDrive/SharePoint

    If Len(docFolder) > 0 Then
        ResolveKeyPath = MostRecentKeyInFolder(docFolder)
        If Len(ResolveKeyPath) > 0 Then Exit Function
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the PDF-Linker key (pseudonym_key.xlsx)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel key", "*.xlsx"
        .Filters.Add "All files", "*.*"
        If Len(docFolder) > 0 Then .InitialFileName = docFolder & "\"
        If .Show = -1 Then
            ResolveKeyPath = .SelectedItems(1)
        Else
            ResolveKeyPath = ""
        End If
    End With
End Function

' Return the full path of the most recently modified pseudonym_key*.xlsx in the
' given folder, or "" if none exists.
Private Function MostRecentKeyInFolder(ByVal folderPath As String) As String
    On Error GoTo Done
    If Len(folderPath) = 0 Then Exit Function

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function

    Dim bestPath As String: bestPath = ""
    Dim bestDate As Date
    Dim f As Object
    For Each f In fso.GetFolder(folderPath).Files
        Dim nm As String: nm = LCase$(f.Name)
        If nm Like KEY_PATTERN Then
            If bestPath = "" Or f.DateLastModified > bestDate Then
                bestPath = f.path
                bestDate = f.DateLastModified
            End If
        End If
    Next f

    MostRecentKeyInFolder = bestPath
Done:
End Function

' Return the document's folder as a local filesystem path. For a document opened
' from a synced OneDrive / SharePoint library (common at work), Word reports its
' Path as an "https://...sharepoint.com/..." URL, which FileSystemObject cannot
' enumerate -- so the key sitting right next to the document went unseen and the
' user got the picker. Map such URLs to the local synced folder; return an
' ordinary local/UNC path unchanged, or "" for an unsaved document.
Private Function DocFolderLocal(ByVal oDoc As Document) As String
    Dim p As String
    p = ""
    On Error Resume Next
    p = oDoc.Path                  ' "" if the document has never been saved
    On Error GoTo 0
    If Len(p) = 0 Then Exit Function

    If LCase$(Left$(p, 7)) = "http://" Or LCase$(Left$(p, 8)) = "https://" Then
        DocFolderLocal = MapUrlToLocalFolder(p)
    Else
        DocFolderLocal = p         ' already local (incl. C:\...\OneDrive\...)
    End If
End Function

' Map a OneDrive / SharePoint folder URL to the local synced folder that mirrors
' it. Everything after ".../Documents/" (personal libraries) or
' ".../Shared Documents/" (team sites) is the path relative to the local sync
' root; try that tail under each OneDrive sync root the shell exposes via
' environment variables. Returns "" if no matching local folder exists.
Private Function MapUrlToLocalFolder(ByVal url As String) As String
    On Error GoTo Done

    Dim rel As String
    Dim marker As Long
    marker = InStr(1, url, "/Documents/", vbTextCompare)
    If marker > 0 Then
        rel = Mid$(url, marker + Len("/Documents/"))
    Else
        marker = InStr(1, url, "/Shared Documents/", vbTextCompare)
        If marker > 0 Then
            rel = Mid$(url, marker + Len("/Shared Documents/"))
        Else
            ' Unknown layout: keep only the trailing folder segment.
            rel = Mid$(url, InStrRev(url, "/") + 1)
        End If
    End If

    rel = Replace(URLDecode(rel), "/", "\")

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim roots(1 To 3) As String
    roots(1) = Environ$("OneDriveCommercial")
    roots(2) = Environ$("OneDrive")
    roots(3) = Environ$("OneDriveConsumer")

    Dim i As Long, cand As String
    For i = 1 To 3
        If Len(roots(i)) > 0 Then
            If Len(rel) > 0 Then cand = roots(i) & "\" & rel Else cand = roots(i)
            If fso.FolderExists(cand) Then
                MapUrlToLocalFolder = cand
                Exit Function
            End If
        End If
    Next i
Done:
End Function

' Decode the %XX escapes (chiefly %20 for a space) that appear in SharePoint
' folder URLs, so the reconstructed local path matches the real folder name.
Private Function URLDecode(ByVal s As String) As String
    Dim i As Long, ch As String, res As String
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "%" And i + 2 <= Len(s) Then
            res = res & ChrW$(CLng("&H" & Mid$(s, i + 1, 2)))
            i = i + 3
        Else
            res = res & ch
            i = i + 1
        End If
    Loop
    URLDecode = res
End Function

'==============================================================================
' READ THE KEY SPREADSHEET  (real <-> fake)
'==============================================================================
' Opens the workbook read-only via Excel automation, finds the "Real Value"
' and "Replacement" columns by header, and fills maps(1..nMaps). Returns False
' on any error or if no usable rows were found.
Private Function ReadPseudonymKey(ByVal path As String, _
                                   ByRef maps() As Mapping, _
                                   ByRef nMaps As Long) As Boolean
    On Error GoTo Fail
    nMaps = 0

    Dim xl As Object
    Dim startedXl As Boolean: startedXl = False
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    On Error GoTo Fail
    If xl Is Nothing Then
        Set xl = CreateObject("Excel.Application")
        startedXl = True
        ' Hide and silence only OUR private instance. Never touch a running
        ' Excel the user already has open: setting Visible=False there hides
        ' their own workbooks in a background process.
        xl.Visible = False
        xl.DisplayAlerts = False
    End If

    ' If the key is already open in that instance (user eyeballing mappings),
    ' read the open copy and leave it open rather than closing it under them.
    Dim wb As Object
    Dim wasOpen As Boolean: wasOpen = False
    On Error Resume Next
    Set wb = xl.Workbooks(Mid$(path, InStrRev(path, "\") + 1))
    If Not wb Is Nothing Then
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            wasOpen = True
        Else
            Set wb = Nothing
        End If
    End If
    On Error GoTo Fail
    If wb Is Nothing Then
        Set wb = xl.Workbooks.Open(FileName:=path, ReadOnly:=True, AddToMRU:=False)
    End If

    Dim ws As Object
    Set ws = wb.Worksheets(1)

    ' Pull the whole used range into a 2-D variant array in one COM round-trip.
    Dim data As Variant
    data = ws.UsedRange.Value

    ' A single-cell used range comes back as a scalar, not an array -> no data.
    If Not IsArray(data) Then GoTo CleanFail

    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    rLo = LBound(data, 1): rHi = UBound(data, 1)
    cLo = LBound(data, 2): cHi = UBound(data, 2)

    ' Locate the two columns we need from the header row.
    Dim realCol As Long, fakeCol As Long, c As Long
    realCol = 0: fakeCol = 0
    For c = cLo To cHi
        Dim hd As String: hd = LCase$(Trim$(CStr(NzText(data(rLo, c)))))
        If hd = "real value" Then realCol = c
        If hd = "replacement" Then fakeCol = c
    Next c
    If realCol = 0 Or fakeCol = 0 Then GoTo CleanFail

    ReDim maps(1 To (rHi - rLo + 1))
    Dim r As Long
    For r = rLo + 1 To rHi
        Dim rv As String, fk As String
        rv = Trim$(CStr(NzText(data(r, realCol))))
        fk = Trim$(CStr(NzText(data(r, fakeCol))))
        If Len(rv) > 0 And Len(fk) > 0 And StrComp(rv, fk, vbBinaryCompare) <> 0 Then
            nMaps = nMaps + 1
            maps(nMaps).real = rv
            maps(nMaps).fake = fk
        End If
    Next r

    If Not wasOpen Then wb.Close SaveChanges:=False
    If startedXl Then xl.Quit
    Set wb = Nothing: Set xl = Nothing

    ReadPseudonymKey = (nMaps > 0)
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wasOpen Then wb.Close SaveChanges:=False
    If startedXl Then xl.Quit
    On Error GoTo 0
    ReadPseudonymKey = False
    Exit Function

Fail:
    On Error Resume Next
    If Not wb Is Nothing Then
        If Not wasOpen Then wb.Close SaveChanges:=False
    End If
    If startedXl And Not xl Is Nothing Then xl.Quit
    On Error GoTo 0
    ReadPseudonymKey = False
End Function

' Empty cells arrive as Null/Empty; fold to "" so CStr never errors.
Private Function NzText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NzText = ""
    Else
        NzText = CStr(v)
    End If
End Function

'==============================================================================
' REPLACEMENT
'==============================================================================
' Replace findText with replaceText across the document's stable stories: the
' main body, each section's headers/footers, footnotes/endnotes when present,
' and each shape's own text frame. Returns the number of those ranges in which
' a replacement was made.
'
' This deliberately does NOT walk StoryRanges/NextStoryRange: enumerating that
' chain while replacing inside the loop can destabilize and crash Word. The
' collections below (Sections, Footnotes, Shapes with per-shape TextRange) stay
' valid across text replacement, so iterating them is safe.
' protectCitations (re-anonymize only): leave any match that sits in italic text
' untouched, so a real party name that also appears inside a cited case name
' (e.g. "Nash v. Superior Court") is not rewritten in the shared copy. This
' mirrors the PDF-Linker pseudonymizer's cardinal invariant -- renaming a cited
' decision is a worse failure than leaving a party name in -- and its caption
' exemption: a brief italicizes cited authorities but not its own caption/prose,
' so the current parties still get replaced while published cites are preserved.
Private Function ReplaceEverywhere(ByVal oDoc As Document, _
                                    ByVal findText As String, _
                                    ByVal replaceText As String, _
                                    Optional ByVal protectCitations As Boolean = False) As Long
    Dim total As Long: total = 0
    If Len(findText) = 0 Then Exit Function
    ' Word's Find raises on search terms longer than 255 characters; under the
    ' resume-next handling below that used to fall through in a dangerous state.
    ' Skip such mappings outright (re-anonymize warns about them up front).
    If Len(findText) > 255 Then Exit Function

    Dim whole As Boolean: whole = ShouldWholeWord(findText)

    ' Main body (caption, party block, and prose are all here in a plain draft).
    If ReplaceInRange(oDoc.content, findText, replaceText, whole, protectCitations) Then total = total + 1

    ' Headers and footers, section by section.
    Dim sec As Section
    Dim hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                If ReplaceInRange(hf.Range, findText, replaceText, whole, protectCitations) Then total = total + 1
            End If
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then
                If ReplaceInRange(hf.Range, findText, replaceText, whole, protectCitations) Then total = total + 1
            End If
        Next hf
    Next sec

    ' Footnotes / endnotes, only when present (accessing the story otherwise errors).
    On Error Resume Next
    If oDoc.Footnotes.count > 0 Then
        If ReplaceInRange(oDoc.StoryRanges(wdFootnotesStory), findText, replaceText, whole, protectCitations) Then total = total + 1
    End If
    If oDoc.Endnotes.count > 0 Then
        If ReplaceInRange(oDoc.StoryRanges(wdEndnotesStory), findText, replaceText, whole, protectCitations) Then total = total + 1
    End If
    On Error GoTo 0

    ' Text boxes / shapes, each one's own text frame directly. (Walking the
    ' wdTextFrameStory NextStoryRange chain is what crashes Word; touching each
    ' shape's TextRange individually is stable.) Without this, a name in a
    ' text box survived re-anonymize into the shared copy.
    On Error Resume Next
    Dim shp As Shape
    For Each shp In oDoc.Shapes
        If shp.TextFrame.HasText Then
            If ReplaceInRange(shp.TextFrame.TextRange, findText, replaceText, whole, protectCitations) Then total = total + 1
        End If
    Next shp
    On Error GoTo 0

    ReplaceEverywhere = total
End Function

' Replace every occurrence of findText with replaceText in one range. Returns
' True if at least one replacement was made.
'
' Casing is handled in two parts:
'
'  1. .MatchCase = False, so we catch every occurrence of the name regardless of
'     the casing it appears in -- including casings the key has no row for (e.g.
'     an all-caps caption when the key only carries the title-case variant).
'
'  2. We do NOT use .Replacement.Text + wdReplaceAll. With MatchCase off Word
'     applies its own "smart case" to the replacement, mangling it (a title-case
'     name matched in an all-caps caption comes back all caps; a two-word
'     replacement loses the second word's capital). Instead we find each match
'     and assign its Range.Text directly, then MatchCasing recases the
'     replacement to mirror the casing the fake actually appeared in.
Private Function ReplaceInRange(ByVal rng As Range, _
                                 ByVal findText As String, _
                                 ByVal replaceText As String, _
                                 ByVal whole As Boolean, _
                                 Optional ByVal protectCitations As Boolean = False) As Boolean
    On Error Resume Next
    Dim scan As Range: Set scan = rng.Duplicate
    Dim madeChange As Boolean
    Do
        With scan.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = findText
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False
            .MatchWholeWord = whole
            .MatchWildcards = False
            ' Capture the result BEFORE testing it: this function runs under
            ' On Error Resume Next, and an error raised inside the old
            ' "If Not .Execute Then Exit Do" skipped the whole statement --
            ' falling through to the replacement below with scan still
            ' spanning the entire story, which would overwrite it wholesale.
            Dim bHit As Boolean
            bHit = False
            Err.Clear
            bHit = .Execute
            If Err.Number <> 0 Then Err.Clear: Exit Do
            If Not bHit Then Exit Do
        End With
        ' scan now spans the matched text. Skip it (leave the real name in place)
        ' when it sits in a cited authority -- italic text -- so re-anonymize
        ' never rewrites a published case name that shares a party's surname.
        If Not (protectCitations And scan.Font.Italic = True) Then
            ' assign directly (no smart-case) after recasing the replacement
            ' to the casing the fake appeared in.
            scan.text = MatchCasing(scan.text, replaceText)
            madeChange = True
        End If
        ' Continue after this match, out to the (live) end of the range.
        scan.Collapse Direction:=wdCollapseEnd
        scan.End = rng.End
        If scan.start >= rng.End Then Exit Do
    Loop
    ReplaceInRange = madeChange
End Function

' Recase replaceText to mirror the casing of the matched fake text, so a name is
' substituted in whatever casing it appeared in -- even a casing the key has no
' dedicated row for:
'   ALL CAPS found (e.g. a caption)   -> all-caps replacement
'   all lowercase found               -> lowercase replacement
'   title/mixed found                 -> title case, but only recovered when the
'                                        stored value is itself mono-case (so an
'                                        all-caps key row can still yield "John
'                                        Smith"); an already-mixed stored value
'                                        like "McDonald" is left untouched.
'   no cased letters (e.g. a number)  -> stored value untouched
Private Function MatchCasing(ByVal matched As String, _
                              ByVal replaceText As String) As String
    Dim u As String: u = UCase$(matched)
    Dim l As String: l = LCase$(matched)
    If u = l Then                       ' no cased letters (e.g. a case number)
        MatchCasing = replaceText
    ElseIf matched = u Then             ' ALL CAPS
        MatchCasing = UCase$(replaceText)
    ElseIf matched = l Then             ' all lowercase
        MatchCasing = LCase$(replaceText)
    Else                                ' title / mixed case
        Dim ru As String: ru = UCase$(replaceText)
        Dim rl As String: rl = LCase$(replaceText)
        If replaceText = ru Or replaceText = rl Then
            ' Stored value is mono-case (ALL CAPS or all lowercase): rebuild
            ' title case so an all-caps key row still reads as a proper name.
            MatchCasing = ProperCase(replaceText)
        Else
            ' Already mixed (e.g. "McDonald", "John Smith") -- leave as authored.
            MatchCasing = replaceText
        End If
    End If
End Function

' Capitalize the first letter of each word and lowercase the rest. Word breaks
' are spaces, hyphens, and apostrophes, so "O'BRIEN" -> "O'Brien" and
' "SMITH-JONES" -> "Smith-Jones". Intercaps like "McDonald" cannot be recovered
' from an all-caps source and become "Mcdonald"; those are rare and only occur
' in the un-keyed-casing fallback.
Private Function ProperCase(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim atStart As Boolean: atStart = True
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            If atStart Then result = result & UCase$(ch) Else result = result & LCase$(ch)
            atStart = False
        Else
            result = result & ch
            atStart = (ch = " " Or ch = "-" Or ch = "'" Or ch = ChrW$(8217))
        End If
    Next i
    ProperCase = result
End Function

' Whole-word matching is safe (and wanted) only for single alphanumeric tokens
' -- name tokens and case numbers -- where a fake could otherwise match inside a
' larger word. Multi-word names, emails, and addresses contain spaces or
' punctuation that Word's word-boundary logic handles poorly, so match those
' literally instead.
Private Function ShouldWholeWord(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If Not (ch Like "[A-Za-z0-9]") Then
            ShouldWholeWord = False
            Exit Function
        End If
    Next i
    ShouldWholeWord = True
End Function

'==============================================================================
' COURT IDENTITY  (Department / Judge / courtroom staff header block)
'==============================================================================
' The tentative's header names Department 515, the judge, and the courtroom
' staff. These are fixed court facts, not matter-specific, so they live here as
' constants rather than in the pseudonym key. A shared copy must not reveal them:
'   de-anonymize (restore = True)  fills them back in  (blank -> real)
'   re-anonymize (restore = False) blanks them out     (real  -> blank)
' Each field is a (real, blank) pair; the blank keeps the label/anchor so the
' header layout is preserved and the toggle round-trips exactly.
Private Sub ApplyCourtIdentity(ByVal oDoc As Document, ByVal restore As Boolean)
    SwapCourtField oDoc, restore, "Courthouse, Department 515", "Courthouse, Department"
    SwapCourtField oDoc, restore, "Judge: Honorable Alison Mackenzie", "Judge:"
    SwapCourtField oDoc, restore, "Judicial Assistant: Steve Temblador", "Judicial Assistant:"
    SwapCourtField oDoc, restore, "Courtroom Assistant: Nancy Quintanilla", "Courtroom Assistant:"
End Sub

' Toggle one court-identity field across the body and headers/footers.
'   restore = True  (de-anonymize): blank -> real, but only where the real value
'                   isn't already present, so re-running never doubles it.
'   restore = False (re-anonymize): real -> blank; idempotent on its own.
Private Sub SwapCourtField(ByVal oDoc As Document, ByVal restore As Boolean, _
                            ByVal realText As String, ByVal blankText As String)
    Dim findText As String, replText As String
    If restore Then
        findText = blankText: replText = realText
    Else
        findText = realText: replText = blankText
    End If

    CourtSwapInRange oDoc.content, findText, replText, restore, realText

    Dim sec As Section, hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then CourtSwapInRange hf.Range, findText, replText, restore, realText
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then CourtSwapInRange hf.Range, findText, replText, restore, realText
        Next hf
    Next sec
End Sub

' One field swap in one range. On restore, skip when realText is already present:
' blankText is a prefix of realText, so replacing then would double the value.
Private Sub CourtSwapInRange(ByVal rng As Range, ByVal findText As String, _
                              ByVal replText As String, ByVal restore As Boolean, _
                              ByVal realText As String)
    On Error Resume Next
    If restore Then
        If RangeContains(rng, realText) Then Exit Sub
    End If
    Dim r As Range: Set r = rng.Duplicate
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = replText
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' True if s occurs in rng (case-sensitive).
Private Function RangeContains(ByVal rng As Range, ByVal s As String) As Boolean
    On Error Resume Next
    Dim r As Range: Set r = rng.Duplicate
    With r.Find
        .ClearFormatting
        .text = s
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        RangeContains = .Execute
    End With
End Function

'==============================================================================
' GENERIC COURT-PERSONNEL SCRUB  (ported from PDF-Linker register_court_names)
'==============================================================================
' Blank court-identity values by PATTERN rather than by exact string, so the
' scrub works whoever is named:
'   - role labels:   "Judicial Assistant: <Name Words>" -> "Judicial Assistant:"
'   - judicial titles: "Hon./Honorable/Judge/Justice/Commissioner <Name Words>"
'                      -> the title alone (title kept, name blanked)
'   - department numbers, label-anchored: "Department 515" -> "Department",
'     "Dept. 72" -> "Dept." (a bare number elsewhere is never touched)
' Runs across the same stories the replacement pass covers. Over-blanking a
' following capitalized word (e.g. "Judge Presiding" -> "Judge") is accepted:
' for anonymity, blanking too much is the safe direction.
Private Sub ScrubCourtIdentityGeneric(ByVal oDoc As Document)
    On Error Resume Next
    ScrubCourtInRange oDoc.content

    Dim sec As Section, hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then ScrubCourtInRange hf.Range
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then ScrubCourtInRange hf.Range
        Next hf
    Next sec

    If oDoc.Footnotes.count > 0 Then ScrubCourtInRange oDoc.StoryRanges(wdFootnotesStory)
    If oDoc.Endnotes.count > 0 Then ScrubCourtInRange oDoc.StoryRanges(wdEndnotesStory)

    Dim shp As Shape
    For Each shp In oDoc.Shapes
        If shp.TextFrame.HasText Then ScrubCourtInRange shp.TextFrame.TextRange
    Next shp
End Sub

Private Sub ScrubCourtInRange(ByVal rng As Range)
    On Error Resume Next
    ' One capitalized name word (letters, apostrophes -- straight and curly --
    ' and hyphens), for the wildcard patterns below.
    Dim nm As String
    nm = "[A-Z][a-zA-Z'" & ChrW(8217) & "-]@"

    ' Department / courtroom numbers, label-anchored.
    WildcardBlankInRange rng, "Department No. [0-9]{1,3}", "Department"
    WildcardBlankInRange rng, "Department [0-9]{1,3}", "Department"
    WildcardBlankInRange rng, "DEPARTMENT [0-9]{1,3}", "DEPARTMENT"
    WildcardBlankInRange rng, "Dept. [0-9]{1,3}", "Dept."
    WildcardBlankInRange rng, "Dept [0-9]{1,3}", "Dept"

    ' Role labels: blank 1-3 capitalized name words after the label. Longest
    ' first so a two-word name isn't half-eaten by the one-word pattern.
    Dim labels As Variant
    labels = Array("Judge:", "Judicial Assistant:", "Courtroom Assistant:", _
                   "Courtroom Clerk:", "Court Clerk:", "Deputy Clerk:", _
                   "Court Reporter:", "Bailiff:", "Court Attendant:", _
                   "Research Attorney:", "Law Clerk:")
    Dim i As Long, k As Long
    For i = LBound(labels) To UBound(labels)
        For k = 3 To 1 Step -1
            WildcardBlankInRange rng, CStr(labels(i)) & NameWordsPattern(nm, k), CStr(labels(i))
        Next k
    Next i

    ' Judicial titles anywhere in the text: keep the title, blank the name.
    Dim titles As Variant
    titles = Array("Honorable", "Hon.", "Judge", "Justice", "Commissioner")
    For i = LBound(titles) To UBound(titles)
        For k = 3 To 1 Step -1
            WildcardBlankInRange rng, CStr(titles(i)) & NameWordsPattern(nm, k), CStr(titles(i))
        Next k
    Next i
End Sub

' " <name> <name> ..." -- k space-separated capitalized-name-word patterns.
Private Function NameWordsPattern(ByVal nm As String, ByVal k As Long) As String
    Dim s As String, i As Long
    For i = 1 To k
        s = s & " " & nm
    Next i
    NameWordsPattern = s
End Function

Private Sub WildcardBlankInRange(ByVal rng As Range, ByVal findPat As String, _
                                  ByVal repl As String)
    On Error Resume Next
    Dim r As Range: Set r = rng.Duplicate
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findPat
        .Replacement.text = repl
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

'==============================================================================
' LEAK SCAN  (ported from PDF-Linker's quarantine/LEAK machinery)
'==============================================================================
' Count occurrences of the key's REAL values still present after the real->fake
' pass -- each one is a would-be leak in the shared export. Matches inside
' italic text are exempt (protected citations, mirroring the replacement pass's
' protectCitations). Fills sList with up to 8 offending values for the dialog.
Private Function CountRealLeaks(ByVal oDoc As Document, ByRef maps() As Mapping, _
                                 ByVal nMaps As Long, ByRef sList As String) As Long
    Dim total As Long, i As Long, nDistinct As Long
    sList = ""
    For i = 1 To nMaps
        Dim n As Long: n = 0
        n = n + LeaksInRange(oDoc.content, maps(i).real)

        Dim sec As Section, hf As HeaderFooter
        For Each sec In oDoc.Sections
            For Each hf In sec.Headers
                If hf.Exists Then n = n + LeaksInRange(hf.Range, maps(i).real)
            Next hf
            For Each hf In sec.Footers
                If hf.Exists Then n = n + LeaksInRange(hf.Range, maps(i).real)
            Next hf
        Next sec

        On Error Resume Next
        If oDoc.Footnotes.count > 0 Then n = n + LeaksInRange(oDoc.StoryRanges(wdFootnotesStory), maps(i).real)
        If oDoc.Endnotes.count > 0 Then n = n + LeaksInRange(oDoc.StoryRanges(wdEndnotesStory), maps(i).real)
        Dim shp As Shape
        For Each shp In oDoc.Shapes
            If shp.TextFrame.HasText Then n = n + LeaksInRange(shp.TextFrame.TextRange, maps(i).real)
        Next shp
        On Error GoTo 0

        If n > 0 Then
            total = total + n
            nDistinct = nDistinct + 1
            If nDistinct <= 8 Then
                sList = sList & "  - " & maps(i).real & "  (" & n & ")" & vbCrLf
            End If
        End If
    Next i
    If nDistinct > 8 Then
        sList = sList & "  ...and " & (nDistinct - 8) & " more value(s)" & vbCrLf
    End If
    CountRealLeaks = total
End Function

' Occurrences of realText in one range, skipping italic (protected) matches.
Private Function LeaksInRange(ByVal rng As Range, ByVal realText As String) As Long
    On Error Resume Next
    Dim n As Long: n = 0
    If Len(realText) = 0 Then Exit Function
    Dim scan As Range: Set scan = rng.Duplicate
    Do
        With scan.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = realText
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False
            .MatchWholeWord = ShouldWholeWord(realText)
            .MatchWildcards = False
            If Not .Execute Then Exit Do
        End With
        If Not (scan.Font.Italic = True) Then n = n + 1
        scan.Collapse Direction:=wdCollapseEnd
        scan.End = rng.End
        If scan.start >= rng.End Then Exit Do
    Loop
    LeaksInRange = n
End Function

'==============================================================================
' RESIDUAL PSEUDONYM HIGHLIGHTING  (leak safety net)
'==============================================================================
' The pseudonymizer draws every fake from a fixed pool of ~140 words. After
' de-anonymize has swapped the keyed fakes back to real values, highlight any
' pool word STILL present in the document -- even embedded inside a larger word
' -- so a fake the key missed (an odd inflection, a stray occurrence) can't slip
' through unnoticed. Returns the number of occurrences highlighted.
'
' Substring matching (MatchWholeWord = False) is intentional and per the user's
' request. Matching is case-sensitive: only a first-capital form ("Nash") or an
' all-caps form ("NASH") is flagged, so a lowercase occurrence -- whether a
' stray "nash" or the "vance" buried in "advance" -- is left alone. Capitalized
' prose that happens to be a pool word (e.g. "Cedar", "Granite") can still be
' flagged, which is fine for a review aid -- the user clears those by eye.
Private Function HighlightResidualPseudonyms(ByVal oDoc As Document) As Long
    Dim pool As Variant: pool = PseudonymPool()
    Dim total As Long: total = 0

    ' Main body.
    total = total + HighlightPoolInRange(oDoc.content, pool)

    ' Headers and footers, section by section.
    Dim sec As Section, hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then total = total + HighlightPoolInRange(hf.Range, pool)
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then total = total + HighlightPoolInRange(hf.Range, pool)
        Next hf
    Next sec

    ' Footnotes / endnotes, only when present.
    On Error Resume Next
    If oDoc.Footnotes.count > 0 Then _
        total = total + HighlightPoolInRange(oDoc.StoryRanges(wdFootnotesStory), pool)
    If oDoc.Endnotes.count > 0 Then _
        total = total + HighlightPoolInRange(oDoc.StoryRanges(wdEndnotesStory), pool)
    On Error GoTo 0

    HighlightResidualPseudonyms = total
End Function

' Highlight every occurrence of every pool word in one range. Returns the count.
Private Function HighlightPoolInRange(ByVal rng As Range, ByVal pool As Variant) As Long
    Dim total As Long, k As Long
    For k = LBound(pool) To UBound(pool)
        total = total + HighlightWordInRange(rng, CStr(pool(k)))
    Next k
    HighlightPoolInRange = total
End Function

' Highlight occurrences of one pool word in a range, case-sensitively, in either
' its stored first-capital form ("Nash") or an all-caps form ("NASH"); a
' lowercase occurrence is not flagged. Still matches inside larger words.
' Returns the number of occurrences highlighted.
Private Function HighlightWordInRange(ByVal rng As Range, ByVal word As String) As Long
    Dim n As Long
    n = HighlightExact(rng, word)                   ' first-capital form, e.g. "Nash"
    If UCase$(word) <> word Then
        n = n + HighlightExact(rng, UCase$(word))    ' all-caps form, e.g. "NASH"
    End If
    HighlightWordInRange = n
End Function

' One case-sensitive highlight pass for an exact term, matching even inside a
' larger word. Returns the number of occurrences highlighted.
Private Function HighlightExact(ByVal rng As Range, ByVal term As String) As Long
    On Error Resume Next
    Dim r As Range: Set r = rng.Duplicate
    Dim n As Long: n = 0
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = term
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = False        ' flag the word even inside a larger word
        .MatchWildcards = False
        Do While .Execute
            r.HighlightColorIndex = wdPink
            n = n + 1
        Loop
    End With
    HighlightExact = n
End Function

' The fixed pool of fake words the pseudonymizer assigns: person surnames,
' entity/company words, street names, and city/locality names. Built in chunks
' (each under VBA's line-length limit) and split on spaces. Juniper and Larkspur
' appear in more than one category upstream; listed once here.
Private Function PseudonymPool() As Variant
    Dim s As String
    ' Person surnames
    s = "Ashford Bennett Calder Danforth Ellery Fenwick Garrick Halloran Ingram Jarrett Keswick Langley Marlowe Nash Orwell Prescott Quill Radley Sable Thorne Underwood Vance Whitlock Yardley"
    s = s & " Ashby Brandt Corwin Delacroix Everts Fairfax Grantham Holloway Isley Jennings Kingsley Lathrop Merrick Norwood Ackerly Bramble Colfax Denning Emmett Forsythe Gable Hendry Ivers Joplin Kessler Lorne Mabry Nolan Ondine Pruett Renwick Sterling Tolliver Ursin Verity Waverly Alden Beaumont Carrow Delane"
    ' Entity / company words
    s = s & " Aldrin Brightwater Cascadia Dunmore Everline Foxglen Granite Havenwood Ironbridge Juniper Kestrel Lumen Meridian Northgate Oakmont Pinnacle Quarry Redwood Silverpeak Torchlight Umbra Vantage Westmark Zephyr Ambrose Beacon Cobalt Drayton Emberly Falcon Gladstone Harborview Ivory Jetstream Kaldor Larkspur Monarch Nimbus Orion Pembroke"
    ' Street names
    s = s & " Cedar Birch Willow Aspen Laurel Poplar Hawthorn Linden Chestnut Sequoia Cypress Alder Dogwood Hickory Rosewood Foxglove Tamarack Sorrel"
    ' City / locality names
    s = s & " Fairview Brookfield Rosedale Elmwood Kingsbury Northvale Westbrook Clearwater Havenport Stonebridge Marlow Redhill Glenmore Oakhurst Bridgeton"
    PseudonymPool = Split(s)
End Function

'==============================================================================
' SORT  (search term length, descending)
'==============================================================================
' byFake = True sorts by the fake length (de-anonymize searches for fakes);
' byFake = False sorts by the real-value length (re-anonymize searches for
' reals). Longest search term first so a bare token never rewrites part of a
' longer full name.
Private Sub SortMappingsByLenDesc(ByRef maps() As Mapping, ByVal nMaps As Long, _
                                   ByVal byFake As Boolean)
    Dim i As Long, j As Long, tmp As Mapping
    For i = 1 To nMaps - 1
        For j = 1 To nMaps - i
            Dim a As Long, b As Long
            If byFake Then
                a = Len(maps(j).fake): b = Len(maps(j + 1).fake)
            Else
                a = Len(maps(j).real): b = Len(maps(j + 1).real)
            End If
            If a < b Then
                tmp = maps(j)
                maps(j) = maps(j + 1)
                maps(j + 1) = tmp
            End If
        Next j
    Next i
End Sub
