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
'     Category | Real Value | Replacement | Status | Source | Occurrences
' (columns are located by HEADER NAME, so a key from an older version that lacks
' Status still reads). Two Status values change what a row means here:
'   "alt spelling"  the Real Value is a synthetic spelling PDF-Linker invented
'                   to widen matching, sharing the CANONICAL row's fake. Usable
'                   real -> fake only; see Mapping.forwardOnly.
'   "no match"      the fake was never written into the exported text -- the key
'                   pins the binding for a party this batch of filings never
'                   mentioned -- so it cannot appear in a draft written from
'                   those exports.
' Neither can be a de-anonymize MISS, so both are counted out of the result
' dialog's tally rather than left to read as failures.
' "Replacement" is USUALLY the fake that appears in the anonymized draft.
' It can instead hold an operator KEEP instruction -- "no" or "never" (leave
' this Real Value verbatim), or a "[bracketed]" / "{braced}" keep-spec -- which
' is NOT a pseudonym and never appeared in the document. IsKeepDecisionCell
' recognizes those and ReadPseudonymKey drops them; see that function for why
' reading them literally corrupts the text in both directions.
'
' MACROS YOU RUN:
'   DeAnonymizeTentative - locate the key, then replace every fake with its
'                          real value throughout the document (in place).
'   ReAnonymizeTentative - the reverse: replace every real value with its fake,
'                          then export the anonymized text as a NEW Markdown
'                          (.md) file so it is safe to share. Hyperlinks are
'                          stripped (keeping their display text) before the
'                          replacement pass, and the anonymized export is named
'                          with the faked version of the document's own title.
'                          It writes TWO files: alongside the anonymized one it
'                          exports the SAME Markdown with the real names still
'                          in it, under the document's own (real) title, so the
'                          two are a matched pair that differ only in the names.
'                          The real-names copy is for local use -- do NOT share
'                          it; the faked one is the shareable copy.
'                          Both carry the document's markup: tracked insertions
'                          and deletions as {++text++} / {--text--}, and comments
'                          as [c1], [c2] ... with author and text collected at
'                          the end. Commenters are NAMED only in the real-names
'                          file; see the note below.
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
'   - MARKUP. Replacement covers the comments story too, so a comment that names
'     a party is scrubbed like the prose. A comment AUTHOR is document metadata,
'     which no pseudonym key covers and no replacement pass can reach, so the
'     anonymized export labels commenters "Reviewer 1", "Reviewer 2" (stable
'     within the file) and only the real-names export prints their names.
'     Re-anonymize forces the window to show all markup inline for the duration
'     of the run: with deletions drawn in balloons, Word hides them from Find AND
'     from Range.Text, so a real name inside a tracked deletion would be neither
'     replaced nor exported. See ForceInlineMarkup.
'   - AUTOSAVE. Re-anonymize turns AutoSave off before its first edit and back on
'     at the end. This is not about what the macro writes -- it only ever writes
'     .md files -- but about where it does its work: the real -> fake scrub runs
'     ON the open document, so for the length of the run that window holds
'     pseudonyms, and AutoSave would push them into the real cloud file within
'     seconds. The .docx itself is never saved after the scrub (the window is
'     closed unsaved and reopened from disk), so the file keeps its real names.
'     AutoSave is restored on the REOPENED document, and only if it was on to
'     begin with. The one path that deliberately leaves it off is the one where
'     the window still holds fake content -- a failed close, or an error mid-run.
'     Each ending says which happened.
'   - Reads .xlsx via Excel automation. The rare JSON fallback that PDF-Linker
'     writes only when openpyxl is missing is not supported.
'   - SHARED WITH ExportMarkdown.bas. The Markdown reader below is the only one
'     in this template, and the plain (no-anonymization) export macro wants the
'     same output, so a short block of public entry points -- see "SHARED EXPORT
'     ENTRY POINTS" -- hands it the reader, the writer, and the two path helpers.
'     Everything anonymization-specific stays private to this module.
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

' Safety valve for the pseudonym highlight passes. A pool word can be ordinary
' English ("Sterling", "Cedar", "Mercer"), so a long document could in principle
' hold thousands of hits for one term. Highlighting is a review aid -- once a
' term has this many flags the point is already made -- so each pass stops here
' rather than grinding on. Far above any real leak count.
Private Const MAX_HITS_PER_TERM As Long = 500

' How many embedded-only matches make a key row worth REPORTING as extraction
' debris (see LooksLikeFragment). Not a correctness threshold: an embedded-only
' term is skipped however rare it is, because for a whole-token term replacing it
' is a no-op either way. This only decides what the result dialog mentions.
Private Const FRAGMENT_MAX_HITS As Long = 8

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
    ' FORWARD-ONLY row (key Status = "alt spelling"): a synthetic spelling
    ' PDF-Linker invented to widen matching -- a hyphenated surname a line wrap
    ' split open ("Ardeshirpour- Zartoshti"), an OCR near-miss ("Sarra" for
    ' "Sara") -- registered against the CANONICAL value's fake so that every
    ' spelling gets scrubbed. Real -> fake is right, so re-anonymize uses it.
    ' Fake -> real is not: two rows then claim one pseudonym and there is no way
    ' to know which spelling to restore. PDF-Linker marks the non-canonical one
    ' so exactly one row owns each reversal; de-anonymize skips these.
    forwardOnly As Boolean
End Type

' Word settings that silently multiply the cost of a bulk edit run. Background
' repagination re-flows the document after edits, and check-as-you-type re-proofs
' every range we touch; with hundreds of replacements those two cost far more
' than the replacement work itself, and ScreenUpdating = False does NOT disable
' them. Saved so the user's own preferences are restored exactly.
Private Type PerfState
    saved As Boolean
    pagination As Boolean
    spell As Boolean
    grammar As Boolean
End Type

' One searchable story (main body, a non-empty header/footer, the footnote or
' endnote story, or a shape's text frame) plus a lowercased copy of its text.
' Collected ONCE per replacement run so the per-mapping loop neither re-walks the
' Sections/Shapes collections nor sweeps stories the search term isn't in.
Private Type StoryRef
    rng As Range
    lower As String
End Type

' What a tracked change did to a run of text, as the Markdown export reports it.
' Word's own revision types are richer (formatting, paragraph properties, moves);
' RevKindOf folds them down to the two that changed words.
Private Const REV_NONE   As Long = 0
Private Const REV_INSERT As Long = 1
Private Const REV_DELETE As Long = 2

' The window's markup display, saved so the run can force it and put it back.
' Two things depend on it, and both fail SILENTLY when deletions sit in balloons
' instead of inline: Word's Find (so the real -> fake pass would miss a party
' name inside a tracked deletion) and Range.Text (so the export would miss it
' too). See ForceInlineMarkup.
Private Type MarkupView
    saved As Boolean
    showRevisions As Boolean
    showInsDel As Boolean
    mode As Long
    view As Long
End Type

' One tracked change, held as a LIVE Range rather than a pair of offsets. Both
' exports mark the same spans, and the real->fake pass runs between them: Word
' keeps a Range anchored to its text across edits, so a captured span still
' brackets the same words after every replacement, whereas a stored offset would
' point somewhere else the moment a fake was a different length than the real
' name. Capturing also survives Word DROPPING the revision mark itself when the
' text inside it is rewritten -- the shareable copy then still shows the change.
Private Type RevSpan
    rng As Range
    kind As Long
End Type

' Everything one Markdown export accumulates as it walks the body: the footnote
' and comment blocks it is building, the tracked-change index the walk consults
' per character, and the commenters it has seen. Passed ByRef through the export
' so the walk's helpers stay one argument wide.
Private Type ExportState
    fnCount As Long                 ' footnotes/endnotes consumed so far
    fnList As String                ' the trailing "[^n]: ..." block
    cmtCount As Long                ' comments consumed so far
    cmtList As String               ' the trailing comment block
    hideAuthors As Boolean          ' shareable copy: "Reviewer 1", not a name
    authors() As String             ' commenters in first-seen order
    nAuthors As Long
    revStart() As Long              ' tracked-change index, sorted by start
    revEnd() As Long
    revKind() As Long
    nRevs As Long
    revCursor As Long               ' the walk only ever moves this forward
    markedRevs As Long              ' tracked changes that actually reached the file
    lastMarked As Long              ' index of the last one marked (counts once)
End Type

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub DeAnonymizeTentative()
    On Error GoTo ErrH

    Dim oDoc As Document
    Set oDoc = ActiveDocument
    If oDoc Is Nothing Then Exit Sub

    ' Per-phase timings, reported at the end when the run was slow. Guessing at
    ' which phase dominates has been unreliable; this measures it.
    Dim sTimes As String, tMark As Single, tRun As Single
    tMark = Timer

    ' Timed too: locating the key enumerates every file in the document's folder
    ' and stats each one. On a OneDrive/SharePoint folder that can stall on cloud
    ' metadata long before any document work starts, which the earlier timings
    ' would have blamed on nothing at all.
    Dim keyPath As String
    keyPath = ResolveKeyPath(oDoc)
    If Len(keyPath) = 0 Then Exit Sub          ' user cancelled the picker
    sTimes = sTimes & "  Locate key file: " & PhaseSecs(tMark) & vbCrLf
    tMark = Timer

    Dim maps() As Mapping
    Dim nMaps As Long, nKeepRows As Long, nUnusedRows As Long
    If Not ReadPseudonymKey(keyPath, maps, nMaps, nKeepRows, nUnusedRows) Then
        MsgBox "Could not read any real/fake mappings from:" & vbCrLf & vbCrLf & _
               keyPath & vbCrLf & vbCrLf & _
               "Make sure this is the pseudonym_key.xlsx PDF-Linker wrote " & _
               "(with 'Real Value' and 'Replacement' columns).", _
               vbExclamation, "De-Anonymize"
        Exit Sub
    End If
    sTimes = sTimes & "  Read key (Excel): " & PhaseSecs(tMark) & vbCrLf

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
    tRun = Timer                     ' start of the timed work (after the prompts)

    ' Background repagination and check-as-you-type re-process the document after
    ' EVERY edit; with hundreds of replacements they cost more than the whole
    ' replacement pass. ScreenUpdating alone does not disable them.
    Dim perf As PerfState: perf = PerfSuppress()
    Dim nRevs As Long: nRevs = RevisionLoad(oDoc)

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
    SetPhase "Removing hyperlinks"
    tMark = Timer
    StripHyperlinksEverywhere oDoc
    sTimes = sTimes & "  Strip hyperlinks: " & PhaseSecs(tMark) & vbCrLf

    ' Deliberately NO custom UndoRecord: wrapping every replacement across a large
    ' document (dozens of terms, each many hits) into one custom undo record
    ' overflows and crashes Word. Word still records normal (multi-step) undo.
    Dim distinctHits As Long, i As Long, nFragmentSkips As Long, nAmbiguous As Long
    tMark = Timer
    distinctHits = ReplaceAllMappings(oDoc, maps, nMaps, True, False, "De-anonymizing", _
                                      nFragmentSkips, nAmbiguous)
    sTimes = sTimes & "  Replace " & nMaps & " mapping(s): " & PhaseSecs(tMark) & vbCrLf

    On Error Resume Next
    oDoc.AutoSaveOn = prevAutoSave
    On Error GoTo ErrH
    oDoc.TrackRevisions = prevTrack

    ' Restore the court-identity header (Department 515, judge, courtroom staff).
    SetPhase "Restoring court header"
    tMark = Timer
    ApplyCourtIdentity oDoc, True
    sTimes = sTimes & "  Court header: " & PhaseSecs(tMark) & vbCrLf

    ' Safety net: flag any pseudonym-pool word still present (even inside a
    ' larger word) in pink, so a fake the key missed doesn't slip through.
    SetPhase "Scanning for leftover pseudonyms"
    tMark = Timer
    Dim nFlags As Long
    nFlags = HighlightResidualPseudonyms(oDoc)
    sTimes = sTimes & "  Pseudonym scan: " & PhaseSecs(tMark) & vbCrLf

    SetPhase ""
    PerfRestore perf
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
           "(department, judge, staff)." & UnusedRowsNote(nUnusedRows) & _
           AmbiguousNote(nAmbiguous) & KeepRowsNote(nKeepRows) & _
           FragmentNote(nFragmentSkips) & sFlagLine & _
           TimingNote(sTimes, tRun, nRevs) & vbCrLf & vbCrLf & _
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
    PerfRestore perf                     ' never leave proofing/pagination off
    Application.ScreenUpdating = True
    Application.StatusBar = False        ' don't strand a progress message
    MsgBox "De-Anonymize hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & eN & ": " & eD, vbExclamation, "De-Anonymize"
End Sub

'==============================================================================
' RE-ANONYMIZE  (reverse: real -> fake, exported as a clean Markdown file,
'                plus a matching real-names Markdown file)
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
    Dim nMaps As Long, nKeepRows As Long
    If Not ReadPseudonymKey(keyPath, maps, nMaps, nKeepRows) Then
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

    Dim i As Long

    ' Run-and-done: no Save-As dialog and no confirmation. Both .md files are
    ' written automatically next to the document (its local synced folder).
    ' The anonymized one is named with the FAKED version of the document's own
    ' title, so the export is recognizable but carries pseudonyms, not real
    ' party names; the real-names one keeps the document's own title. Each
    ' file's name therefore says which version it holds. DocFolderLocal maps a
    ' SharePoint/OneDrive URL to the writable local sync folder; if the folder
    ' can't be resolved (never-saved doc) it falls back to Documents.
    Dim savePath As String, realPath As String
    Dim outFolder As String: outFolder = ExportFolderFor(oDoc)
    savePath = outFolder & "\" & FakedDocTitle(oDoc, maps, nMaps) & ".md"
    realPath = outFolder & "\" & RealDocTitle(oDoc) & ".md"

    ' A title that holds no real value from the key fakes to itself, and then
    ' both exports claim the same path: one would silently overwrite the other,
    ' leaving a file whose name says nothing about which version survived. Same
    ' hazard if the export would land on the open document itself. Either way,
    ' push the real-names copy aside rather than clobber anything.
    If StrComp(savePath, realPath, vbTextCompare) = 0 Or _
       StrComp(realPath, SafeFullName(oDoc), vbTextCompare) = 0 Then
        realPath = outFolder & "\" & RealDocTitle(oDoc) & " (real names).md"
    End If

    ' From this point on, no automatic de-anonymize for the rest of the Word
    ' session (set even if the run errors out partway -- fail safe). This also
    ' keeps the close hook from firing when we discard the scratch edits below.
    g_ReAnonThisSession = True

    ' Save the user's pending edits FIRST, while the document still holds only
    ' real-name content. The run ends by discarding the scratch window and
    ' reloading the original from disk, so anything unsaved would otherwise be
    ' lost -- this saves it so the user doesn't have to remember to. Safe: no
    ' fake content exists yet. Skipped for a never-saved document (nothing to
    ' save to) and if the doc is unchanged.
    On Error Resume Next
    If Len(oDoc.path) > 0 And Not oDoc.Saved Then oDoc.Save
    On Error GoTo ErrH

    Application.ScreenUpdating = False
    ' See PerfSuppress: repagination and check-as-you-type re-process the
    ' document after every edit and dominate a bulk run.
    Dim perf As PerfState: perf = PerfSuppress()

    ' The real->fake scrub runs IN MEMORY on the open document ONLY -- so italic
    ' cited authorities can be detected and preserved -- and the result is then
    ' read out as Markdown. The Word file itself is NEVER written: AutoSave is
    ' disabled before the first edit, no Save/SaveAs is ever issued, and the
    ' in-memory edits are discarded at the end (the window is reloaded from the
    ' untouched original on disk).
    Dim prevTrack As Boolean: prevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False
    On Error Resume Next
    ' Remember whether AutoSave was on, so the reloaded window can be handed back
    ' in the state the user keeps it in rather than in the state this macro needed
    ' it in. Reading it can raise (older Word, or a document somewhere AutoSave
    ' doesn't cover); False then means "turn nothing on later", the safe answer.
    Dim prevAutoSave As Boolean: prevAutoSave = oDoc.AutoSaveOn
    oDoc.AutoSaveOn = False              ' must precede edits: AutoSave would
    On Error GoTo ErrH                   ' push real->fake edits to the ORIGINAL
    Dim bStateSaved As Boolean: bStateSaved = True   ' ErrH may now restore

    ' Show every tracked change inline before anything reads or replaces text.
    ' With deletions in balloons, a real name inside one is invisible to Word's
    ' Find AND to Range.Text -- it would neither be replaced nor exported, which
    ' is a leak in the first case and a silent omission in the second.
    Dim mv As MarkupView: mv = ForceInlineMarkup()
    Dim nComments As Long
    On Error Resume Next
    nComments = oDoc.Comments.count
    On Error GoTo ErrH

    ' Strip hyperlinks (keeping display text) before replacing and exporting:
    ' link targets can carry real names/paths the Markdown must not contain,
    ' and a real name inside a link's display text is replaced more reliably
    ' once the link is gone. Runs on the in-memory scratch copy only -- the
    ' original file (reloaded below) keeps its links.
    StripHyperlinksEverywhere oDoc

    ' Capture the tracked changes ONCE, here: after the hyperlink strip and before
    ' the first replacement. Both exports mark this same set, so the anonymized
    ' copy shows exactly the changes the real-names copy shows -- it does not
    ' re-read Word's revision marks, which the replacement pass can rewrite or
    ' drop when it edits the text inside one.
    Dim spans() As RevSpan
    Dim nSpans As Long: nSpans = CaptureRevisions(oDoc, spans)

    ' Export the real-names copy FIRST, before a single value is swapped. Same
    ' Markdown reader, same hyperlink-stripped body, same tracked changes, so the
    ' pair differs only in the names themselves -- the anonymized file can be read
    ' against this one line for line. It costs a second walk of the document,
    ' which is cheap next to the replacement pass.
    SetPhase "Writing the real-names export", "Re-Anonymize"
    Dim markedReal As Long
    WriteUtf8NoBom realPath, DocToMarkdown(oDoc, False, spans, nSpans, markedReal)

    ' Reverse direction: replace each real value with its fake. protectCitations
    ' leaves names inside italic cited authorities alone, so a party surname that
    ' also names a published case isn't rewritten in the shared copy. No custom
    ' undo record (it overflows and crashes Word on large documents).
    Dim distinctHits As Long
    distinctHits = ReplaceAllMappings(oDoc, maps, nMaps, False, True, "Re-anonymizing")

    ' Blank the court-identity header (Department 515, judge, courtroom staff) so
    ' the shared copy doesn't reveal them. ApplyCourtIdentity also scrubs the
    ' body, which is what the Markdown export reads.
    ApplyCourtIdentity oDoc, False

    ' Read the now-anonymized body out as Markdown and write it to disk (UTF-8,
    ' no BOM). This and the real-names export above are the only files the macro
    ' writes; the .docx is still never touched.
    SetPhase "Writing the anonymized export", "Re-Anonymize"
    Dim md As String, markedFake As Long
    ' True: commenters as "Reviewer n". Same spans as the real-names export.
    md = DocToMarkdown(oDoc, True, spans, nSpans, markedFake)
    WriteUtf8NoBom savePath, md
    SetPhase ""                          ' don't strand a progress message

    ' Read the finished file back for real values that survived inside a tracked
    ' change -- the one span the replacement pass can quietly fail to reach.
    Dim nResidual As Long
    nResidual = ResidualsInMarkup(md, maps, nMaps)

    RestoreMarkupView mv                 ' hand the user's markup view back

    ' Discard the in-memory fake edits: reload the window from the untouched
    ' original so the user is back on the real-names document and a stray Ctrl+S
    ' can never push fakes into it. If the document was never saved to disk
    ' (no path to reopen), leave the scratch window in place with AutoSave off
    ' and warn instead.
    Dim origPath As String: origPath = ""
    On Error Resume Next
    If Len(oDoc.path) > 0 Then origPath = oDoc.FullName
    On Error GoTo ErrH

    PerfRestore perf
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
            Dim reDoc As Document
            Set reDoc = Documents.Open(FileName:=origPath, AddToRecentFiles:=False)
            reloaded = (Err.Number = 0) And Not (reDoc Is Nothing)
            Err.Clear

            ' Hand AutoSave back. The document it was turned off on is gone --
            ' closed without saving -- and this is a fresh open of the untouched
            ' original, so there is no fake content left anywhere for AutoSave to
            ' push. It has to be set explicitly rather than left to Word: the
            ' toggle is remembered per file, so without this the user's own
            ' tentative quietly stops auto-saving from here on, and the macro
            ' would have changed a setting it was only ever meant to borrow.
            ' Only turned back ON, never off: someone who keeps AutoSave off
            ' (prevAutoSave False) keeps it off.
            If reloaded And prevAutoSave Then
                reDoc.AutoSaveOn = True
                Err.Clear
            End If
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

    ' AutoSave is off for the whole run because the scrub happens IN the open
    ' document (see above), and AutoSave would push those fake names straight into
    ' the real cloud file. Say what became of it, in each of the three endings:
    ' silently borrowing a user's setting and not saying whether it came back is
    ' how someone ends up working for a week with AutoSave off.
    Dim tail As String
    If reloaded Then
        tail = "Your Word window has been reloaded from the original file " & _
               "(real names), which was never modified."
        If prevAutoSave Then
            tail = tail & " AutoSave was off while the scrub ran and has been " & _
                   "turned back on for the reloaded document."
        End If
    ElseIf closedOK Then
        tail = "The scratch window (fake names) was discarded, but the original " & _
               "could not be reopened automatically -- open it yourself from:" & _
               vbCrLf & origPath & vbCrLf & "It was never modified."
        If prevAutoSave Then
            tail = tail & " Check that AutoSave is on once you reopen it: it was " & _
                   "turned off while the scrub ran, and this run never got the " & _
                   "chance to turn it back on."
        End If
    Else
        tail = "This window still holds the re-anonymized (fake) content and was " & _
               "NOT saved -- close it WITHOUT saving to discard those edits and " & _
               "get back to the untouched original."
        If prevAutoSave Then
            tail = tail & " AutoSave has been LEFT off on purpose: with fake " & _
                   "names in this window, turning it back on now is what would " & _
                   "push them into the real file. Turn it on again once you have " & _
                   "closed this window without saving and reopened the original."
        End If
    End If

    MsgBox "Re-anonymized: replaced " & distinctHits & " of " & nMaps & _
           " value(s) and blanked the court-identity header." & _
           KeepRowsNote(nKeepRows) & vbCrLf & vbCrLf & _
           "Names inside italic cited case names were left as-is so a party " & _
           "surname that also names a published case wasn't rewritten -- check " & _
           "any italicized cites if a real party name should have been replaced." & _
           MarkupNote(nComments, markedReal, markedFake, nResidual) & vbCrLf & vbCrLf & _
           "Saved an anonymized Markdown file (safe to share) to:" & vbCrLf & _
           savePath & vbCrLf & vbCrLf & _
           "and the same text with the REAL names (keep this one local) to:" & _
           vbCrLf & realPath & vbCrLf & vbCrLf & _
           tail, vbInformation, "Re-Anonymize"
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
    PerfRestore perf                     ' never leave proofing/pagination off
    RestoreMarkupView mv                 ' never leave the markup view forced
    Application.ScreenUpdating = True
    Application.StatusBar = False        ' don't strand a progress message
    MsgBox "Re-Anonymize hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & reN & ": " & reD & vbCrLf & vbCrLf & _
           "If the error happened before the .md files were written, this window " & _
           "may hold partial re-anonymize edits that were NOT saved anywhere -- " & _
           "close it WITHOUT saving to get back to the untouched original. " & _
           "(AutoSave was left off for the same reason -- turn it back on once " & _
           "you have reopened the original.) The real-names export " & _
           "is written before any name is swapped, so it may exist even though " & _
           "the anonymized one does not -- check for both before sharing " & _
           "anything.", _
           vbExclamation, "Re-Anonymize"
End Sub

' The document's title (filename without extension) with every real value from
' the key replaced by its fake, longest real first (the maps are already sorted
' that way when this is called). Matching is case-insensitive and the fake is
' recased to mirror the casing found, same as the body replacement. Characters
' Windows forbids in filenames are folded to "-" as a safety net.
Private Function FakedDocTitle(ByVal oDoc As Document, _
                                ByRef maps() As Mapping, _
                                ByVal nMaps As Long) As String
    Dim t As String: t = DocBaseTitle(oDoc)

    Dim i As Long
    For i = 1 To nMaps
        t = ReplaceCIString(t, maps(i).real, maps(i).fake)
    Next i

    FakedDocTitle = SanitizeFileTitle(t, "Anonymized Draft")
End Function

' The document's own title, real names and all: the name for the export that
' keeps them. No key values are swapped -- that is the whole point of the pair,
' the faked title names the shareable copy and this one names the local copy.
Private Function RealDocTitle(ByVal oDoc As Document) As String
    RealDocTitle = DocumentExportTitle(oDoc, "Original Draft")
End Function

' The document's filename without its extension; "" if the name can't be read.
Private Function DocBaseTitle(ByVal oDoc As Document) As String
    Dim t As String
    On Error Resume Next
    t = oDoc.name
    On Error GoTo 0

    Dim dotPos As Long: dotPos = InStrRev(t, ".")
    If dotPos > 1 Then t = Left$(t, dotPos - 1)

    DocBaseTitle = t
End Function

' Fold the characters Windows forbids in a filename to "-", trimming the result;
' falls back to the caller's placeholder when nothing usable is left.
Private Function SanitizeFileTitle(ByVal t As String, _
                                    ByVal fallback As String) As String
    Dim k As Long, ch As String
    For k = 1 To Len(t)
        ch = Mid$(t, k, 1)
        If InStr(1, "\/:*?""<>|", ch) > 0 Then Mid$(t, k, 1) = "-"
    Next k

    SanitizeFileTitle = Trim$(t)
    If Len(SanitizeFileTitle) = 0 Then SanitizeFileTitle = fallback
End Function

' The document's full path, or "" if it can't be read. A never-saved document
' answers with a bare name ("Document1"), which is fine: the caller only uses
' this to make sure an export doesn't land on the open file itself. Public for
' ExportMarkdown.bas, which makes the same check.
Public Function SafeFullName(ByVal oDoc As Document) As String
    Dim s As String
    On Error Resume Next
    s = oDoc.FullName
    On Error GoTo 0
    SafeFullName = s
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
' SHARED EXPORT ENTRY POINTS  (used by ExportMarkdown.bas)
'==============================================================================
' The Markdown reader below is this template's only one, and the plain export
' macro in ExportMarkdown.bas wants exactly what it produces -- same headings,
' same {++inserted++} / {--deleted--} notation, same comment block. These four
' wrappers (plus SafeFullName above) are the whole of what that module may call;
' the reader's internals, the key reading, and the replacement passes all stay
' private here.
'
' They are functions with arguments, so none of them shows up on Word's Alt+F8
' macro list: the only macros a user runs from this module are still
' DeAnonymizeTentative and ReAnonymizeTentative.

' Read a document out as Markdown, with the tracked changes and comments it
' carries. Wraps the three steps every export needs: force all markup inline
' (Word hides deletions in balloons from Range.Text, so a change would silently
' vanish from the file), capture the tracked changes, walk the body -- and hand
' the user's markup view back afterwards, error or not. Makes no edit to the
' document. markedOut reports how many tracked changes reached the file.
'
' hideAuthors labels commenters "Reviewer 1", "Reviewer 2" instead of naming
' them; see DocToMarkdown for why the shareable copy needs that.
Public Function BuildMarkdownFromDocument(ByVal oDoc As Document, _
                                          ByVal hideAuthors As Boolean, _
                                          Optional ByRef markedOut As Long) As String
    Dim mv As MarkupView: mv = ForceInlineMarkup()

    On Error GoTo Fail
    Dim spans() As RevSpan
    Dim nSpans As Long: nSpans = CaptureRevisions(oDoc, spans)
    BuildMarkdownFromDocument = DocToMarkdown(oDoc, hideAuthors, spans, nSpans, _
                                              markedOut)
    RestoreMarkupView mv
    Exit Function

Fail:
    Dim n As Long, d As String: n = Err.Number: d = Err.Description
    RestoreMarkupView mv                 ' never leave the markup view forced
    Err.Raise n, "BuildMarkdownFromDocument", d
End Function

' Write one export to disk as UTF-8 without a BOM.
Public Sub WriteMarkdownFile(ByVal path As String, ByVal text As String)
    WriteUtf8NoBom path, text
End Sub

' The folder an export belongs in: the document's own folder, mapped to the local
' synced copy when Word reports a SharePoint/OneDrive URL (ADODB.Stream cannot
' write to a URL). A never-saved document has no folder, so it falls back to
' Documents.
Public Function ExportFolderFor(ByVal oDoc As Document) As String
    Dim f As String: f = DocFolderLocal(oDoc)
    If Len(f) = 0 Then f = Environ$("USERPROFILE") & "\Documents"
    ExportFolderFor = f
End Function

' The document's own title as a filename: its name without the extension, with
' the characters Windows forbids folded to "-", or the caller's fallback when
' nothing usable is left.
Public Function DocumentExportTitle(ByVal oDoc As Document, _
                                    ByVal fallback As String) As String
    DocumentExportTitle = SanitizeFileTitle(DocBaseTitle(oDoc), fallback)
End Function

'==============================================================================
' MARKDOWN EXPORT  (read the in-memory, already-anonymized body out as Markdown)
'==============================================================================
' Convert the document's main body to Markdown:
'   - paragraph styles Heading 1..6 / Title  ->  # .. ###### / #
'   - list paragraphs                        ->  "- " (bullet) or the number label
'   - bold / italic runs                     ->  **bold**, *italic*, ***both***
'   - footnote/endnote reference marks        ->  [^n], with the note texts
'                                                 collected into a trailing block
'   - tracked insertions / deletions         ->  {++inserted++} / {--deleted--}
'   - comments                               ->  [cn] at the anchor, with author
'                                                 and text in a trailing block
' Only the body is exported: headers/footers carry the court identity (already
' blanked) and have no place in Markdown. Formatting Markdown can't express
' (alignment, tab leaders in the caption, tables) is dropped but the text is
' kept. Text boxes/shapes are not exported.
'
' hideAuthors is set for the shareable copy: a commenter's name is document
' METADATA, which no pseudonym key covers and the real -> fake pass cannot reach,
' so naming them there would leak a real person past everything this module does.
' The anonymized copy gets "Reviewer 1", "Reviewer 2" -- stable within the file,
' so a reader can still tell two commenters apart -- and the real-names copy gets
' the actual names.
Private Function DocToMarkdown(ByVal oDoc As Document, _
                               ByVal hideAuthors As Boolean, _
                               ByRef spans() As RevSpan, _
                               ByVal nSpans As Long, _
                               Optional ByRef markedOut As Long) As String
    Dim st As ExportState
    st.hideAuthors = hideAuthors
    BuildRevisionIndex st, spans, nSpans

    Dim sb As String
    Dim firstBlock As Boolean: firstBlock = True

    Dim p As Paragraph
    For Each p In oDoc.content.Paragraphs
        Dim line As String
        line = ParagraphToMarkdown(oDoc, p, st)
        If Len(line) > 0 Then
            If Not firstBlock Then sb = sb & vbCrLf & vbCrLf
            sb = sb & line
            firstBlock = False
        End If
    Next p

    AppendUnanchoredComments oDoc, st

    If Len(st.fnList) > 0 Then sb = sb & vbCrLf & vbCrLf & st.fnList
    If Len(st.cmtList) > 0 Then
        sb = sb & vbCrLf & vbCrLf & "## Comments" & vbCrLf & vbCrLf & st.cmtList
    End If

    markedOut = st.markedRevs
    DocToMarkdown = MarkupLegend(st) & sb & vbCrLf
End Function

' A one-line note at the top of the file saying how markup is written, so the
' braces read as notation rather than as something the judge typed. Emitted only
' when the document actually carried markup.
Private Function MarkupLegend(ByRef st As ExportState) As String
    Dim parts As String
    If st.markedRevs > 0 Then
        parts = "tracked changes as {++insertions++} and {--deletions--}"
    End If
    If st.cmtCount > 0 Then
        If Len(parts) > 0 Then parts = parts & "; "
        parts = parts & "comments as [c1], [c2] ... with the text collected at the end"
    End If
    If Len(parts) = 0 Then Exit Function

    MarkupLegend = "> This export marks " & parts & "." & vbCrLf & vbCrLf
End Function

' One body paragraph -> one Markdown block (or "" for an empty paragraph, which
' just becomes block separation). List prefix wins over heading prefix.
Private Function ParagraphToMarkdown(ByVal oDoc As Document, ByVal p As Paragraph, _
                                     ByRef st As ExportState) As String
    ' Paragraph content without the trailing paragraph mark.
    Dim wr As Range: Set wr = p.Range.Duplicate
    If wr.Characters.count >= 1 Then wr.MoveEnd wdCharacter, -1

    Dim inner As String
    inner = InlineMarkdown(oDoc, wr, st)
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
' character by character. A paragraph carrying a tracked change is always walked:
' the mark has to land on the changed run, which needs per-character positions.
' That question goes to the captured index, never to wr.Revisions -- by the time
' the anonymized export runs, Word may have dropped its own mark from a span this
' export still means to show, and the paragraph would take the fast path and lose
' the change.
Private Function InlineMarkdown(ByVal oDoc As Document, ByVal wr As Range, _
                                ByRef st As ExportState) As String
    Dim txt As String: txt = wr.text
    If Len(txt) = 0 Then Exit Function

    Dim boldUniform As Boolean, italicUniform As Boolean
    boldUniform = (wr.Font.Bold <> wdUndefined)
    italicUniform = (wr.Font.Italic <> wdUndefined)

    If boldUniform And italicUniform And Not SpanTouches(st, wr.Start, wr.End) Then
        InlineMarkdown = Emph(MapText(oDoc, txt, st), _
                              (wr.Font.Bold = True), (wr.Font.Italic = True))
    Else
        InlineMarkdown = WalkRuns(oDoc, wr, st)
    End If
End Function

' Walk a mixed-format range character by character, grouping consecutive
' same-format characters into runs and wrapping each run in its emphasis markers
' and its tracked-change markers. A run breaks on a change of ANY of the three,
' so an insertion inside an italic case name comes out nested, not flattened.
Private Function WalkRuns(ByVal oDoc As Document, ByVal wr As Range, _
                          ByRef st As ExportState) As String
    Dim result As String, runText As String
    Dim curB As Long, curI As Long, curR As Long
    curB = -1: curI = -1: curR = -1         ' -1 = no run started yet
    Dim n As Long: n = wr.Characters.count
    Dim i As Long
    For i = 1 To n
        Dim ch As Range: Set ch = wr.Characters(i)
        Dim c As String: c = ch.text
        If c = Chr$(2) Or c = Chr$(5) Then  ' footnote/endnote or comment mark
            If Len(runText) > 0 Then
                result = result & EmitRun(runText, curB, curI, curR)
                runText = ""
            End If
            If c = Chr$(2) Then
                result = result & EmitNote(oDoc, st)
            Else
                result = result & EmitComment(oDoc, st)
            End If
            curB = -1: curI = -1: curR = -1
        Else
            Dim b As Long, it As Long, rv As Long
            b = IIf(ch.Font.Bold = True, 1, 0)
            it = IIf(ch.Font.Italic = True, 1, 0)
            rv = RevKindAt(st, ch.Start)
            If b <> curB Or it <> curI Or rv <> curR Then
                If Len(runText) > 0 Then
                    result = result & EmitRun(runText, curB, curI, curR)
                    runText = ""
                End If
                curB = b: curI = it: curR = rv
            End If
            runText = runText & MapChar(c)
        End If
    Next i
    If Len(runText) > 0 Then result = result & EmitRun(runText, curB, curI, curR)
    WalkRuns = result
End Function

' One finished run: emphasis markers first, then the tracked-change wrapper
' around them, so the change marker always brackets the whole run.
Private Function EmitRun(ByVal s As String, ByVal b As Long, _
                         ByVal it As Long, ByVal rv As Long) As String
    EmitRun = MarkRevision(Emph(s, b = 1, it = 1), rv)
End Function

' Wrap a run in CriticMarkup so both a reader and Claude can see what the tracked
' change was, rather than reading an edit-in-progress as settled text. Text with
' no revision on it comes back untouched.
Private Function MarkRevision(ByVal s As String, ByVal kind As Long) As String
    If Len(s) = 0 Then
        MarkRevision = s
    ElseIf kind = REV_INSERT Then
        MarkRevision = "{++" & s & "++}"
    ElseIf kind = REV_DELETE Then
        MarkRevision = "{--" & s & "--}"
    Else
        MarkRevision = s                 ' REV_NONE, or a run that never started
    End If
End Function

' Map a plain (single-format) text run to Markdown, translating footnote and
' comment reference marks and per-character specials.
Private Function MapText(ByVal oDoc As Document, ByVal s As String, _
                         ByRef st As ExportState) As String
    Dim res As String, i As Long
    For i = 1 To Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        If c = Chr$(2) Then
            res = res & EmitNote(oDoc, st)
        ElseIf c = Chr$(5) Then
            res = res & EmitComment(oDoc, st)
        Else
            res = res & MapChar(c)
        End If
    Next i
    MapText = res
End Function

' Consume the next footnote/endnote reference (in document order): append its
' text to the trailing block and return the "[^n]" inline marker.
Private Function EmitNote(ByVal oDoc As Document, ByRef st As ExportState) As String
    st.fnCount = st.fnCount + 1
    Dim body As String: body = FlattenNote(NoteText(oDoc, st.fnCount))
    If Len(st.fnList) > 0 Then st.fnList = st.fnList & vbCrLf
    st.fnList = st.fnList & "[^" & st.fnCount & "]: " & body
    EmitNote = "[^" & st.fnCount & "]"
End Function

'------------------------------------------------------------------------------
' COMMENTS
'------------------------------------------------------------------------------
' Consume the next comment (document order, the same way notes are consumed) and
' return its "[cn]" inline marker. Word marks each comment's anchor in the body
' text with Chr(5), replies included, so the n-th mark is the n-th comment.
Private Function EmitComment(ByVal oDoc As Document, ByRef st As ExportState) As String
    EmitComment = "[c" & AddCommentEntry(oDoc, st, "") & "]"
End Function

' Any comment the walk never met -- one anchored in a header, a text box, or left
' unanchored by an edit -- still goes into the block, flagged. A comment silently
' missing from the export is the one failure worth being noisy about: the reader
' has no way to notice it isn't there.
Private Sub AppendUnanchoredComments(ByVal oDoc As Document, ByRef st As ExportState)
    Dim total As Long
    On Error Resume Next
    total = oDoc.Comments.count
    On Error GoTo 0

    Do While st.cmtCount < total
        AddCommentEntry oDoc, st, " (not anchored in the body)"
    Loop
End Sub

' Append one comment to the trailing block and return its number. The block is a
' Markdown list, not "[cn]: text" -- that shape is a link reference definition,
' which a renderer swallows whole.
Private Function AddCommentEntry(ByVal oDoc As Document, ByRef st As ExportState, _
                                 ByVal note As String) As Long
    st.cmtCount = st.cmtCount + 1
    AddCommentEntry = st.cmtCount

    Dim body As String, who As String
    On Error Resume Next
    If st.cmtCount <= oDoc.Comments.count Then
        body = FlattenNote(oDoc.Comments(st.cmtCount).Range.text)
        who = AuthorLabel(st, oDoc.Comments(st.cmtCount).Author)
        If IsCommentReply(oDoc.Comments(st.cmtCount)) Then note = note & " (reply)"
    End If
    On Error GoTo 0
    If Len(who) = 0 Then who = "Unknown"

    If Len(st.cmtList) > 0 Then st.cmtList = st.cmtList & vbCrLf
    st.cmtList = st.cmtList & "- **[c" & st.cmtCount & "]** " & who & note & _
                 " -- " & body
End Function

' The name to print for a commenter: the real one locally, a stable "Reviewer n"
' in the shareable copy (see DocToMarkdown). Either way the author is registered
' in first-seen order, so the same person keeps the same label all file long.
Private Function AuthorLabel(ByRef st As ExportState, ByVal author As String) As String
    author = Trim$(author)
    If Len(author) = 0 Then author = "Unknown"
    If st.nAuthors = 0 Then ReDim st.authors(1 To 8)

    Dim i As Long
    For i = 1 To st.nAuthors
        If StrComp(st.authors(i), author, vbTextCompare) = 0 Then Exit For
    Next i

    If i > st.nAuthors Then
        If i > UBound(st.authors) Then ReDim Preserve st.authors(1 To UBound(st.authors) + 8)
        st.authors(i) = author
        st.nAuthors = i
    End If

    If st.hideAuthors Then
        AuthorLabel = "Reviewer " & i
    Else
        AuthorLabel = author
    End If
End Function

' True when the comment is a reply in a threaded conversation. Comment.Ancestor
' only exists in newer Word object models, so it is reached late-bound (the
' parameter is Object): an older Word raises here instead of refusing to compile
' the whole project.
Private Function IsCommentReply(ByVal cmt As Object) As Boolean
    Dim anc As Object
    On Error Resume Next
    Set anc = cmt.Ancestor
    On Error GoTo 0
    IsCommentReply = Not (anc Is Nothing)
End Function

'------------------------------------------------------------------------------
' TRACKED CHANGES
'------------------------------------------------------------------------------
' Capture the body's tracked insertions and deletions as live Ranges, ONCE, before
' anything replaces a name. Both exports then mark the same set of changes: see
' RevSpan for why the span has to be a Range and not a pair of offsets.
'
' Only content changes are captured; RevKindOf drops the rest. The array is
' always dimensioned, so a document with no tracked changes still passes a usable
' (empty) array down to the export.
Private Function CaptureRevisions(ByVal oDoc As Document, _
                                  ByRef spans() As RevSpan) As Long
    Dim n As Long: n = 0
    ReDim spans(1 To 64)

    On Error Resume Next
    Dim rev As Revision
    For Each rev In oDoc.content.Revisions
        Dim kind As Long: kind = RevKindOf(rev.Type)
        If kind <> REV_NONE Then
            n = n + 1
            If n > UBound(spans) Then ReDim Preserve spans(1 To UBound(spans) + 64)
            Set spans(n).rng = rev.Range.Duplicate
            spans(n).kind = kind
        End If
    Next rev
    On Error GoTo 0

    CaptureRevisions = n
End Function

' Turn the captured spans into a position index for one export, reading each
' Range's CURRENT bounds -- which is the point: by the time the anonymized export
' runs, every span has moved by however much the replacements before it changed
' the text, and Word has already done that arithmetic for us.
'
' The walk then asks RevKindAt per character instead of reaching into the
' Revisions collection a character at a time, which is the difference between one
' pass over the revisions and one COM call per character of the document.
Private Sub BuildRevisionIndex(ByRef st As ExportState, _
                               ByRef spans() As RevSpan, ByVal nSpans As Long)
    st.nRevs = 0
    st.revCursor = 1
    ReDim st.revStart(1 To 64)
    ReDim st.revEnd(1 To 64)
    ReDim st.revKind(1 To 64)

    On Error Resume Next
    Dim i As Long
    For i = 1 To nSpans
        Dim sStart As Long, sEnd As Long
        sStart = -1: sEnd = -1
        sStart = spans(i).rng.Start
        sEnd = spans(i).rng.End

        ' A span whose text was replaced away collapses to nothing; there is no
        ' run left to mark, so drop it rather than index an empty range. Spans
        ' that overlap (an insertion one reviewer later deleted) are resolved by
        ' the walk in start order: the first covers its own extent, the second
        ' picks up from where the first ends. Every character still gets marked.
        If sEnd > sStart Then
            If st.nRevs + 1 > UBound(st.revStart) Then
                ReDim Preserve st.revStart(1 To UBound(st.revStart) + 64)
                ReDim Preserve st.revEnd(1 To UBound(st.revEnd) + 64)
                ReDim Preserve st.revKind(1 To UBound(st.revKind) + 64)
            End If
            st.nRevs = st.nRevs + 1

            ' Keep the index sorted by start. Word returns revisions in story
            ' order and edits preserve that order, so this shifts nothing in
            ' practice -- but RevKindAt's cursor only ever moves forward, and one
            ' out-of-order entry would be stepped past and lost from the export.
            Dim j As Long: j = st.nRevs
            Do While j > 1
                If st.revStart(j - 1) <= sStart Then Exit Do
                st.revStart(j) = st.revStart(j - 1)
                st.revEnd(j) = st.revEnd(j - 1)
                st.revKind(j) = st.revKind(j - 1)
                j = j - 1
            Loop
            st.revStart(j) = sStart
            st.revEnd(j) = sEnd
            st.revKind(j) = spans(i).kind
        End If
    Next i
    On Error GoTo 0
End Sub

' True when a captured tracked change overlaps this range. Advances the cursor
' only past spans that end before the range does -- the same step RevKindAt would
' take at the range's first character -- so asking never costs the walk anything.
Private Function SpanTouches(ByRef st As ExportState, _
                             ByVal rStart As Long, ByVal rEnd As Long) As Boolean
    Do While st.revCursor <= st.nRevs
        If st.revEnd(st.revCursor) > rStart Then Exit Do
        st.revCursor = st.revCursor + 1
    Loop
    If st.revCursor > st.nRevs Then Exit Function

    SpanTouches = (st.revStart(st.revCursor) < rEnd)
End Function

' The tracked change covering a character position, or REV_NONE. The export walks
' the body front to back, so the cursor only ever moves forward. Each span is
' counted the first time it marks anything, so the tally is changes carried into
' the file -- which the two exports are then compared on.
Private Function RevKindAt(ByRef st As ExportState, ByVal pos As Long) As Long
    Do While st.revCursor <= st.nRevs
        If st.revEnd(st.revCursor) > pos Then Exit Do
        st.revCursor = st.revCursor + 1
    Loop
    If st.revCursor > st.nRevs Then Exit Function

    If pos >= st.revStart(st.revCursor) Then
        RevKindAt = st.revKind(st.revCursor)
        If st.lastMarked <> st.revCursor Then
            st.lastMarked = st.revCursor
            st.markedRevs = st.markedRevs + 1
        End If
    End If
End Function

' Fold Word's revision types down to the two the export shows. Only content
' changes are marked: a formatting or paragraph-property revision changed no
' words, and marking it would wrap untouched prose in {++...++}. A move is shown
' as the insertion and deletion it is made of.
Private Function RevKindOf(ByVal revType As Long) As Long
    Select Case revType
        Case wdRevisionInsert, wdRevisionMovedTo:   RevKindOf = REV_INSERT
        Case wdRevisionDelete, wdRevisionMovedFrom: RevKindOf = REV_DELETE
        Case Else:                                  RevKindOf = REV_NONE
    End Select
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
            ' Inline object / comment / field control marks. A comment mark only
            ' reaches here from inside a note (the body's are consumed by the
            ' walk, which turns them into "[cn]" markers).
            MapChar = ""
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
    ' ADODB.Stream can only write to a local/UNC path. Catch a URL (or a missing
    ' folder) here and raise a plain-English error instead of a bare 3004.
    If LCase$(Left$(path, 7)) = "http://" Or LCase$(Left$(path, 8)) = "https://" Then
        Err.Raise vbObjectError + 3004, "WriteUtf8NoBom", _
            "Cannot write to a cloud URL: " & path & vbCrLf & _
            "Choose a local folder for the .md file."
    End If
    Dim parent As String
    Dim sl As Long: sl = InStrRev(path, "\")
    If sl > 0 Then parent = Left$(path, sl - 1)
    If Len(parent) > 0 And Dir$(parent, vbDirectory) = "" Then
        Err.Raise vbObjectError + 3004, "WriteUtf8NoBom", _
            "The target folder does not exist:" & vbCrLf & parent
    End If

    Dim st As Object, bin As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                     ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText text

    st.Position = 0
    st.Type = 1                     ' adTypeBinary
    st.Position = 3                 ' skip the 3-byte UTF-8 BOM
    Dim bytes As Variant: bytes = st.Read
    st.Close: Set st = Nothing

    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1                    ' adTypeBinary
    bin.Open
    bin.Write bytes
    bin.SaveToFile path, 2          ' adSaveCreateOverWrite
    bin.Close: Set bin = Nothing
    Exit Sub

Fail:
    Dim n As Long, d As String: n = Err.Number: d = Err.Description
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    If Not bin Is Nothing Then bin.Close
    On Error GoTo 0
    Err.Raise n, "WriteUtf8NoBom", d
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
    ' See PerfSuppress: these dominate a bulk edit run, and this hook runs while
    ' the user is waiting on a close.
    Dim perf As PerfState: perf = PerfSuppress()
    Dim prevTrack As Boolean: prevTrack = Doc.TrackRevisions
    Doc.TrackRevisions = False
    Dim prevAutoSave As Boolean: prevAutoSave = False
    prevAutoSave = Doc.AutoSaveOn
    Doc.AutoSaveOn = False

    ' Strip hyperlinks first for the same reason as the manual macro: this
    ' hook runs BEFORE the close review's own link removal, and a fake name
    ' inside a link's display text has been missed on exactly this first pass.
    StripHyperlinksEverywhere Doc

    ReplaceAllMappings Doc, maps, nMaps, True, False, "De-anonymizing"

    ' Restore the court-identity header (Department 515, judge, courtroom staff).
    ApplyCourtIdentity Doc, True

    Doc.AutoSaveOn = prevAutoSave
    Doc.TrackRevisions = prevTrack
    PerfRestore perf
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
' on any error or if no usable rows were found. nKeepRows reports how many rows
' were operator KEEP instructions rather than pseudonyms (see IsKeepDecisionCell).
Private Function ReadPseudonymKey(ByVal path As String, _
                                   ByRef maps() As Mapping, _
                                   ByRef nMaps As Long, _
                                   Optional ByRef nKeepRows As Long, _
                                   Optional ByRef nUnusedRows As Long) As Boolean
    On Error GoTo Fail
    nMaps = 0
    nKeepRows = 0
    nUnusedRows = 0

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

    ' Locate the columns we need from the header row. Real Value and Replacement
    ' are required; Status is optional (a key from an older version of
    ' PDF-Linker may not have it) and carries the two verdicts below.
    Dim realCol As Long, fakeCol As Long, statCol As Long, c As Long
    realCol = 0: fakeCol = 0: statCol = 0
    For c = cLo To cHi
        Dim hd As String: hd = LCase$(Trim$(CStr(NzText(data(rLo, c)))))
        If hd = "real value" Then realCol = c
        If hd = "replacement" Then fakeCol = c
        If hd = "status" Then statCol = c
    Next c
    If realCol = 0 Or fakeCol = 0 Then GoTo CleanFail

    ReDim maps(1 To (rHi - rLo + 1))
    Dim r As Long
    For r = rLo + 1 To rHi
        Dim rv As String, fk As String, st As String
        rv = Trim$(CStr(NzText(data(r, realCol))))
        fk = Trim$(CStr(NzText(data(r, fakeCol))))
        If Len(rv) > 0 And Len(fk) > 0 And StrComp(rv, fk, vbBinaryCompare) <> 0 Then
            If IsKeepDecisionCell(rv, fk) Then
                nKeepRows = nKeepRows + 1       ' operator instruction, not a fake
            Else
                nMaps = nMaps + 1
                maps(nMaps).real = rv
                maps(nMaps).fake = fk
                ' PDF-Linker's own verdict on the row. Two of its values say
                ' this row cannot be a de-anonymize MISS, so the result dialog
                ' can stop them reading as failures:
                '   "alt spelling" - a synthetic spelling of another row's
                '                    value, sharing that row's fake. Forward
                '                    only; see Mapping.forwardOnly.
                '   "no match"     - the fake was never written into the
                '                    exported text (the key pins the binding for
                '                    a party this batch never mentioned), so it
                '                    cannot be in a draft written from those
                '                    exports.
                If statCol > 0 Then
                    st = LCase$(Trim$(CStr(NzText(data(r, statCol)))))
                    maps(nMaps).forwardOnly = (st = "alt spelling")
                    If st = "alt spelling" Or st = "no match" Then _
                        nUnusedRows = nUnusedRows + 1
                End If
            End If
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
' BULK-EDIT PERFORMANCE
'==============================================================================
' Turn off the two background services that re-process the document after every
' single edit, and return the previous settings for PerfRestore. ScreenUpdating
' = False does not cover these: background repagination still re-flows the whole
' document and check-as-you-type still re-proofs each range touched, so a run of
' hundreds of replacements pays both costs hundreds of times over.
' Yield to Windows so a long sweep doesn't get marked "Not Responding" -- EXCEPT
' while the close hook is on the stack. There, Word has already committed to
' closing the document, and pumping the queue mid-run lets a queued close (an
' impatient second click on the X) re-enter a document that is half torn down.
' The sweeps that call this are the same ones the close hook drives, so the
' choice has to be made here rather than at each call site.
Private Sub PumpQueue()
    If modMain.gInCloseReview Then Exit Sub
    DoEvents
End Sub

Private Function PerfSuppress() As PerfState
    Dim p As PerfState
    On Error Resume Next
    p.pagination = Application.Options.Pagination
    p.spell = Application.Options.CheckSpellingAsYouType
    p.grammar = Application.Options.CheckGrammarAsYouType
    p.saved = True
    Application.Options.Pagination = False
    Application.Options.CheckSpellingAsYouType = False
    Application.Options.CheckGrammarAsYouType = False
    On Error GoTo 0
    PerfSuppress = p
End Function

' Put the user's own settings back exactly. Safe to call twice and safe when
' PerfSuppress never ran (saved = False).
Private Sub PerfRestore(ByRef p As PerfState)
    If Not p.saved Then Exit Sub
    On Error Resume Next
    Application.Options.Pagination = p.pagination
    Application.Options.CheckSpellingAsYouType = p.spell
    Application.Options.CheckGrammarAsYouType = p.grammar
    p.saved = False
    On Error GoTo 0
End Sub

' Show all markup inline for the duration of a re-anonymize run, remembering what
' the user had. Word hides tracked deletions from BOTH Find and Range.Text when
' they are drawn in balloons rather than in the text, so with the wrong view a
' real name inside a deletion is neither replaced nor exported: the shareable
' copy would silently carry it, and both exports would silently drop the change.
' Setting the view is not a document edit -- nothing here dirties the file.
Private Function ForceInlineMarkup() As MarkupView
    Dim mv As MarkupView
    On Error Resume Next
    Err.Clear                    ' saved is read off Err below; Err survives calls
    With ActiveWindow.view
        mv.showRevisions = .ShowRevisionsAndComments
        mv.showInsDel = .ShowInsertionsAndDeletions
        mv.mode = .RevisionsMode
        mv.view = .RevisionsView
        mv.saved = (Err.Number = 0)

        .ShowRevisionsAndComments = True
        .ShowInsertionsAndDeletions = True
        .RevisionsView = wdRevisionsViewFinal
        .RevisionsMode = wdInLineRevisions
    End With
    On Error GoTo 0
    ForceInlineMarkup = mv
End Function

' Put the user's markup view back. Safe to call twice and safe when
' ForceInlineMarkup never ran or couldn't read the window (saved = False).
Private Sub RestoreMarkupView(ByRef mv As MarkupView)
    If Not mv.saved Then Exit Sub
    On Error Resume Next
    With ActiveWindow.view
        .RevisionsMode = mv.mode
        .RevisionsView = mv.view
        .ShowInsertionsAndDeletions = mv.showInsDel
        .ShowRevisionsAndComments = mv.showRevisions
    End With
    mv.saved = False
    On Error GoTo 0
End Sub

' Count of tracked revisions across the document. A document carrying thousands
' of revision marks makes every Find and every text assignment dramatically more
' expensive, which no amount of tuning here can undo -- so the timing note
' surfaces it rather than leaving the user to wonder.
Private Function RevisionLoad(ByVal oDoc As Document) As Long
    On Error Resume Next
    RevisionLoad = oDoc.Revisions.count
End Function

'==============================================================================
' PHASE TIMING  (diagnostics)
'==============================================================================
' Seconds elapsed since tFrom, formatted for the result dialog. Timer is seconds
' since midnight, so a run spanning midnight would go negative -- fold that.
Private Function PhaseSecs(ByVal tFrom As Single) As String
    Dim d As Single: d = Timer - tFrom
    If d < 0 Then d = d + 86400#
    PhaseSecs = Format$(d, "0.0") & "s"
End Function

' Show the phase breakdown only when the run was actually slow, so a normal fast
' run keeps a clean dialog. The threshold is deliberately low enough that anyone
' who notices a wait gets the numbers.
Private Function TimingNote(ByVal sTimes As String, ByVal tRun As Single, _
                            Optional ByVal nRevs As Long = 0) As String
    Dim d As Single: d = Timer - tRun
    If d < 0 Then d = d + 86400#
    If d < 8 Then Exit Function
    Dim sRev As String
    If nRevs > 0 Then
        sRev = "  (Document carries " & nRevs & " tracked revision(s), which " & _
               "makes every search and edit markedly slower -- accepting or " & _
               "rejecting them first will speed this up.)" & vbCrLf
    End If
    TimingNote = vbCrLf & vbCrLf & "Took " & Format$(d, "0.0") & "s. Where the " & _
                 "time went:" & vbCrLf & sTimes & sRev
End Function

' Name the running phase in the status bar so a long run shows what it is doing.
' Pass "" to hand the status bar back to Word. macroName says which direction is
' running: the re-anonymize half calls this too, and labelling its phases
' "De-Anonymize" would name the opposite macro.
Private Sub SetPhase(ByVal s As String, _
                     Optional ByVal macroName As String = "De-Anonymize")
    On Error Resume Next
    If Len(s) = 0 Then
        Application.StatusBar = False
    Else
        Application.StatusBar = macroName & ": " & s & " ..."
    End If
End Sub

' One line for the result dialog about the markup both exports carried over, so
' the notation isn't a surprise when the file is opened -- and so the "Reviewer"
' labels are explained where the user will see them. Empty when the document had
' neither tracked changes nor comments.
'
' markedReal/markedFake are what each export actually wrote. They agree by
' construction (one captured set of spans, marked twice), so a mismatch means a
' span's text was replaced away between the two writes -- reported rather than
' left for the user to notice by diffing the files.
Private Function MarkupNote(ByVal nComments As Long, _
                            ByVal markedReal As Long, ByVal markedFake As Long, _
                            ByVal nResidual As Long) As String
    If markedReal <= 0 And markedFake <= 0 And nComments <= 0 Then Exit Function

    Dim s As String
    s = vbCrLf & vbCrLf & "Carried " & markedReal & " tracked change(s) and " & _
        nComments & " comment(s) into both files: insertions as {++text++}, " & _
        "deletions as {--text--}, comments as [c1], [c2] ... with the text " & _
        "collected at the end."
    If nComments > 0 Then
        s = s & " Commenters are named in the real-names file only -- the " & _
            "anonymized one labels them ""Reviewer 1"", ""Reviewer 2"", because " & _
            "an author name is document metadata that no pseudonym key covers."
    End If
    If markedFake <> markedReal Then
        s = s & vbCrLf & vbCrLf & "NOTE: the anonymized file marks " & markedFake & _
            " of those " & markedReal & " change(s) -- the rest were edits whose " & _
            "text the pseudonym pass replaced entirely. Compare the two files at " & _
            "those spots before relying on the shared copy's markup."
    End If
    If nResidual > 0 Then
        s = s & vbCrLf & vbCrLf & "CHECK BEFORE SHARING: " & nResidual & _
            " key value(s) still appear INSIDE the anonymized file's tracked-change " & _
            "markup. Open it and read those {++...++} / {--...--} spans. (A party " & _
            "name that also names a cited case is left alone on purpose and shows " & _
            "up here too.)"
    End If
    MarkupNote = s
End Function

' How many of the key's real values still appear inside the anonymized export's
' tracked-change markup. This is the one place a replacement can fail quietly:
' text inside a tracked deletion is a run Word does not always let Find rewrite,
' and unlike ordinary prose nobody re-reads a deleted sentence before sharing the
' file. Checked on the finished string, so it reports what the file actually says
' rather than what the replacement pass believed it did.
Private Function ResidualsInMarkup(ByVal md As String, ByRef maps() As Mapping, _
                                   ByVal nMaps As Long) As Long
    Dim hay As String: hay = LCase$(MarkedText(md))
    If Len(hay) = 0 Then Exit Function

    Dim i As Long, n As Long
    For i = 1 To nMaps
        If Len(maps(i).real) > 0 Then
            If InStr(1, hay, LCase$(maps(i).real), vbBinaryCompare) > 0 Then
                n = n + 1
            End If
        End If
    Next i
    ResidualsInMarkup = n
End Function

' Just the text inside the export's tracked-change markers, one span per line.
' An unterminated marker ends the scan rather than running to the end of the
' file: half a span is not evidence of anything.
Private Function MarkedText(ByVal md As String) As String
    Dim res As String
    Dim pos As Long: pos = 1
    Do
        Dim aIns As Long, aDel As Long, a As Long, closer As String
        aIns = InStr(pos, md, "{++", vbBinaryCompare)
        aDel = InStr(pos, md, "{--", vbBinaryCompare)
        If aIns = 0 And aDel = 0 Then Exit Do

        If aIns = 0 Or (aDel > 0 And aDel < aIns) Then
            a = aDel: closer = "--}"
        Else
            a = aIns: closer = "++}"
        End If

        Dim e As Long: e = InStr(a + 3, md, closer, vbBinaryCompare)
        If e = 0 Then Exit Do

        res = res & Mid$(md, a + 3, e - a - 3) & vbLf
        pos = e + 3
    Loop
    MarkedText = res
End Function

' One line for the result dialog when key rows were skipped as extraction debris
' -- a term that matched many times but never once stood on its own. Reported so
' a skip is visible rather than silent; empty when there were none.
Private Function FragmentNote(ByVal nSkips As Long) As String
    If nSkips <= 0 Then Exit Function
    FragmentNote = vbCrLf & vbCrLf & "Skipped " & nSkips & " key row(s) whose " & _
                   "search term never appeared on its own -- only buried inside " & _
                   "longer words (e.g. ""ES"" inside ""cases""). Those are " & _
                   "extraction fragments, not pseudonyms; replacing them would " & _
                   "have rewritten ordinary prose."
End Function

' One line for the result dialog when key rows were retired as AMBIGUOUS -- one
' fake claimed by two different real values, so there is no way to know which to
' restore (see ReplaceAllMappings). The mapping is skipped, which is safe but
' leaves that pseudonym in the document, so say so: the pink highlight will show
' where, and the user has to pick the right name by hand. Empty when none.
Private Function AmbiguousNote(ByVal nAmbiguous As Long) As String
    If nAmbiguous <= 0 Then Exit Function
    AmbiguousNote = vbCrLf & vbCrLf & "Left " & nAmbiguous & " key row(s) " & _
                    "alone: their pseudonym is claimed by more than one real " & _
                    "value, so there is no way to tell which to restore. Those " & _
                    "names are still pseudonyms in the document -- look for the " & _
                    "pink highlights and correct them by hand."
End Function

' One line for the result dialog when the key holds rows that cannot be restored
' HERE and are not misses: a binding the anonymizer never wrote into the exported
' text (Status "no match" -- a party this batch never mentioned) or an alternate
' spelling whose fake another row already reverses (Status "alt spelling").
' Without this they read as a large de-anonymize failure in the "restored X of Y"
' count. Empty when there were none.
Private Function UnusedRowsNote(ByVal nUnusedRows As Long) As String
    If nUnusedRows <= 0 Then Exit Function
    UnusedRowsNote = vbCrLf & vbCrLf & "That total includes " & nUnusedRows & _
                     " mapping(s) that could not apply here -- the key pins them " & _
                     "for parties this batch of filings never mentioned, or they " & _
                     "are alternate spellings another row already restores -- so " & _
                     "they are not misses."
End Function

' One line for the result dialog when the key carried operator KEEP rows, so it
' is clear those were recognized and deliberately not treated as pseudonyms
' rather than silently lost. Empty when there were none.
Private Function KeepRowsNote(ByVal nKeepRows As Long) As String
    If nKeepRows <= 0 Then Exit Function
    KeepRowsNote = vbCrLf & vbCrLf & "Skipped " & nKeepRows & " KEEP row(s) in " & _
                   "the key (a ""no"", a ""never"", or a [bracketed] or " & _
                   "{braced} keep-spec). Those are " & _
                   "operator instructions to leave a value alone, not " & _
                   "pseudonyms, so there is nothing to swap for them."
End Function

' True when the key's Replacement cell holds an operator KEEP INSTRUCTION rather
' than a pseudonym, in which case the row is not a mapping at all and must be
' skipped in BOTH directions.
'
' PDF-Linker's key loader (_pn_load_key) accepts the same control vocabulary the
' LEAKS "Fix?" column does. Such a row builds NO faking term -- the Real Value
' was deliberately left VERBATIM in the anonymized text:
'
'   "no" / "n"      KEEP: leave this Real Value alone, never fake it.
'   "never"         the same, everywhere: the NUCLEAR keep of the whole value,
'                   honored in this and every other case folder. It is the same
'                   instruction as bracing the entire value, without retyping
'                   the value inside braces.
'   "[bracketed]"   keep-spec: keep the bracketed part(s) verbatim and auto-fake
'                   the rest. The cell records the operator's INSTRUCTION, not
'                   any fake that was used.
'   "{braced}"      the same cut with a stronger promise (the braced words are
'                   never faked in any folder, not even inside a party name).
'
' Read literally these look like pseudonyms, which is exactly how a real key went
' wrong here. Six rows reading "no" made de-anonymize try to restore the word
' "no" -- it would rewrite a caption's standalone "No." and crawl every "not" and
' "notice" on the way -- and rows like "[HONORABLE]" or "[DEMURRER TO]" would
' have been written INTO the shared copy by re-anonymize as literal bracketed
' text. Neither value was ever a fake. "never" is the same hazard and a worse
' word for it: unguarded, one such row rewrites every "never" in the document
' into somebody's name.
'
' Mirrors _pn_bracket_keep's fall-through exactly: brackets and braces are cut
' the same way and read together, and when ANY delimited part is not a substring
' of the Real Value, PDF-Linker abandons the keep reading and treats the cell as
' an ordinary explicit replacement, so we do too.
Private Function IsKeepDecisionCell(ByVal realValue As String, ByVal cell As String) As Boolean
    Dim c As String: c = Trim$(cell)
    If Len(c) = 0 Then Exit Function

    Dim lc As String: lc = LCase$(c)
    If lc = "no" Or lc = "n" Or lc = "never" Then
        IsKeepDecisionCell = True
        Exit Function
    End If

    If InStr(1, c, "[") = 0 And InStr(1, c, "{") = 0 Then Exit Function

    Dim found As Boolean: found = False
    If Not KeepPartsArePresent(realValue, c, "[", "]", found) Then Exit Function
    If Not KeepPartsArePresent(realValue, c, "{", "}", found) Then Exit Function

    IsKeepDecisionCell = found
End Function

' Every part of `cell` delimited by opener/closer must occur in the Real Value.
' Returns False the moment one does not -- the cell is a literal replacement
' after all -- and sets `found` when at least one real part was seen.
Private Function KeepPartsArePresent(ByVal realValue As String, ByVal cell As String, _
                                     ByVal opener As String, ByVal closer As String, _
                                     ByRef found As Boolean) As Boolean
    KeepPartsArePresent = True
    Dim p As Long, q As Long, part As String
    p = InStr(1, cell, opener)
    Do While p > 0
        q = InStr(p + 1, cell, closer)
        If q = 0 Then Exit Do
        part = Trim$(Mid$(cell, p + 1, q - p - 1))
        If Len(part) > 0 Then
            If InStr(1, realValue, part, vbTextCompare) = 0 Then
                KeepPartsArePresent = False
                Exit Function
            End If
            found = True
        End If
        p = InStr(q + 1, cell, opener)
    Loop
End Function

' Run every mapping against the document, skipping the ones whose search term
' isn't in it. Shared by all three callers (manual de-anonymize, re-anonymize,
' and the on-close hook) so they all get the same pre-filter and progress
' reporting. Returns the number of mappings that actually changed something.
'
'   useFake  = True  -> search maps().fake, write maps().real  (de-anonymize)
'              False -> search maps().real, write maps().fake  (re-anonymize)
'   protect  passes through to ReplaceInStories' protectCitations.
'
' Two rounds, not one: replacing a fake inserts a real value, and an inserted
' value can in principle contain another mapping's search term that wasn't in the
' document when the filter ran. A second round re-reads the text and picks up any
' mapping that was skipped the first time. Each mapping is still processed AT
' MOST ONCE (the done() flags), so this always terminates and never re-replaces
' its own output -- and it strictly covers more than the old single pass, which
' would have missed such a reveal for any mapping earlier in the sort order.
Private Function ReplaceAllMappings(ByVal oDoc As Document, ByRef maps() As Mapping, _
                                     ByVal nMaps As Long, ByVal useFake As Boolean, _
                                     ByVal protect As Boolean, _
                                     ByVal progressLabel As String, _
                                     Optional ByRef nFragmentSkips As Long = 0, _
                                     Optional ByRef nAmbiguous As Long = 0) As Long
    ' "handled", not "done": Done: is used as a label elsewhere in this module.
    ' "pass", not "round": Round is a built-in VBA function.
    Dim handled() As Boolean
    ReDim handled(1 To nMaps)

    ' Defenses against degenerate key rows, both observed in a real key.
    ' (Operator KEEP instructions -- "no" and "[bracketed]" keep-specs -- are a
    ' different thing entirely and never reach here: ReadPseudonymKey drops them
    ' at load time via IsKeepDecisionCell, because they were never fakes.)
    '
    ' De-anonymize: pre-retire every mapping whose FAKE is claimed by rows with
    ' DIFFERENT real values. A key has shipped with one fake address bound to two
    ' distinct real ones ("1330 N Cahuenga Blvd." and "1331 N. Cahuenga Blvd."),
    ' and there is no way to know which such a token should restore to. The real
    ' values are compared CASE-INSENSITIVELY: most duplicate fakes are casing
    ' pairs of the same name ("GARDELLA"/"Gardella", from the caption and the
    ' body), which restore identically -- MatchCasing recases from the matched
    ' text -- and must NOT be retired.
    '
    ' Retiring is FAIL-SAFE (the fake is left standing and the pink residual
    ' scan flags it) but it is still a name the user must fix by hand, so the
    ' count is reported rather than swallowed.
    '
    ' Forward-only rows are dropped BEFORE that grouping, not caught by it.
    ' PDF-Linker registers SYNTHETIC spellings of a name against that name's own
    ' fake so that every spelling scrubs -- a wrap-split hyphen ("Ardeshirpour-
    ' Zartoshti"), an OCR near-miss ("Sarra" for "Sara") -- and each is its own
    ' key row. Left in, they made the fake of every hyphenated party look
    ' ambiguous and retired the real mapping with them, so no hyphenated name
    ' de-anonymized at all. PDF-Linker marks the non-canonical row (Status =
    ' "alt spelling"); skipping those first leaves exactly one row per fake and
    ' the guard sees a clean key. The guard still matters: it is the only thing
    ' standing between a genuine collision -- or a key written before the marker
    ' existed -- and a wrong restore.
    '
    ' Re-anonymize: skip a real value shorter than 3 characters. Rows like
    ' real "JR" and real "TO" are extraction junk, and replacing every standalone
    ' "to" with a pseudonym would corrupt the shared copy's prose.
    Dim da As Long
    If useFake Then
        ' Group by lowercased fake in ONE pass with a Dictionary. The previous
        ' nested scan was O(n^2) locale-aware StrComp -- ~40,000 comparisons on a
        ' 200-row key, each one slow.
        Dim seenAt As Object, ambiguous As Object
        Set seenAt = CreateObject("Scripting.Dictionary")
        Set ambiguous = CreateObject("Scripting.Dictionary")
        Dim fk As String
        ' Forward-only rows are out of this direction entirely, so retire them
        ' before grouping -- otherwise they are the duplicate.
        For da = 1 To nMaps
            If maps(da).forwardOnly Then handled(da) = True
        Next da
        For da = 1 To nMaps
            If Not handled(da) Then
                fk = LCase$(maps(da).fake)
                If seenAt.Exists(fk) Then
                    ' Same fake claimed again: ambiguous only when the REAL
                    ' values differ. Duplicate fakes are usually casing pairs of
                    ' one name ("GARDELLA"/"Gardella", from the caption and the
                    ' body) which restore identically -- MatchCasing recases
                    ' from the matched text -- and must NOT be retired.
                    If StrComp(maps(da).real, maps(seenAt(fk)).real, vbTextCompare) <> 0 Then
                        ambiguous(fk) = True
                    End If
                Else
                    seenAt(fk) = da
                End If
            End If
        Next da
        If ambiguous.count > 0 Then
            For da = 1 To nMaps
                If Not handled(da) Then
                    If ambiguous.Exists(LCase$(maps(da).fake)) Then
                        handled(da) = True
                        nAmbiguous = nAmbiguous + 1
                    End If
                End If
            Next da
        End If
    Else
        For da = 1 To nMaps
            If Len(maps(da).real) < 3 Then handled(da) = True
        Next da
    End If

    Dim distinctHits As Long: distinctHits = 0
    Dim pass As Long, passHits As Long, i As Long
    Dim stories() As StoryRef, nStories As Long

    For pass = 1 To 2
        ' Collect the stories (and their text) ONCE per pass, not once per
        ' mapping. Rebuilding on pass 2 also refreshes the cached text so a term
        ' revealed by an earlier insertion is seen.
        nStories = CollectStories(oDoc, stories)
        If nStories = 0 Then Exit For
        passHits = 0

        For i = 1 To nMaps
            If Not handled(i) Then
                handled(i) = True           ' set before the call: process once, period
                Dim n As Long
                If useFake Then
                    n = ReplaceInStories(stories, nStories, maps(i).fake, maps(i).real, _
                                         protect, nFragmentSkips)
                Else
                    n = ReplaceInStories(stories, nStories, maps(i).real, maps(i).fake, _
                                         protect, nFragmentSkips)
                End If
                If n > 0 Then
                    distinctHits = distinctHits + 1
                    passHits = passHits + 1
                End If
            End If

            ' Keep Word's queue serviced so the window stays responsive, and show
            ' progress so a long key doesn't look like a hang.
            If i Mod 10 = 0 Then
                On Error Resume Next
                Application.StatusBar = progressLabel & ": " & i & " of " & nMaps & " ..."
                On Error GoTo 0
                PumpQueue
            End If
        Next i

        ' Nothing changed this pass, so no insertion can have revealed a new
        ' term: another pass would re-collect the stories for nothing.
        If passHits = 0 Then Exit For
    Next pass

    On Error Resume Next
    Application.StatusBar = False           ' hand the status bar back to Word
    On Error GoTo 0

    ReplaceAllMappings = distinctHits
End Function

'==============================================================================
' PRESENCE PRE-FILTER  (performance)
'==============================================================================
' Shared superset-safe containment test. Used by the residual-pseudonym
' highlighter to skip a native Find sweep for a term that isn't in the story at
' all; the replacement path does the same test inline, per story, in
' ReplaceInStories.
'
' The filter only ever has to be a SUPERSET of what Word's Find would match. It
' lowercases both sides and compares binary, matching Find's MatchCase = False
' behavior for the ASCII names and case numbers a key holds. An empty haystack
' means the text could not be read, so answer True and let the real Find decide
' -- work is never skipped on a guess.
Private Function TermMaybePresent(ByVal hayLower As String, ByVal term As String) As Boolean
    If Len(hayLower) = 0 Then
        TermMaybePresent = True
        Exit Function
    End If
    If Len(term) = 0 Then Exit Function
    TermMaybePresent = (InStr(1, hayLower, LCase$(term), vbBinaryCompare) > 0)
End Function

'==============================================================================
' REPLACEMENT
'==============================================================================
' Collect every searchable story ONCE: the main body, each section's non-empty
' headers/footers, footnotes/endnotes when present, and each shape's own text
' frame -- each paired with a lowercased copy of its text.
'
' Doing this once per pass instead of once per mapping is the whole point. The
' old ReplaceEverywhere re-walked oDoc.Sections (building three Header and three
' Footer objects per section, then sweeping even the empty ones) and oDoc.Shapes
' for EVERY mapping, so a 300-row key meant ~3,000 story sweeps plus 300 rounds
' of collection-walking. With the stories cached, a mapping only pays for the
' stories whose text actually contains its search term -- for almost every name
' that is the body alone.
'
' This deliberately does NOT walk StoryRanges/NextStoryRange: enumerating that
' chain while replacing inside the loop can destabilize and crash Word. The
' collections used here (Sections, Footnotes, Shapes with per-shape TextRange)
' stay valid across text replacement, so iterating them is safe.
Private Function CollectStories(ByVal oDoc As Document, ByRef arr() As StoryRef) As Long
    Dim n As Long: n = 0
    ReDim arr(1 To 64)

    On Error Resume Next

    ' Main body (caption, party block, and prose are all here in a plain draft).
    AddStory arr, n, oDoc.content

    ' Headers and footers, section by section. Empty ones are dropped by AddStory.
    Dim sec As Section, hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then AddStory arr, n, hf.Range
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then AddStory arr, n, hf.Range
        Next hf
    Next sec

    ' Footnotes / endnotes, only when present (accessing the story otherwise errors).
    If oDoc.Footnotes.count > 0 Then AddStory arr, n, oDoc.StoryRanges(wdFootnotesStory)
    If oDoc.Endnotes.count > 0 Then AddStory arr, n, oDoc.StoryRanges(wdEndnotesStory)

    ' Comments. A reviewer's comment names the parties as freely as the prose
    ' does, and the Markdown export now carries comments into the shareable copy
    ' -- so this story has to be scrubbed with the rest. (The author NAME is
    ' metadata, not story text; nothing here can reach it, which is why the
    ' anonymized export labels commenters "Reviewer n" instead.)
    If oDoc.Comments.count > 0 Then AddStory arr, n, oDoc.StoryRanges(wdCommentsStory)

    ' Text boxes / shapes, each one's own text frame directly.
    Dim shp As Shape
    For Each shp In oDoc.Shapes
        If shp.TextFrame.HasText Then AddStory arr, n, shp.TextFrame.TextRange
    Next shp

    On Error GoTo 0
    CollectStories = n
End Function

' Append one story and its lowercased text. Stories with no text (the usual case
' for most of a section's six header/footer slots) are skipped outright -- they
' can't contain a search term, and sweeping them was pure cost.
Private Sub AddStory(ByRef arr() As StoryRef, ByRef n As Long, ByVal r As Range)
    On Error Resume Next
    If r Is Nothing Then Exit Sub
    Dim t As String
    t = r.text
    ' Nothing but paragraph/cell marks is an EMPTY story -- most of a section's
    ' six header/footer slots -- and can hold no search term. Len(t) = 0 alone
    ' missed these: an empty header's text is a bare vbCr, so they were kept and
    ' swept for every mapping.
    If Len(Trim$(Replace(Replace(t, vbCr, ""), Chr$(7), ""))) = 0 Then Exit Sub
    n = n + 1
    If n > UBound(arr) Then ReDim Preserve arr(1 To UBound(arr) + 64)
    Set arr(n).rng = r
    arr(n).lower = LCase$(t)
End Sub

' Replace findText with replaceText in every collected story that contains it.
' Returns the number of stories in which a replacement was made.
'
' Each qualifying story gets two phases:
'   1. NativeReplacePasses -- Word's own wdReplaceAll handles ALL occurrences of
'      the common casing forms in one COM call per form. This is where the bulk
'      of the hits go, at native speed. The old per-hit VBA loop cost ~15 COM
'      round trips per occurrence (Find setup, boundary probes, text assignment),
'      so a key whose names appear hundreds of times each froze Word for the
'      whole stretch -- and DoEvents never ran inside a mapping, so the window
'      was marked "Not Responding" on top of being slow.
'   2. ReplaceInRange -- the careful per-hit sweep, now only mopping up what the
'      native passes' stricter matching missed: possessives ("Thorne's"),
'      tokens Word's whole-word logic skips against parentheses/quotes, and odd
'      mixed casings. Typically zero or a handful of hits.
'
' The per-story containment test is the same superset-safe filter used elsewhere:
' lowercase both sides and compare binary, matching Find's MatchCase = False for
' the ASCII names and case numbers a key holds. The cached text can only go stale
' by losing occurrences (replacement removes the search term), so a stale hit
' costs one wasted sweep -- it can never skip a story that still needs one.
'
' protectCitations (re-anonymize only): leave any match that sits in italic text
' untouched, so a real party name that also appears inside a cited case name
' (e.g. "Nash v. Superior Court") is not rewritten in the shared copy. This
' mirrors the PDF-Linker pseudonymizer's cardinal invariant -- renaming a cited
' decision is a worse failure than leaving a party name in -- and its caption
' exemption: a brief italicizes cited authorities but not its own caption/prose,
' so the current parties still get replaced while published cites are preserved.
Private Function ReplaceInStories(ByRef stories() As StoryRef, ByVal nStories As Long, _
                                   ByVal findText As String, ByVal replaceText As String, _
                                   Optional ByVal protectCitations As Boolean = False, _
                                   Optional ByRef nFragmentSkips As Long = 0) As Long
    Dim total As Long: total = 0
    If Len(findText) = 0 Then Exit Function
    ' Word's Find raises on search terms longer than 255 characters; under the
    ' resume-next handling in ReplaceInRange that used to fall through in a
    ' dangerous state. Skip such mappings outright.
    If Len(findText) > 255 Then Exit Function

    Dim whole As Boolean: whole = ShouldWholeWord(findText)
    Dim needle As String: needle = LCase$(findText)

    Dim k As Long
    Dim changed As Boolean
    For k = 1 To nStories
        If InStr(1, stories(k).lower, needle, vbBinaryCompare) > 0 Then

            ' Contained in other words, never on its own -> skip this story
            ' entirely. For a whole-token term every embedded hit would be
            ' refused by the boundary check anyway, so the replacement is a
            ' no-op and this removes only the scanning. It is what makes a
            ' debris row like "ES" (151 hits inside "cases"/"issues", none
            ' standalone) cost nothing instead of a full sweep, while a term
            ' that DOES stand alone somewhere is still replaced there.
            If whole Then
                If OnlyEmbedded(stories(k).lower, needle) Then
                    If LooksLikeFragment(stories(k).lower, needle) Then
                        nFragmentSkips = nFragmentSkips + 1
                    End If
                    GoTo NextStory
                End If
            End If

            ' Reset per story: this tracks whether THIS story changed.
            changed = NativeReplacePasses(stories(k).rng, findText, replaceText, _
                                          whole, protectCitations)

            ' Only run the careful per-hit sweep when it actually has work. Both
            ' phases can only ever REMOVE occurrences, so one in-memory read of
            ' the story's current text settles it -- and for a whole-token term
            ' it answers the real question (is there a STANDALONE occurrence
            ' left?), not merely whether the letters appear somewhere. That is
            ' the sweep's whole cost: a Find plus a ~8-COM-call boundary probe
            ' for every occurrence buried inside a larger word.
            If NeedsManualSweep(stories(k).rng, needle, whole) Then
                If ReplaceInRange(stories(k).rng, findText, replaceText, whole, protectCitations) Then
                    changed = True
                End If
            End If
            If changed Then total = total + 1
        End If
NextStory:
    Next k

    ReplaceInStories = total
End Function

' True when the careful per-hit sweep still has work to do in this story, decided
' from ONE in-memory read of the story's current text.
'
' Read fresh, not from the cached snapshot, because the native passes just edited
' it. On a read failure answer True so the sweep still runs -- never skip
' replacement work on a guess.
'
' For a whole-token term this asks the real question: is there an occurrence with
' NON-ALPHANUMERIC characters on both sides? Plain containment is not enough and
' was the expensive mistake. A key can bind fragment fakes like "COU" or "ES",
' which occur inside "court", "counsel", "cases" hundreds of times but never
' stand alone; containment said "present", so the sweep ran and spent ~8 COM
' calls per occurrence discovering each was embedded. HasWholeToken settles all
' of them from the string already in hand.
'
' It mirrors WholeTokenBoundaries exactly, so the cases the sweep exists for
' still qualify: a possessive ("Thorne's") and a quote- or paren-hugged token
' ("(Nash)") both have non-alphanumeric neighbours and still return True.
Private Function NeedsManualSweep(ByVal rng As Range, ByVal needleLower As String, _
                                  ByVal whole As Boolean) As Boolean
    On Error GoTo Assume
    Dim t As String: t = LCase$(rng.text)
    If whole Then
        NeedsManualSweep = HasWholeToken(t, needleLower)
    Else
        NeedsManualSweep = (InStr(1, t, needleLower, vbBinaryCompare) > 0)
    End If
    Exit Function
Assume:
    NeedsManualSweep = True
End Function

' True when needleLower occurs in textLower as a whole alphanumeric token.
Private Function HasWholeToken(ByVal textLower As String, ByVal needleLower As String) As Boolean
    HasWholeToken = (CountWholeToken(textLower, needleLower, 1) > 0)
End Function

' Number of whole-alphanumeric-token occurrences of needleLower in textLower.
' Both arguments must already be lowercased; the module has no Option Compare, so
' Like is binary and "[a-z0-9]" tests exactly the lowercase class.
'
' stopAt > 0 returns as soon as that many have been found -- callers only ever
' need "any?" or "more than N?", so a fragment that appears everywhere never
' costs a full scan.
Private Function CountWholeToken(ByVal textLower As String, ByVal needleLower As String, _
                                 Optional ByVal stopAt As Long = 0) As Long
    Dim L As Long, n As Long, p As Long, c As Long
    L = Len(textLower): n = Len(needleLower)
    If n = 0 Or L = 0 Then Exit Function

    p = InStr(1, textLower, needleLower, vbBinaryCompare)
    Do While p > 0
        Dim okL As Boolean, okR As Boolean
        okL = (p = 1)
        If Not okL Then okL = Not (Mid$(textLower, p - 1, 1) Like "[a-z0-9]")
        okR = (p + n > L)
        If Not okR Then okR = Not (Mid$(textLower, p + n, 1) Like "[a-z0-9]")
        If okL And okR Then
            c = c + 1
            If stopAt > 0 And c >= stopAt Then Exit Do
        End If
        p = InStr(p + 1, textLower, needleLower, vbBinaryCompare)
    Loop
    CountWholeToken = c
End Function

' True when the term occurs in this story ONLY inside larger words -- never on
' its own. "ES" buried in "cases"/"issues", "COU" in "court"/"counsel".
'
' For a whole-token term this makes the story skippable for FREE, not merely
' cheaply: WholeTokenBoundaries already refuses every embedded hit, so with no
' standalone occurrence the replacement is a no-op and skipping removes only the
' scan. Nothing that would have been replaced stops being replaced -- "COU" is
' still replaced in the stories where it does stand alone; only the wholly
' embedded ones drop out.
'
' This is a better discriminator than term length, which was the first attempt.
' In the user's own keys the legitimate party name "Bernards" occurs 36 times and
' the debris row "HNT" 14, so no frequency cutoff separates them and no length
' cutoff is principled -- but "stands alone somewhere?" separates them exactly.
Private Function OnlyEmbedded(ByVal storyLower As String, ByVal needleLower As String) As Boolean
    OnlyEmbedded = (CountWholeToken(storyLower, needleLower, 1) = 0)
End Function

' True when the term matched more than FRAGMENT_MAX_HITS times in the story.
' Used only to decide whether an OnlyEmbedded skip is worth REPORTING: a term
' that appeared once inside one longer word is unremarkable, whereas one that
' appeared dozens of times and never once stood alone is extraction debris the
' user should know about. Stops counting at the threshold.
Private Function LooksLikeFragment(ByVal storyLower As String, ByVal needleLower As String) As Boolean
    If Len(needleLower) = 0 Then Exit Function
    Dim p As Long, n As Long
    p = InStr(1, storyLower, needleLower, vbBinaryCompare)
    Do While p > 0
        n = n + 1
        If n > FRAGMENT_MAX_HITS Then
            LooksLikeFragment = True
            Exit Function
        End If
        p = InStr(p + 1, storyLower, needleLower, vbBinaryCompare)
    Loop
End Function

' Bulk phase: run Word's native wdReplaceAll once per DISTINCT casing form of
' findText -- as stored ("Nash"), ALL CAPS ("NASH"), all lowercase ("nash") --
' each with MatchCase = True and a replacement PRE-cased through the same
' MatchCasing used by the per-hit loop, so the output is identical for those
' forms. MatchCase = True also means Word inserts the replacement literally: the
' smart-case mangling that ruled wdReplaceAll out under MatchCase = False (the
' original reason for the per-hit loop) never engages.
'
' Correctness relative to the per-hit sweep that follows:
'   - Boundaries: the native passes use Word's MatchWholeWord for single-token
'     terms, which is only ever STRICTER than our WholeTokenBoundaries (it
'     treats an apostrophe as a word character and balks at some
'     parenthesis/quote-hugged tokens). Stricter = it can only MISS occurrences,
'     never rewrite inside a larger word that our own check would protect; the
'     manual sweep still catches everything it misses.
'   - Italic protection: with protectCitations the native Find carries
'     .Font.Italic = False, so a fully-italic cited case name is never matched.
'     A partially-italic span fails the uniform-format criterion and is left for
'     the manual sweep, whose existing rule (skip only when the WHOLE match is
'     italic) then decides it -- same outcome as before.
'   - Odd mixed casings ("nAsH") match none of the three forms and fall through
'     to the manual sweep, exactly as they always did.
' Returns True when any pass replaced something.
Private Function NativeReplacePasses(ByVal rng As Range, ByVal findText As String, _
                                     ByVal replaceText As String, ByVal whole As Boolean, _
                                     ByVal protectCitations As Boolean) As Boolean
    On Error Resume Next

    ' Distinct casing forms only: duplicate forms with MatchCase = True would be
    ' the same search twice. A term with no cased letters (a case number) yields
    ' a single form.
    Dim forms(1 To 3) As String
    Dim nF As Long: nF = 1
    forms(1) = findText
    If UCase$(findText) <> findText Then
        nF = nF + 1: forms(nF) = UCase$(findText)
    End If
    If LCase$(findText) <> findText And LCase$(findText) <> UCase$(findText) Then
        nF = nF + 1: forms(nF) = LCase$(findText)
    End If

    Dim k As Long, repl As String, r As Range
    For k = 1 To nF
        repl = MatchCasing(forms(k), replaceText)
        ' Find's Replacement.Text caps at 255 characters; longer values fall
        ' through to the manual sweep, which assigns Range.Text directly.
        If Len(repl) <= 255 Then
            Set r = rng.Duplicate
            With r.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                If protectCitations Then .Font.Italic = False
                .text = forms(k)
                .Replacement.text = repl
                .Forward = True
                .Wrap = wdFindStop
                .MatchCase = True
                .MatchWholeWord = whole
                .MatchWildcards = False
                If .Execute(Replace:=wdReplaceAll) Then NativeReplacePasses = True
            End With
        End If
    Next k
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
    Dim swept As Long: swept = 0
    Do
        ' Pump the message queue every few hits WITHIN a term, not just between
        ' terms: a single name with hundreds of occurrences used to run this
        ' whole loop without a DoEvents, and Windows marks Word "Not Responding"
        ' after a few silent seconds even though work is progressing.
        swept = swept + 1
        If swept Mod 20 = 0 Then PumpQueue
        With scan.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = findText
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False
            ' Whole-token restriction is enforced by WholeTokenBoundaries below,
            ' NOT by Word's MatchWholeWord: the latter counts an apostrophe as a
            ' word character (so "Thorne's" never matched and the name leaked)
            ' and mishandles tokens hugged by parentheses/quotation marks. Search
            ' as a plain substring; our own boundary check decides what to keep.
            .MatchWholeWord = False
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
        ' scan now spans the matched text. Two independent reasons to leave a
        ' match in place:
        '
        '  1. Whole-token names/case numbers (whole = True) must not be rewritten
        '     inside a larger word. WholeTokenBoundaries requires the characters
        '     on either side of the match to be non-alphanumeric, so "(Nash)",
        '     a quoted name, and the possessive "Thorne's" are all caught while
        '     "Nash" inside "Nashville" is not. Multi-word / punctuated finds
        '     (whole = False) keep the old literal-substring behavior.
        '
        '  2. protectCitations (re-anonymize): a match in italic text is a cited
        '     authority, left as-is so a published case name that shares a party
        '     surname isn't rewritten in the shared copy.
        Dim doReplace As Boolean: doReplace = True
        If whole Then
            If Not WholeTokenBoundaries(scan) Then doReplace = False
        End If
        If doReplace And protectCitations And scan.Font.Italic = True Then doReplace = False
        If doReplace Then
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

' True when the matched range is a whole alphanumeric token -- the characters
' immediately before AND after it are not letters or digits. This is our own
' word-boundary test, used instead of Word's MatchWholeWord for single-token
' names and case numbers. Word's version treats an apostrophe as a word
' character, so a possessive ("Thorne's") read as one word and never matched --
' the real name then leaked through both redaction and de-redaction -- and it
' has been observed to skip tokens pressed against parentheses or quotation
' marks. Here EVERY non-alphanumeric character counts as a boundary: "(", ")",
' "[", "]", straight or curly quotes, apostrophes, spaces, punctuation, and the
' story edge. So "(Nash)", "“Nash”", and "Thorne's" all qualify, while
' "Nash" inside "Nashville" does not.
Private Function WholeTokenBoundaries(ByVal matched As Range) As Boolean
    On Error Resume Next
    WholeTokenBoundaries = True

    ' Character immediately before the match (empty at the story start).
    Dim pb As Range: Set pb = matched.Duplicate
    pb.Collapse Direction:=wdCollapseStart
    If pb.MoveStart(wdCharacter, -1) <> 0 Then
        If IsAlnumChar(pb.text) Then
            WholeTokenBoundaries = False
            Exit Function
        End If
    End If

    ' Character immediately after the match (empty at the story end).
    Dim pa As Range: Set pa = matched.Duplicate
    pa.Collapse Direction:=wdCollapseEnd
    If pa.MoveEnd(wdCharacter, 1) <> 0 Then
        If IsAlnumChar(pa.text) Then WholeTokenBoundaries = False
    End If
End Function

' True for a single ASCII letter or digit. A one-character range at a boundary
' may also come back as a paragraph mark, field control, or curly quote -- all
' correctly reported as non-alphanumeric (a boundary).
Private Function IsAlnumChar(ByVal c As String) As Boolean
    If Len(c) <> 1 Then Exit Function
    IsAlnumChar = (c Like "[A-Za-z0-9]")
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
' Casing is decided WORD BY WORD when the two sides have the same word count,
' and only phrase-wide when they do not.
'
' The whole-phrase rule below has one branch that cannot be right for a
' multi-word row: "title / mixed case". A phrase like "that Yardley" is neither
' all-caps nor all-lowercase, so it lands there, and because the stored
' replacement "That Lin" is itself mixed the rule says "already mixed -- leave
' as authored" and writes the KEY's capitalization into the document. A key
' carrying "That Lin -> That Yardley" therefore capitalized "that" mid-sentence
' everywhere the phrase appeared:
'
'     "...on the ground that Lin was a sham defendant"     (as filed)
'     "...on the ground That Lin was a sham defendant"     (as restored)
'
' Word by word the same pair is decided correctly: "that" is all-lowercase, so
' its counterpart lowercases; "Yardley" is title case, so "Lin" is left as
' authored. The document's own capitalization survives the round trip, which is
' the property that matters -- the key says what a name IS, never how a
' sentence capitalizes it.
'
' Falls back to the phrase-wide rule when the word counts differ (a one-word
' fake for a two-word real, an address, an e-mail), where there is no
' correspondence to walk.
Private Function MatchCasing(ByVal matched As String, _
                              ByVal replaceText As String) As String
    Dim mw() As String, rw() As String
    mw = Split(matched, " ")
    rw = Split(replaceText, " ")
    If UBound(mw) > 0 And UBound(mw) = UBound(rw) Then
        Dim i As Long, out As String
        For i = 0 To UBound(mw)
            If i > 0 Then out = out & " "
            out = out & MatchCasingWord(mw(i), rw(i))
        Next i
        MatchCasing = out
        Exit Function
    End If
    MatchCasing = MatchCasingWord(matched, replaceText)
End Function

' The original phrase-wide rule, now applied per word by MatchCasing above.
Private Function MatchCasingWord(ByVal matched As String, _
                              ByVal replaceText As String) As String
    Dim u As String: u = UCase$(matched)
    Dim l As String: l = LCase$(matched)
    If u = l Then                       ' no cased letters (e.g. a case number)
        MatchCasingWord = replaceText
    ElseIf matched = u Then             ' ALL CAPS
        MatchCasingWord = UCase$(replaceText)
    ElseIf matched = l Then             ' all lowercase
        MatchCasingWord = LCase$(replaceText)
    Else                                ' title / mixed case
        Dim ru As String: ru = UCase$(replaceText)
        Dim rl As String: rl = LCase$(replaceText)
        If replaceText = ru Or replaceText = rl Then
            ' Stored value is mono-case (ALL CAPS or all lowercase): rebuild
            ' title case so an all-caps key row still reads as a proper name.
            MatchCasingWord = ProperCase(replaceText)
        Else
            ' Already mixed (e.g. "McDonald", "John Smith") -- leave as authored.
            MatchCasingWord = replaceText
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
' RESIDUAL PSEUDONYM HIGHLIGHTING  (leak safety net)
'==============================================================================
' The pseudonymizer draws every fake from a fixed pool of words plus a handful
' of placeholder email domains. Highlight any of them STILL present in the
' document -- even embedded inside a larger word -- so a fake that leaked in (a
' key the tool missed, an odd inflection, a stray occurrence) can't slip through
' unnoticed. Returns the number of occurrences highlighted. Public so the
' close-before-review pass (modMain.RunAllDocumentChecks) can run the same
' safety net on every reviewed document, not only right after de-anonymize.
'
' Substring matching (MatchWholeWord = False) is intentional and per the user's
' request. Name/place words are matched case-sensitively: only a first-capital
' form ("Nash") or an all-caps form ("NASH") is flagged, so a lowercase
' occurrence -- whether a stray "nash" or the "vance" buried in "advance" -- is
' left alone. Capitalized prose that happens to be a pool word (e.g. "Cedar",
' "Granite") can still be flagged, which is fine for a review aid -- the user
' clears those by eye. The email domains are lowercase with dots, so they are
' matched literally and case-insensitively instead.
'
' bodyOnly: kept for callers that want the body alone. The close review used to
' pass True, because its highlight clearers (ClearCheckHighlights /
' ClearAllHighlightsExceptYellow) swept only the body and a flag left in a header
' could never be removed. Those clearers now sweep every reviewed story, so the
' close review leaves this False and gets the full coverage the manual
' de-anonymize caller has always had.
Public Function HighlightResidualPseudonyms(ByVal oDoc As Document, _
                                            Optional ByVal bodyOnly As Boolean = False) As Long
    Dim pool As Variant: pool = PseudonymPool()
    Dim doms As Variant: doms = EmailDomainPool()
    Dim total As Long: total = 0

    ' Main body.
    total = total + HighlightFakesInRange(oDoc.content, pool, doms)
    If bodyOnly Then
        HighlightResidualPseudonyms = total
        Exit Function
    End If

    ' Headers and footers, section by section.
    Dim sec As Section, hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then total = total + HighlightFakesInRange(hf.Range, pool, doms)
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then total = total + HighlightFakesInRange(hf.Range, pool, doms)
        Next hf
    Next sec

    ' Footnotes / endnotes, only when present.
    On Error Resume Next
    If oDoc.Footnotes.count > 0 Then _
        total = total + HighlightFakesInRange(oDoc.StoryRanges(wdFootnotesStory), pool, doms)
    If oDoc.Endnotes.count > 0 Then _
        total = total + HighlightFakesInRange(oDoc.StoryRanges(wdEndnotesStory), pool, doms)
    On Error GoTo 0

    HighlightResidualPseudonyms = total
End Function

' Highlight every pool word and every email domain in one range. Returns the
' count. Words use the case-sensitive first-capital/all-caps rule; domains are
' matched literally and case-insensitively.
'
' Read this story's text ONCE and use a fast in-memory InStr to decide which
' terms are worth a real Find. Without that filter the pass is ~870 pool words in
' two case forms plus the domains -- around 1,800 native scans of every story,
' per document -- which is a large part of what made a long key feel like a hang.
' Now a term that isn't in the story costs a string search instead of a scan.
' That filter is also what makes the pool's SIZE cheap: it more than doubled when
' PDF-Linker's pools were enlarged for a case that needed 305 distinct name
' words, and the cost of the words nobody used is one InStr each.
'
' The filter is deliberately case-INSENSITIVE while the highlighters below are
' case-sensitive: it only has to be a superset. A word absent case-insensitively
' is absent in every case form, and an empty haystack (text unreadable) falls
' through to scanning everything, so nothing is ever missed.
Private Function HighlightFakesInRange(ByVal rng As Range, ByVal pool As Variant, _
                                        ByVal doms As Variant) As Long
    Dim total As Long, k As Long

    Dim raw As String, hayLower As String
    Dim readOK As Boolean: readOK = False
    On Error Resume Next
    raw = rng.text
    readOK = (Err.Number = 0)
    On Error GoTo 0
    hayLower = LCase$(raw)

    ' An EMPTY story (most of a section's six header/footer slots hold nothing
    ' but a paragraph mark) contains no term, so skip it outright. Without this
    ' it fell into TermMaybePresent's "couldn't read the text, assume present"
    ' fallback -- which is right for a READ FAILURE but wrong for a genuinely
    ' empty story, and it made every blank header pay all ~1,800 native sweeps.
    If readOK And Len(Trim$(Replace(Replace(hayLower, vbCr, ""), Chr$(7), ""))) = 0 Then
        Exit Function
    End If

    Dim term As String
    For k = LBound(pool) To UBound(pool)
        term = CStr(pool(k))
        If TermMaybePresent(hayLower, term) Then
            total = total + HighlightWordInRange(rng, term)
        End If
        If k Mod 50 = 0 Then PumpQueue
    Next k

    For k = LBound(doms) To UBound(doms)
        term = CStr(doms(k))
        If TermMaybePresent(hayLower, term) Then
            total = total + HighlightLiteralCI(rng, term)
        End If
    Next k

    HighlightFakesInRange = total
End Function

' Highlight every occurrence of a lowercase literal fake -- an email domain such
' as "example.com" -- in a range, case-insensitively and even inside a larger
' token, in pink. Returns the count. Domains are lowercase and contain dots, so
' the capitalized-word rule used for name fakes does not apply: match verbatim.
Private Function HighlightLiteralCI(ByVal rng As Range, ByVal term As String) As Long
    On Error Resume Next
    If Len(term) = 0 Then Exit Function
    Dim r As Range: Set r = rng.Duplicate
    Dim n As Long: n = 0
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = term
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False        ' flag the domain even inside a larger token
        .MatchWildcards = False
        Do While .Execute
            r.HighlightColorIndex = wdPink
            n = n + 1
            If n >= MAX_HITS_PER_TERM Then Exit Do
            If n Mod 50 = 0 Then PumpQueue  ' keep Word's queue serviced mid-term
        Loop
    End With
    HighlightLiteralCI = n
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
            If n >= MAX_HITS_PER_TERM Then Exit Do
            If n Mod 50 = 0 Then PumpQueue  ' keep Word's queue serviced mid-term
        Loop
    End With
    HighlightExact = n
End Function

' The fixed pool of fake words the pseudonymizer assigns: person surnames,
' entity/company words, street names, and city/locality names. Built in chunks
' (each under VBA's line-length limit) and split on spaces. The four categories
' are disjoint upstream, so no word is listed twice.
'
' This is a COPY of PDF-Linker's pools -- _PN_NAME_WORDS, _PN_ENTITY_WORDS,
' _PN_STREET_NAMES, _PN_CITY_NAMES, in that order -- and the copy is what makes
' the leak net work: a fake drawn from a word this list does not have is a
' pseudonym that ships through the review unflagged. So it has to be re-synced
' whenever those pools grow, which they do when a case runs the pool out and the
' tool starts numbering its stand-ins ("Deverell5"). Last synced at 695 name
' words (the largest folder seen needed 305 distinct ones), 108 entity, 55
' street, 15 city -- 873 in all, and every one of them distinct.
Private Function PseudonymPool() As Variant
    Dim s As String
    ' Person surnames
    s = "Ashford Bennett Calder Danforth Ellery Fenwick Garrick Halloran Ingram Jarrett Keswick Langley Marlowe Nash Orwell Prescott Quill Radley Sable Thorne Underwood Vance Whitlock Yardley Ashby Brandt Corwin Delacroix Everts Fairfax Grantham Holloway Isley Jennings Kingsley Lathrop Merrick Norwood Ackerly Bramble Colfax Denning Emmett Forsythe Gable Hendry Ivers Joplin Kessler Lorne Mabry Nolan Ondine Pruett Renwick Sterling"
    s = s & " Tolliver Ursin Verity Waverly Alden Beaumont Carrow Delane Abernathy Alcott Amberly Ashcombe Atwater Balfour Bancroft Barlowe Bexley Blackwood Braddock Brimley Cadwell Calloway Carden Cartwright Chadwick Chamberlin Chetwood Clarendon Cleary Cranston Cresswell Darrow Davenport Deverell Doran Dunhill Eastwick Edgerton Ellsworth Fairbank Fallon Farraday Fenmore Finnegan Gaskell Gearhart Goddard Hadley Halstead Hargrove Hartwell"
    s = s & " Hawkridge Hemsley Hollis Huxley Ingersoll Jarrow Kimball Kinsley Larkin Ledbetter Linford Lockwood Ludlow Mallory Mansfield Marsden Mayhew Milburn Montrose Mowbray Oakley Ormsby Paget Parrish Pemberton Penrose Prentiss Quenby Ramsey Rathbone Redmond Ridley Rockwell Rutherford Sackett Selwyn Sheridan Sinclair Stanhope Stockton Swinton Thorpe Trafford Tremaine Underhill Vickers Wadsworth Waldron Warwick Wescott Wexford Whitby"
    s = s & " Winslow Wolcott Wycliffe Yates Yorke Ainsworth Braxton Denholm Harrell Kimbrell Lassiter Mercer Northcott Ravenscar Stroud Thackeray Weatherby Aldous Birkett Crandall Eldridge Fanshawe Grimwood Harkness Loxley Merton Pennington Rushton Sedgwick Tennant Waverley Wharton Yeardley Abbotsford Ackworth Adderley Alverstone Anstruther Applewhite Arbuthnot Ardleigh Armitage Ashdown Astley Atherton Attwell Aylesworth Babbington"
    s = s & " Baddeley Bagshaw Bainbridge Balcombe Bardsley Barkworth Bartholomew Battersby Beckford Bedingfield Bellingham Benfield Beresford Berrington Bickerton Biddulph Billingsley Birchall Blakeney Blandford Bolingbroke Bosworth Bracewell Braithwaite Brancaster Brandreth Brereton Bridgewater Brindley Bristow Broadbent Brockhurst Broughton Buckminster Bulstrode Burghley Burnaby Burstall Cadogan Callender Camberwell Canfield Cantrell"
    s = s & " Cardington Carmichael Carnforth Carstairs Cathcart Cavendish Caxton Chalmers Chandos Charnley Chatterton Chelmsford Cheriton Chilcott Chillingworth Chorley Claverhouse Cliveden Coldwell Colgrave Collingwood Conistone Cotterill Coverdale Crawshaw Creighton Crossfield Culpepper Cunliffe Curzon Dalgleish Dalrymple Danbury Darlington Dashwood Daventry Delamere Denbigh Derwent Devereaux Dinsdale Ditchfield Dorrington Doughty"
    s = s & " Drakeford Drummond Dunstable Durward Dysart Easterbrook Eccleston Edgecombe Elderfield Elphinstone Elverton Endicott Erskine Etherington Everard Ewbank Fairweather Farnham Farrington Fearnley Felsham Fennimore Ferrers Fetherston Fitzalan Fitzhugh Fleetwood Flintham Fordyce Forrester Fothergill Framley Frobisher Fulmer Gadsby Gainsford Galbraith Gallimore Gatliff Gawthorpe Gedney Gilliatt Glanville Glossop Godalming Goodenough"
    s = s & " Grafton Greenhalgh Gresham Greville Grimshaw Grosvenor Gulliver Guthrie Hackforth Haddington Halliwell Hambleton Hammersley Hardcastle Harmsworth Harrowgate Haslemere Hatherleigh Haverford Hawksmoor Haythorne Headington Heathcote Henshaw Hepworth Herriot Hexham Hillingdon Hindmarsh Hobhouse Holbrooke Hollingsworth Holmwood Hopcroft Hornby Horrocks Houghton Hovingham Howarth Hulbert Hunsdon Huntingdon Hurstwood Hutchings"
    s = s & " Ilchester Ingleby Inskip Isherwood Jacoby Jardine Jeavons Jellicoe Jephson Jesmond Jevington Jocelyn Jolliffe Kearsley Keighley Kelsall Kendrick Kenilworth Kenmare Kentridge Kerrigan Kettering Kilbride Kilmartin Kingscote Kinnaird Kirkbride Knatchbull Knowlton Kyneston Lambourne Lanchester Landseer Langdale Lanyon Latimer Laverock Lavington Leconfield Leighton Lennox Lethbridge Leveson Lilburne Lindisfarne Linthorpe Livesey"
    s = s & " Llewellyn Lonsdale Lovelace Lowther Lumsden Luscombe Lyndhurst Lyttelton Macaulay Maidstone Mainwaring Malvern Manningham Marchmont Markham Marchbanks Marlborough Marston Masterman Maudsley Maulden Maynard Melbury Meriwether Micklethwaite Middleton Mildmay Millington Milverton Minchin Mirfield Molyneux Monkton Moorcroft Mordaunt Morecambe Mortimer Mountjoy Muirhead Mulcaster Murchison Musgrave Nasmyth Neville Newcombe"
    s = s & " Newington Nightingale Norbury Northam Nuneaton Nuttall Oakenshaw Oldbury Oldcastle Ollerton Orchardson Ormerod Osbaldeston Osgood Otterbourne Oughtred Overbury Oxenham Padgett Paignton Palliser Pargeter Parminter Patchett Pattinson Peachey Pelham Pendlebury Penhaligon Pentreath Percival Perriman Petherbridge Pevensey Pickersgill Pilkington Plackett Plumstead Polkinghorne Pomeroy Ponsonby Poulton Prendergast Prestbury Prideaux"
    s = s & " Pringle Purcell Pyecroft Quarrington Quennell Quimby Quintrell Radcliffe Ranelagh Rathmore Ravensworth Rawlinson Rayburn Redfern Redgrave Rendlesham Restarick Ribblesdale Rickerby Riddington Rimmington Rivenhall Robsart Rockingham Rolleston Romilly Rookwood Roscommon Rossiter Rothbury Rowntree Ruddock Rushbrooke Rutledge Saddlington Salkeld Saltonstall Sandbrook Sandringham Satterthwaite Saunderson Savernake Scarborough"
    s = s & " Seabright Seagrave Seaton Selborne Sempill Severn Shackleton Sharnbrook Shawcross Shelmerdine Shenstone Sherbourne Shipley Shrewsbury Skelton Smallwood Snelgrove Somerville Southwell Spofforth Stapleton Staveley Stebbing Stenhouse Stopford Stourton Stowell Stradbroke Strangeways Stretton Studholme Sudeley Sunderland Sutcliffe Swaffield Swanwick Sydenham Symington Talbot Tarleton Tattersall Teasdale Templeton Thelwall Thirlwall"
    s = s & " Thistlewood Thornbury Thrapston Throckmorton Thurlow Tilbury Tindall Tiverley Todhunter Tollemache Towneley Trelawney Trenholme Trevelyan Trewin Trumbull Tunstall Turnbull Twyford Tyndale Tyrwhitt Ullswater Upcott Uppingham Urquhart Uxbridge Vandeleur Vansittart Varley Vaughan Ventris Verinder Vesey Villiers Vinall Voysey Wakefield Walsingham Wanstead Warburton Wardlaw Warnford Wavertree Weddell Welbeck Wellesley Wemyss"
    s = s & " Wendover Wentworth Westenra Westerham Wetherall Whalley Wheatcroft Whichcote Whitcombe Whittaker Wickham Widdecombe Wilberforce Wilbraham Willoughby Wimborne Winchcombe Windlesham Wingrave Winstanley Winterbourne Wisbech Withington Wivenhoe Woburn Wollaston Woodbridge Woolnough Wootton Worsley Wrenbury Wrightson Wykeham Wyndham Yaxley Yeovil Youlgrave Younghusband Zouche"
    ' Entity / company words
    s = s & " Aldrin Brightwater Cascadia Dunmore Everline Foxglen Granite Havenwood Ironbridge Fernvale Kestrel Lumen Meridian Northgate Oakmont Pinnacle Quarry Redwood Silverpeak Torchlight Umbra Vantage Westmark Zephyr Ambrose Beacon Cobalt Drayton Emberly Falcon Gladstone Harborview Ivory Jetstream Kaldor Kingsmere Monarch Nimbus Orion Pembroke Arclight Brookstone Cairnwood Clearspring Crestline Dovewood Eastmark Eldergrove Ferncliff"
    s = s & " Fieldstone Foxbridge Glenrock Goldcrest Graystone Highpoint Hollowmere Ironwood Kirkwall Lakemont Lanternwood Ledgewood Marbury Millbrook Moorland Oakspire Overland Parkhurst Pinehurst Ravenwood Riverton Rockhaven Sablewood Sandpiper Shorewood Silvergate Solstice Springvale Starling Stonehaven Thornfield Timberline Wexmoor Whitfield Wildmere Windermere Ashcroft Brightmoor Coppervale Dawnfield Emberton Frostgate Greenhollow"
    s = s & " Hartland Ironclad Keystone Lightwell Meadowgate Northwind Opalridge Pinecrest Quillmark Rosemont Stormont Truenorth Umberwood Vanguard Wellspring Yarrowvale"
    ' Street names
    s = s & " Cedar Birch Willow Aspen Juniper Laurel Poplar Hawthorn Linden Chestnut Sequoia Cypress Alder Dogwood Hickory Rosewood Foxglove Larkspur Tamarack Sorrel Bayberry Bilberry Blackthorn Bracken Buckthorn Catalpa Chicory Cinquefoil Clematis Comfrey Elderberry Fernbrake Gentian Hazelnut Hornbeam Ironbark Ivywood Marjoram Mulberry Myrtlewood Persimmon Quince Sagebrush Silverbell Snowberry Sumac Tanglewood Teasel Thornapple Trefoil"
    s = s & " Vervain Wintergreen Witchhazel Yarrowleaf Yewberry"
    ' City / locality names
    s = s & " Fairview Brookfield Rosedale Elmwood Kingsbury Northvale Westbrook Clearwater Havenport Stonebridge Marlow Redhill Glenmore Oakhurst Bridgeton"
    PseudonymPool = Split(s)
End Function

' The placeholder email domains the anonymizer assigns. Matched literally and
' case-insensitively (they are lowercase and contain dots), unlike the
' capitalized name/place fakes in PseudonymPool.
Private Function EmailDomainPool() As Variant
    Dim s As String
    s = "example.com mailhaven.net postbox.org letterbox.co postbay.org mailglen.net inboxvale.com mailcrest.net penbox.org quillbox.net mailridge.com postgate.net letterfield.org inboxharbor.com mailbrook.net postvale.org letterglen.co mailstead.com postholm.net inboxmere.org mailwick.co postloft.com letterdale.net mailbourne.org postcairn.co inboxfern.com mailthorne.net postbriar.org letterhollow.co mailquarry.com postbeacon.net"
    s = s & " inboxlantern.org"
    EmailDomainPool = Split(s)
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
