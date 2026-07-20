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
'                          then save a metadata-free copy as a NEW document
'                          (fresh file = no version history) so it is safe to
'                          share. The original file is left unchanged.
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
    Application.ScreenUpdating = True
    MsgBox "De-Anonymize hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & eN & ": " & eD, vbExclamation, "De-Anonymize"
End Sub

'==============================================================================
' RE-ANONYMIZE  (reverse: real -> fake, saved as a clean new document)
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

    ' Choose where to save the clean copy BEFORE changing anything, so the run
    ' can be cancelled with nothing touched. Give it a neutral default name so
    ' the real party names are never carried in the filename.
    Dim savePath As String
    savePath = PickReAnonSavePath(oDoc)
    If Len(savePath) = 0 Then Exit Sub

    If MsgBox("Re-anonymize using " & nMaps & " mapping(s) and save a " & _
              "metadata-free copy to:" & vbCrLf & vbCrLf & savePath & vbCrLf & vbCrLf & _
              "The original document on disk is left unchanged.", _
              vbYesNo + vbQuestion, "Re-Anonymize") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    ' ORDER MATTERS: every scrub below runs IN MEMORY on the open document,
    ' and only then is the result written out via SaveAs2. Saving first meant
    ' version 1 of "Anonymized Draft.docx" hit the disk -- and any synced
    ' OneDrive/SharePoint folder's server-side version history -- with every
    ' real name still in it, where RemoveDocumentInformation can't reach.
    ' The original FILE is still never written: AutoSave is disabled before
    ' the first edit, and the only saves target savePath.
    Dim prevTrack As Boolean: prevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False
    Dim prevAutoSave As Boolean: prevAutoSave = False
    On Error Resume Next
    prevAutoSave = oDoc.AutoSaveOn
    oDoc.AutoSaveOn = False              ' must precede edits: AutoSave would
    On Error GoTo ErrH                   ' push real->fake edits to the ORIGINAL

    ' Clear any pink residual-pseudonym flags a prior de-anonymize left: they
    ' mark exactly which tokens were fakes, which the shared copy must not show.
    ClearResidualFlags oDoc

    ' Reverse direction: replace each real value with its fake. protectCitations
    ' leaves names inside italic cited authorities alone, so a party surname that
    ' also names a published case isn't rewritten in the shared copy. No custom
    ' undo record (it overflows and crashes Word on large documents).
    Dim distinctHits As Long, i As Long
    For i = 1 To nMaps
        If ReplaceEverywhere(oDoc, maps(i).real, maps(i).fake, True) > 0 Then
            distinctHits = distinctHits + 1
        End If
        If i Mod 5 = 0 Then DoEvents
    Next i

    ' Blank the court-identity header (Department 515, judge, courtroom staff) so
    ' the shared copy doesn't reveal them.
    ApplyCourtIdentity oDoc, False

    ' Strip metadata: comments, revisions, versions, and personal/document
    ' information. Combined with the fresh file below, this leaves no trail back
    ' to the real matter.
    On Error Resume Next
    oDoc.RemoveDocumentInformation wdRDIAll
    On Error GoTo ErrH

    ' Mark as re-anonymize output so the close hook never tries to de-anonymize
    ' it back to real names. Set before the save so it rides into the new file.
    SetDocFlag oDoc, REANON_CREATED_VAR

    ' First and only disk write: the fully scrubbed content.
    oDoc.SaveAs2 FileName:=savePath, FileFormat:=wdFormatXMLDocument, _
                 AddToRecentFiles:=False
    ' oDoc is now bound to savePath; the original file was never written.

    oDoc.TrackRevisions = prevTrack
    On Error Resume Next
    oDoc.AutoSaveOn = prevAutoSave
    On Error GoTo ErrH

    oDoc.Save
    Application.ScreenUpdating = True

    MsgBox "Re-anonymized: replaced " & distinctHits & " of " & nMaps & _
           " value(s). The court-identity header (department, judge, staff) " & _
           "was blanked." & vbCrLf & vbCrLf & _
           "Names inside italic cited case names were left as-is so a party " & _
           "surname that also names a published case wasn't rewritten -- check " & _
           "any italicized cites if a real party name should have been replaced." & _
           vbCrLf & vbCrLf & _
           "Saved a metadata-free copy to:" & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "This window is now that copy; the original file is unchanged.", _
           vbInformation, "Re-Anonymize"
    Exit Sub

ErrH:
    Dim reN As Long: reN = Err.Number
    Dim reD As String: reD = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    MsgBox "Re-Anonymize hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & reN & ": " & reD & vbCrLf & vbCrLf & _
           "If the error happened before the save, this window holds partial " & _
           "re-anonymize edits that were NOT saved anywhere -- close it " & _
           "WITHOUT saving to get back to the untouched original.", _
           vbExclamation, "Re-Anonymize"
End Sub

' Ask where to save the anonymized copy. Defaults to the document's folder with
' a neutral name (so real party names aren't carried in the filename). Returns
' "" if cancelled. Ensures a .docx extension.
Private Function PickReAnonSavePath(ByVal oDoc As Document) As String
    Dim folder As String
    folder = ""
    On Error Resume Next
    folder = oDoc.path
    On Error GoTo 0
    If Len(folder) = 0 Then folder = Environ$("USERPROFILE") & "\Documents"

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Save the anonymized copy as"
        .InitialFileName = folder & "\Anonymized Draft.docx"
        If .Show <> -1 Then
            PickReAnonSavePath = ""
            Exit Function
        End If
        PickReAnonSavePath = .SelectedItems(1)
    End With

    If LCase$(Right$(PickReAnonSavePath, 5)) <> ".docx" Then
        PickReAnonSavePath = PickReAnonSavePath & ".docx"
    End If
End Function

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
    If HasDocFlag(Doc, REANON_CREATED_VAR) Then Exit Sub

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
            If Not .Execute Then Exit Do
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

' Remove every pink residual-pseudonym flag from the document (body, headers/
' footers, notes, text boxes). Run at the start of re-anonymize: surviving pink
' highlights in the shared copy would advertise exactly which tokens were fakes.
' Uses a highlight-seeking Find (fast) rather than walking Characters (O(n) COM
' calls). Other highlight colors -- the user's yellow, the close-review's green/
' turquoise -- are left alone.
Private Sub ClearResidualFlags(ByVal oDoc As Document)
    On Error Resume Next
    ClearPinkInRange oDoc.content

    Dim sec As Section, hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then ClearPinkInRange hf.Range
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then ClearPinkInRange hf.Range
        Next hf
    Next sec

    If oDoc.Footnotes.count > 0 Then ClearPinkInRange oDoc.StoryRanges(wdFootnotesStory)
    If oDoc.Endnotes.count > 0 Then ClearPinkInRange oDoc.StoryRanges(wdEndnotesStory)

    Dim shp As Shape
    For Each shp In oDoc.Shapes
        If shp.TextFrame.HasText Then ClearPinkInRange shp.TextFrame.TextRange
    Next shp
End Sub

Private Sub ClearPinkInRange(ByVal rng As Range)
    On Error Resume Next
    Dim r As Range: Set r = rng.Duplicate
    With r.Find
        .ClearFormatting
        .text = ""
        .Highlight = True
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            If r.HighlightColorIndex = wdPink Then r.HighlightColorIndex = wdNoHighlight
            If r.End >= rng.End Then Exit Do
            r.Collapse Direction:=wdCollapseEnd
            r.End = rng.End
        Loop
    End With
End Sub

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
