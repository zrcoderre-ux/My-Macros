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
    Application.ScreenUpdating = True

    SetDocFlag oDoc, DEANON_DONE_VAR      ' don't auto-run again on close

    MsgBox "De-anonymized: restored " & distinctHits & " of " & nMaps & _
           " pseudonym(s)." & vbCrLf & vbCrLf & _
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

    ' Save the current document AS the new file first, so all edits and the
    ' metadata strip happen on a fresh file (no inherited version history) and
    ' the original file is never written to.
    oDoc.SaveAs2 FileName:=savePath, FileFormat:=wdFormatXMLDocument, _
                 AddToRecentFiles:=False
    ' oDoc is now bound to savePath.

    ' Mark this file as re-anonymize output so the close hook never tries to
    ' de-anonymize it back to real names.
    SetDocFlag oDoc, REANON_CREATED_VAR

    Dim prevTrack As Boolean: prevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False
    Dim prevAutoSave As Boolean: prevAutoSave = False
    On Error Resume Next
    prevAutoSave = oDoc.AutoSaveOn
    oDoc.AutoSaveOn = False
    On Error GoTo ErrH

    ' Reverse direction: replace each real value with its fake. No custom undo
    ' record (it overflows and crashes Word on large documents).
    Dim distinctHits As Long, i As Long
    For i = 1 To nMaps
        If ReplaceEverywhere(oDoc, maps(i).real, maps(i).fake) > 0 Then
            distinctHits = distinctHits + 1
        End If
        If i Mod 5 = 0 Then DoEvents
    Next i

    ' Strip metadata: comments, revisions, versions, and personal/document
    ' information. Combined with the new file, this leaves no trail back to the
    ' real matter.
    On Error Resume Next
    oDoc.RemoveDocumentInformation wdRDIAll
    On Error GoTo ErrH

    oDoc.TrackRevisions = prevTrack
    On Error Resume Next
    oDoc.AutoSaveOn = prevAutoSave
    On Error GoTo ErrH

    oDoc.Save
    Application.ScreenUpdating = True

    MsgBox "Re-anonymized: replaced " & distinctHits & " of " & nMaps & _
           " value(s)." & vbCrLf & vbCrLf & _
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
           "Error " & reN & ": " & reD, vbExclamation, "Re-Anonymize"
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

    Dim folder As String: folder = Doc.path
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
' (the key travels with the document -- often Downloads, but not always).
' Fall back to a file picker, starting in that folder, if none is found.
Private Function ResolveKeyPath(ByVal oDoc As Document) As String
    Dim docFolder As String
    docFolder = ""
    On Error Resume Next
    docFolder = oDoc.Path          ' "" if the document has never been saved
    On Error GoTo 0

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
    End If
    xl.Visible = False
    xl.DisplayAlerts = False

    Dim wb As Object
    Set wb = xl.Workbooks.Open(FileName:=path, ReadOnly:=True, AddToMRU:=False)

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

    wb.Close SaveChanges:=False
    If startedXl Then xl.Quit
    Set wb = Nothing: Set xl = Nothing

    ReadPseudonymKey = (nMaps > 0)
    Exit Function

CleanFail:
    On Error Resume Next
    wb.Close SaveChanges:=False
    If startedXl Then xl.Quit
    On Error GoTo 0
    ReadPseudonymKey = False
    Exit Function

Fail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
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
' main body, each section's headers/footers, and footnotes/endnotes when
' present. Returns the number of those ranges in which a replacement was made.
'
' This deliberately does NOT walk StoryRanges/NextStoryRange (including text
' frames): enumerating that collection while doing wdReplaceAll inside the loop
' can destabilize and crash Word. The collections below stay valid across text
' replacement, so iterating them is safe.
Private Function ReplaceEverywhere(ByVal oDoc As Document, _
                                    ByVal findText As String, _
                                    ByVal replaceText As String) As Long
    Dim total As Long: total = 0
    If Len(findText) = 0 Then Exit Function

    Dim whole As Boolean: whole = ShouldWholeWord(findText)

    ' Main body (caption, party block, and prose are all here in a plain draft).
    If ReplaceInRange(oDoc.content, findText, replaceText, whole) Then total = total + 1

    ' Headers and footers, section by section.
    Dim sec As Section
    Dim hf As HeaderFooter
    For Each sec In oDoc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                If ReplaceInRange(hf.Range, findText, replaceText, whole) Then total = total + 1
            End If
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then
                If ReplaceInRange(hf.Range, findText, replaceText, whole) Then total = total + 1
            End If
        Next hf
    Next sec

    ' Footnotes / endnotes, only when present (accessing the story otherwise errors).
    On Error Resume Next
    If oDoc.Footnotes.count > 0 Then
        If ReplaceInRange(oDoc.StoryRanges(wdFootnotesStory), findText, replaceText, whole) Then total = total + 1
    End If
    If oDoc.Endnotes.count > 0 Then
        If ReplaceInRange(oDoc.StoryRanges(wdEndnotesStory), findText, replaceText, whole) Then total = total + 1
    End If
    On Error GoTo 0

    ReplaceEverywhere = total
End Function

' Replace every occurrence of findText with replaceText in one range, in a
' single pass. Returns True if at least one replacement was made. Execute with
' wdReplaceAll returns that flag directly, so no counting loop is needed.
Private Function ReplaceInRange(ByVal rng As Range, _
                                 ByVal findText As String, _
                                 ByVal replaceText As String, _
                                 ByVal whole As Boolean) As Boolean
    On Error Resume Next
    Dim r As Range: Set r = rng.Duplicate
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = whole
        .MatchWildcards = False
        ReplaceInRange = .Execute(Replace:=wdReplaceAll)
    End With
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
