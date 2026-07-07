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
' MACRO YOU RUN:
'   DeAnonymizeTentative - locate the key, then replace every fake with its
'                          real value throughout the document body.
'
' NOTES:
'   - The draft is a regular document (no live mail-merge fields): every fake
'     -- caption, party block, and body prose -- is plain text and gets
'     restored. Replacement runs across all stories: main text, headers,
'     footers, footnotes, and text boxes.
'   - Longest fakes are replaced first so a bare-surname fake never rewrites
'     part of a longer full-name fake.
'   - Reads .xlsx via Excel automation. The rare JSON fallback that PDF-Linker
'     writes only when openpyxl is missing is not supported.
'==============================================================================
Option Explicit

' PDF-Linker writes "pseudonym_key.xlsx"; match that plus any de-duplicated
' copies Windows may create (e.g. "pseudonym_key (1).xlsx"). Newest wins.
Private Const KEY_PATTERN As String = "pseudonym_key*.xlsx"

Private Type Mapping
    real As String
    fake As String
End Type

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub DeAnonymizeTentative()
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
    SortMappingsByFakeLenDesc maps, nMaps

    If MsgBox("Restore real names using " & nMaps & " mapping(s) from:" & vbCrLf & vbCrLf & _
              keyPath & vbCrLf & vbCrLf & _
              "This replaces every pseudonym throughout the document with its " & _
              "real value. Ctrl+Z undoes it.", _
              vbYesNo + vbQuestion, "De-Anonymize") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Dim prevTrack As Boolean: prevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False

    Dim oUndo As UndoRecord: Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "De-Anonymize Tentative"

    Dim totalHits As Long, distinctHits As Long, i As Long
    For i = 1 To nMaps
        Dim hits As Long
        hits = ReplaceEverywhere(oDoc, maps(i).fake, maps(i).real)
        If hits > 0 Then
            totalHits = totalHits + hits
            distinctHits = distinctHits + 1
        End If
    Next i

    oUndo.EndCustomRecord
    oDoc.TrackRevisions = prevTrack
    Application.ScreenUpdating = True

    MsgBox "De-anonymized: " & totalHits & " replacement(s) across " & _
           distinctHits & " pseudonym(s)." & vbCrLf & vbCrLf & _
           "Review the result before finalizing.", vbInformation, "De-Anonymize"
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
' Replace findText with replaceText across every story in the document -- main
' text, plus each header/footer/footnote/text-box story reached via
' NextStoryRange. Returns the total number replaced.
Private Function ReplaceEverywhere(ByVal oDoc As Document, _
                                    ByVal findText As String, _
                                    ByVal replaceText As String) As Long
    Dim total As Long: total = 0
    If Len(findText) = 0 Then Exit Function

    Dim whole As Boolean: whole = ShouldWholeWord(findText)

    Dim story As Range
    For Each story In oDoc.StoryRanges
        Dim s As Range: Set s = story
        Do While Not s Is Nothing
            total = total + ReplaceInRange(s, findText, replaceText, whole)
            ' NextStoryRange raises an error past the last linked story on some
            ' builds; trap it and end the walk rather than looping forever.
            Dim nxt As Range: Set nxt = Nothing
            On Error Resume Next
            Set nxt = s.NextStoryRange
            On Error GoTo 0
            Set s = nxt
        Loop
    Next story

    ReplaceEverywhere = total
End Function

' Count occurrences of findText within one story range, then replace them all.
' Counting first (on a duplicate, so the original range is untouched) keeps the
' report accurate without a replace-loop.
Private Function ReplaceInRange(ByVal rng As Range, _
                                 ByVal findText As String, _
                                 ByVal replaceText As String, _
                                 ByVal whole As Boolean) As Long
    ReplaceInRange = 0

    Dim cnt As Long: cnt = 0
    Dim cRng As Range: Set cRng = rng.Duplicate
    With cRng.Find
        .ClearFormatting
        .text = findText
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = whole
        .MatchWildcards = False
        Do While .Execute
            cnt = cnt + 1
            If cnt > 100000 Then Exit Do        ' runaway guard
        Loop
    End With
    If cnt = 0 Then Exit Function

    Dim rRng As Range: Set rRng = rng.Duplicate
    With rRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = whole
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

    ReplaceInRange = cnt
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
' SORT  (fake length, descending)
'==============================================================================
Private Sub SortMappingsByFakeLenDesc(ByRef maps() As Mapping, ByVal nMaps As Long)
    Dim i As Long, j As Long, tmp As Mapping
    For i = 1 To nMaps - 1
        For j = 1 To nMaps - i
            If Len(maps(j).fake) < Len(maps(j + 1).fake) Then
                tmp = maps(j)
                maps(j) = maps(j + 1)
                maps(j + 1) = tmp
            End If
        Next j
    Next i
End Sub
