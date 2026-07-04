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
'   - The caption / party block is filled by the mail-merge fields, which
'     already carry the real names, so this only rewrites the body prose.
'   - Longest fakes are replaced first so a bare-surname fake never rewrites
'     part of a longer full-name fake.
'   - Reads .xlsx via Excel automation. The rare JSON fallback that PDF-Linker
'     writes only when openpyxl is missing is not supported.
'==============================================================================
Option Explicit

' Default file name PDF-Linker writes (see write_key in pdf_linker.py).
Private Const KEY_FILENAME As String = "pseudonym_key.xlsx"

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
    keyPath = ResolveKeyPath()
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
              "This replaces every pseudonym in the document body with its real " & _
              "value. Ctrl+Z undoes it.", _
              vbYesNo + vbQuestion, "De-Anonymize") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Dim prevTrack As Boolean: prevTrack = oDoc.TrackRevisions
    oDoc.TrackRevisions = False

    Dim oUndo As UndoRecord: Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "De-Anonymize Tentative"

    Dim totalHits As Long, distinctHits As Long, i As Long
    For i = 1 To nMaps
        Dim hits As Long
        hits = ReplaceAllInBody(oDoc, maps(i).fake, maps(i).real)
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
           "The caption / merge fields already carry the real names; review " & _
           "them separately.", vbInformation, "De-Anonymize"
End Sub

'==============================================================================
' KEY-FILE LOCATION
'==============================================================================
' Prefer the default pseudonym_key.xlsx in Downloads (where PDF-Linker writes
' it); otherwise let the user pick the file.
Private Function ResolveKeyPath() As String
    Dim defaultPath As String
    defaultPath = Environ$("USERPROFILE") & "\Downloads\" & KEY_FILENAME
    If Dir(defaultPath) <> "" Then
        ResolveKeyPath = defaultPath
        Exit Function
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the PDF-Linker key (pseudonym_key.xlsx)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel key", "*.xlsx"
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then
            ResolveKeyPath = .SelectedItems(1)
        Else
            ResolveKeyPath = ""
        End If
    End With
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
' Count occurrences of findText in the body, then replace them all with
' replaceText. Returns the number replaced. Counting first keeps the report
' accurate without a replace-loop (real never contains its own fake, but a
' count-then-replaceAll pass is simplest and safe).
Private Function ReplaceAllInBody(ByVal oDoc As Document, _
                                   ByVal findText As String, _
                                   ByVal replaceText As String) As Long
    ReplaceAllInBody = 0
    If Len(findText) = 0 Then Exit Function

    Dim whole As Boolean: whole = ShouldWholeWord(findText)

    ' 1. Count.
    Dim cnt As Long: cnt = 0
    Dim cRng As Range: Set cRng = oDoc.content
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

    ' 2. Replace all in one pass.
    Dim rRng As Range: Set rRng = oDoc.content
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

    ReplaceAllInBody = cnt
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
