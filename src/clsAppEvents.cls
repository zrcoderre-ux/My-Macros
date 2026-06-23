VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As word.Application
Attribute App.VB_VarHelpID = -1

'=============================================================================
' SELECTION CHANGE  (new)
' Fires after every cursor movement, including each character typed.
' Delegates to the autocomplete module, which decides whether anything
' needs to happen. Wrapped in error handling so a bad event can never
' break the user's editing session.
'=============================================================================
Private Sub App_WindowSelectionChange(ByVal sel As Selection)
    On Error Resume Next
    Debug.Print "WSC: mode=" & ParentheticalAutocomplete.m_Mode & _
                " anchor=" & ParentheticalAutocomplete.m_AnchorPos & _
                " last=" & ParentheticalAutocomplete.m_LastSeenPos & _
                " curPos=" & Selection.start & _
                " typed=[" & ParentheticalAutocomplete.m_TypedSoFar & "]"
    OnSelectionChanged
    On Error GoTo 0
End Sub

'=============================================================================
' DOCUMENT BEFORE CLOSE  (existing, unchanged)
'=============================================================================
Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    If Not IsInOneDrive(Doc) Then
        MsgBox "Skipped: not in OneDrive. Path = " & Doc.FullName
        Exit Sub
    End If

    ' Only run checks on documents saved inside OneDrive or SharePoint
    If Not IsInOneDrive(Doc) Then Exit Sub

    ' Skip checks when the mail merge macro is closing documents mid-workflow
    If modMain.gSkipCloseChecks Then Exit Sub

    ' Self-healing guard: re-initialize if the event handler was lost
    If gAppEvents Is Nothing Then InitializeAppEvents

    ' Ask first - run no checks unless the user wants to review
    Dim respReview As Integer
    respReview = MsgBox("Review document before closing?", _
                        vbQuestion + vbYesNo, "Document Review")
    If respReview = vbNo Then Exit Sub

    ' --- Phase 1: Run all checks, highlight everything found ---
    Dim issues         As Boolean
    Dim userHighlights As Boolean
    RunAllDocumentChecks Doc, issues, userHighlights

    If issues Then
        Dim resp As Integer
        resp = MsgBox("Document contains possible issues." & vbCrLf & vbCrLf & _
                      "Highlighted items:" & vbCrLf & vbCrLf & _
                      "  Green = unmatched punctuation or double space" & vbCrLf & _
                      "  Cyan  = placeholder word ""blank""" & vbCrLf & vbCrLf & _
                      "Yes = stay open to review" & vbCrLf & _
                      "No  = close and remove highlights (keep yellow)", _
                      vbExclamation + vbYesNo, "Document Check")
        If resp = vbYes Then
            Cancel = True
        Else
            ClearAllHighlightsExceptYellow Doc
        End If
        Exit Sub
    End If

    ' --- Phase 2: User highlight check (only reached if no issues found) ---
    If userHighlights Then
        Dim resp2 As Integer
        resp2 = MsgBox("Document contains highlighted text." & vbCrLf & vbCrLf & _
                       "You may have left reminders for yourself." & vbCrLf & vbCrLf & _
                       "Close anyway?", _
                       vbExclamation + vbYesNo, "Highlighted Text Found")
        If resp2 = vbNo Then Cancel = True
    End If
End Sub


