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
    OnSelectionChanged
    On Error GoTo 0
End Sub

'=============================================================================
' DOCUMENT BEFORE CLOSE
'=============================================================================
' The event itself does nothing but fence the review: set the mid-close flag,
' run it, and make sure neither the flag nor an error can outlive the call.
' Nothing may escape this handler -- an unhandled runtime error inside a Word
' application event fires with the document already half closed, and Word has no
' safe way to unwind it, so the user gets a crash rather than an error dialog.
Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    On Error GoTo Fail

    ' Tell the long sweeps below they are running mid-close, so none of them
    ' yields to the message pump while Word is tearing the document down.
    modMain.gInCloseReview = True
    CloseReview Doc, Cancel
    modMain.gInCloseReview = False
    Exit Sub

Fail:
    Dim eN As Long: eN = Err.Number
    Dim eD As String: eD = Err.Description
    On Error Resume Next
    modMain.gInCloseReview = False
    ' Put the screen back before saying anything: a failure partway through a
    ' check leaves redraw off, and a MsgBox over a frozen window reads as a hang.
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "The document review hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & eN & ": " & eD & vbCrLf & vbCrLf & _
           "The document will close normally.", _
           vbExclamation, "Document Review"
End Sub

' The unhighlightable findings, as a block for the issues dialog, or "" when
' there are none -- so the common case reads exactly as it always did.
Private Function NoteBlock(ByVal notes As String) As String
    If Len(notes) = 0 Then Exit Function
    NoteBlock = vbCrLf & vbCrLf & "Also found, NOT highlighted:" & vbCrLf & notes
End Function

Private Sub CloseReview(ByVal Doc As Document, ByRef Cancel As Boolean)
    ' Silent gates: the close review runs ONLY on dated work documents saved
    ' in OneDrive. If either condition fails, do nothing at all -- no popup,
    ' no checks. (A document whose title doesn't end in M.D.YYYY, or that
    ' isn't in OneDrive, just closes normally.)
    If Not IsInOneDrive(Doc) Then Exit Sub
    If Not TitleEndsWithDate(Doc) Then Exit Sub

    ' Skip checks when the mail merge macro is closing documents mid-workflow
    If modMain.gSkipCloseChecks Then Exit Sub

    ' Self-healing guard: re-initialize if the event handler was lost
    If gAppEvents Is Nothing Then InitializeAppEvents

    ' Auto-restore real names on close. Runs only if de-anonymize has not already
    ' run on this document AND the document was not produced by the re-anonymize
    ' macro (both tracked via document variables). Silently does nothing when
    ' there is no pseudonym key in the document's folder.
    DeAnonymize.RunDeAnonymizeOnClose Doc

    ' Ask first - run no checks unless the user wants to review
    Dim respReview As Integer
    respReview = MsgBox("Review document before closing?", _
                        vbQuestion + vbYesNo, "Document Review")
    If respReview = vbNo Then Exit Sub

    ' --- Auto-cleanup: strip every hyperlink as part of the review (silent) ---
    CitationLinker.RemoveAllHyperlinks_Quiet Doc

    ' --- Phase 1: Run all checks, highlight everything found ---
    Dim issues         As Boolean
    Dim userHighlights As Boolean
    Dim notes          As String
    RunAllDocumentChecks Doc, issues, userHighlights, notes

    If issues Then
        Dim resp As Integer
        resp = MsgBox("Document contains possible issues." & vbCrLf & vbCrLf & _
                      "Highlighted items:" & vbCrLf & vbCrLf & _
                      "  Green = unmatched punctuation or double space" & vbCrLf & _
                      "  Cyan  = placeholder word ""blank""" & vbCrLf & _
                      "  Pink  = possible leftover pseudonym (anonymizer fake)" & vbCrLf & vbCrLf & _
                      "The review covers headers, footers, and footnotes as well " & _
                      "as the body, so a highlight may not be in view." & _
                      NoteBlock(notes) & vbCrLf & vbCrLf & _
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

    ' Findings that could not be highlighted, with nothing else to report. Their
    ' own message: the highlight dialog above would send the user looking for a
    ' mark that does not exist.
    If Len(notes) > 0 Then
        Dim respN As Integer
        respN = MsgBox("The review found something it could not highlight:" & _
                       vbCrLf & vbCrLf & notes & vbCrLf & _
                       "Nothing is marked in the document." & vbCrLf & vbCrLf & _
                       "Yes = stay open to look" & vbCrLf & _
                       "No  = close anyway", _
                       vbExclamation + vbYesNo, "Document Check")
        If respN = vbYes Then
            Cancel = True
            Exit Sub
        End If
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


