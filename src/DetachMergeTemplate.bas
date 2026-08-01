Attribute VB_Name = "DetachMergeTemplate"
' ============================================================
' DetachMergeTemplate
'
' Repairs a document that is still attached to a macro-bearing
' template. Tentative drafts produced by the Mail Merge Order
' Template inherit AttachedTemplate from the merge main
' document, so every saved draft points back at that .dotm.
' Word loads an attached template's VBA project with the
' document and gives it precedence over this add-in, so the
' template's macros and key bindings win.
'
' UpdateStylesOnOpen is cleared BEFORE the reattach, so
' swapping to Normal cannot pull Normal's styles into the
' document the next time it opens.
'
' AutoOpen handles this automatically for documents it sees;
' this is the manual override for one document on demand.
' ============================================================
Option Explicit

Public Sub DetachFromMergeTemplate()
    Dim sOld As String

    If Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "Detach Template"
        Exit Sub
    End If

    On Error GoTo Fail

    sOld = ActiveDocument.AttachedTemplate.Name

    If LCase(sOld) = "normal.dotm" Then
        MsgBox "Already attached to Normal.dotm. Nothing to do.", _
               vbInformation, "Detach Template"
        Exit Sub
    End If

    ActiveDocument.UpdateStylesOnOpen = False
    ActiveDocument.AttachedTemplate = ""
    If ActiveDocument.ReadOnly = False Then ActiveDocument.Save

    MsgBox "Detached from " & sOld & "." & vbCrLf & _
           "Now attached to Normal.dotm.", _
           vbInformation, "Detach Template"
    Exit Sub

Fail:
    MsgBox "Could not detach." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Detach Template"
End Sub
