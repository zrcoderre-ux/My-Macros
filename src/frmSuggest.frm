VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSuggest 
   ClientHeight    =   405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSuggest.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSuggest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =====================================================================
' frmSuggest v3 -- single-suggestion label + counter
'
' FORM SETUP (designer):
'   1. Open frmSuggest in the VBA editor.
'   2. DELETE the existing lstSuggestions ListBox.
'   3. Properties on the form itself:
'        Name           = frmSuggest
'        Caption        = (blank)
'        BorderStyle    = 0 - fmBorderStyleNone
'        ShowModal      = False
'        StartUpPosition = 0 - Manual
'        Width          = 320
'        Height         = 22       <-- much shorter than before; just a strip
'        BackColor      = &H80000018  (window background) or pick a soft yellow
'   4. Add a Label control:
'        Name           = lblSuggestion
'        Left           = 4
'        Top            = 2
'        Width          = 240
'        Height         = 18
'        Font           = match document body (e.g. Times New Roman 11 or Calibri 11)
'        ForeColor      = &H80000012  (window text)
'        BackStyle      = 1 - Transparent
'        Caption        = (blank)
'   5. Add a second Label control:
'        Name           = lblCounter
'        Left           = 248
'        Top            = 4
'        Width          = 68
'        Height         = 14
'        Font           = same family, 8pt, italic
'        ForeColor      = &H808080  (gray)
'        BackStyle      = 1 - Transparent
'        TextAlign      = 3 - fmTextAlignRight
'        Caption        = (blank)
'
' Optional: add a thin border by setting the form's BorderStyle to
'   1 - fmBorderStyleSingle and BorderColor to a soft gray.
'
' Then right-click the form -> View Code, and paste the code below,
' replacing everything that's currently there.
' =====================================================================
Option Explicit

Private Sub UserForm_Initialize()
    ' Module owns all state; nothing to do here.
End Sub

' Clicking the suggestion accepts it (mouse fallback; not the primary path).
Private Sub lblSuggestion_Click()
    ParentheticalAutocomplete.AcceptCurrentMatch
End Sub

' With the non-activating-window fix, this form should never receive
' keystrokes -- the document does. These handlers are a defensive
' backstop in case focus lands here somehow (e.g. user clicks the label).
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                             ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab, vbKeyRight
            KeyCode = 0
            ParentheticalAutocomplete.AcceptCurrentMatch
        Case vbKeyEscape
            KeyCode = 0
            ParentheticalAutocomplete.DismissSuggest
        Case Else
            On Error Resume Next
            Application.ActiveWindow.SetFocus
            On Error GoTo 0
    End Select
End Sub


