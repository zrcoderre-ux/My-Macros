Attribute VB_Name = "Module4"
Option Explicit

' REMOVED: Dim oApp As New AutoExecClass (No longer needed)

Sub AutoExec()
    ' REMOVED: Set oApp.WordApp = word.Application (No longer needed)
    
    CustomizationContext = NormalTemplate
    
    ' This is what makes it "Global" - it maps the keys to your wrapping script
    Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:="WrapCitations.CheckAndWrap", _
        KeyCode:=BuildKeyCode(wdKeySpacebar)
        
    Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:="WrapCitations.CheckAndWrapEnter", _
        KeyCode:=BuildKeyCode(wdKeyReturn)
End Sub

Sub AutoClose()
    On Error Resume Next
    FindKey(BuildKeyCode(wdKeySpacebar)).Clear
    FindKey(BuildKeyCode(wdKeyReturn)).Clear
End Sub

