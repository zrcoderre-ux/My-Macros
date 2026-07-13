Attribute VB_Name = "Module4"
Option Explicit
' Internal: RegisterWrapKeyBindings is called by AutoExec, never by the user.
' Hide it from the Alt+F8 list while keeping it callable within the project.
Option Private Module

' Spacebar/Return citation-wrapping bindings. Called at startup by AutoExec in
' the MacroKeyBindings module. Bound in THIS template's context (not Normal),
' so the bindings stay contained to My_Macros.dotm, reapply every launch, and
' never need an AutoClose teardown.
Public Sub RegisterWrapKeyBindings()
    CustomizationContext = ThisDocument

    Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:="WrapCitations.CheckAndWrap", _
        KeyCode:=BuildKeyCode(wdKeySpacebar)

    Application.KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:="WrapCitations.CheckAndWrapEnter", _
        KeyCode:=BuildKeyCode(wdKeyReturn)

    ' Don't mark the template dirty; bindings reapply on the next launch anyway.
    ThisDocument.Saved = True
End Sub

