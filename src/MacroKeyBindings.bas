Attribute VB_Name = "MacroKeyBindings"
' ============================================================
' MacroKeyBindings
'
' Applies macro keyboard shortcuts at every Word startup, so they survive
' template rebuilds (including a from-scratch rebuild) and travel to any
' machine. Modeled on RegisterKeyBindings in the ParentheticalAutocomplete
' module.
'
' KeyCode modifiers: Control = 512, Shift = 256.
' Letter/keys used here: V = 86, T = 84, C = 67, Spacebar = 32.
'
' Note: Ctrl+Shift+V, Ctrl+Shift+C, and Ctrl+Shift+Space are built-in Word
' shortcuts; these assignments deliberately override them in this template.
' ============================================================
Option Explicit

' Runs automatically when Word starts and loads this global template.
Sub AutoExec()
    ApplyMacroKeyBindings
    ' Reapply the autocomplete shortcuts too, quietly (no scan, no popup).
    RegisterParenKeyBindings
End Sub

' Can also be run on demand via Alt+F8 to reapply the shortcuts immediately.
Public Sub ApplyMacroKeyBindings()
    On Error Resume Next
    CustomizationContext = ThisDocument

    ' Ctrl+Shift+V -> PasteLegalQuotation   (quote legal citation)
    KeyBindings.Add KeyCode:=BuildKeyCode(86, 512, 256), _
                    KeyCategory:=1, Command:="PasteLegalQuotation"

    ' Ctrl+Shift+T -> ApplyTitleCase
    KeyBindings.Add KeyCode:=BuildKeyCode(84, 512, 256), _
                    KeyCategory:=1, Command:="ApplyTitleCase"

    ' Ctrl+Shift+C -> ConvertToShortCitations   (short cite)
    KeyBindings.Add KeyCode:=BuildKeyCode(67, 512, 256), _
                    KeyCategory:=1, Command:="ConvertToShortCitations"

    ' Ctrl+Shift+Spacebar -> ReplaceParagraphMarksWithSpaces
    KeyBindings.Add KeyCode:=BuildKeyCode(32, 512, 256), _
                    KeyCategory:=1, Command:="ReplaceParagraphMarksWithSpaces"

    ' Don't mark the template dirty; bindings reapply on the next launch anyway.
    ThisDocument.Saved = True
End Sub
