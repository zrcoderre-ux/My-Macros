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
' Letter/keys used here: V = 86, T = 84, C = 67, H = 72, D = 68, K = 75,
' R = 82, Spacebar = 32.
'
' Note: Ctrl+Shift+V, Ctrl+Shift+C, Ctrl+Shift+D, Ctrl+Shift+K, and
' Ctrl+Shift+Space are built-in Word shortcuts; these assignments deliberately
' override them in this template. (Ctrl+Shift+K is Word's small caps, which
' nothing here uses; the letter is worth more as "Keep with next".)
' ============================================================
Option Explicit

' Runs automatically when Word starts and loads this global template.
Sub AutoExec()
    ApplyMacroKeyBindings
    ' Spacebar / Return citation-wrapping bindings (WrapCitations via Module4).
    RegisterWrapKeyBindings
    ' Reapply the autocomplete shortcuts too, quietly (no scan, no popup).
    RegisterParenKeyBindings
End Sub

' Can also be run on demand via Alt+F8 to reapply the shortcuts immediately.
' Private (called only by AutoExec below) so it stays off the Alt+F8 list.
Private Sub ApplyMacroKeyBindings()
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

    ' Ctrl+Shift+H -> ToggleCitationLinks   (apply/remove citation hyperlinks)
    KeyBindings.Add KeyCode:=BuildKeyCode(72, 512, 256), _
                    KeyCategory:=1, Command:="ToggleCitationLinks"

    ' Ctrl+Shift+D -> DeAnonymizeTentative   (restore real names from the key)
    KeyBindings.Add KeyCode:=BuildKeyCode(68, 512, 256), _
                    KeyCategory:=1, Command:="DeAnonymizeTentative"

    ' Ctrl+Shift+R -> ReAnonymizeTentative   (the other direction: real -> fake,
    ' exported as Markdown). Next to Ctrl+Shift+D on purpose -- the two halves of
    ' the same toggle sit one key apart.
    KeyBindings.Add KeyCode:=BuildKeyCode(82, 512, 256), _
                    KeyCategory:=1, Command:="ReAnonymizeTentative"

    ' Ctrl+Shift+K -> FormatHeadings   (Keep with next + underline roman titles)
    KeyBindings.Add KeyCode:=BuildKeyCode(75, 512, 256), _
                    KeyCategory:=1, Command:="FormatHeadings"

    ' Don't mark the template dirty; bindings reapply on the next launch anyway.
    ThisDocument.Saved = True
End Sub
