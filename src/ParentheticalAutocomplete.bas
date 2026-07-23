Attribute VB_Name = "ParentheticalAutocomplete"
' NOTE: do NOT add Option Private Module here. Several of this module's macros
' (OnOpenParen, AcceptOrTab, etc.) are assigned to keys via KeyBindings.Add, and
' a macro in an Option Private Module module cannot be resolved as a key-binding
' target -- RegisterParenKeyBindings would fail with runtime error 5346.
'=============================================================================
' ParentheticalAutocomplete.bas
' Word VBA Macro -- Legal Citation Autocomplete
' v3: single-suggestion label, counter, Ctrl+Up/Down cycling, ";" + space
'
' WHAT CHANGED FROM v2:
'   1. ";" now triggers ONLY after typing "; " (semicolon + space). The
'      bound handler types both characters before checking, so the cite
'      template is inserted at the position after the space, matching
'      Bluebook style for string cites.
'
'   2. The popup shows ONE suggestion at a time as a label, with a small
'      "(2 of 5)" counter beside it. No dropdown list visible.
'
'   3. Ctrl+Up / Ctrl+Down cycle through alternatives without leaving
'      the keyboard or losing your typing context. Bare Up/Down still
'      go to the document (so normal editing isn't disrupted).
'
'   4. Narrowing-as-you-type is the same mechanism as v2, but with
'      additional diagnostics: if the WindowSelectionChange hook is
'      somehow not wired, you'll see the popup but it won't update --
'      run DiagnoseHook to confirm.
'
' BEHAVIORAL CONTRACT:
'   - Trigger 1: type "(" right after sentence-end punctuation + space
'     (or ." (). Popup shows top match for any prior cite in the doc.
'   - Trigger 2: type "; " inside an existing open paren. Popup shows
'     top match, excluding cites already in this paren.
'   - Keep typing to narrow; Ctrl+Down/Up to cycle among alternatives.
'   - Tab / Enter / Right-Arrow accepts the displayed suggestion.
'   - Escape, or typing a prefix that matches nothing, dismisses.
'   - The popup NEVER inserts a closing ")".
'   - A parenthetical counts as a citation template if it contains a
'     section/paragraph symbol (§ / ¶) OR matches the page-reference
'     pattern "... at p" (e.g. "(Mot. at p 6:5)" or "(Opp. at pp. 4-5)").
'     Either way, the page/line numbers after the marker are stripped so
'     the reusable template ends right after the symbol or the "at p".
'
' INSTALL:
'   1. Alt+F11; replace the ParentheticalAutocomplete module with this file.
'   2. Open frmSuggest, REMOVE the existing listbox if any, and add the
'      controls described in the companion file frmSuggest_code_v3.txt:
'        - lblSuggestion : Label across most of the form's width
'        - lblCounter    : small Label at the far right
'      Then paste the form code from frmSuggest_code_v3.txt into the
'      form's code window.
'   3. Make sure clsAppEvents.App_WindowSelectionChange calls
'      ParentheticalAutocomplete.OnSelectionChanged (unchanged).
'   4. Save (Ctrl+S).
'   5. Alt+F8 -> RegisterKeyBindings (once).
'
' BACKOUT:
'   Alt+F8 -> UnregisterKeyBindings.
'=============================================================================

Option Explicit

'--- Win32 declarations for the non-activating-form fix ----------------------
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLongA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
         ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
        (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" _
        (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetFocusAPI Lib "user32" Alias "GetFocus" () As LongPtr
#Else
    Private Declare Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongA Lib "user32" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
         ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal wFlags As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" _
        (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" _
        (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long
    Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" _
        (ByVal hWnd As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetFocusAPI Lib "user32" Alias "GetFocus" () As Long
#End If

' GetDeviceCaps indices
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

' Cached form HWND so we don't have to call FindWindowA repeatedly
' (and risk picking up some other ThunderDFrame if a second form appears).
' Cached document HWND (Word's top-level "OpusApp" window) lets us push
' focus back to it via Win32 SetForegroundWindow after Show steals it.
#If VBA7 Then
    Private m_FormHwnd As LongPtr
    Private m_DocHwnd As LongPtr
#Else
    Private m_FormHwnd As Long
    Private m_DocHwnd As Long
#End If

Private Const GWL_EXSTYLE       As Long = -20
Private Const GWL_STYLE         As Long = -16
Private Const WS_EX_NOACTIVATE  As Long = &H8000000
Private Const WS_EX_TOOLWINDOW  As Long = &H80&
Private Const WS_CAPTION        As Long = &HC00000
Private Const WS_SYSMENU        As Long = &H80000
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_MINIMIZE       As Long = &H20000000
Private Const WS_MAXIMIZE       As Long = &H1000000
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_SHOWWINDOW    As Long = &H40
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOZORDER      As Long = &H4

'--- Configuration -----------------------------------------------------------
Private Const FORM_OFFSET_X     As Long = 4     ' pixels right of caret
Private Const FORM_OFFSET_Y     As Long = 18    ' pixels below caret baseline
Private Const MAX_TYPED         As Long = 80    ' safety cap on filter length

'--- Trigger modes -----------------------------------------------------------
Public Const MODE_NONE  As Integer = 0
Public Const MODE_OPEN  As Integer = 1
Public Const MODE_SEMI  As Integer = 2

'--- Public state ------------------------------------------------------------
Public m_Mode           As Integer
Public m_TypedSoFar     As String
Public m_AnchorPos      As Long
Public m_LastSeenPos    As Long

Public m_ExcludeCites() As String
Public m_ExcludeCount   As Long

Public m_AllCites()     As String
Public m_AllFreqs()     As Long
Public m_CiteCount      As Long

' v3: current filtered match list and cycle index
Public m_Matches()      As String
Public m_MatchCount     As Long
Public m_MatchIndex     As Long

' Guard: when True, OnSelectionChanged should NOT process events. Set during
' the window-of-vulnerability between BeginSession and the form being
' fully visible, because the act of showing the form, applying Win32 style
' changes, and repositioning can fire spurious selection-change events
' that would otherwise be mistaken for "cursor jumped" and dismiss the
' session before the user has typed anything.
Public m_InSetup        As Boolean

' Polling: WindowSelectionChange doesn't fire for keyboard typing on some
' Word builds. We work around it by checking Selection.Start every
' POLL_INTERVAL_MS while a session is active. m_PollPending tracks whether
' a timer callback is already scheduled, so we don't stack callbacks.
Private Const POLL_INTERVAL_MS As Long = 75
Private m_PollPending As Boolean
Private m_NextPollAt As Date
' Count of OnTime callbacks queued but not yet fired. Word's OnTime cannot
' be cancelled, so PollTick uses this to drain stale ticks: only the most
' recently armed tick may process and re-arm (see PollTick).
Private m_TicksQueued As Long

'--- Symbol helpers ----------------------------------------------------------
' paragraph sign = Chr(182), section sign = Chr(167)
Public Function PARA() As String
    PARA = Chr(182)
End Function
Public Function SECT() As String
    SECT = Chr(167)
End Function

'=============================================================================
' KEY BINDING REGISTRATION
'=============================================================================
Public Sub RegisterParenKeyBindings()
    ' Bindings + engine init only. No document scan, no popup, so this is safe
    ' to call at startup. OnOpenParen / OnSemicolon scan on demand when typed.
    CustomizationContext = ThisDocument

    ' "(" -> OnOpenParen
    KeyBindings.Add KeyCode:=BuildKeyCode(57, 256), _
                    KeyCategory:=1, Command:="OnOpenParen"

    ' ";" -> OnSemicolon
    KeyBindings.Add KeyCode:=BuildKeyCode(186), _
                    KeyCategory:=1, Command:="OnSemicolon"

    ' Ctrl+Down -> CycleNext  (Down=40, Ctrl modifier=512)
    KeyBindings.Add KeyCode:=BuildKeyCode(40, 512), _
                    KeyCategory:=1, Command:="CycleNext"

    ' Ctrl+Up -> CyclePrev  (Up=38)
    KeyBindings.Add KeyCode:=BuildKeyCode(38, 512), _
                    KeyCategory:=1, Command:="CyclePrev"

    ' Tab -> AcceptOrTab  (Tab=9)
    KeyBindings.Add KeyCode:=BuildKeyCode(9), _
                    KeyCategory:=1, Command:="AcceptOrTab"

    ' Enter -> AcceptOrEnter  (Return=13)
    KeyBindings.Add KeyCode:=BuildKeyCode(13), _
                    KeyCategory:=1, Command:="AcceptOrEnter"

    ' Right Arrow -> AcceptOrRight  (Right=39)
    KeyBindings.Add KeyCode:=BuildKeyCode(39), _
                    KeyCategory:=1, Command:="AcceptOrRight"

    ' Escape -> DismissOrPass  (Esc=27)
    KeyBindings.Add KeyCode:=BuildKeyCode(27), _
                    KeyCategory:=1, Command:="DismissOrPass"

    If gAppEvents Is Nothing Then InitializeAppEvents

    ' Don't mark the template dirty; bindings reapply on the next launch anyway.
    ThisDocument.Saved = True
End Sub

Public Sub RegisterKeyBindings()
    ' Manual activation (Alt+F8): apply bindings, scan the active document, and
    ' report how many citation templates were found.
    RegisterParenKeyBindings

    ScanDocument
    MsgBox "Parenthetical Autocomplete v3 is active." & vbCr & _
           m_CiteCount & " unique citation template(s) found.", _
           vbInformation, "Autocomplete Ready"
End Sub

Public Sub UnregisterKeyBindings()
    CustomizationContext = ThisDocument

    ClearBinding BuildKeyCode(57, 256)
    ClearBinding BuildKeyCode(186)
    ClearBinding BuildKeyCode(40, 512)
    ClearBinding BuildKeyCode(38, 512)
    ClearBinding BuildKeyCode(9)
    ClearBinding BuildKeyCode(13)
    ClearBinding BuildKeyCode(39)
    ClearBinding BuildKeyCode(27)

    DismissSuggest
    MsgBox "Default behavior restored.", vbInformation, "Autocomplete Disabled"
End Sub

Private Sub ClearBinding(ByVal kc As Long)
    Dim kb As KeyBinding
    Set kb = FindKeyBinding(kc)
    If Not kb Is Nothing Then kb.Clear
End Sub

Private Function FindKeyBinding(ByVal kc As Long) As KeyBinding
    Dim k As KeyBinding
    For Each k In KeyBindings
        If k.KeyCode = kc Then
            Set FindKeyBinding = k
            Exit Function
        End If
    Next k
    Set FindKeyBinding = Nothing
End Function

' Diagnostic: confirms the WindowSelectionChange hook is alive so the
' popup actually narrows when you type. Run via Alt+F8.
Public Sub DiagnoseHook()
    If gAppEvents Is Nothing Then
        MsgBox "gAppEvents is Nothing -- run InitializeAppEvents.", vbExclamation
    Else
        MsgBox "gAppEvents is initialized. If narrowing still doesn't work, " & _
               "verify clsAppEvents.App_WindowSelectionChange calls " & _
               "ParentheticalAutocomplete.OnSelectionChanged.", vbInformation
    End If
End Sub

'=============================================================================
' BOUND KEY HANDLERS
'=============================================================================
Public Sub OnOpenParen()
    Selection.TypeText "("

    ' Main body only: the session records Selection offsets but reads them
    ' back through ActiveDocument.Range (the main text story), so a session
    ' begun in a footnote/header/text box narrowed against unrelated BODY
    ' text and an accept deleted the wrong characters.
    If Selection.StoryType <> wdMainTextStory Then Exit Sub

    If IsSentenceEndSpaceBehindParen() Then
        ScanDocument
        If m_CiteCount > 0 Then
            BeginSession MODE_OPEN
        End If
    End If
End Sub


Public Sub OnSemicolon()
    ' Type the bare ";" first and add the trailing space ONLY in trigger
    ' context (inside an open parenthetical). The old version always typed
    ' "; ", which made it impossible to type a bare semicolon anywhere --
    ' end of line, before a close paren, "id.;" -- without deleting the
    ' unwanted space. Anchor still sits after the space in trigger context,
    ' so accepting a suggestion inserts cleanly into "...; <cite> ".
    Selection.TypeText ";"

    ' Main body only -- same story-offset mismatch as OnOpenParen.
    If Selection.StoryType <> wdMainTextStory Then Exit Sub

    If IsInsideOpenParen() Then
        Selection.TypeText " "
        ScanDocument
        If m_CiteCount > 0 Then
            CollectExistingCites
            BeginSession MODE_SEMI
        End If
    End If
End Sub

Public Sub CycleNext()
    If m_Mode = MODE_NONE Then
        ' Pass through: emulate default Ctrl+Down (next paragraph)
        On Error Resume Next
        Selection.MoveDown Unit:=wdParagraph, count:=1
        On Error GoTo 0
        Exit Sub
    End If
    If m_MatchCount = 0 Then Exit Sub
    m_MatchIndex = (m_MatchIndex + 1) Mod m_MatchCount
    RefreshDisplay
End Sub

Public Sub CyclePrev()
    If m_Mode = MODE_NONE Then
        On Error Resume Next
        Selection.MoveUp Unit:=wdParagraph, count:=1
        On Error GoTo 0
        Exit Sub
    End If
    If m_MatchCount = 0 Then Exit Sub
    m_MatchIndex = m_MatchIndex - 1
    If m_MatchIndex < 0 Then m_MatchIndex = m_MatchCount - 1
    RefreshDisplay
End Sub

' --- Conditional handlers: accept if popup is showing, otherwise emulate
'     the key's default behavior so normal editing isn't disrupted. ---

Public Sub AcceptOrTab()
    If m_Mode <> MODE_NONE And m_MatchCount > 0 Then
        AcceptCurrentMatch
    Else
        ' Default Tab behavior: insert a tab character.
        ' (In tables this normally moves to the next cell; we accept the
        ' compromise of inserting a literal tab since perfect emulation
        ' is hard and the popup is the dominant context for this key.)
        Selection.TypeText vbTab
    End If
End Sub

Public Sub AcceptOrEnter()
    If m_Mode <> MODE_NONE And m_MatchCount > 0 Then
        AcceptCurrentMatch
    Else
        ' Delegate to the WrapCitations Enter handler, NOT a bare
        ' TypeParagraph. Only one macro can own the Return key binding, and
        ' this module registers after Module4's RegisterWrapKeyBindings at
        ' startup, so this binding wins -- a bare TypeParagraph here silently
        ' killed citation-wrap-on-Enter for the whole session. CheckAndWrapEnter
        ' runs the wrap check and then types the paragraph itself.
        WrapCitations.CheckAndWrapEnter
    End If
End Sub

Public Sub AcceptOrRight()
    If m_Mode <> MODE_NONE And m_MatchCount > 0 Then
        AcceptCurrentMatch
    Else
        ' Default Right Arrow: move cursor right one character.
        Selection.MoveRight Unit:=wdCharacter, count:=1
    End If
End Sub

Public Sub DismissOrPass()
    If m_Mode <> MODE_NONE Then
        DismissSuggest
    End If
    ' If no session active, do nothing -- Esc has no standard typing
    ' behavior to emulate. (Word uses it for various UI cancellations,
    ' but those aren't accessible via VBA's TypeText.)
End Sub

Private Sub BeginSession(ByVal mode As Integer)
    m_InSetup = True            ' suppress OnSelectionChanged during show
    m_Mode = mode
    m_TypedSoFar = ""
    m_AnchorPos = Selection.start
    m_LastSeenPos = m_AnchorPos
    RebuildMatches

    If m_MatchCount = 0 Then
        m_Mode = MODE_NONE
        m_InSetup = False
        Exit Sub
    End If

    m_MatchIndex = 0
    ShowSuggestForm
    RefreshDisplay
    m_InSetup = False           ' setup done; events can now drive narrowing

    ' WindowSelectionChange doesn't fire reliably for keyboard typing on
    ' some Word builds, so we drive narrowing by polling the selection
    ' position on a short timer instead. The timer self-rearms while the
    ' session is active and stops itself on DismissSuggest.
    StartPollTimer
End Sub

'=============================================================================
' SELECTION-CHANGE HOOK  (called from clsAppEvents.App_WindowSelectionChange)
' Still wired as a backup — it fires reliably for MOUSE clicks even when
' it doesn't fire for keyboard typing, so it's our dismiss-on-click path.
'=============================================================================
Public Sub OnSelectionChanged()
    ProcessSelectionUpdate
End Sub

' Shared narrowing logic, called both from the WSC hook (when it fires)
' and from the poll timer (which is what actually drives keyboard
' narrowing on builds where WSC doesn't fire for typing).
Private Sub ProcessSelectionUpdate()
    If m_InSetup Then Exit Sub
    If m_Mode = MODE_NONE Then Exit Sub

    Dim curPos As Long
    curPos = Selection.start

    ' No movement since last check — nothing to do (timer no-op case).
    If curPos = m_LastSeenPos Then Exit Sub

    If curPos < m_AnchorPos Then
        DismissSuggest
        Exit Sub
    End If

    ' Allow forward motion of any size up to MAX_TYPED; a big *backward*
    ' jump was already caught above, and a big forward jump means the user
    ' clicked deeper into the doc — dismiss.
    If curPos - m_LastSeenPos > 16 Then
        DismissSuggest
        Exit Sub
    End If

    Dim typed As String
    typed = ReadDocRange(m_AnchorPos, curPos)

    If Len(typed) > MAX_TYPED Then
        DismissSuggest
        Exit Sub
    End If

    ' Detect actual change in the typed prefix (vs. cursor moving via
    ' selection-collapse or some other no-op). If the prefix didn't
    ' change, no need to rebuild matches — just sync the position.
    Dim prefixChanged As Boolean
    prefixChanged = (typed <> m_TypedSoFar)

    m_TypedSoFar = typed
    m_LastSeenPos = curPos

    If Not prefixChanged Then Exit Sub

    ' Remember which cite we were on before rebuilding, so we can keep
    ' the user pointed at it if it's still in the (possibly enlarged or
    ' shrunken) match list. This matters most for backspace: typing "Sm"
    ' then backspacing to "S" should keep "Smith Decl." highlighted if it
    ' was selected, not jump back to index 0.
    Dim prevSelection As String
    If m_MatchCount > 0 And m_MatchIndex >= 0 And m_MatchIndex < m_MatchCount Then
        prevSelection = m_Matches(m_MatchIndex)
    End If

    RebuildMatches

    If m_MatchCount = 0 Then
        ' Hide the form but DON'T dismiss the session. The user has typed
        ' a prefix that matches nothing -- but they may be about to
        ' backspace, in which case we want the popup to come back. Only
        ' DismissSuggest tears down state; here we just visually hide the
        ' form and let the poll timer keep watching. The session ends
        ' only when the cursor leaves the typing region (backward past
        ' anchor, large forward jump, MAX_TYPED exceeded, or Esc).
        HideFormPreserveSession
    Else
        ' Form may have been hidden by a previous no-match state; show
        ' it again. Re-running ShowSuggestForm would re-apply chrome
        ' stripping and focus return, both of which are idempotent but
        ' unnecessary; just unhide if we already have the HWND, otherwise
        ' do the full show. Either path re-returns focus to the document
        ' afterward, since unhiding can re-activate the form on some
        ' Office builds.
        '
        ' Note: UserForm.Visible is read-only -- assigning to it is a
        ' compile error. The MSForms way to unhide a loaded form is
        ' .Show, which is a no-op for visibility if already visible and
        ' just makes it visible again if it was Hidden.
        If m_FormHwnd = 0 Then
            ShowSuggestForm
        Else
            On Error Resume Next
            frmSuggest.Show vbModeless
            ReturnFocusToDocument
            On Error GoTo 0
        End If

        ' Try to keep the user on the same cite they had highlighted.
        ' If it's gone from the new match set (because narrowing dropped
        ' it, not because backspace brought items back), reset to 0.
        m_MatchIndex = 0
        If Len(prevSelection) > 0 Then
            Dim i As Long
            For i = 0 To m_MatchCount - 1
                If m_Matches(i) = prevSelection Then
                    m_MatchIndex = i
                    Exit For
                End If
            Next i
        End If
        RefreshDisplay
        PositionFormAtCaret
    End If
End Sub

' Hide the popup visually but keep the session, timer, anchor, and HWND
' caches alive. Used when the typed prefix matches nothing but the user
' might still backspace back into a matching range. Distinct from
' DismissSuggest, which is the actual session teardown.
Private Sub HideFormPreserveSession()
    On Error Resume Next
    If frmSuggest.Visible Then frmSuggest.Hide
    On Error GoTo 0
End Sub

Private Function ReadDocRange(ByVal startPos As Long, ByVal endPos As Long) As String
    On Error GoTo ErrExit
    If endPos <= startPos Then
        ReadDocRange = ""
        Exit Function
    End If
    Dim r As Range
    Set r = ActiveDocument.Range(startPos, endPos)
    ReadDocRange = r.text
    Exit Function
ErrExit:
    ReadDocRange = ""
End Function

'=============================================================================
' DOCUMENT SCANNER  (stack-based pairing; nesting- and paragraph-aware)
'=============================================================================
Private Sub ScanDocument()
    Dim sText    As String
    Dim k        As Long
    Dim ch       As String
    Dim inner    As String

    ' Stack of "(" positions. Pushing on "(" and popping the MOST RECENT
    ' "(" on ")" means the INNERMOST pair is the candidate examined -- for
    ' "(see (Mot. at p. 5))" that's "(Mot. at p. 5)", not the outer span.
    ' A stray unmatched "(" is never popped, so it can no longer capture
    ' everything up to some ")" paragraphs later. The enclosing pair of a
    ' nested candidate still pops on its own ")", but its inner text keeps
    ' the consumed "(...)" characters, so the "(" rejection below drops it.
    Dim openPos(0 To 199) As Long
    Dim openTop  As Long
    openTop = 0

    Dim tmpCites(0 To 4999) As String
    Dim tmpFreqs(0 To 4999) As Long
    Dim tmpCount As Long
    tmpCount = 0

    sText = ActiveDocument.content.text

    For k = 1 To Len(sText)
        ch = Mid(sText, k, 1)
        If ch = "(" Then
            If openTop < 199 Then
                openTop = openTop + 1
                openPos(openTop) = k
            End If
        ElseIf ch = ")" Then
            If openTop > 0 Then
                inner = Mid(sText, openPos(openTop) + 1, k - openPos(openTop) - 1)
                openTop = openTop - 1

                ' Reject candidates that cross a paragraph mark or still
                ' contain a "(": the stored template is replayed verbatim
                ' via sel.TypeText on accept, so it must be a clean,
                ' single-paragraph segment with no leftover nesting.
                If InStr(inner, vbCr) = 0 And InStr(inner, "(") = 0 Then

                    If IsCiteSegment(inner) Then
                        Dim parts() As String
                        parts = Split(inner, ";")

                        Dim p As Long
                        For p = 0 To UBound(parts)
                            Dim rawCite As String
                            rawCite = Trim(parts(p))

                            If IsCiteSegment(rawCite) And Len(rawCite) > 0 Then
                                rawCite = NormalizeCiteSegment(rawCite)

                                If Len(Trim(rawCite)) > 0 Then
                                    Dim found As Boolean
                                    found = False
                                    Dim j As Long
                                    For j = 0 To tmpCount - 1
                                        If tmpCites(j) = rawCite Then
                                            tmpFreqs(j) = tmpFreqs(j) + 1
                                            found = True
                                            Exit For
                                        End If
                                    Next j
                                    If Not found And tmpCount < 5000 Then
                                        tmpCites(tmpCount) = rawCite
                                        tmpFreqs(tmpCount) = 1
                                        tmpCount = tmpCount + 1
                                    End If
                                End If
                            End If
                        Next p
                    End If

                End If
            End If
        End If
    Next k

    Dim si As Long, sj As Long, stC As String, stF As Long
    For si = 1 To tmpCount - 1
        stC = tmpCites(si)
        stF = tmpFreqs(si)
        sj = si - 1
        Do While sj >= 0
            If tmpFreqs(sj) >= stF Then Exit Do
            tmpCites(sj + 1) = tmpCites(sj)
            tmpFreqs(sj + 1) = tmpFreqs(sj)
            sj = sj - 1
        Loop
        tmpCites(sj + 1) = stC
        tmpFreqs(sj + 1) = stF
    Next si

    m_CiteCount = tmpCount
    If tmpCount > 0 Then
        ReDim m_AllCites(0 To tmpCount - 1)
        ReDim m_AllFreqs(0 To tmpCount - 1)
        Dim i As Long
        For i = 0 To tmpCount - 1
            m_AllCites(i) = tmpCites(i)
            m_AllFreqs(i) = tmpFreqs(i)
        Next i
    Else
        ReDim m_AllCites(0 To 0)
        ReDim m_AllFreqs(0 To 0)
    End If
End Sub

'=============================================================================
' COLLECT CITES ALREADY IN THE CURRENT PARENTHETICAL  (semicolon mode)
'=============================================================================
Private Sub CollectExistingCites()
    On Error GoTo ErrExit
    m_ExcludeCount = 0
    ReDim m_ExcludeCites(0 To 99)

    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    rng.MoveStart Unit:=wdCharacter, count:=-2000

    Dim buf As String
    buf = rng.text

    Dim k As Long, depth As Long
    depth = 0
    Dim lastOpen As Long
    lastOpen = 0
    For k = Len(buf) To 1 Step -1
        Dim c As String
        c = Mid(buf, k, 1)
        If c = ")" Then depth = depth + 1
        If c = "(" Then
            If depth = 0 Then
                lastOpen = k
                Exit For
            Else
                depth = depth - 1
            End If
        End If
    Next k

    If lastOpen = 0 Then Exit Sub

    Dim inner As String
    inner = Mid(buf, lastOpen + 1)
    Dim parts() As String
    parts = Split(inner, ";")

    Dim p As Long
    For p = 0 To UBound(parts)
        Dim seg As String
        seg = Trim(parts(p))
        If IsCiteSegment(seg) Then
            seg = NormalizeCiteSegment(seg)
            If Len(Trim(seg)) > 0 And m_ExcludeCount < 100 Then
                m_ExcludeCites(m_ExcludeCount) = seg
                m_ExcludeCount = m_ExcludeCount + 1
            End If
        End If
    Next p
    Exit Sub
ErrExit:
End Sub

Public Function IsExcluded(ByVal s As String) As Boolean
    Dim i As Long
    For i = 0 To m_ExcludeCount - 1
        If m_ExcludeCites(i) = s Then
            IsExcluded = True
            Exit Function
        End If
    Next i
    IsExcluded = False
End Function

'=============================================================================
' MATCH LIST  (v3 -- now stores filtered matches in m_Matches array)
'=============================================================================
Private Sub RebuildMatches()
    ReDim m_Matches(0 To IIf(m_CiteCount = 0, 0, m_CiteCount - 1))
    m_MatchCount = 0

    Dim prefix As String
    prefix = LCase(m_TypedSoFar)

    Dim i As Long
    For i = 0 To m_CiteCount - 1
        If m_Mode = MODE_SEMI Then
            If IsExcluded(m_AllCites(i)) Then GoTo NextCite
        End If
        If prefix = "" Or InStr(1, LCase(m_AllCites(i)), prefix) = 1 Then
            m_Matches(m_MatchCount) = m_AllCites(i)
            m_MatchCount = m_MatchCount + 1
        End If
NextCite:
    Next i
End Sub

' Update the form's label + counter to reflect m_Matches(m_MatchIndex).
Private Sub RefreshDisplay()
    On Error Resume Next
    If m_MatchCount = 0 Then
        frmSuggest.lblSuggestion.Caption = ""
        frmSuggest.lblCounter.Caption = ""
        Exit Sub
    End If
    frmSuggest.lblSuggestion.Caption = m_Matches(m_MatchIndex)
    If m_MatchCount = 1 Then
        frmSuggest.lblCounter.Caption = ""
    Else
        frmSuggest.lblCounter.Caption = "(" & (m_MatchIndex + 1) & " of " & m_MatchCount & ")"
    End If
    On Error GoTo 0
End Sub

'=============================================================================
' POPUP FORM SHOW / HIDE / POSITION
'=============================================================================
Private Sub ShowSuggestForm()
    ' Order matters here. WS_EX_NOACTIVATE has to be applied to the form's
    ' HWND, which doesn't exist until Show creates it. The first Show() will
    ' steal focus regardless of the ex-style; we accept that and explicitly
    ' return focus to the document afterward. We also park the form off-
    ' screen before Show so the user never sees a frame at the wrong spot,
    ' then SetWindowPos it to the caret with SWP_NOACTIVATE.
    On Error Resume Next

    ' 0. Capture Word's top-level window NOW, before Show steals foreground.
    '    Word's main window has class "OpusApp". We use FindWindowA because
    '    Application.hWnd (the Word object-model equivalent) isn't available
    '    on all Word builds. Caching the HWND lets us push focus back via
    '    Win32 SetForegroundWindow, which actually moves keyboard focus —
    '    unlike Application.ActiveWindow.SetFocus (a Word object-model call
    '    that may not move Win32 focus until the next idle tick).
    m_DocHwnd = FindWindowA("OpusApp", vbNullString)

    ' 1. Park off-screen so the initial Show doesn't flash at (0,0).
    frmSuggest.Top = -10000
    frmSuggest.Left = -10000

    ' 2. Show. This is what creates the HWND. Focus gets stolen briefly.
    frmSuggest.Show vbModeless

    ' 3. Capture the HWND now, while we know which ThunderDFrame is ours
    '    (the most recently created one — search top-down from the desktop).
    m_FormHwnd = FindFormHwnd()

    ' 4. Strip title bar / make non-activating. Both modify window styles
    '    on m_FormHwnd.
    MakeFormChromeless
    MakeFormNonActivating

    ' 5. Shrink to content height now that the chrome is gone.
    ShrinkFormToContents

    ' 6. Return focus to the document via Win32. Order: foreground first
    '    (puts the document window in front and gives it activation), then
    '    SetFocus to make sure the keyboard input target is the document.
    '    Critically, we now set m_InSetup back to False AFTER this so the
    '    selection-change that fires from focus moving doesn't dismiss us.
    ReturnFocusToDocument

    ' 7. NOW position. Using SetWindowPos with SWP_NOACTIVATE avoids the
    '    re-activation that setting .Left/.Top can cause on some Office
    '    versions.
    PositionFormAtCaret

    On Error GoTo 0
End Sub

' Walk the top-level windows looking for the ThunderDFrame whose HWND we
' just created. FindWindowA with a NULL caption returns the first match in
' Z-order, which is *usually* the newest top-level window — but if Word
' has any other ThunderDFrame in the process (which can happen with other
' add-ins), we'd get the wrong one. Kept as a single helper so we have
' one place to harden later if needed.
'
' NOTE on the #If: conditional-compilation directives can't split a single
' declaration (e.g. they can't appear inside a function signature). So we
' declare two complete function bodies under the #If/#Else and let the
' compiler pick one. Only the matching branch is compiled.
#If VBA7 Then
Private Function FindFormHwnd() As LongPtr
    Dim hWnd As LongPtr
    hWnd = FindWindowA("ThunderDFrame", vbNullString)
    FindFormHwnd = hWnd
End Function
#Else
Private Function FindFormHwnd() As Long
    Dim hWnd As Long
    hWnd = FindWindowA("ThunderDFrame", vbNullString)
    FindFormHwnd = hWnd
End Function
#End If

Private Sub ReturnFocusToDocument()
    On Error Resume Next
    If m_DocHwnd = 0 Then Exit Sub

    ' Win32 directly. Application.ActiveWindow.SetFocus (the Word object-
    ' model call) doesn't reliably move keyboard focus on Win10/Win11 when
    ' a UserForm just activated — Word treats it as a hint and may defer
    ' until the next message-pump tick, by which point the form still has
    ' focus and the user has to click into the document. SetForegroundWindow
    ' followed by SetFocus moves both the activation AND the focus target.
    SetForegroundWindow m_DocHwnd
    SetFocusAPI m_DocHwnd

    ' Belt-and-suspenders: also ask Word to make this its active window,
    ' which keeps Word's internal state in sync with the Win32 focus state.
    Application.ActiveWindow.SetFocus
    On Error GoTo 0
End Sub

' At design time the form may be much taller than needed (Word forces a
' minimum height because of the title bar). Once the chrome is stripped
' at runtime, we can shrink the form to just enclose its controls so the
' visible popup is a tight strip beside the caret.
Private Sub ShrinkFormToContents()
    On Error Resume Next
    Dim ctL As MSForms.Control
    Dim maxBottom As Double
    Dim maxRight As Double
    maxBottom = 0
    maxRight = 0
    For Each ctL In frmSuggest.Controls
        If ctL.Visible Then
            If ctL.Top + ctL.Height > maxBottom Then maxBottom = ctL.Top + ctL.Height
            If ctL.Left + ctL.Width > maxRight Then maxRight = ctL.Left + ctL.Width
        End If
    Next ctL
    ' Height/Width are read/write on MSForms UserForms; InsideHeight/InsideWidth
    ' are read-only. After MakeFormChromeless strips the title bar/border, the
    ' form's outer Height equals its content area, so setting Height directly
    ' achieves the same goal as setting InsideHeight would have.
    If maxBottom > 0 Then frmSuggest.Height = maxBottom + 2
    If maxRight > 0 Then frmSuggest.Width = maxRight + 4
    On Error GoTo 0
End Sub

Private Sub PositionFormAtCaret()
    If m_FormHwnd = 0 Then Exit Sub

    ' GetPoint returns the screen-pixel bounding box of the given range.
    ' For a collapsed selection (caret) it returns leftPx as the caret's
    ' horizontal position and topPx as the top of the caret line.
    Dim leftPx As Long, topPx As Long, widthPx As Long, heightPx As Long
    On Error Resume Next
    ActiveWindow.GetPoint leftPx, topPx, widthPx, heightPx, Selection.Range

    ' GetPoint returns zeros if the range isn't currently rendered (e.g.
    ' scrolled out of view or measurement not yet ready). If that happens,
    ' try once more after asking Word to scroll the selection into view.
    If leftPx = 0 And topPx = 0 Then
        ActiveWindow.ScrollIntoView Selection.Range, True
        ActiveWindow.GetPoint leftPx, topPx, widthPx, heightPx, Selection.Range
    End If
    On Error GoTo 0

    If leftPx = 0 And topPx = 0 Then Exit Sub

    Dim targetXPx As Long, targetYPx As Long
    targetXPx = leftPx + widthPx + FORM_OFFSET_X
    targetYPx = topPx + heightPx + FORM_OFFSET_Y

    ' Position with SetWindowPos in pixels. This bypasses two pitfalls:
    '   (a) Application.PixelsToPoints doesn't exist in Word (only Excel
    '       and PowerPoint expose it), so the previous code was silently
    '       setting Left/Top to 0.
    '   (b) Setting frmSuggest.Left/.Top can re-activate the form on some
    '       Office builds; SWP_NOACTIVATE prevents that.
    SetWindowPos m_FormHwnd, 0, targetXPx, targetYPx, 0, 0, _
                 SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub

' Convert a pixel measurement to points using the actual screen DPI.
' Word's Application object doesn't expose PixelsToPoints (that's Excel/
' PowerPoint only), so we query GDI directly. Kept here in case anything
' else in the module wants form-coordinate (points) positioning.
Private Function PixelsToPointsManual(ByVal px As Double, _
                                      ByVal horizontal As Boolean) As Double
    On Error Resume Next
    #If VBA7 Then
        Dim hDC As LongPtr
    #Else
        Dim hDC As Long
    #End If
    Dim dpi As Long
    hDC = GetDC(0)                          ' DC for the entire screen
    If horizontal Then
        dpi = GetDeviceCaps(hDC, LOGPIXELSX)
    Else
        dpi = GetDeviceCaps(hDC, LOGPIXELSY)
    End If
    ReleaseDC 0, hDC
    If dpi <= 0 Then dpi = 96               ' fallback to standard DPI
    PixelsToPointsManual = px * 72# / CDbl(dpi)
End Function

Private Sub MakeFormNonActivating()
    On Error Resume Next
    If m_FormHwnd = 0 Then Exit Sub

    Dim ex As Long
    ex = GetWindowLongA(m_FormHwnd, GWL_EXSTYLE)
    ex = ex Or WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW
    SetWindowLongA m_FormHwnd, GWL_EXSTYLE, ex

    ' Important: do NOT include SWP_SHOWWINDOW here. SWP_SHOWWINDOW
    ' implicitly activates on some Windows versions even with NOACTIVATE
    ' on the same call. We just want the style change to take effect.
    SetWindowPos m_FormHwnd, 0, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or _
                 SWP_NOZORDER Or SWP_FRAMECHANGED
    On Error GoTo 0
End Sub

' Strip the title bar / X button / sizing border from the form via Win32.
' fmBorderStyleNone in the designer only removes the thin outer border; the
' title bar is governed by WS_CAPTION/WS_SYSMENU/WS_THICKFRAME which the
' designer doesn't expose. This rips those bits out of the form's window
' style and tells Windows to recompute the non-client area
' (SWP_FRAMECHANGED), so the title bar actually disappears.
Private Sub MakeFormChromeless()
    On Error Resume Next
    If m_FormHwnd = 0 Then Exit Sub

    Dim styl As Long
    styl = GetWindowLongA(m_FormHwnd, GWL_STYLE)
    ' Clear caption, system menu, sizing border, and any min/max state
    styl = styl And (Not WS_CAPTION)
    styl = styl And (Not WS_SYSMENU)
    styl = styl And (Not WS_THICKFRAME)
    styl = styl And (Not WS_MINIMIZE)
    styl = styl And (Not WS_MAXIMIZE)
    SetWindowLongA m_FormHwnd, GWL_STYLE, styl

    ' Force the frame to be recomputed -- without this, the title bar
    ' stays drawn until the next resize or repaint.
    SetWindowPos m_FormHwnd, 0, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
                 SWP_NOACTIVATE Or SWP_FRAMECHANGED
    On Error GoTo 0
End Sub

Public Sub DismissSuggest()
    StopPollTimer
    m_Mode = MODE_NONE
    m_TypedSoFar = ""
    m_ExcludeCount = 0
    m_AnchorPos = 0
    m_LastSeenPos = 0
    m_MatchCount = 0
    m_MatchIndex = 0
    m_FormHwnd = 0          ' invalidate; next ShowSuggestForm will recapture
    m_DocHwnd = 0
    On Error Resume Next
    If frmSuggest.Visible Then frmSuggest.Hide
    On Error GoTo 0
End Sub

'=============================================================================
' POLLING TIMER
' WindowSelectionChange doesn't fire for keyboard typing on some Word builds
' (we confirmed it doesn't fire here even though the typed character does
' appear in the document and the cursor does advance). To drive narrowing,
' we schedule an Application.OnTime callback every POLL_INTERVAL_MS while
' a session is active. The callback re-runs the same narrowing logic that
' OnSelectionChanged uses, then re-arms itself if the session is still
' active. Word's OnTime cannot be cancelled (no Schedule:=False, unlike
' Excel), so StopPollTimer just clears m_PollPending; PollTick drains
' stale queued ticks via m_TicksQueued so only the most recently armed
' tick can process and re-arm (see PollTick).
'=============================================================================
Private Sub StartPollTimer()
    ' Belt-and-suspenders: make sure nothing's already scheduled.
    StopPollTimer
    SchedulePoll
End Sub

Private Sub StopPollTimer()
    On Error Resume Next
    ' Word's Application.OnTime has no cancel mechanism (unlike Excel's
    ' Schedule:=False -- Word's signature is just When, Name, Tolerance), so
    ' an already-queued tick WILL still fire. Clearing m_PollPending makes
    ' that orphan tick a no-op; the m_TicksQueued drain in PollTick keeps a
    ' stale tick from ever re-arming a second chain on top of a new
    ' session's own timer.
    m_PollPending = False
    On Error GoTo 0
End Sub

Private Sub SchedulePoll()
    On Error Resume Next
    ' Word's Application.OnTime is positional: When, Name [, Tolerance].
    ' POLL_INTERVAL_MS milliseconds from now. Date arithmetic in VBA is in
    ' days; 1 day = 86,400,000 ms, so ms / 86400000 gives us the offset.
    m_NextPollAt = Now + CDbl(POLL_INTERVAL_MS) / 86400000#
    Application.OnTime m_NextPollAt, "ParentheticalAutocomplete.PollTick"
    m_PollPending = True
    m_TicksQueued = m_TicksQueued + 1
    On Error GoTo 0
End Sub

' Public because Application.OnTime requires a Public procedure to call.
' Re-runs narrowing, then re-arms the timer if a session is still active.
'
' Queued ticks cannot be cancelled, so a stale tick from a previous
' session can fire after a new session has armed its own tick. We guard
' against that here so an orphan tick never restarts polling on top of a
' fresh session's own timer chain.
Public Sub PollTick()
    ' Drain stale ticks: every SchedulePoll increments m_TicksQueued and
    ' every firing decrements it. If ticks remain queued after this one,
    ' this is an OLD tick and a newer one is still coming -- exit without
    ' processing or re-arming, so exactly one chain survives.
    If m_TicksQueued > 0 Then m_TicksQueued = m_TicksQueued - 1
    If m_TicksQueued > 0 Then Exit Sub

    ' If we're not the timer the current session is waiting on, bail. The
    ' check is "is there a poll pending right now and is it me?": if
    ' m_PollPending is False, either no session is active or StopPollTimer
    ' was called -- either way, this tick is orphaned and must not re-arm.
    If Not m_PollPending Then Exit Sub
    m_PollPending = False

    If m_Mode = MODE_NONE Then Exit Sub

    ' A tick can fire after the document is gone (user closed the last window
    ' before the pending OnTime callback ran). Unguarded, Selection.Start
    ' raised an unhandled runtime-error dialog AND left the session alive with
    ' the popup orphaned -- pressing Enter in the NEXT document then injected
    ' the stale suggestion into it. Any failure here tears the session down.
    On Error GoTo Dead
    If Documents.count = 0 Then GoTo Dead
    ProcessSelectionUpdate
    ' ProcessSelectionUpdate may have dismissed the session (e.g. no matches);
    ' check again before re-arming.
    If m_Mode <> MODE_NONE Then SchedulePoll
    Exit Sub

Dead:
    On Error Resume Next
    DismissSuggest
End Sub

'=============================================================================
' ACCEPT  (called from form Tab/Enter/Right or programmatically)
' v3: takes no argument; reads the currently-displayed match.
'=============================================================================
Public Sub AcceptCurrentMatch()
    ' Sync the typed prefix FIRST. m_TypedSoFar is normally updated by the
    ' poll timer, whose real granularity is ~1 second -- typing "Sm" and
    ' hitting Tab inside that window meant deleteLen was computed from a
    ' stale prefix and the raw "Sm" stayed in front of the inserted
    ' suggestion ("(SmSmith Decl. ..."). ProcessSelectionUpdate brings
    ' m_TypedSoFar current (and may legitimately dismiss the session).
    ProcessSelectionUpdate
    If m_Mode = MODE_NONE Then Exit Sub

    If m_MatchCount = 0 Then
        DismissSuggest
        Exit Sub
    End If

    Dim suggestion As String
    suggestion = m_Matches(m_MatchIndex)

    Dim deleteLen As Long
    deleteLen = Len(m_TypedSoFar)

    Dim sel As Selection
    Set sel = Selection

    If deleteLen > 0 Then
        sel.MoveLeft Unit:=wdCharacter, count:=deleteLen, Extend:=wdExtend
        sel.Delete
    End If

    sel.TypeText suggestion
    DismissSuggest
End Sub

'=============================================================================
' HELPERS
'=============================================================================

Public Function ContainsSymbol(ByVal s As String) As Boolean
    ContainsSymbol = (InStr(s, PARA()) > 0 Or InStr(s, SECT()) > 0)
End Function

' A parenthetical segment qualifies as a citation template if it carries a
' section/paragraph symbol OR matches the "... at p" page-reference pattern
' (e.g. "Mot. at p 6:5", "Opp. at pp. 4-5").
Public Function IsCiteSegment(ByVal s As String) As Boolean
    IsCiteSegment = (ContainsSymbol(s) Or ContainsAtP(s))
End Function

' Turn a raw segment into its reusable template: collapse doubled symbols and
' spaces, then strip the page/line text after whichever marker is present.
Public Function NormalizeCiteSegment(ByVal s As String) As String
    Dim r As String
    r = NormalizeDoubleSymbols(s)
    r = NormalizeSpaces(r)
    If ContainsSymbol(r) Then
        r = StripAfterSymbol(r)
    ElseIf ContainsAtP(r) Then
        r = StripAfterAtP(r)
    End If
    NormalizeCiteSegment = r
End Function

' True when the string contains an "at p" page reference. The marker must be
' a standalone "at" (preceded by start-of-string or a non-letter) so words
' like "treat patient" don't trip it.
Public Function ContainsAtP(ByVal s As String) As Boolean
    ContainsAtP = (FindAtPMarker(s) > 0)
End Function

' Returns the 1-based index of the "p" in the first qualifying "at p" marker,
' or 0 if none. ("at p" is a/t/space/p, so the "p" sits 3 chars past "at".)
Private Function FindAtPMarker(ByVal s As String) As Long
    Dim lc As String
    lc = LCase(s)

    Dim startAt As Long
    startAt = 1
    Dim atPos As Long
    Do
        atPos = InStr(startAt, lc, "at p")
        If atPos = 0 Then Exit Do
        If atPos = 1 Then
            FindAtPMarker = atPos + 3
            Exit Function
        ElseIf Not IsLetterChar(Mid(lc, atPos - 1, 1)) Then
            FindAtPMarker = atPos + 3
            Exit Function
        End If
        startAt = atPos + 1
    Loop
    FindAtPMarker = 0
End Function

' Trim a segment down to "... at p" (or "at pp.", "at p.") and append a space,
' dropping the page/line numbers that follow. Mirrors StripAfterSymbol.
Public Function StripAfterAtP(ByVal s As String) As String
    Dim pPos As Long
    pPos = FindAtPMarker(s)
    If pPos = 0 Then
        StripAfterAtP = s
        Exit Function
    End If

    Dim endPos As Long
    endPos = pPos
    ' Consume an additional "p" so "at pp." is kept intact.
    Do While endPos < Len(s) And LCase(Mid(s, endPos + 1, 1)) = "p"
        endPos = endPos + 1
    Loop
    ' Consume an optional abbreviating period.
    If endPos < Len(s) And Mid(s, endPos + 1, 1) = "." Then
        endPos = endPos + 1
    End If

    StripAfterAtP = Left(s, endPos) & " "
End Function

Private Function IsLetterChar(ByVal c As String) As Boolean
    Dim lc As String
    lc = LCase(c)
    IsLetterChar = (lc >= "a" And lc <= "z")
End Function

Public Function NormalizeDoubleSymbols(ByVal s As String) As String
    Dim r As String
    r = s
    r = Replace(r, PARA() & PARA(), PARA())
    r = Replace(r, SECT() & SECT(), SECT())
    NormalizeDoubleSymbols = r
End Function

Public Function StripAfterSymbol(ByVal s As String) As String
    Dim symPos As Long
    symPos = 0
    Dim k As Long
    For k = 1 To Len(s)
        Dim ch As String
        ch = Mid(s, k, 1)
        If ch = PARA() Or ch = SECT() Then symPos = k
    Next k

    If symPos = 0 Then
        StripAfterSymbol = s
    Else
        StripAfterSymbol = Left(s, symPos) & " "
    End If
End Function

Public Function NormalizeSpaces(ByVal s As String) As String
    Dim r As String
    r = Trim(s)
    Do While InStr(r, "  ") > 0
        r = Replace(r, "  ", " ")
    Loop
    NormalizeSpaces = r
End Function

Public Function IsInsideOpenParen() As Boolean
    On Error GoTo ErrHandler
    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    rng.MoveStart Unit:=wdCharacter, count:=-500

    Dim buf As String
    buf = rng.text

    Dim k As Long, depth As Long
    depth = 0
    For k = Len(buf) To 1 Step -1
        Dim c As String
        c = Mid(buf, k, 1)
        If c = ")" Then depth = depth + 1
        If c = "(" Then
            If depth = 0 Then
                IsInsideOpenParen = True
                Exit Function
            Else
                depth = depth - 1
            End If
        End If
    Next k
    IsInsideOpenParen = False
    Exit Function
ErrHandler:
    IsInsideOpenParen = False
End Function

Public Function IsSentenceEndSpaceBehindParen() As Boolean
    On Error GoTo ErrHandler

    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    rng.MoveEnd Unit:=wdCharacter, count:=-1
    rng.Collapse Direction:=wdCollapseEnd
    rng.MoveStart Unit:=wdCharacter, count:=-4

    Dim tail As String
    tail = rng.text
    Dim n As Long
    n = Len(tail)

    If n = 0 Then
        IsSentenceEndSpaceBehindParen = True
        Exit Function
    End If

    Dim last As String
    last = Right(tail, 1)
    If last = vbCr Or last = Chr(11) Or last = Chr(12) Then
        IsSentenceEndSpaceBehindParen = True
        Exit Function
    End If

    If n < 2 Then
        IsSentenceEndSpaceBehindParen = False
        Exit Function
    End If

    If last <> " " Then
        IsSentenceEndSpaceBehindParen = False
        Exit Function
    End If

    Dim c2 As String
    c2 = Mid(tail, n - 1, 1)

    If IsSentenceEndPunct(c2) Then
        IsSentenceEndSpaceBehindParen = True
        Exit Function
    End If

    If IsClosingQuote(c2) And n >= 3 Then
        Dim c3 As String
        c3 = Mid(tail, n - 2, 1)
        If IsSentenceEndPunct(c3) Then
            IsSentenceEndSpaceBehindParen = True
            Exit Function
        End If
    End If

    IsSentenceEndSpaceBehindParen = False
    Exit Function
ErrHandler:
    IsSentenceEndSpaceBehindParen = False
End Function

Private Function IsSentenceEndPunct(ByVal c As String) As Boolean
    IsSentenceEndPunct = (c = "." Or c = "!" Or c = "?")
End Function

Private Function IsClosingQuote(ByVal c As String) As Boolean
    IsClosingQuote = (c = Chr(34) Or c = ChrW(8221) Or _
                      c = Chr(39) Or c = ChrW(8217))
End Function


