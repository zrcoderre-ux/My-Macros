Attribute VB_Name = "ReplaceBody"
'==============================================================================
' ReplaceBody.bas
'------------------------------------------------------------------------------
' Replace the whole body of a merged document without destroying its header.
'
' THE PROBLEM. Word stores a section's header, footer and page setup in the mark
' that ENDS the section -- the section break for every section but the last, the
' document's final paragraph mark for the last one. A mail-merged order keeps its
' header in that break, so the break is not decoration: it IS the header. Ctrl+A
' selects it along with everything else, and pasting over the selection takes the
' break out and the header with it. Keeping it by hand means selecting the body
' and stopping one character short of the break, which on a ten-page draft is a
' lot of careful scrolling for one keystroke's worth of work.
'
' MACRO YOU RUN:
'   ReplaceBodyKeepHeader - Ctrl+Shift+B: paste the clipboard over the body of one
'                           section, leaving that section's terminating mark --
'                           and so its header, footer and page setup -- exactly
'                           where it was. What Ctrl+A, Ctrl+V does, minus the
'                           collateral damage. Must stay a no-argument Public Sub
'                           to remain key-bindable.
'
' WHICH SECTION IT REPLACES. The one the cursor is in. If the selection spans
' sections, or sits in a section holding no text (clicking past the break of a
' two-section merge lands there), it falls back to the section with the most text
' in it, which in a merged order is the body. Nothing outside that one section is
' touched, so a caption page living in its own section survives the replace too.
'
' NOTES:
'   - Clipboard content ending in a paragraph mark would leave the terminating
'     mark stranded on an empty paragraph of its own: a blank line, and with a
'     next-page break, a blank page. That one trailing mark is removed, and the
'     last pasted paragraph's style and formatting are put back on the merged
'     paragraph so the terminator's own formatting does not win.
'   - Headers are never read or written here. They survive by not being in the
'     way of the paste.
'   - The whole run is one undo record: Ctrl+Z puts the old body back.
'==============================================================================
Option Explicit

' Word's errors for a paste with nothing pastable on the clipboard.
Private Const ERR_NO_CLIPBOARD As Long = 4605
Private Const ERR_CMD_FAILED   As Long = 4198

Public Sub ReplaceBodyKeepHeader()
    Dim oDoc As Document
    Dim oSec As Section
    Dim oBody As Range
    Dim oUndo As UndoRecord
    Dim lStart As Long
    Dim lTerm As Long
    Dim lAfter As Long
    Dim bUndoOpen As Boolean
    Dim lErr As Long
    Dim sErr As String

    If Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "Replace Body"
        Exit Sub
    End If

    Set oDoc = ActiveDocument

    If oDoc.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected, so its body cannot be replaced." & vbCrLf & _
               "Unprotect it and run this again.", vbExclamation, "Replace Body"
        Exit Sub
    End If

    Set oSec = TargetSection(oDoc)
    If oSec Is Nothing Then
        MsgBox "Could not work out which section holds the body.", _
               vbExclamation, "Replace Body"
        Exit Sub
    End If

    ' A section's last character is its terminator, and the terminator carries the
    ' header. So the range we hand to the paste stops one character short of it.
    lStart = oSec.Range.start
    lTerm = oSec.Range.End - 1
    Set oBody = oDoc.Range(lStart, lTerm)

    Set oUndo = Application.UndoRecord
    oUndo.StartCustomRecord "Replace Body (Keep Header)"
    bUndoOpen = True

    On Error GoTo Fail
    Application.ScreenUpdating = False

    oBody.Select
    Selection.Paste                     ' raises if the clipboard holds nothing
    lAfter = Selection.End

    TrimStrandedTerminator oDoc, oSec, lStart, lAfter

    Application.ScreenUpdating = True
    oUndo.EndCustomRecord
    bUndoOpen = False

    Application.StatusBar = "Body replaced. Header, footer and page setup kept."
    Exit Sub

Fail:
    lErr = Err.Number
    sErr = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = True
    ' An undo record left open crashes the next run, so close it whatever happened.
    If bUndoOpen Then oUndo.EndCustomRecord
    On Error GoTo 0

    If lErr = ERR_NO_CLIPBOARD Or lErr = ERR_CMD_FAILED Then
        MsgBox "There was nothing on the clipboard to paste." & vbCrLf & _
               "Copy the new body first, then run this again.", _
               vbExclamation, "Replace Body"
    Else
        MsgBox "Could not replace the body. Nothing was changed, or Ctrl+Z will " & _
               "put it back." & vbCrLf & vbCrLf & _
               "Error " & lErr & ": " & sErr, vbCritical, "Replace Body"
    End If
End Sub

' The section whose body gets replaced. The cursor decides when it can: one
' section, in the body text, with something in it. Otherwise the biggest section
' wins, which is the body in every merged order this runs on.
Private Function TargetSection(ByVal oDoc As Document) As Section
    Dim oPick As Section
    Dim oSec As Section
    Dim oBest As Section
    Dim lBest As Long
    Dim lLen As Long

    On Error Resume Next
    If Selection.StoryType = wdMainTextStory Then
        If Selection.Sections.Count = 1 Then Set oPick = Selection.Sections(1)
    End If
    Err.Clear
    On Error GoTo 0

    If Not oPick Is Nothing Then
        If BodyLength(oPick) > 0 Then
            Set TargetSection = oPick
            Exit Function
        End If
    End If

    lBest = -1
    For Each oSec In oDoc.Sections
        lLen = BodyLength(oSec)
        If lLen > lBest Then
            lBest = lLen
            Set oBest = oSec
        End If
    Next oSec

    Set TargetSection = oBest
End Function

' Characters in a section, not counting the terminating mark.
Private Function BodyLength(ByVal oSec As Section) As Long
    BodyLength = oSec.Range.End - oSec.Range.start - 1
End Function

' The terminator is still in place. If the paste ended with a paragraph mark, the
' terminator is now alone on an empty paragraph -- a blank line the old body did
' not have, and a blank page when the break is next-page. Delete that one mark,
' carrying the last pasted paragraph's style and formatting onto the paragraph the
' terminator ends, so the merge does not hand the terminator's formatting to text
' that just arrived.
'
' Every check here guards the same thing: never delete a mark unless it is one the
' paste itself put down. lStart is where the paste began, and the section's own
' range says where its terminator ended up, so a paste that inserted nothing -- or
' one that shifted the sections around by carrying its own breaks -- fails a check
' and leaves the blank line rather than risk the break behind it.
Private Sub TrimStrandedTerminator(ByVal oDoc As Document, ByVal oSec As Section, _
                                   ByVal lStart As Long, ByVal lAfter As Long)
    Dim oMark As Range
    Dim oLast As Paragraph
    Dim oFmt As ParagraphFormat
    Dim sStyle As String

    ' Nothing was pasted, so the mark in front of lAfter is not ours to touch.
    If lAfter - 1 < lStart Then Exit Sub

    ' The paste has to have ended on a paragraph mark.
    Set oMark = oDoc.Range(lAfter - 1, lAfter)
    If oMark.Text <> vbCr Then Exit Sub

    ' And lAfter has to be the section's terminating mark, stranded on the empty
    ' paragraph that mark now opens.
    If oSec.Range.End - 1 <> lAfter Then Exit Sub

    Set oLast = oMark.Paragraphs(1)
    sStyle = oLast.Style
    Set oFmt = oLast.Format.Duplicate

    oMark.Delete

    On Error Resume Next
    With oDoc.Range(lAfter - 1, lAfter - 1).Paragraphs(1)
        .Style = sStyle
        .Format = oFmt
    End With
    Err.Clear
    On Error GoTo 0
End Sub
