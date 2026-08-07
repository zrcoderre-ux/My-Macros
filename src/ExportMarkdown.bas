Attribute VB_Name = "ExportMarkdown"
'==============================================================================
' ExportMarkdown.bas
'------------------------------------------------------------------------------
' Export the active document to a Markdown (.md) file that carries its tracked
' changes and comments. The file is NAMED AFTER THE DOCUMENT ("Smith Tentative
' .docx" -> "Smith Tentative.md") and written next to it, in the document's own
' folder.
'
' Nothing here is anonymized and nothing is replaced: this reads the document and
' writes one file. It shares the Markdown reader with the re-anonymize macro
' (DeAnonymize.BuildMarkdownFromDocument) because that reader is the only one in
' this template and both macros want the same output:
'   - Heading 1..6 / Title paragraphs   ->  # .. ###### / #
'   - list paragraphs                   ->  "- " or the number label
'   - bold / italic runs                ->  **bold**, *italic*, ***both***
'   - footnotes / endnotes              ->  [^n], texts collected at the end
'   - tracked insertions / deletions    ->  {++inserted++} / {--deleted--}
'   - comments                          ->  [c1], [c2] at the end of the passage
'                                           each marks, with that passage, the
'                                           author, and the comment text collected
'                                           at the end
'
' MACRO YOU RUN:
'   ExportDocumentToMarkdown - write <document name>.md next to the document.
'
' NOTES:
'   - Commenters are NAMED. This is a local export of your own document, not the
'     shareable copy: the "Reviewer 1" labelling exists to keep a real person out
'     of an anonymized file, and there is nothing to keep out of this one. Use
'     ReAnonymizeTentative for anything that leaves the building.
'   - Only the body is exported (see DocToMarkdown): headers and footers carry
'     the caption furniture and have no place in Markdown, and text boxes and
'     other shapes aren't walked. Formatting Markdown can't express -- alignment,
'     the caption's tab leaders, tables -- is dropped, but its text is kept.
'   - The document is never modified. The run forces all markup inline while it
'     reads (Word hides tracked deletions from Range.Text when they sit in
'     balloons, so a deleted passage would silently miss the file) and puts the
'     window's markup view back afterwards; that is a view setting, not an edit.
'     A document that was unmodified when the macro started is left unmodified,
'     so an export never turns into a "save changes?" prompt at closing time.
'   - Re-running REPLACES the previous export of the same document. That is the
'     point of naming the file after the document -- the .md tracks the .docx --
'     and the result dialog says when a file was replaced rather than created.
'   - For a document in a synced OneDrive / SharePoint folder the export lands in
'     the local synced copy of that folder (Word reports the folder as an https
'     URL, which can't be written to). A never-saved document falls back to
'     Documents, since it has no folder of its own yet.
'==============================================================================
Option Explicit

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub ExportDocumentToMarkdown()
    On Error GoTo ErrH

    Dim oDoc As Document
    Set oDoc = ActiveDocument
    If oDoc Is Nothing Then Exit Sub

    ' Reading a document can dirty it (Word sets Saved = False for repagination
    ' and field housekeeping of its own). This macro makes no edit, so a clean
    ' document is put back to clean at the end rather than left asking to be
    ' saved because it was exported.
    Dim wasSaved As Boolean
    On Error Resume Next
    wasSaved = oDoc.Saved
    On Error GoTo ErrH

    Dim outPath As String
    outPath = DeAnonymize.ExportFolderFor(oDoc) & "\" & _
              DeAnonymize.DocumentExportTitle(oDoc, "Document") & ".md"

    ' A .md opened in Word would export onto itself: same folder, same base name,
    ' same extension. Step aside rather than overwrite the file being read.
    If StrComp(outPath, DeAnonymize.SafeFullName(oDoc), vbTextCompare) = 0 Then
        outPath = DeAnonymize.ExportFolderFor(oDoc) & "\" & _
                  DeAnonymize.DocumentExportTitle(oDoc, "Document") & " (export).md"
    End If

    ' Whether this run creates the file or replaces an earlier export of the same
    ' document. Read before writing, obviously, and only to phrase the dialog.
    Dim replaced As Boolean
    On Error Resume Next
    replaced = (Len(Dir$(outPath)) > 0)
    On Error GoTo ErrH

    Application.ScreenUpdating = False
    Application.StatusBar = "Export to Markdown: reading the document ..."

    Dim nComments As Long
    On Error Resume Next
    nComments = oDoc.Comments.count
    On Error GoTo ErrH

    ' False: name the commenters. The reader forces markup inline, captures the
    ' tracked changes, and restores the view; marked is what reached the file.
    Dim md As String, marked As Long
    md = DeAnonymize.BuildMarkdownFromDocument(oDoc, False, marked)
    DeAnonymize.WriteMarkdownFile outPath, md

    Application.StatusBar = False         ' don't strand a progress message
    Application.ScreenUpdating = True

    On Error Resume Next
    If wasSaved Then oDoc.Saved = True
    On Error GoTo 0

    MsgBox IIf(replaced, "Replaced the Markdown export at:", _
                         "Exported this document to Markdown:") & vbCrLf & _
           outPath & vbCrLf & vbCrLf & _
           MarkupSummary(marked, nComments) & vbCrLf & vbCrLf & _
           "The document itself was not changed. Headers and footers are not " & _
           "part of the export.", _
           vbInformation, "Export to Markdown"
    Exit Sub

ErrH:
    Dim eN As Long: eN = Err.Number
    Dim eD As String: eD = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.StatusBar = False          ' don't strand a progress message
    On Error GoTo 0
    MsgBox "Export to Markdown hit an error and stopped:" & vbCrLf & vbCrLf & _
           "Error " & eN & ": " & eD & vbCrLf & vbCrLf & _
           "The document was not changed; at worst the .md file is missing or " & _
           "incomplete.", _
           vbExclamation, "Export to Markdown"
End Sub

' One line for the result dialog about the markup the export carried over, so the
' braces read as notation rather than as something someone typed. Says so plainly
' when the document had neither -- a silent dialog would leave the user wondering
' whether the tracked changes were dropped or were never there.
Private Function MarkupSummary(ByVal marked As Long, _
                               ByVal nComments As Long) As String
    If marked <= 0 And nComments <= 0 Then
        MarkupSummary = "The document carried no tracked changes or comments."
        Exit Function
    End If

    MarkupSummary = "Carried " & marked & " tracked change(s) and " & nComments & _
                    " comment(s) into the file: insertions as {++text++}, " & _
                    "deletions as {--text--}, comments as [c1], [c2] at the end " & _
                    "of the passage each one marks, with that passage, the " & _
                    "author, and the comment text collected at the end."
End Function
