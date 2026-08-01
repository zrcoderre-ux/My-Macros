Attribute VB_Name = "AutoOpen"
Sub AutoOpen()
    Dim Doc As Document
    Dim docPath As String
    ' Never let this macro block an open (locked files, Outlook
    ' attachments, and ReadOnly batch opens all make Save raise).
    On Error Resume Next
    Set Doc = ActiveDocument
    docPath = LCase(Doc.FullName)

    ' Merged tentative drafts leave the mail merge with AttachedTemplate still
    ' pointing at the Mail Merge Order Template, a macro-bearing .dotm. Word
    ' loads that template's VBA project alongside the document and gives it
    ' precedence over this add-in, so its macros and key bindings shadow ours.
    ' Detach on open. Scoped by name so no other template is ever touched.
    If InStr(LCase(Doc.AttachedTemplate.Name), "mail merge order template") = 1 Then
        Doc.UpdateStylesOnOpen = False
        Doc.AttachedTemplate = ""
        If Doc.ReadOnly = False Then Doc.Save
    End If

    ' Add or remove folder paths as needed
    Dim folders(1) As String
    folders(0) = LCase("C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court\")
    folders(1) = LCase("C:\Users\ZCoderre\Los Angeles Superior Court\Research Attorney and Law Clerk Unit - Zachary Coderre\Workups\")
    
    Dim i As Integer
    For i = 0 To UBound(folders)
        If InStr(1, docPath, folders(i)) = 1 Then
            If Doc.UpdateStylesOnOpen = True Then
                ' Always clear the flag in memory (stops style stomping this
                ' session even when read-only); only persist it when we can.
                Doc.UpdateStylesOnOpen = False
                If Doc.ReadOnly = False Then Doc.Save
            End If
            Exit For
        End If
    Next i
End Sub
