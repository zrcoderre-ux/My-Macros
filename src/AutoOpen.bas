Attribute VB_Name = "AutoOpen"
Sub AutoOpen()
    Dim Doc As Document
    Dim docPath As String
    Set Doc = ActiveDocument
    docPath = LCase(Doc.FullName)
    
    ' Add or remove folder paths as needed
    Dim folders(1) As String
    folders(0) = LCase("C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court\")
    folders(1) = LCase("C:\Users\ZCoderre\Los Angeles Superior Court\Research Attorney and Law Clerk Unit - Zachary Coderre\Workups\")
    
    Dim i As Integer
    For i = 0 To UBound(folders)
        If InStr(1, docPath, folders(i)) = 1 Then
            If Doc.UpdateStylesOnOpen = True Then
                Doc.UpdateStylesOnOpen = False
                Doc.Save
            End If
            Exit For
        End If
    Next i
End Sub
