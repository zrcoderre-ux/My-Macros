Attribute VB_Name = "HyphenateSelection"
Sub HyphenateSelection()
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select the words you want to hyphenate first."
        Exit Sub
    End If
    Dim s As String
    Dim sTail As String
    s = Selection.text
    ' Peel off trailing spaces the selection grabbed so they aren't
    ' hyphenated, but keep them: the replacement overwrites the FULL
    ' selection, and dropping them fused the result with the next word.
    sTail = ""
    Do While Right(s, 1) = " "
        sTail = sTail & " "
        s = Left(s, Len(s) - 1)
    Loop
    s = Replace(s, " ", "-")
    Selection.TypeText text:=s & sTail
End Sub
