Attribute VB_Name = "HyphenateSelection"
Sub HyphenateSelection()
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select the words you want to hyphenate first."
        Exit Sub
    End If
    Dim s As String
    s = Selection.text
    ' Trim trailing space if the selection grabbed one
    Do While Right(s, 1) = " "
        s = Left(s, Len(s) - 1)
    Loop
    s = Replace(s, " ", "-")
    Selection.TypeText text:=s
End Sub
