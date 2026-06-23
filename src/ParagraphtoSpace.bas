Attribute VB_Name = "ParagraphtoSpace"

Sub ReplaceParagraphMarksWithSpaces()
    Dim oRange As Range
    
    ' If no text is selected, paste clipboard and apply macro to pasted text
    If Selection.Type = wdSelectionIP Then
        Dim oStart As Long
        oStart = Selection.Range.start
        
        ' Paste to match destination formatting
        Selection.PasteAndFormat wdFormatSurroundingFormattingWithEmphasis
        
        ' Select the pasted text
        Dim oDoc As Document
        Set oDoc = ActiveDocument
        Set oRange = oDoc.Range(oStart, Selection.Range.End)
        oRange.Select
    Else
        Set oRange = Selection.Range
    End If
    
    ' Replace paragraph marks with spaces
    With oRange.Find
        .text = "^p"
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Re-set the range
    Set oRange = Selection.Range
    
    ' Pass 1: Single digit numbers 1-9 followed by period and space
    With oRange.Find
        .text = "([1-9])(\. )"
        .Replacement.text = "(\1) "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Re-set the range
    Set oRange = Selection.Range
    
    ' Pass 2: Two digit numbers 10-29 followed by period and space
    With oRange.Find
        .text = "([1-2][0-9])(\. )"
        .Replacement.text = "(\1) "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

End Sub
