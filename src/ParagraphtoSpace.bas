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

    ' Pass 1: TWO-digit numbers 10-99 followed by period and space.
    ' Two digits MUST run before one digit: the single-digit pattern has no
    ' left boundary, so run first it matched the "2. " inside "12. " and
    ' produced "1(2) " -- corrupting every list number 11-29 before the
    ' two-digit pass could see it. (Same largest-first rule TitleCase uses.)
    With oRange.Find
        .text = "([1-9][0-9])(\. )"
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

    ' Pass 2: single-digit numbers 1-9 followed by period and space.
    ' Runs second; by now every two-digit "N. " is already "(N) ".
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

End Sub
