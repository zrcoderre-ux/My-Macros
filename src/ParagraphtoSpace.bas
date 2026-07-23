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
    ' Both passes require a preceding space so digits inside a larger number
    ' can't match ("decided in 2023. The" must not become "in 20(23) The").
    ' Paragraph marks are already spaces, so every list number is space-
    ' preceded -- except one at the very start of the range, handled after
    ' the passes. Two digits still run before one digit so the single-digit
    ' pass never sees the "2. " inside " 12. " first (largest-first rule).
    With oRange.Find
        .text = "( )([1-9][0-9])(\. )"
        .Replacement.text = "\1(\2) "
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

    ' Pass 2: single-digit numbers 1-9, same preceding-space boundary.
    ' Runs second; by now every two-digit " N. " is already " (N) ".
    With oRange.Find
        .text = "( )([1-9])(\. )"
        .Replacement.text = "\1(\2) "
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

    ' Edge case: a list number at the very start of the range has no
    ' preceding space, so neither wildcard pass can match it. Convert a
    ' leading "N. " or "NN. " (first digit 1-9) directly.
    Dim sHead As String
    Dim nDigits As Long
    sHead = oRange.text
    nDigits = 0
    If Len(sHead) >= 3 Then
        If Mid(sHead, 1, 1) Like "[1-9]" Then
            If Mid(sHead, 2, 2) = ". " Then
                nDigits = 1
            ElseIf Len(sHead) >= 4 Then
                If Mid(sHead, 2, 1) Like "[0-9]" And Mid(sHead, 3, 2) = ". " Then
                    nDigits = 2
                End If
            End If
        End If
    End If
    If nDigits > 0 Then
        Dim oHead As Range
        Set oHead = ActiveDocument.Range(oRange.start, oRange.start + nDigits + 2)
        oHead.text = "(" & Left(sHead, nDigits) & ") "
    End If

End Sub
