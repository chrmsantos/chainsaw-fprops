Sub SubstSugerirPorIndicarMain()
'
' Replace "Sugerir" and its variants with "Indicar" and its variants throughout the document.
'
    Dim findArray As Variant
    Dim replaceArray As Variant
    Dim i As Integer

    ' Add Word constants if not available
    Const wdFindContinue As Long = 1
    Const wdReplaceAll As Long = 2

    ' List of variants to replace (aligned and corrected)
    findArray = Array( _
        "Sugerir", "sugerir", _
        "Sugiro", "sugiro", _
        "Sugerido", "sugerido", _
        "Sugerida", "sugerida", _
        "Sugeridos", "sugeridos", _
        "Sugeridas", "sugeridas", _
        "Sugere", "sugere", _
        "Sugeri", "sugeri", _
        "Sugerimos", "sugerimos" _
    )
    replaceArray = Array( _
        "Indicar", "indicar", _
        "Indico", "indico", _
        "Indicado", "indicado", _
        "Indicada", "indicada", _
        "Indicados", "indicados", _
        "Indicadas", "indicadas", _
        "Indica", "indica", _
        "Indiquei", "indiquei", _
        "Indicamos", "indicamos" _
    )

    For i = LBound(findArray) To UBound(findArray)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = findArray(i)
            .Replacement.Text = replaceArray(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = True
            .Execute Replace:=wdReplaceAll
        End With
    Next i
End Sub
