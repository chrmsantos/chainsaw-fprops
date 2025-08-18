Sub SubstSugerirPorIndicar()
'
' Replace "Sugerir" and its variants with "Indicar" and its variants throughout the document.
'
    Dim findArray As Variant
    Dim replaceArray As Variant
    Dim i As Integer

    ' List of variants to replace (add more as needed)
    findArray = Array("Sugerir", "sugerir, ""sugiro", "Sugiro", "sugerido", "Sugerido", "Sugerida", "sugerida", "Sugeridos", "sugeridos", "Sugeridas", "sugeridas", "Sugere", "sugere", "Sugeri", "sugeri", "Sugerimos", "sugerimos")
    replaceArray = Array("Indicar", "indicar", "indico", "Indico", "indicado", "Indicado", "Indicada", "indicada", "Indicados", "indicados", "Indicadas", "indicadas", "Indica", "indica", "Indiquei", "indiquei", "Indicamos", "indicamos")

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
