Sub IndicaoEmenta()
    Dim par As Paragraph
    Dim texto As String

    Set par = Selection.Paragraphs(1)
    
    ' 1. Recuo de 9 cm em todas as linhas
    par.LeftIndent = 255.15 ' 9 cm em pontos
    par.FirstLineIndent = 0
    
    ' 2. Alinhamento justificado
    par.Alignment = 3 ' wdAlignParagraphJustify
    
    ' 3. Substituir "Sugere" por "Indica" se for a primeira palavra do parágrafo
    texto = par.Range.Text
    texto = Trim(texto)
    If LCase(Left(texto, 6)) = "sugere" Then
        texto = "Indica" & Mid(texto, 7)
        par.Range.Text = texto
    End If
End Sub