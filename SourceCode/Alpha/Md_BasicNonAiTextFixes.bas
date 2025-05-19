'================================================================================
' PONTO DE ENTRADA: Main_RBNAF
' Orquestra as correções textuais básicas não-AI no documento informado.
'================================================================================
Public Sub Main_RBNAF(doc As Document)
    On Error GoTo ErrorHandler

    ApplyStandardReplacements doc ' Substituições padrão
    FixConsiderandoEnding doc    ' Ajuste de "Considerando"
    ReplaceLastWordFirstLine doc ' Substituição da última palavra da primeira linha

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar as correções textuais: " & Err.Description, vbCritical, "Erro"
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: FixConsiderandoEnding
' Finalidade: Garante que parágrafos iniciados com "Considerando" terminem com ponto e vírgula (;).
'--------------------------------------------------------------------------------
Private Sub FixConsiderandoEnding(doc As Document)
    On Error Resume Next

    If doc.Paragraphs.Count = 0 Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String

    For Each para In doc.Paragraphs
        paraText = Trim(Replace(para.Range.Text, vbCr, "")) ' Remove marcas de parágrafo

        If LCase(Left(paraText, 11)) = "considerando" Then
            ' Se termina com ponto, troca por ponto e vírgula
            If Right(paraText, 1) = "." Then
                paraText = Left(paraText, Len(paraText) - 1) & ";"
            ' Se não termina com ponto e vírgula, adiciona
            ElseIf Right(paraText, 1) <> ";" And Len(paraText) > 0 Then
                paraText = paraText & ";"
            End If
            ' Atualiza o texto do parágrafo sem perder a marca de parágrafo
            para.Range.Text = paraText & vbCr
        End If
    Next para
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ApplyStandardReplacements
' Realiza substituições de texto no documento com base em padrões predefinidos.
'--------------------------------------------------------------------------------
Private Sub ApplyStandardReplacements(doc As Document)
    On Error Resume Next

    Dim replacements As Variant
    replacements = Array( _
        Array("[!.\?\n] Rua", "rua", True), _
        Array("[!.\?\n] Bairro", "bairro", True), _
        Array("[Dd][´`][Oo]este", "d'Oeste", True), _
        Array("([0-9]@ de [A-Za-z]@ de )([0-9]{4})", Format(Date, "dd 'de' mmmm 'de' yyyy"), True))

    Dim i As Integer
    For i = LBound(replacements) To UBound(replacements)
        With doc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = replacements(i)(0)
            .Replacement.Text = replacements(i)(1)
            .MatchWildcards = replacements(i)(2)
            .Execute Replace:=wdReplaceAll
        End With
    Next i
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ReplaceLastWordFirstLine
' Finalidade: Substitui a última palavra da primeira linha por "$NUMERO$/$ANO$".
'--------------------------------------------------------------------------------
Private Sub ReplaceLastWordFirstLine(doc As Document)
    On Error Resume Next

    If doc.Paragraphs.Count = 0 Then Exit Sub

    Dim para As Paragraph
    Set para = doc.Paragraphs(1)
    Dim paraText As String
    paraText = Trim(Replace(para.Range.Text, vbCr, ""))

    If Len(paraText) = 0 Then Exit Sub

    Dim words() As String
    words = Split(paraText, " ")

    If UBound(words) >= 0 Then
        words(UBound(words)) = "$NUMERO$/$ANO$"
        paraText = Join(words, " ")
        para.Range.Text = paraText & vbCr
    End If
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: UpdateDateBeforeSignature
' Atualiza a linha de data (terceira acima da palavra-chave de assinatura) para a data atual.
'--------------------------------------------------------------------------------
Private Sub UpdateDateBeforeSignature(doc As Document)
    On Error Resume Next

    If doc.Paragraphs.Count < 4 Then Exit Sub

    Dim i As Long
    Dim paraText As String
    Dim keywords As Variant
    keywords = Array("vereador", "presidente", "vice-presidente", "1º secretário", "2º secretário")

    For i = doc.Paragraphs.Count To 4 Step -1
        paraText = LCase(Trim(doc.Paragraphs(i).Range.Text))
        Dim k As Integer
        For k = LBound(keywords) To UBound(keywords)
            If InStr(paraText, keywords(k)) > 0 Then
                ' Atualiza a terceira linha acima da palavra-chave
                doc.Paragraphs(i - 3).Range.Text = Format(Date, "d 'de' mmmm 'de' yyyy") & vbCr
                Exit Sub
            End If
        Next k
    Next i
End Sub


