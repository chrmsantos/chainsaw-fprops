Sub ModelImporterMain()
    Dim docAtivo As Document
    Dim docFonte As Document
    Dim caminhoFonte As String
    Dim rngFinal As Range
    Dim rngOrigem As Range
    Dim rngDestino As Range
    Dim userName As String

    ' 1. Pule uma linha no final do documento ativo
    Set docAtivo = ActiveDocument
    Set rngFinal = docAtivo.Content
    rngFinal.Collapse Direction:=wdCollapseEnd
    rngFinal.InsertParagraphAfter

    ' 2. Insira uma quebra de página no novo final do documento ativo
    rngFinal.Collapse Direction:=wdCollapseEnd
    rngFinal.InsertBreak Type:=wdPageBreak

    ' 3. Abra o documento fonte e obtenha o range do conteúdo
    userName = Environ("USERNAME")
    caminhoFonte = "C:\Users\" & userName & "\Documents\INDICACAO OFICIAL.docx"
    Set docFonte = Documents.Open(FileName:=caminhoFonte, ReadOnly:=True)
    Set rngOrigem = docFonte.Content

    ' 4. Insira o conteúdo formatado logo após a quebra de página
    Set rngDestino = docAtivo.Content
    rngDestino.Collapse Direction:=wdCollapseEnd
    rngDestino.FormattedText = rngOrigem.FormattedText

    docFonte.Close SaveChanges:=False

    ' 5. Visualização em duas páginas com zoom de 80%
    With ActiveWindow.View
        .Type = wdPrintView
        .Zoom.PageColumns = 2
        .Zoom.Percentage = 80
    End With
End Sub
