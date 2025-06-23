Sub InserirDocumentoComFormatacao()
    Dim docAtivo As Document
    Dim docFonte As Document
    Dim caminhoFonte As String
    Dim rngFinal As Range
    Dim userName As String

    ' 1. Pule uma linha no final do documento ativo
    Set docAtivo = ActiveDocument
    Set rngFinal = docAtivo.Content
    rngFinal.Collapse Direction:=wdCollapseEnd
    rngFinal.InsertParagraphAfter

    ' 2. Insira uma quebra de página no novo final do documento ativo
    rngFinal.Collapse Direction:=wdCollapseEnd
    rngFinal.InsertBreak Type:=wdPageBreak

    ' 3. Copie todo o conteúdo e formatação do documento fonte
    userName = Environ("USERNAME")
    caminhoFonte = "C:\Users\" & userName & "\Documentos\INDICACAO OFICIAL (rc).docx"
    Set docFonte = Documents.Open(FileName:=caminhoFonte, ReadOnly:=True)
    docFonte.Content.Copy

    ' 4. Cole o conteúdo formatado após a quebra de página
    rngFinal.Collapse Direction:=wdCollapseEnd
    rngFinal.PasteAndFormat (wdFormatOriginalFormatting)
    docFonte.Close SaveChanges:=False

    ' 5. Execute a subrotina Formatter (já existente no projeto)
    Call Formatter

    ' 6. Visualização em duas páginas com zoom de 80%
    With ActiveWindow.View
        .Type = wdPrintView
        .Zoom.PageColumns = 2
        .Zoom.Percentage = 80
    End With
End Sub
