Option Explicit

'================================================================================
' MÓDULO PRINCIPAL DE RETIFICAÇÃO DE PROPOSITURAS - VERSÃO 0.3.0
' Requisitos:
' - Microsoft Word 2010 ou superior
' - Referências habilitadas:
'   * Microsoft Word Object Library
'   * Microsoft Office Object Library
'================================================================================
'
' Autor: Christian Martin dos Santos
' Data de criação: 08/04/2025
' Data da última atualização: 21/04/2025
' Descrição: Este módulo automatiza o processamento de proposituras no Microsoft Word.
'
'================================================================================

'--------------------------------------------------------------------------------
' DECLARAÇÃO DE VARIÁVEIS GLOBAIS
'--------------------------------------------------------------------------------
Private Const BACKUP_BASE_PATH As String = "Documentos\_VbaCmsbo\RetPropLegis\OriginaisBackup\"
Private Const HEADER_IMAGE_PATH As String = "Documents\_VbaCmsbo\RetPropLegis\ConfigFiles\Cabecalho.png"

'--------------------------------------------------------------------------------
' Botão: Retificar o documento
'--------------------------------------------------------------------------------
' SUBROTINA PRINCIPAL: RetifyTheDocument
' Executa a retificação do documento ativo, incluindo backup, validações e processamento.
'--------------------------------------------------------------------------------
Sub RetifyTheDocument()
    On Error GoTo ErrorHandler

    ' Verifica se há documentos abertos
    If Documents.Count = 0 Then
        MsgBox "Nenhum documento está aberto. Por favor, abra um documento antes de executar a retificação.", _
               vbExclamation, "Documento não encontrado"
        Exit Sub
    End If

    Dim originalDoc As Document: Set originalDoc = ActiveDocument
    Dim backupPath As String
    Dim editCount As Integer
    Dim startTime As Double: startTime = Timer

    ' Desativa a atualização da tela para melhorar a performance
    Application.ScreenUpdating = False

    ' Criação de backup
    backupPath = CreateBackup(originalDoc)
    If backupPath = "" Then
        MsgBox "O backup não foi criado. A execução continuará.", vbExclamation, "Aviso de Backup"
    End If

    ' Processamento do documento
    editCount = ProcessDocument(originalDoc)

    ' Exibe mensagem de conclusão
    ShowCompletionMessage backupPath, originalDoc.FullName, editCount, Timer - startTime

Cleanup:
    Application.ScreenUpdating = True

    ' Executa validações após o processamento
    RunValidations originalDoc

    ' Ajusta o zoom do documento para 110%
    ActiveWindow.View.Zoom.Percentage = 110
    Exit Sub

ErrorHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro no Processamento"
    Resume Cleanup
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: CreateBackup
' Cria um backup do documento ativo, garantindo que o diretório seja criado se não existir.
'--------------------------------------------------------------------------------
Private Function CreateBackup(doc As Document) As String
    On Error GoTo ErrorHandler

    Dim backupFolder As String
    Dim backupPath As String
    Dim docName As String

    ' Define o nome do documento e o caminho do backup
    docName = IIf(doc.Path = "", "Documento1.docx", doc.Name)
    backupFolder = Environ("USERPROFILE") & "\" & BACKUP_BASE_PATH & Format(Date, "yyyy-mm-dd") & "\"

    ' Garante que o diretório do backup seja criado
    If Not CreateDirectoryIfNeeded(backupFolder) Then
        MsgBox "Não foi possível criar o diretório de backup: " & backupFolder & vbCrLf & _
               "A execução continuará sem o backup.", vbExclamation, "Aviso de Backup"
        CreateBackup = "" ' Retorna vazio, mas permite continuar
        Exit Function
    End If

    ' Define o caminho completo do backup
    backupPath = backupFolder & SanitizeFileName(docName)

    ' Salva o documento no caminho do backup
    doc.SaveAs FileName:=backupPath, FileFormat:=wdFormatDocumentDefault ' Compatível com Word 2007 e posterior
    CreateBackup = backupPath
    Exit Function

ErrorHandler:
    MsgBox "Erro ao criar o backup: " & Err.Description & vbCrLf & _
           "A execução continuará sem o backup.", vbExclamation, "Aviso de Backup"
    CreateBackup = "" ' Retorna vazio, mas permite continuar
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: ProcessDocument
' Realiza o processamento do documento ativo, incluindo formatação, validações e ajustes.
'--------------------------------------------------------------------------------
Private Function ProcessDocument(doc As Document) As Integer
    Dim editCount As Integer: editCount = 0

    ' Remove linhas em branco antes do título
    RemoveBlankLinesBeforeTitle doc
    editCount = editCount + 1

    ' Remove espaços duplicados
    RemoveExtraSpaces doc
    editCount = editCount + 1

    ' Insere a imagem no cabeçalho
    InsertImageInHeader doc
    editCount = editCount + 1

    ' Limpa os metadados do documento
    ClearDocumentMetadata doc
    editCount = editCount + 1

    ' Remove toda a formatação
    ClearDocumentFormatting doc
    editCount = editCount + 1

    ' Aplica formatação padrão
    FormatDocument doc
    editCount = editCount + 1

    ' Ajusta o recuo do primeiro parágrafo (ementa)
    IndentFirstParagraph doc
    editCount = editCount + 1

    ' Substitui a primeira palavra da ementa, se necessário
    ReplaceFirstWordInEmenta doc
    editCount = editCount + 1

    ' Corrige o final de parágrafos iniciados com "Considerando"
    FixConsiderandoEnding doc
    editCount = editCount + 1

    ' Substitui termos legais
    ReplaceLegalTerms doc
    editCount = editCount + 1

    ' Realiza substituições padrão
    ApplyStandardReplacements doc
    editCount = editCount + 1

    ' Remove marcas d'água
    RemoveDocumentWatermarks doc
    editCount = editCount + 1

    ' Formata linhas específicas
    FormatSpecificLines doc
    editCount = editCount + 1

    ProcessDocument = editCount
End Function

'--------------------------------------------------------------------------------
' SUBROTINA: InsertImageInHeader
' Insere a imagem "cabecalho.png" no cabeçalho, dimensionada para 17 cm de largura,
' centralizada horizontalmente e verticalmente, e adiciona uma linha em branco alinhada à direita.
'--------------------------------------------------------------------------------
Private Sub InsertImageInHeader(doc As Document)
    On Error Resume Next

    Dim imagePath As String
    imagePath = Environ("USERPROFILE") & "\" & HEADER_IMAGE_PATH

    If Dir(imagePath) = "" Then
        MsgBox "O arquivo de imagem não foi encontrado: " & imagePath, vbExclamation, "Erro ao Inserir Imagem"
        Exit Sub
    End If

    Dim section As Section
    Dim header As HeaderFooter
    Dim shape As Shape
    Dim inlineShape As InlineShape
    Dim aspectRatio As Double
    Dim originalWidth As Double
    Dim originalHeight As Double

    ' Define a largura da imagem para 17 cm
    Dim imageWidth As Double
    imageWidth = Application.CentimetersToPoints(17)

    For Each section In doc.Sections
        Set header = section.Headers(wdHeaderFooterPrimary)

        ' Remove elementos existentes no cabeçalho
        header.Range.Text = ""

        ' Insere a imagem como InlineShape
        Set inlineShape = header.Range.InlineShapes.AddPicture(FileName:=imagePath, LinkToFile:=False, SaveWithDocument:=True)

        ' Calcula a altura proporcional com base na largura de 17 cm
        originalWidth = inlineShape.Width
        originalHeight = inlineShape.Height
        aspectRatio = originalHeight / originalWidth
        inlineShape.Width = imageWidth
        inlineShape.Height = inlineShape.Width * aspectRatio

        ' Converte InlineShape para Shape para ajustar a posição
        Set shape = inlineShape.ConvertToShape

        ' Centraliza a imagem horizontalmente e verticalmente
        shape.LockAnchor = True
        shape.Left = wdShapeCenter
        shape.Top = wdShapeCenter

        ' Adiciona uma linha em branco alinhada à direita
        With header.Range
            .Collapse wdCollapseEnd
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .InsertAfter vbCrLf
        End With
    Next section
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ClearDocumentMetadata
' Remove os metadados do documento ativo.
'--------------------------------------------------------------------------------
Private Sub ClearDocumentMetadata(doc As Document)
    On Error Resume Next

    Dim prop As DocumentProperty
    doc.BuiltInDocumentProperties("Title") = ""
    doc.BuiltInDocumentProperties("Subject") = ""
    doc.BuiltInDocumentProperties("Keywords") = ""
    doc.BuiltInDocumentProperties("Comments") = ""
    doc.BuiltInDocumentProperties("Author") = "Anônimo"
    doc.BuiltInDocumentProperties("Last Author") = "Anônimo"
    doc.BuiltInDocumentProperties("Manager") = ""
    doc.BuiltInDocumentProperties("Company") = ""

    For Each prop In doc.CustomDocumentProperties
        prop.Delete
    Next prop
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ClearDocumentFormatting
' Remove toda a formatação do documento ativo.
'--------------------------------------------------------------------------------
Private Sub ClearDocumentFormatting(doc As Document)
    On Error Resume Next
    doc.Content.Font.Reset
    doc.Content.ParagraphFormat.Reset
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: FormatDocument
' Ajusta a formatação do documento ativo para um padrão predefinido.
'--------------------------------------------------------------------------------
Private Sub FormatDocument(doc As Document)
    On Error Resume Next

    With doc.PageSetup
        .TopMargin = Application.CentimetersToPoints(4.5)
        .BottomMargin = Application.CentimetersToPoints(3)
        .LeftMargin = Application.CentimetersToPoints(3)
        .RightMargin = Application.CentimetersToPoints(3)
        .HeaderDistance = Application.CentimetersToPoints(0.7)
        .FooterDistance = Application.CentimetersToPoints(0.7)
    End With

    Dim para As Paragraph
    For Each para In doc.Paragraphs
        With para.Range.Font
            .Name = "Arial"
            .Size = 12
        End With
        With para.Format
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
        End With
    Next para
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: IndentFirstParagraph
' Ajusta o recuo do primeiro parágrafo (ementa) logo abaixo do título,
' garantindo que ele contenha pelo menos uma letra.
'--------------------------------------------------------------------------------
Private Sub IndentFirstParagraph(doc As Document)
    On Error Resume Next

    ' Verifica se o documento possui pelo menos dois parágrafos
    If doc.Paragraphs.Count < 2 Then Exit Sub

    ' Obtém o texto do segundo parágrafo
    Dim secondParagraphText As String
    secondParagraphText = Trim(doc.Paragraphs(2).Range.Text)

    ' Verifica se o segundo parágrafo contém pelo menos uma letra
    If secondParagraphText Like "*[A-Za-z]*" Then
        doc.Paragraphs(2).Format.LeftIndent = Application.CentimetersToPoints(9)
    End If
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
            .Text = replacements(i)(0)
            .Replacement.Text = replacements(i)(1)
            .MatchWildcards = replacements(i)(2)
            .Execute Replace:=wdReplaceAll
        End With
    Next i
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: RemoveDocumentWatermarks
' Remove todas as marcas d'água do documento ativo.
'--------------------------------------------------------------------------------
Private Sub RemoveDocumentWatermarks(doc As Document)
    On Error Resume Next

    Dim section As Section
    Dim header As HeaderFooter
    Dim shape As Shape

    For Each section In doc.Sections
        For Each header In section.Headers
            For Each shape In header.Shapes
                If shape.Type = msoTextEffect Then shape.Delete
            Next shape
        Next header
        For Each header In section.Footers
            For Each shape In header.Shapes
                If shape.Type = msoTextEffect Then shape.Delete
            Next shape
        Next header
    Next section
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ShowCompletionMessage
' Exibe uma mensagem de conclusão com informações sobre o backup e o processamento.
'--------------------------------------------------------------------------------
Private Sub ShowCompletionMessage(backupPath As String, docPath As String, editCount As Integer, executionTime As Double)
    MsgBox "Retificação concluída com sucesso!" & vbCrLf & vbCrLf & _
           "Backup criado em: " & backupPath & vbCrLf & _
           "Número de edições realizadas: " & editCount & vbCrLf & _
           "Tempo de execução: " & Format(executionTime, "0.00") & " segundos", _
           vbInformation, "Retificação Completa"
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: CreateDirectoryIfNeeded
' Cria o diretório especificado, incluindo subdiretórios, se necessário.
'--------------------------------------------------------------------------------
Private Function CreateDirectoryIfNeeded(folderPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim parentFolder As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Verifica se o diretório já existe
    If Not fso.FolderExists(folderPath) Then
        ' Cria diretórios recursivamente
        parentFolder = fso.GetParentFolderName(folderPath)
        If parentFolder <> "" And Not fso.FolderExists(parentFolder) Then
            CreateDirectoryIfNeeded parentFolder
        End If
        fso.CreateFolder folderPath
    End If

    ' Verifica novamente se o diretório foi criado com sucesso
    CreateDirectoryIfNeeded = fso.FolderExists(folderPath)
    Exit Function

ErrorHandler:
    MsgBox "Erro ao criar o diretório: " & folderPath & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro no Backup"
    CreateDirectoryIfNeeded = False
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: SanitizeFileName
' Remove caracteres inválidos do nome do arquivo.
'--------------------------------------------------------------------------------
Private Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String: invalidChars = "\/:*?""<>|"
    Dim i As Integer
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i
    SanitizeFileName = fileName
End Function

'--------------------------------------------------------------------------------
' Botão: Desfazer todas as alterações
'--------------------------------------------------------------------------------
' SUBROTINA: UndoAllChanges
' Finalidade: Desfaz todas as alterações realizadas no documento ativo.
'--------------------------------------------------------------------------------
Sub UndoAllChanges()
    On Error GoTo ErrorHandler

    ' Verifica se há um documento ativo
    If Documents.Count = 0 Then
        MsgBox "Nenhum documento está aberto para desfazer alterações.", vbExclamation, "Documento não encontrado"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Verifica se há alterações a serem desfeitas
    If doc.Saved Then
        MsgBox "Nenhuma alteração foi realizada desde o último salvamento.", vbInformation, "Nada a desfazer"
        Exit Sub
    End If

    ' Desativa a atualização da tela para melhorar a performance
    Application.ScreenUpdating = False

    ' Desfaz todas as alterações realizadas no documento
    While doc.Undo
        ' Continua desfazendo até que não haja mais alterações
    Wend

    ' Restaura a atualização da tela
    Application.ScreenUpdating = True

    ' Mensagem de conclusão
    MsgBox "Todas as alterações realizadas desde a abertura do arquivo ou do último salvamento foram desfeitas com sucesso.", _
           vbInformation, "Alterações Desfeitas"
    Exit Sub

ErrorHandler:
    ' Tratamento de erros
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro ao Desfazer Alterações"
End Sub

'--------------------------------------------------------------------------------
' Botão: Salvar, fechar e minimizar o Word
'--------------------------------------------------------------------------------
' SUBROTINA: SaveCloseAndMinimizeWord
' Finalidade: Salva o documento ativo, fecha-o e minimiza a janela do Microsoft Word.
'--------------------------------------------------------------------------------
Sub SaveCloseAndMinimizeWord()
    On Error GoTo ErrorHandler

    ' Verifica se há um documento ativo
    If Documents.Count = 0 Then
        MsgBox "Nenhum documento está aberto para salvar e fechar.", vbExclamation, "Documento não encontrado"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Salva o documento ativo
    If Not doc.Saved Then
        doc.Save
        MsgBox "Documento salvo com sucesso.", vbInformation, "Salvamento Concluído"
    Else
        MsgBox "Nenhuma alteração foi detectada. O documento já está salvo.", vbInformation, "Sem Alterações"
    End If

    ' Fecha o documento ativo
    doc.Close SaveChanges:=wdDoNotSaveChanges

    ' Minimiza a janela do Microsoft Word
    Application.WindowState = wdWindowStateMinimize

    Exit Sub

ErrorHandler:
    ' Tratamento de erros
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro ao Salvar, Fechar e Minimizar"
End Sub

'--------------------------------------------------------------------------------
' FUNÇÃO: ValidateFirstLineFormat
' Verifica se a primeira linha do documento termina com o texto literal "$NUMERO$/$ANO$".
'--------------------------------------------------------------------------------
Private Function ValidateFirstLineFormat(doc As Document) As Boolean
    On Error Resume Next

    Dim firstLine As String
    firstLine = Trim(doc.Paragraphs(1).Range.Text)

    ' Verifica se a primeira linha termina com o texto literal "$NUMERO$/$ANO$"
    If Not firstLine Like "* $NUMERO$/$ANO$" Then
        MsgBox "A primeira linha do documento não termina com o texto literal '$NUMERO$/$ANO$'.", _
               vbExclamation, "Formato Inválido"
        ValidateFirstLineFormat = False
    Else
        ValidateFirstLineFormat = True
    End If
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: RunValidations
' Finalidade: Executa todas as validações necessárias no documento ativo.
'--------------------------------------------------------------------------------
Private Sub RunValidations(doc As Document)
    On Error Resume Next

    Dim isValidFirstLine As Boolean

    ' Validação da primeira linha
    isValidFirstLine = ValidateFirstLineFormat(doc)

    ' Outras validações podem ser adicionadas aqui no futuro
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ReplaceFirstWordInEmenta
' Finalidade: Substitui a primeira palavra da ementa (parágrafo logo abaixo do título)
'             quando for "Sugiro", "Sugere" ou "Sugestão", por "Indico", "Sugiro" ou "Indicação".
'--------------------------------------------------------------------------------
Private Sub ReplaceFirstWordInEmenta(doc As Document)
    On Error Resume Next

    ' Verifica se o documento possui pelo menos dois parágrafos (título + ementa)
    If doc.Paragraphs.Count < 2 Then Exit Sub

    ' Obtém o texto do segundo parágrafo (ementa)
    Dim ementa As Range
    Set ementa = doc.Paragraphs(2).Range
    Dim firstWord As String
    firstWord = Trim(Split(ementa.Text, " ")(0)) ' Obtém a primeira palavra

    ' Substitui a primeira palavra conforme necessário
    Select Case LCase(firstWord)
        Case "sugiro"
            ementa.Text = Replace(ementa.Text, firstWord, "Indico", 1, 1, vbTextCompare)
        Case "sugere"
            ementa.Text = Replace(ementa.Text, firstWord, "Sugiro", 1, 1, vbTextCompare)
        Case "sugestão"
            ementa.Text = Replace(ementa.Text, firstWord, "Indicação", 1, 1, vbTextCompare)
    End Select
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: RemoveBlankLinesBeforeTitle
' Finalidade: Remove todas as linhas em branco antes do título do documento, garantindo que ele esteja na primeira linha.
'--------------------------------------------------------------------------------
Private Sub RemoveBlankLinesBeforeTitle(doc As Document)
    On Error Resume Next

    ' Verifica se o documento possui pelo menos um parágrafo
    If doc.Paragraphs.Count = 0 Then Exit Sub

    ' Loop para remover linhas em branco no início do documento
    Do While Trim(doc.Paragraphs(1).Range.Text) = ""
        doc.Paragraphs(1).Range.Delete
    Loop
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: RemoveExtraSpaces
' Finalidade: Remove espaços duplicados, triplicados, etc., em todo o texto do documento.
'--------------------------------------------------------------------------------
Private Sub RemoveExtraSpaces(doc As Document)
    On Error Resume Next

    ' Verifica se há um documento ativo
    If Documents.Count = 0 Then
        Exit Sub
    End If

    ' Remove espaços duplicados no texto do documento
    With doc.Content.Find
        .Text = "  " ' Dois espaços
        .Replacement.Text = " " ' Um espaço
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceAll)
            ' Continua substituindo até que não haja mais espaços duplicados
        Loop
    End With
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: FixConsiderandoEnding
' Finalidade: Garante que parágrafos iniciados com "Considerando" terminem com ponto e vírgula (;).
'--------------------------------------------------------------------------------
Private Sub FixConsiderandoEnding(doc As Document)
    On Error Resume Next

    ' Verifica se há parágrafos no documento
    If doc.Paragraphs.Count = 0 Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String

    ' Itera por todos os parágrafos do documento
    For Each para In doc.Paragraphs
        paraText = Trim(para.Range.Text) ' Obtém o texto do parágrafo sem espaços extras

        ' Verifica se o parágrafo começa com "Considerando" (case-insensitive)
        If LCase(Left(paraText, 11)) = "considerando" Then
            ' Substitui o ponto final no final do parágrafo por ponto e vírgula
            If Right(paraText, 1) = "." Then
                para.Range.Text = Left(paraText, Len(paraText) - 1) & ";"
            End If
        End If
    Next para
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ReplaceLegalTerms
' Finalidade: Substitui "artigo" por "art.", "inciso" por "inc." e "alínea" por "al."
'             no texto do documento, independentemente da caixa.
'--------------------------------------------------------------------------------
Private Sub ReplaceLegalTerms(doc As Document)
    On Error Resume Next

    ' Verifica se há um documento ativo
    If Documents.Count = 0 Then
        Exit Sub
    End If

    ' Substituições de termos legais
    Dim replacements As Variant
    replacements = Array( _
        Array("artigo", "art."), _
        Array("inciso", "inc."), _
        Array("alínea", "al.") _
    )

    Dim i As Integer
    For i = LBound(replacements) To UBound(replacements)
        With doc.Content.Find
            .Text = replacements(i)(0)
            .Replacement.Text = replacements(i)(1)
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = False ' Ignora a caixa (maiúsculas/minúsculas)
            .Execute Replace:=wdReplaceAll
        End With
    Next i
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: FormatSpecificLines
' Formata a primeira linha, a linha com "justificativa(s)" e a linha com "anexo(s)".
'--------------------------------------------------------------------------------
Private Sub FormatSpecificLines(doc As Document)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim paraIndex As Integer
    Dim maxParagraphs As Integer

    ' Define a quantidade máxima de parágrafos a serem processados (para evitar crashes)
    maxParagraphs = 1000 ' Ajuste conforme necessário

    ' Itera por todos os parágrafos do documento
    paraIndex = 1
    For Each para In doc.Paragraphs
        If paraIndex > maxParagraphs Then Exit For ' Limita o número de parágrafos processados

        paraText = Trim(para.Range.Text)

        ' Formata a primeira linha do texto
        If paraIndex = 1 Then
            With para.Range
                .Font.Bold = True
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.LeftIndent = 0
            End With
        End If

        ' Formata a linha com "justificativa" ou "justificativas"
        If LCase(paraText) = "justificativa" Or LCase(paraText) = "justificativas" Then
            With para.Range
                .Font.Bold = True
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.LeftIndent = 0
            End With
        End If

        ' Formata a linha com "anexo" ou "anexos"
        If InStr(LCase(paraText), "anexo") > 0 Then
            With para.Range
                .Font.Bold = True
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
                .ParagraphFormat.LeftIndent = 0
            End With
        End If

        paraIndex = paraIndex + 1
    Next para

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao formatar linhas específicas: " & Err.Description, vbCritical, "Erro"
End Sub
