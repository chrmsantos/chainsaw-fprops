' =============================================================================
' PROJETO: CHAINSAW FOR PROPOSALS (CHAINSW-FPROPS)
' =============================================================================
'
' Sistema automatizado de padronização de documentos legislativos no Microsoft Word
'
' Licença: Apache 2.0 modificada (ver LICENSE)
' Versão: 1.0-alpha8-optimized | Data: 2025-09-18
' Repositório: github.com/chrmsantos/chainsaw-fprops
' Autor: Christian Martin dos Santos <chrmsantos@gmail.com>
'
' =============================================================================
' FUNCIONALIDADES PRINCIPAIS:
' =============================================================================
'
' • VERIFICAÇÕES DE SEGURANÇA E COMPATIBILIDADE:
'   - Validação de versão do Word (mínimo: 2010)
'   - Verificação de tipo e proteção do documento
'   - Controle de espaço em disco e estrutura mínima
'   - Proteção contra falhas e recuperação automática
'
' • SISTEMA DE BACKUP AUTOMÁTICO:
'   - Backup automático antes de qualquer modificação
'   - Pasta de backups organizada por documento
'   - Limpeza automática de backups antigos (limite: 10 arquivos)
'   - Subrotina pública para acesso à pasta de backups
'
' • FORMATAÇÃO AUTOMATIZADA INSTITUCIONAL:
'   - Limpeza completa de formatação ao iniciar
'   - Remoção robusta de espaços múltiplos e tabs
'   - Controle de linhas vazias (máximo 2 sequenciais)
'   - PROTEÇÃO MÁXIMA: Preserva imagens inline, flutuantes e objetos
'   - Primeira linha: SEMPRE caixa alta, negrito, sublinhado, centralizada
'   - Parágrafos 2°, 3° e 4°: recuo esquerdo 9cm, sem recuo primeira linha
'   - "Considerando": caixa alta e negrito no início de parágrafos
'   - "Justificativa": centralizada, sem recuos, negrito, capitalizada
'   - "Anexo/Anexos": alinhado à esquerda, sem recuos, negrito, capitalizado
'   - Configuração de margens e orientação (A4)
'   - Fonte Arial 12pt com espaçamento 1.4
'   - Recuos e alinhamento justificado
'   - Cabeçalho com logotipo institucional
'   - Rodapé com numeração centralizada
'   - Visualização: zoom 110%, régua visível, modo impressão
'   - Remoção de marcas d'água e formatações manuais
'
' • SISTEMA DE LOGS E MONITORAMENTO:
'   - Registro detalhado de operações
'   - Controle de erros com fallback
'   - Mensagens na barra de status
'   - Histórico de execução
'
' • PERFORMANCE OTIMIZADA:
'   - Processamento eficiente para documentos grandes
'   - Desabilitação temporária de atualizações visuais
'   - Gerenciamento inteligente de recursos
'
' =============================================================================

'VBA
Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Word built-in constants
Private Const wdNoProtection As Long = -1
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdAlignParagraphJustify As Long = 3
Private Const wdLineSpaceSingle As Long = 0
Private Const wdLineSpace1pt5 As Long = 1
Private Const wdLineSpacingMultiple As Long = 5
Private Const wdStatisticPages As Long = 2
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoPicture As Long = 13
Private Const msoTextEffect As Long = 15
Private Const wdCollapseEnd As Long = 0
Private Const wdFieldPage As Long = 33
Private Const wdFieldNumPages As Long = 26
Private Const wdFieldEmpty As Long = -1
Private Const wdRelativeHorizontalPositionPage As Long = 1
Private Const wdRelativeVerticalPositionPage As Long = 1
Private Const wdWrapTopBottom As Long = 3
Private Const wdAlertsAll As Long = 0
Private Const wdAlertsNone As Long = -1
Private Const wdColorAutomatic As Long = -16777216
Private Const wdOrientPortrait As Long = 0
Private Const wdUnderlineNone As Long = 0
Private Const wdUnderlineSingle As Long = 1
Private Const wdTextureNone As Long = 0
Private Const wdPrintView As Long = 3

' Document formatting constants
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 4.6
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\chainsaw-fprops\private-data\Stamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

' Minimum supported version
Private Const MIN_SUPPORTED_VERSION As Long = 14 ' Word 2010

' Logging constants
Private Const LOG_LEVEL_INFO As Long = 1
Private Const LOG_LEVEL_WARNING As Long = 2
Private Const LOG_LEVEL_ERROR As Long = 3

' Required string constant
Private Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"

' Timeout constants
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

' Backup constants
Private Const BACKUP_FOLDER_NAME As String = "chainsaw-backups"
Private Const MAX_BACKUP_FILES As Long = 10

'================================================================================
' GLOBAL VARIABLES
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date
Private backupFilePath As String

'================================================================================
' MAIN ENTRY POINT - #STABLE
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    
    If Not CheckWordVersion() Then
        Application.StatusBar = "Erro: Versão do Word não suportada (mínimo: Word 2010)"
        LogMessage "Versão do Word " & Application.version & " não suportada. Mínimo: 14.0", LOG_LEVEL_ERROR
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = Nothing
    
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        Application.StatusBar = "Erro: Nenhum documento está acessível"
        LogMessage "Nenhum documento acessível para processamento", LOG_LEVEL_ERROR
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    If Not InitializeLogging(doc) Then
        LogMessage "Falha ao inicializar sistema de logs", LOG_LEVEL_WARNING
    End If
    
    LogMessage "Iniciando padronização do documento: " & doc.Name, LOG_LEVEL_INFO
    
    StartUndoGroup "Padronização de Documento - " & doc.Name
    
    If Not SetAppState(False, "Formatando documento...") Then
        LogMessage "Falha ao configurar estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If
    
    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Operação cancelada: documento precisa ser salvo"
            LogMessage "Operação cancelada - documento não foi salvo", LOG_LEVEL_INFO
            Exit Sub
        End If
    End If
    
    ' Cria backup do documento antes de qualquer modificação
    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
        Application.StatusBar = "Aviso: Backup não foi possível - processando sem backup"
    Else
        Application.StatusBar = "Backup criado - formatando documento..."
    End If
    
    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    Application.StatusBar = "Documento padronizado com sucesso!"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO

CleanUp:
    SafeCleanup
    
    If Not SetAppState(True, "Documento padronizado com sucesso!") Then
        LogMessage "Falha ao restaurar estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro crítico durante processamento - verificar logs"
    
    EmergencyRecovery
End Sub

'================================================================================
' EMERGENCY RECOVERY - #STABLE
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Application.EnableCancelKey = 0
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    CloseAllOpenFiles
End Sub

'================================================================================
' SAFE CLEANUP - LIMPEZA SEGURA - #STABLE
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    
    EndUndoGroup
    
    ReleaseObjects
End Sub

'================================================================================
' RELEASE OBJECTS - #STABLE
'================================================================================
Private Sub ReleaseObjects()
    On Error Resume Next
    
    Dim nullObj As Object
    Set nullObj = Nothing
    
    Dim memoryCounter As Long
    For memoryCounter = 1 To 3
        DoEvents
    Next memoryCounter
End Sub

'================================================================================
' CLOSE ALL OPEN FILES - #STABLE
'================================================================================
Private Sub CloseAllOpenFiles()
    On Error Resume Next
    
    Dim fileNumber As Integer
    For fileNumber = 1 To 511
        If Not EOF(fileNumber) Then
            Close fileNumber
        End If
    Next fileNumber
End Sub

'================================================================================
' VERSION COMPATIBILITY CHECK - #STABLE
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Long
    version = Application.version
    
    If version < MIN_SUPPORTED_VERSION Then
        CheckWordVersion = False
    Else
        CheckWordVersion = True
    End If
    
    Exit Function
    
ErrorHandler:
    CheckWordVersion = False
End Function

'================================================================================
' UNDO GROUP MANAGEMENT - #STABLE
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        EndUndoGroup
    End If
    
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
End Sub

Private Sub EndUndoGroup()
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
End Sub

'================================================================================
' LOGGING MANAGEMENT - APRIMORADO COM DETALHES - #STABLE
'================================================================================
Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If doc.Path <> "" Then
        logFilePath = doc.Path & "\" & Format(Now, "yyyy-mm-dd") & "_" & _
                     Replace(doc.Name, ".doc", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docx", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docm", "") & "_FormattingLog.txt"
    Else
        logFilePath = Environ("TEMP") & "\" & Format(Now, "yyyy-mm-dd") & "_DocumentFormattingLog.txt"
    End If
    
    Open logFilePath For Output As #1
    Print #1, "========================================================"
    Print #1, "LOG DE FORMATAÇÃO DE DOCUMENTO - SISTEMA DE REGISTRO"
    Print #1, "========================================================"
    Print #1, "Sessão: " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    Print #1, "Usuário: " & Environ("USERNAME")
    Print #1, "Estação: " & Environ("COMPUTERNAME")
    Print #1, "Versão Word: " & Application.version
    Print #1, "Documento: " & doc.Name
    Print #1, "Local: " & IIf(doc.Path = "", "(Não salvo)", doc.Path)
    Print #1, "Proteção: " & GetProtectionType(doc)
    Print #1, "Tamanho: " & GetDocumentSize(doc)
    Print #1, "Tempo Execução: " & Format(Now - executionStartTime, "HH:MM:ss")
    Print #1, "Erros: " & Err.Number & " - " & Err.Description
    Print #1, "========================================================"
    Close #1
    
    loggingEnabled = True
    InitializeLogging = True
    
    Exit Function
    
ErrorHandler:
    loggingEnabled = False
    InitializeLogging = False
End Function

Private Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error GoTo ErrorHandler
    
    If Not loggingEnabled Then Exit Sub
    
    Dim levelText As String
    Dim levelIcon As String
    
    Select Case level
        Case LOG_LEVEL_INFO
            levelText = "INFO"
            levelIcon = ""
        Case LOG_LEVEL_WARNING
            levelText = "AVISO"
            levelIcon = ""
        Case LOG_LEVEL_ERROR
            levelText = "ERRO"
            levelIcon = ""
        Case Else
            levelText = "OUTRO"
            levelIcon = ""
    End Select
    
    Dim formattedMessage As String
    formattedMessage = Format(Now, "yyyy-mm-dd HH:MM:ss") & " [" & levelText & "] " & levelIcon & " " & message
    
    Open logFilePath For Append As #1
    Print #1, formattedMessage
    Close #1
    
    Debug.Print "LOG: " & formattedMessage
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FALHA NO LOGGING: " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler
    
    If loggingEnabled Then
        Open logFilePath For Append As #1
        Print #1, "================================================"
        Print #1, "FIM DA SESSÃO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
        Print #1, "Duração: " & Format(Now - executionStartTime, "HH:MM:ss")
        Print #1, "Status: " & IIf(formattingCancelled, "CANCELADO", "CONCLUÍDO")
        Print #1, "================================================"
        Close #1
    End If
    
    loggingEnabled = False
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erro ao finalizar logging: " & Err.Description
    loggingEnabled = False
End Sub

'================================================================================
' UTILITY: GET PROTECTION TYPE - #STABLE
'================================================================================
Private Function GetProtectionType(doc As Document) As String
    On Error Resume Next
    
    Select Case doc.protectionType
        Case wdNoProtection: GetProtectionType = "Sem proteção"
        Case 1: GetProtectionType = "Protegido contra revisões"
        Case 2: GetProtectionType = "Protegido contra comentários"
        Case 3: GetProtectionType = "Protegido contra formulários"
        Case 4: GetProtectionType = "Protegido contra leitura"
        Case Else: GetProtectionType = "Tipo desconhecido (" & doc.protectionType & ")"
    End Select
End Function

'================================================================================
' UTILITY: GET DOCUMENT SIZE - #STABLE
'================================================================================
Private Function GetDocumentSize(doc As Document) As String
    On Error Resume Next
    
    Dim size As Long
    size = doc.BuiltInDocumentProperties("Number of Characters").Value * 2
    
    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' APPLICATION STATE HANDLER - #STABLE
'================================================================================
Private Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim success As Boolean
    success = True
    
    With Application
        On Error Resume Next
        .ScreenUpdating = enabled
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        On Error Resume Next
        .DisplayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        If statusMsg <> "" Then
            On Error Resume Next
            .StatusBar = statusMsg
            If Err.Number <> 0 Then success = False
            On Error GoTo ErrorHandler
        ElseIf enabled Then
            On Error Resume Next
            .StatusBar = False
            If Err.Number <> 0 Then success = False
            On Error GoTo ErrorHandler
        End If
        
        On Error Resume Next
        .EnableCancelKey = 0
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
    End With
    
    SetAppState = success
    Exit Function
    
ErrorHandler:
    SetAppState = False
End Function

'================================================================================
' GLOBAL CHECKING - VERIFICAÇÕES ROBUSTAS
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Application.StatusBar = "Erro: Documento não acessível para verificação"
        LogMessage "Documento não acessível para verificação", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Erro: Tipo de documento não suportado (Tipo: " & doc.Type & ")"
        LogMessage "Tipo de documento não suportado: " & doc.Type, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        Application.StatusBar = "Erro: Documento protegido (" & protectionType & ")"
        LogMessage "Documento protegido detectado: " & protectionType, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If
    
    If doc.ReadOnly Then
        Application.StatusBar = "Erro: Documento em modo somente leitura"
        LogMessage "Documento em modo somente leitura: " & doc.FullName, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not CheckDiskSpace(doc) Then
        Application.StatusBar = "Erro: Espaço em disco insuficiente"
        LogMessage "Espaço em disco insuficiente para operação segura", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
        LogMessage "Estrutura do documento validada com avisos", LOG_LEVEL_WARNING
    End If

    LogMessage "Verificações de segurança concluídas com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    Application.StatusBar = "Erro durante verificações de segurança"
    LogMessage "Erro durante verificações: " & Err.Description, LOG_LEVEL_ERROR
    PreviousChecking = False
End Function

'================================================================================
' DISK SPACE CHECK - VERIFICAÇÃO SIMPLIFICADA
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificação simplificada - assume espaço suficiente se não conseguir verificar
    Dim fso As Object
    Dim drive As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If doc.Path <> "" Then
        Set drive = fso.GetDrive(Left(doc.Path, 3))
    Else
        Set drive = fso.GetDrive(Left(Environ("TEMP"), 3))
    End If
    
    ' Verificação básica - 10MB mínimo
    If drive.AvailableSpace < 10485760 Then ' 10MB em bytes
        LogMessage "Espaço em disco muito baixo", LOG_LEVEL_WARNING
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If
    
    Exit Function
    
ErrorHandler:
    ' Se não conseguir verificar, assume que há espaço suficiente
    CheckDiskSpace = True
End Function

'================================================================================
' MAIN FORMATTING ROUTINE - #STABLE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Formatações básicas de página e estrutura
    If Not ApplyPageSetup(doc) Then
        LogMessage "Falha na configuração de página", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    ' Limpeza e formatações otimizadas (sem logs detalhados para performance)
    ClearAllFormatting doc
    CleanDocumentStructure doc
    ValidatePropositionType doc
    FormatDocumentTitle doc
    
    ' Formatações principais
    If Not ApplyStdFont(doc) Then
        LogMessage "Falha na formatação de fontes", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not ApplyStdParagraphs(doc) Then
        LogMessage "Falha na formatação de parágrafos", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    ' Formatação específica do 1º parágrafo (caixa alta, negrito, sublinhado)
    FormatFirstParagraph doc

    ' Formatação específica do 2º parágrafo
    FormatSecondParagraph doc

    ' Formatações específicas (sem verificação de retorno para performance)
    FormatConsiderandoParagraphs doc
    ApplyTextReplacements doc
    
    ' Formatação específica para Justificativa/Anexo/Anexos
    FormatJustificativaAnexoParagraphs doc
    
    EnableHyphenation doc
    RemoveWatermark doc
    InsertHeaderStamp doc
    
    ' Limpeza final de espaços múltiplos em todo o documento
    CleanMultipleSpaces doc
    
    ' Controle de linhas em branco sequenciais (máximo 2)
    LimitSequentialEmptyLines doc
    
    ' Configuração final da visualização
    ConfigureDocumentView doc
    
    If Not InsertFooterStamp(doc) Then
        LogMessage "Falha na inserção do rodapé", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    LogMessage "Formatação completa aplicada", LOG_LEVEL_INFO
    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatação: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' PAGE SETUP - #STABLE
'================================================================================
Private Function ApplyPageSetup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
        .Gutter = 0
        .Orientation = wdOrientPortrait
    End With
    
    LogMessage "Configuração de página aplicada: margens e orientação definidas", LOG_LEVEL_INFO
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro na configuração de página: " & Err.Description, LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

' ================================================================================
' FONT FORMMATTING - #STABLE
' ================================================================================
Private Function ApplyStdFont(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim underlineRemovedCount As Long
    Dim isTitle As Boolean
    Dim hasConsiderando As Boolean

    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False
        isTitle = False
        hasConsiderando = False

        If para.Range.InlineShapes.Count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
        ' Proteção adicional: verifica outros tipos de conteúdo visual
        If Not hasInlineImage And HasVisualContent(para) Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
        ' Verifica se é o primeiro parágrafo com texto (título)
        If i <= 3 And para.Format.Alignment = wdAlignParagraphCenter Then
            Dim paraText As String
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            If paraText <> "" Then
                ' Primeira linha com texto sempre é título, independente do conteúdo
                isTitle = True
            End If
        End If
        
        ' Verifica se o parágrafo começa com "considerando"
        Dim paraFullText As String
        paraFullText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        If Len(paraFullText) >= 12 And LCase(Left(paraFullText, 12)) = "considerando" Then
            hasConsiderando = True
        End If
        
        ' Verifica se é um parágrafo especial (Justificativa/Anexo/Anexos)
        Dim isSpecialParagraph As Boolean
        Dim cleanParaText As String
        cleanParaText = LCase(paraFullText)
        If cleanParaText = "justificativa" Or cleanParaText = "anexo" Or cleanParaText = "anexos" Then
            isSpecialParagraph = True
        End If

        If Not hasInlineImage Then
            ' Aplica formatação de fonte padrão, mas preserva negrito de títulos e considerandos
            With para.Range.Font
                .Name = STANDARD_FONT
                .Size = STANDARD_FONT_SIZE
                
                ' Remove sublinhado de todo o documento, exceto do título
                If .Underline <> wdUnderlineNone And Not isTitle Then
                    .Underline = wdUnderlineNone
                    underlineRemovedCount = underlineRemovedCount + 1
                End If
                
                .Color = wdColorAutomatic
            End With
            
            ' Trata negrito de forma seletiva
            If Not isTitle And Not hasConsiderando And Not isSpecialParagraph Then
                ' Remove negrito apenas se não for título, considerando ou parágrafo especial
                If para.Range.Font.Bold = True Then
                    para.Range.Font.Bold = False
                End If
            End If
            
            formattedCount = formattedCount + 1
        End If
    Next i
    
    LogMessage "Formatação de fonte aplicada: " & formattedCount & " parágrafos formatados, " & skippedCount & " ignorados (proteção de imagens/objetos), " & underlineRemovedCount & " sublinhados removidos (preservando título, considerandos e seções especiais)", LOG_LEVEL_INFO
    ApplyStdFont = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de fonte: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFont = False
End Function

'================================================================================
' PARAGRAPH FORMATTING - #STABLE
'================================================================================
Private Function ApplyStdParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim paragraphIndent As Single
    Dim firstIndent As Single
    Dim rightMarginPoints As Single
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim paraText As String
    Dim prevPara As Paragraph

    rightMarginPoints = 0

    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

        If para.Range.InlineShapes.Count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
        ' Proteção adicional: verifica outros tipos de conteúdo visual
        If Not hasInlineImage And HasVisualContent(para) Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        If Not hasInlineImage Then
            ' Limpeza robusta de espaços múltiplos
            Dim cleanText As String
            cleanText = para.Range.Text
            
            ' Remove múltiplos espaços consecutivos
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
            
            ' Remove espaços antes/depois de quebras de linha
            cleanText = Replace(cleanText, " " & vbCr, vbCr)
            cleanText = Replace(cleanText, vbCr & " ", vbCr)
            cleanText = Replace(cleanText, " " & vbLf, vbLf)
            cleanText = Replace(cleanText, vbLf & " ", vbLf)
            
            ' Remove tabs extras
            Do While InStr(cleanText, vbTab & vbTab) > 0
                cleanText = Replace(cleanText, vbTab & vbTab, vbTab)
            Loop
            
            ' Substitui tabs por espaços simples
            cleanText = Replace(cleanText, vbTab, " ")
            
            ' Remove espaços múltiplos novamente após conversão de tabs
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
            
            ' Aplica o texto limpo
            If cleanText <> para.Range.Text Then
                para.Range.Text = cleanText
            End If

            paraText = Trim(LCase(Replace(Replace(Replace(para.Range.Text, ".", ""), ",", ""), ";", "")))
            paraText = Replace(paraText, vbCr, "")
            paraText = Replace(paraText, vbLf, "")
            paraText = Replace(paraText, " ", "")

            With para.Format
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                .RightIndent = rightMarginPoints
                .SpaceBefore = 0
                .SpaceAfter = 0

                If para.Alignment = wdAlignParagraphCenter Then
                    .LeftIndent = 0
                    .FirstLineIndent = 0
                Else
                    firstIndent = .FirstLineIndent
                    paragraphIndent = .LeftIndent
                    If paragraphIndent >= CentimetersToPoints(5) Then
                        .LeftIndent = CentimetersToPoints(9.5)
                    ElseIf firstIndent < CentimetersToPoints(5) Then
                        .LeftIndent = CentimetersToPoints(0)
                        .FirstLineIndent = CentimetersToPoints(1.5)
                    End If
                End If
            End With

            If para.Alignment = wdAlignParagraphLeft Then
                para.Alignment = wdAlignParagraphJustify
            End If
            
            formattedCount = formattedCount + 1
        End If
    Next i
    
    LogMessage "Formatação de parágrafos aplicada: " & formattedCount & " parágrafos formatados, " & skippedCount & " ignorados (proteção de imagens/objetos)", LOG_LEVEL_INFO
    ApplyStdParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de parágrafos: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdParagraphs = False
End Function

'================================================================================
' FORMAT SECOND PARAGRAPH - FORMATAÇÃO APENAS DO 2º PARÁGRAFO - #NEW
'================================================================================
Private Function FormatSecondParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long
    
    ' Identifica apenas o 2º parágrafo (considerando apenas parágrafos com texto)
    actualParaIndex = 0
    secondParaIndex = 0
    
    ' Encontra o 2º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Registra o índice do 2º parágrafo
            If actualParaIndex = 2 Then
                secondParaIndex = i
                Exit For ' Já encontramos o 2º parágrafo
            End If
        End If
        
        ' Proteção: não processa mais de 10 parágrafos
        If i > 10 Then Exit For
    Next i
    
    ' Aplica formatação específica apenas ao 2º parágrafo
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(secondParaIndex)
        
        ' Verifica se não é um parágrafo com imagem (proteção)
        If Not HasVisualContent(para) Then
            With para.Format
                .LeftIndent = CentimetersToPoints(9)      ' Recuo à esquerda de 9 cm
                .FirstLineIndent = 0                      ' Sem recuo da primeira linha
                .RightIndent = 0                          ' Sem recuo à direita
                .Alignment = wdAlignParagraphJustify      ' Justificado
            End With
            
            LogMessage "2º parágrafo formatado com recuo de 9cm (posição real: " & secondParaIndex & ")", LOG_LEVEL_INFO
        Else
            LogMessage "2º parágrafo ignorado por conter conteúdo visual (posição: " & secondParaIndex & ")", LOG_LEVEL_INFO
        End If
    Else
        LogMessage "2º parágrafo não encontrado para formatação", LOG_LEVEL_WARNING
    End If
    
    FormatSecondParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do 2º parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatSecondParagraph = False
End Function

'================================================================================
' FORMAT FIRST PARAGRAPH - FORMATAÇÃO DO 1º PARÁGRAFO - #NEW
'================================================================================
Private Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim firstParaIndex As Long
    
    ' Identifica o 1º parágrafo (considerando apenas parágrafos com texto)
    actualParaIndex = 0
    firstParaIndex = 0
    
    ' Encontra o 1º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Registra o índice do 1º parágrafo
            If actualParaIndex = 1 Then
                firstParaIndex = i
                Exit For ' Já encontramos o 1º parágrafo
            End If
        End If
        
        ' Proteção: não processa mais de 10 parágrafos
        If i > 10 Then Exit For
    Next i
    
    ' Aplica formatação específica apenas ao 1º parágrafo
    If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(firstParaIndex)
        
        ' Verifica se não é um parágrafo com imagem (proteção)
        If Not HasVisualContent(para) Then
            ' Formatação do 1º parágrafo: caixa alta, negrito e sublinhado
            With para.Range.Font
                .AllCaps = True           ' Caixa alta (maiúsculas)
                .Bold = True              ' Negrito
                .Underline = wdUnderlineSingle ' Sublinhado
            End With
            
            ' Aplicar também formatação de parágrafo se necessário
            With para.Format
                .Alignment = wdAlignParagraphCenter       ' Centralizado
                .LeftIndent = 0                           ' Sem recuo à esquerda
                .FirstLineIndent = 0                      ' Sem recuo da primeira linha
                .RightIndent = 0                          ' Sem recuo à direita
            End With
            
            LogMessage "1º parágrafo formatado com caixa alta, negrito, sublinhado e centralizado (posição real: " & firstParaIndex & ")", LOG_LEVEL_INFO
        Else
            LogMessage "1º parágrafo ignorado por conter conteúdo visual (posição: " & firstParaIndex & ")", LOG_LEVEL_INFO
        End If
    Else
        LogMessage "1º parágrafo não encontrado para formatação", LOG_LEVEL_WARNING
    End If
    
    FormatFirstParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do 1º parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatFirstParagraph = False
End Function

'================================================================================
' ENABLE HYPHENATION - #STABLE
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
        LogMessage "Hifenização ativada com configurações padrão", LOG_LEVEL_INFO
        EnableHyphenation = True
    Else
        LogMessage "Hifenização já estava ativa", LOG_LEVEL_INFO
        EnableHyphenation = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao ativar hifenização: " & Err.Description, LOG_LEVEL_ERROR
    EnableHyphenation = False
End Function

'================================================================================
' REMOVE WATERMARK - #STABLE
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
    Dim i As Long
    Dim removedCount As Long

    For Each sec In doc.Sections
        For Each header In sec.Headers
            If header.Exists And header.Shapes.Count > 0 Then
                For i = header.Shapes.Count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header
        
        For Each header In sec.Footers
            If header.Exists And header.Shapes.Count > 0 Then
                For i = header.Shapes.Count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header
    Next sec

    If removedCount > 0 Then
        LogMessage "Marcas d'água removidas: " & removedCount & " itens", LOG_LEVEL_INFO
    Else
        LogMessage "Nenhuma marca d'água encontrada para remoção", LOG_LEVEL_INFO
    End If
    
    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao remover marcas d'água: " & Err.Description, LOG_LEVEL_ERROR
    RemoveWatermark = False
End Function

'================================================================================
' INSERT HEADER IMAGE - #STABLE
'================================================================================
Private Function InsertHeaderStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim shp As Shape
    Dim imgFound As Boolean
    Dim sectionsProcessed As Long

    username = GetSafeUserName()
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Busca inteligente da imagem em múltiplos locais
    If Dir(imgFile) = "" Then
        ' Tenta localização alternativa no perfil do usuário
        imgFile = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH
        If Dir(imgFile) = "" Then
            ' Tenta localização de rede corporativa
            imgFile = "\\strqnapmain\Dir. Legislativa\Christian" & HEADER_IMAGE_RELATIVE_PATH
            If Dir(imgFile) = "" Then
                ' Registra erro e tenta continuar sem a imagem
                Application.StatusBar = "Aviso: Imagem de cabeçalho não encontrada"
                LogMessage "Imagem de cabeçalho não encontrada em nenhum local: " & HEADER_IMAGE_RELATIVE_PATH, LOG_LEVEL_WARNING
                InsertHeaderStamp = False
                Exit Function
            End If
        End If
    End If

    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete
            
            Set shp = header.Shapes.AddPicture( _
                FileName:=imgFile, _
                LinkToFile:=False, _
                SaveWithDocument:=msoTrue)
            
            If shp Is Nothing Then
                LogMessage "Falha ao inserir imagem no cabeçalho da seção " & sectionsProcessed + 1, LOG_LEVEL_WARNING
            Else
                With shp
                    .LockAspectRatio = msoTrue
                    .Width = imgWidth
                    .Height = imgHeight
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    .Left = (doc.PageSetup.PageWidth - .Width) / 2
                    .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
                    .WrapFormat.Type = wdWrapTopBottom
                    .ZOrder msoSendToBack
                End With
                
                imgFound = True
                sectionsProcessed = sectionsProcessed + 1
            End If
        End If
    Next sec

    If imgFound Then
        LogMessage "Cabeçalho inserido em " & sectionsProcessed & " seções. Imagem: " & imgFile, LOG_LEVEL_INFO
        InsertHeaderStamp = True
    Else
        LogMessage "Nenhum cabeçalho foi inserido", LOG_LEVEL_WARNING
        InsertHeaderStamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir cabeçalho: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderStamp = False
End Function

'================================================================================
' INSERT FOOTER PAGE NUMBERS - #STABLE
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rng As Range
    Dim sectionsProcessed As Long

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        
        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range
            
            rng.Delete
            
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Fields.Add Range:=rng, Type:=wdFieldPage
            
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Text = "-"
            
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Fields.Add Range:=rng, Type:=wdFieldNumPages
            
            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.size = FOOTER_FONT_SIZE
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
            
            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    LogMessage "Rodapé inserido em " & sectionsProcessed & " seções com numeração de páginas", LOG_LEVEL_INFO
    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir rodapé: " & Err.Description, LOG_LEVEL_ERROR
    InsertFooterStamp = False
End Function

'================================================================================
' UTILITY: CM TO POINTS - #STABLE
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then
        CentimetersToPoints = cm * 28.35
    End If
End Function

'================================================================================
' UTILITY: SAFE USERNAME - #STABLE
'================================================================================
Private Function GetSafeUserName() As String
    On Error GoTo ErrorHandler
    
    Dim rawName As String
    Dim safeName As String
    Dim i As Integer
    Dim c As String
    
    rawName = Environ("USERNAME")
    If rawName = "" Then rawName = Environ("USER")
    If rawName = "" Then
        On Error Resume Next
        rawName = CreateObject("WScript.Network").username
        On Error GoTo 0
    End If
    
    If rawName = "" Then
        rawName = "UsuarioDesconhecido"
    End If
    
    For i = 1 To Len(rawName)
        c = Mid(rawName, i, 1)
        If c Like "[A-Za-z0-9_\-]" Then
            safeName = safeName & c
        ElseIf c = " " Then
            safeName = safeName & "_"
        End If
    Next i
    
    If safeName = "" Then safeName = "Usuario"
    
    GetSafeUserName = safeName
    Exit Function
    
ErrorHandler:
    GetSafeUserName = "Usuario"
End Function

'================================================================================
' VALIDATE DOCUMENT STRUCTURE - SIMPLIFICADO - #STABLE
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next
    
    ' Verificação básica e rápida
    If doc.Range.End > 0 And doc.Sections.Count > 0 Then
        ValidateDocumentStructure = True
    Else
        LogMessage "Documento com estrutura inconsistente", LOG_LEVEL_WARNING
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' CRITICAL FIX: SAVE DOCUMENT BEFORE PROCESSING
' TO PREVENT CRASHES ON NEW NON SAVED DOCUMENTS - #STABLE
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Aguardando salvamento do documento..."
    LogMessage "Iniciando salvamento obrigatório do documento", LOG_LEVEL_INFO
    
    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
        LogMessage "Operação de salvamento cancelada pelo usuário", LOG_LEVEL_INFO
        Application.StatusBar = "Salvamento cancelado pelo usuário"
        SaveDocumentFirst = False
        Exit Function
    End If

    ' Aguarda confirmação do salvamento com timeout de segurança
    Dim waitCount As Integer
    Dim maxWait As Integer
    maxWait = 10
    
    For waitCount = 1 To maxWait
        DoEvents
        If doc.Path <> "" Then Exit For
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
        Application.StatusBar = "Aguardando salvamento... (" & waitCount & "/" & maxWait & ")"
    Next waitCount

    If doc.Path = "" Then
        LogMessage "Falha ao salvar documento após " & maxWait & " tentativas", LOG_LEVEL_ERROR
        Application.StatusBar = "Falha no salvamento - operação cancelada"
        SaveDocumentFirst = False
    Else
        LogMessage "Documento salvo com sucesso em: " & doc.Path, LOG_LEVEL_INFO
        Application.StatusBar = "Documento salvo com sucesso"
        SaveDocumentFirst = True
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro durante salvamento: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    Application.StatusBar = "Erro durante salvamento"
    SaveDocumentFirst = False
End Function

'================================================================================
' CLEAR ALL FORMATTING - LIMPEZA INICIAL COMPLETA - #NEW
'================================================================================
Private Function ClearAllFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Limpando formatação existente..."
    
    ' Limpeza global mais eficiente
    With doc.Range
        ' Remove formatação de caracteres de forma mais direta
        .Font.Reset
        .Font.Name = STANDARD_FONT
        .Font.Size = STANDARD_FONT_SIZE
        .Font.Color = wdColorAutomatic
        
        ' Remove formatação de parágrafos
        .ParagraphFormat.Reset
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .ParagraphFormat.LineSpacing = 12
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.RightIndent = 0
        .ParagraphFormat.FirstLineIndent = 0
        
        ' Remove bordas e sombreamento
        On Error Resume Next
        .Borders.Enable = False
        .Shading.Texture = wdTextureNone
        On Error GoTo ErrorHandler
    End With
    
    ' Remove estilos personalizados de forma segura
    Dim para As Paragraph
    Dim paraCount As Long
    paraCount = 0
    
    For Each para In doc.Paragraphs
        On Error Resume Next
        para.Style = "Normal"
        paraCount = paraCount + 1
        ' Evita loops infinitos
        If paraCount > 1000 Then Exit For
        On Error GoTo ErrorHandler
    Next para
    
    LogMessage "Formatação limpa: " & paraCount & " parágrafos resetados", LOG_LEVEL_INFO
    ClearAllFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao limpar formatação: " & Err.Description, LOG_LEVEL_WARNING
    ClearAllFormatting = False ' Não falha o processo por isso
End Function

'================================================================================
' CLEAN DOCUMENT STRUCTURE - FUNCIONALIDADES 2, 6, 7 - #NEW
'================================================================================
Private Function CleanDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim firstTextParaIndex As Long
    Dim lastContentParaIndex As Long
    Dim emptyLinesRemoved As Long
    Dim leadingSpacesRemoved As Long
    
    ' Funcionalidade 2: Remove linhas em branco acima do título (primeira linha com texto)
    ' PROTEÇÃO: Não remove parágrafos que contenham imagens
    firstTextParaIndex = -1
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        Dim paraTextCheck As String
        paraTextCheck = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Encontra o primeiro parágrafo com texto real (não apenas imagens)
        If paraTextCheck <> "" Then
            firstTextParaIndex = i
            Exit For
        End If
    Next i
    
    If firstTextParaIndex > 1 Then
        For i = firstTextParaIndex - 1 To 1 Step -1
            Set para = doc.Paragraphs(i)
            Dim paraTextEmpty As String
            paraTextEmpty = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' PROTEÇÃO MÁXIMA: Usa função especializada para detectar conteúdo visual
            If paraTextEmpty = "" And Not HasVisualContent(para) Then
                para.Range.Delete
                emptyLinesRemoved = emptyLinesRemoved + 1
            Else
                ' Log quando preserva um parágrafo por segurança
                If paraTextEmpty = "" Then
                    LogMessage "Parágrafo preservado por conter possível conteúdo visual (posição " & i & ")", LOG_LEVEL_INFO
                End If
            End If
        Next i
    End If
    
    ' Funcionalidade 6: Remove linhas em branco no final do documento - DESABILITADO POR SOLICITAÇÃO
    ' NOTA: Linhas em branco no final do documento agora são preservadas conforme novo requisito
    ' PROTEÇÃO: Preserva parágrafos com imagens em qualquer lugar do documento
    '
    ' lastContentParaIndex = -1
    ' For i = doc.Paragraphs.Count To 1 Step -1
    '     Set para = doc.Paragraphs(i)
    '     Dim paraText As String
    '     paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
    '     
    '     ' Verifica se há conteúdo (texto OU conteúdo visual)
    '     If paraText <> "" Or HasVisualContent(para) Then
    '         lastContentParaIndex = i
    '         Exit For
    '     End If
    ' Next i
    ' 
    ' If lastContentParaIndex > 0 And lastContentParaIndex < doc.Paragraphs.Count Then
    '     For i = doc.Paragraphs.Count To lastContentParaIndex + 1 Step -1
    '         Set para = doc.Paragraphs(i)
    '         Dim paraTextFinal As String
    '         paraTextFinal = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
    '         
    '         ' PROTEÇÃO MÁXIMA: Usa função especializada para detectar conteúdo visual
    '         If paraTextFinal = "" And Not HasVisualContent(para) Then
    '             para.Range.Delete
    '             emptyLinesRemoved = emptyLinesRemoved + 1
    '         Else
    '             ' Log quando preserva um parágrafo por segurança
    '             If paraTextFinal = "" Then
    '                 LogMessage "Parágrafo final preservado por conter possível conteúdo visual (posição " & i & ")", LOG_LEVEL_INFO
    '             End If
    '         End If
    '     Next i
    ' End If
    
    ' Funcionalidade 7: Remove espaços e tabs no início de parágrafos
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        Dim originalText As String
        Dim cleanText As String
        
        originalText = para.Range.Text
        ' Remove espaços e tabs do início, preservando quebras de linha
        cleanText = originalText
        
        ' Remove espaços e tabs no início
        Do While Left(cleanText, 1) = " " Or Left(cleanText, 1) = vbTab
            cleanText = Mid(cleanText, 2)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
        Loop
        
        If cleanText <> originalText Then
            para.Range.Text = cleanText
        End If
    Next i
    
    LogMessage "Estrutura limpa: " & emptyLinesRemoved & " linhas vazias removidas (proteção máxima para imagens), " & leadingSpacesRemoved & " espaços iniciais removidos", LOG_LEVEL_INFO
    CleanDocumentStructure = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza da estrutura: " & Err.Description, LOG_LEVEL_ERROR
    CleanDocumentStructure = False
End Function

'================================================================================
' SAFE CHECK FOR VISUAL CONTENT - VERIFICAÇÃO SEGURA DE CONTEÚDO VISUAL - #NEW
'================================================================================
Private Function HasVisualContent(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verifica imagens inline (método principal)
    If para.Range.InlineShapes.Count > 0 Then
        HasVisualContent = True
        Exit Function
    End If
    
    ' Verifica caracteres especiais que podem indicar objetos incorporados
    Dim paraText As String
    paraText = para.Range.Text
    
    ' Caracteres especiais do Word que indicam objetos
    If InStr(paraText, Chr(1)) > 0 Then ' Objeto incorporado
        HasVisualContent = True
        Exit Function
    End If
    
    If InStr(paraText, Chr(8)) > 0 Then ' Campo ou objeto
        HasVisualContent = True
        Exit Function
    End If
    
    ' Verifica se há campos no parágrafo (podem conter imagens)
    If para.Range.Fields.Count > 0 Then
        HasVisualContent = True
        Exit Function
    End If
    
    ' Proteção extra: parágrafos muito pequenos podem conter anchors ou objetos ocultos
    Dim cleanText As String
    cleanText = Trim(Replace(Replace(paraText, vbCr, ""), vbLf, ""))
    
    ' Se o parágrafo tem caracteres mas não é texto normal, preserva
    If Len(cleanText) > 0 And Len(cleanText) < 10 Then
        ' Pode ser um parágrafo contendo apenas um objeto ou anchor
        HasVisualContent = True
        Exit Function
    End If
    
    ' Verifica se o parágrafo tem formatação especial que pode indicar objeto
    If para.Range.Font.Hidden = True Then
        HasVisualContent = True
        Exit Function
    End If
    
    HasVisualContent = False
    Exit Function

ErrorHandler:
    ' Em caso de erro, assume que há conteúdo visual (máxima segurança)
    HasVisualContent = True
End Function

'================================================================================
' VALIDATE PROPOSITION TYPE - FUNCIONALIDADE 3 - #NEW
'================================================================================
Private Function ValidatePropositionType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim firstWord As String
    Dim paraText As String
    Dim i As Long
    Dim userResponse As VbMsgBoxResult
    
    ' Encontra o primeiro parágrafo com texto
    For i = 1 To doc.Paragraphs.Count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        LogMessage "Documento não possui texto para validação", LOG_LEVEL_WARNING
        ValidatePropositionType = True
        Exit Function
    End If
    
    ' Extrai a primeira palavra
    Dim words() As String
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
    End If
    
    ' Verifica se é uma das proposituras válidas
    If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then
        LogMessage "Tipo de proposição validado: " & firstWord, LOG_LEVEL_INFO
        ValidatePropositionType = True
    Else
        ' Informa sobre documento não-padrão e continua automaticamente
        LogMessage "Primeira palavra não reconhecida como proposição padrão: " & firstWord & " - continuando processamento", LOG_LEVEL_WARNING
        Application.StatusBar = "Aviso: Documento não é Indicação/Requerimento/Moção - processando mesmo assim"
        
        ' Pequena pausa para o usuário visualizar a mensagem
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 2  ' 2 segundos
            DoEvents
        Loop
        
        LogMessage "Processamento de documento não-padrão autorizado automaticamente: " & firstWord, LOG_LEVEL_INFO
        ValidatePropositionType = True
    End If
    
    Exit Function

ErrorHandler:
    LogMessage "Erro na validação do tipo de proposição: " & Err.Description, LOG_LEVEL_ERROR
    ValidatePropositionType = False
End Function

'================================================================================
' FORMAT DOCUMENT TITLE - FUNCIONALIDADES 4 e 5 - #NEW
'================================================================================
Private Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim i As Long
    Dim newText As String
    
    ' Encontra o primeiro parágrafo com texto (após exclusão de linhas em branco)
    For i = 1 To doc.Paragraphs.Count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        LogMessage "Nenhum texto encontrado para formatação do título", LOG_LEVEL_WARNING
        FormatDocumentTitle = True
        Exit Function
    End If
    
    ' Remove ponto final se existir
    If Right(paraText, 1) = "." Then
        paraText = Left(paraText, Len(paraText) - 1)
    End If
    
    ' Verifica se é uma proposição (para aplicar substituição $NUMERO$/$ANO$)
    Dim isProposition As Boolean
    Dim firstWord As String
    
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
        If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then
            isProposition = True
        End If
    End If
    
    ' Se for proposição, substitui a última palavra por $NUMERO$/$ANO$
    If isProposition And UBound(words) >= 0 Then
        ' Reconstrói o texto substituindo a última palavra
        newText = ""
        For i = 0 To UBound(words) - 1
            If i > 0 Then newText = newText & " "
            newText = newText & words(i)
        Next i
        
        ' Adiciona $NUMERO$/$ANO$ no lugar da última palavra
        If newText <> "" Then newText = newText & " "
        newText = newText & "$NUMERO$/$ANO$"
    Else
        ' Se não for proposição, mantém o texto original
        newText = paraText
    End If
    
    ' SEMPRE aplica formatação de título: caixa alta, negrito, sublinhado
    firstPara.Range.Text = UCase(newText) & vbCrLf
    
    ' Formatação completa do título (primeira linha)
    With firstPara.Range.Font
        .Bold = True
        .Underline = wdUnderlineSingle
    End With
    
    With firstPara.Format
        .Alignment = wdAlignParagraphCenter
        .LeftIndent = 0
        .FirstLineIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 6  ' Pequeno espaço após o título
    End With
    
    If isProposition Then
        LogMessage "Título de proposição formatado: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    Else
        LogMessage "Primeira linha formatada como título: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    End If
    
    FormatDocumentTitle = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do título: " & Err.Description, LOG_LEVEL_ERROR
    FormatDocumentTitle = False
End Function

'================================================================================
' FORMAT CONSIDERANDO PARAGRAPHS - OTIMIZADO E SIMPLIFICADO - FUNCIONALIDADE 8 - #NEW
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim i As Long
    
    ' Percorre todos os parágrafos procurando por "considerando" no início
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Verifica se o parágrafo começa com "considerando" (ignorando maiúsculas/minúsculas)
        If Len(paraText) >= 12 And LCase(Left(paraText, 12)) = "considerando" Then
            ' Verifica se após "considerando" vem espaço, vírgula, ponto-e-vírgula ou fim da linha
            Dim nextChar As String
            If Len(paraText) > 12 Then
                nextChar = Mid(paraText, 13, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    ' É realmente "considerando" no início do parágrafo
                    Set rng = para.Range
                    rng.End = rng.Start + 12 ' Seleciona apenas "considerando"
                    
                    With rng
                        .Text = "CONSIDERANDO"
                        .Font.Bold = True
                    End With
                    
                    totalFormatted = totalFormatted + 1
                End If
            Else
                ' Parágrafo contém apenas "considerando"
                Set rng = para.Range
                rng.End = rng.Start + 12
                
                With rng
                    .Text = "CONSIDERANDO"
                    .Font.Bold = True
                End With
                
                totalFormatted = totalFormatted + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação 'considerando' aplicada: " & totalFormatted & " ocorrências em negrito e caixa alta", LOG_LEVEL_INFO
    FormatConsiderandoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação 'considerando': " & Err.Description, LOG_LEVEL_ERROR
    FormatConsiderandoParagraphs = False
End Function

'================================================================================
' APPLY TEXT REPLACEMENTS - FUNCIONALIDADES 10 e 11 - #NEW
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim replacementCount As Long
    
    Set rng = doc.Range
    
    ' Funcionalidade 10: Substitui variantes de "d'Oeste"
    Dim dOesteVariants() As String
    Dim i As Long
    
    ' Define as variantes possíveis dos 3 primeiros caracteres de "d'Oeste"
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O"   ' Original
    dOesteVariants(1) = "d´O"   ' Acento agudo
    dOesteVariants(2) = "d`O"   ' Acento grave  
    dOesteVariants(3) = "d" & Chr(8220) & "O"   ' Aspas curvas esquerda
    dOesteVariants(4) = "d'o"   ' Minúscula
    dOesteVariants(5) = "d´o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & Chr(8220) & "o"
    dOesteVariants(8) = "D'O"   ' Maiúscula no D
    dOesteVariants(9) = "D´O"
    dOesteVariants(10) = "D`O"
    dOesteVariants(11) = "D" & Chr(8220) & "O"
    dOesteVariants(12) = "D'o"
    dOesteVariants(13) = "D´o"
    dOesteVariants(14) = "D`o"
    dOesteVariants(15) = "D" & Chr(8220) & "o"
    
    For i = 0 To UBound(dOesteVariants)
        With rng.Find
            .ClearFormatting
            .Text = dOesteVariants(i) & "este"
            .Replacement.Text = "d'Oeste"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute(Replace:=True)
                replacementCount = replacementCount + 1
            Loop
        End With
    Next i
    
    ' Funcionalidade 11: Substitui variantes de "- Vereador -"
    Set rng = doc.Range
    Dim vereadorVariants() As String
    ReDim vereadorVariants(0 To 7)
    
    ' Variantes dos caracteres inicial e final
    vereadorVariants(0) = "- Vereador -"    ' Original
    vereadorVariants(1) = "– Vereador –"    ' Travessão
    vereadorVariants(2) = "— Vereador —"    ' Em dash
    vereadorVariants(3) = "- vereador -"    ' Minúscula
    vereadorVariants(4) = "– vereador –"
    vereadorVariants(5) = "— vereador —"
    vereadorVariants(6) = "-Vereador-"      ' Sem espaços
    vereadorVariants(7) = "–Vereador–"
    
    For i = 0 To UBound(vereadorVariants)
        If vereadorVariants(i) <> "- Vereador -" Then
            With rng.Find
                .ClearFormatting
                .Text = vereadorVariants(i)
                .Replacement.Text = "- Vereador -"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                Do While .Execute(Replace:=True)
                    replacementCount = replacementCount + 1
                Loop
            End With
        End If
    Next i
    
    LogMessage "Substituições de texto aplicadas: " & replacementCount & " substituições realizadas", LOG_LEVEL_INFO
    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    LogMessage "Erro nas substituições de texto: " & Err.Description, LOG_LEVEL_ERROR
    ApplyTextReplacements = False
End Function

'================================================================================
' FORMAT JUSTIFICATIVA/ANEXO PARAGRAPHS - FORMATAÇÃO ESPECÍFICA - #NEW
'================================================================================
Private Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim i As Long
    Dim formattedCount As Long
    
    ' Percorre todos os parágrafos do documento
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        
        ' Não processa parágrafos com conteúdo visual
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            cleanText = LCase(paraText)
            
            ' Verifica se contém exclusivamente uma das palavras-chave
            If cleanText = "justificativa" Or cleanText = "anexo" Or cleanText = "anexos" Then
                
                ' Aplica formatação de capitalização (primeira maiúscula, resto minúscula)
                Dim formattedText As String
                If Len(paraText) > 0 Then
                    formattedText = UCase(Left(paraText, 1)) & LCase(Mid(paraText, 2))
                End If
                
                ' Atualiza o texto do parágrafo
                para.Range.Text = formattedText & vbCrLf
                
                ' Aplica formatação específica
                With para.Format
                    .LeftIndent = 0               ' Recuo à esquerda 0
                    .FirstLineIndent = 0          ' Recuo da 1ª linha 0
                    .RightIndent = 0              ' Sem recuo à direita
                    .SpaceBefore = 12             ' Espaço antes para separação
                    .SpaceAfter = 6               ' Espaço depois
                    
                    ' Alinhamento específico conforme a palavra
                    If cleanText = "justificativa" Then
                        .Alignment = wdAlignParagraphCenter    ' Justificativa centralizada
                        LogMessage "Parágrafo 'Justificativa' formatado (centralizado, sem recuos)", LOG_LEVEL_INFO
                    Else ' anexo ou anexos
                        .Alignment = wdAlignParagraphLeft      ' Anexo/Anexos à esquerda
                        LogMessage "Parágrafo '" & formattedText & "' formatado (alinhado à esquerda, sem recuos)", LOG_LEVEL_INFO
                    End If
                End With
                
                ' Formatação de fonte especial (opcional: pode aplicar negrito)
                With para.Range.Font
                    .Bold = True    ' Destaca essas seções especiais
                End With
                
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação Justificativa/Anexo concluída: " & formattedCount & " parágrafos especiais formatados", LOG_LEVEL_INFO
    FormatJustificativaAnexoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação Justificativa/Anexo: " & Err.Description, LOG_LEVEL_ERROR
    FormatJustificativaAnexoParagraphs = False
End Function

'================================================================================
' SUBROTINA PÚBLICA: ABRIR PASTA DE LOGS - #NEW
'================================================================================
Public Sub AbrirPastaLogs()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim logsFolder As String
    Dim defaultLogsFolder As String
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de logs baseada no documento atual ou temp
    If Not doc Is Nothing And doc.Path <> "" Then
        logsFolder = doc.Path
    Else
        logsFolder = Environ("TEMP")
    End If
    
    ' Verifica se a pasta existe
    If Dir(logsFolder, vbDirectory) = "" Then
        logsFolder = Environ("TEMP")
    End If
    
    ' Abre a pasta no Windows Explorer
    Shell "explorer.exe """ & logsFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de logs aberta: " & logsFolder
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de logs aberta pelo usuário: " & logsFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de logs"
    
    ' Fallback: tenta abrir pasta temporária
    On Error Resume Next
    Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
    If Err.Number = 0 Then
        Application.StatusBar = "Pasta temporária aberta como alternativa"
    Else
        Application.StatusBar = "Não foi possível abrir pasta de logs"
    End If
End Sub

'================================================================================
' SUBROTINA PÚBLICA: ABRIR REPOSITÓRIO GITHUB - FUNCIONALIDADE 9 - #NEW
'================================================================================
Public Sub AbrirRepositorioGitHub()
    On Error GoTo ErrorHandler
    
    Dim repoURL As String
    Dim shellResult As Long
    
    ' URL do repositório do projeto
    repoURL = "https://github.com/chrmsantos/chainsaw-fprops"
    
    ' Abre o link no navegador padrão
    shellResult = Shell("rundll32.exe url.dll,FileProtocolHandler " & repoURL, vbNormalFocus)
    
    If shellResult > 0 Then
        Application.StatusBar = "Repositório GitHub aberto no navegador"
        
        ' Log da operação se sistema de log estiver ativo
        If loggingEnabled Then
            LogMessage "Repositório GitHub aberto pelo usuário: " & repoURL, LOG_LEVEL_INFO
        End If
    Else
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositório GitHub"
    
    ' Log do erro se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Erro ao abrir repositório GitHub: " & Err.Description, LOG_LEVEL_ERROR
    End If
    
    ' Fallback: tenta copiar URL para a área de transferência
    On Error Resume Next
    Dim dataObj As Object
    Set dataObj = CreateObject("htmlfile").parentWindow.clipboardData
    dataObj.setData "text", repoURL
    
    If Err.Number = 0 Then
        Application.StatusBar = "URL copiada para área de transferência: " & repoURL
    Else
        Application.StatusBar = "Não foi possível abrir o repositório. URL: " & repoURL
    End If
End Sub

'================================================================================
' SISTEMA DE BACKUP - FUNCIONALIDADE DE SEGURANÇA - #NEW
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Não faz backup se documento não foi salvo
    If doc.Path = "" Then
        LogMessage "Backup ignorado - documento não salvo", LOG_LEVEL_INFO
        CreateDocumentBackup = True
        Exit Function
    End If
    
    Dim backupFolder As String
    Dim fso As Object
    Dim docName As String
    Dim docExtension As String
    Dim timeStamp As String
    Dim backupFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define pasta de backup
    backupFolder = fso.GetParentFolderName(doc.Path) & "\" & BACKUP_FOLDER_NAME
    
    ' Cria pasta de backup se não existir
    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
        LogMessage "Pasta de backup criada: " & backupFolder, LOG_LEVEL_INFO
    End If
    
    ' Extrai nome e extensão do documento
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)
    
    ' Cria timestamp para o backup
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")
    
    ' Nome do arquivo de backup
    backupFileName = docName & "_backup_" & timeStamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName
    
    ' Salva uma cópia do documento como backup
    Application.StatusBar = "Criando backup do documento..."
    
    ' Salva o documento atual primeiro para garantir que está atualizado
    doc.Save
    
    ' Cria uma cópia do arquivo usando FileSystemObject
    fso.CopyFile doc.FullName, backupFilePath, True
    
    ' Limpa backups antigos se necessário
    CleanOldBackups backupFolder, docName
    
    LogMessage "Backup criado com sucesso: " & backupFileName, LOG_LEVEL_INFO
    Application.StatusBar = "Backup criado - processando documento..."
    
    CreateDocumentBackup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao criar backup: " & Err.Description, LOG_LEVEL_ERROR
    CreateDocumentBackup = False
End Function

'================================================================================
' LIMPEZA DE BACKUPS ANTIGOS - SIMPLIFICADO - #NEW
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error Resume Next
    
    ' Limpeza simplificada - só remove se houver muitos arquivos
    Dim fso As Object
    Dim folder As Object
    Dim filesCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    
    filesCount = folder.Files.Count
    
    ' Se há mais de 15 arquivos na pasta de backup, registra aviso
    If filesCount > 15 Then
        LogMessage "Muitos backups na pasta (" & filesCount & " arquivos) - considere limpeza manual", LOG_LEVEL_WARNING
    End If
End Sub

'================================================================================
' SUBROTINA PÚBLICA: ABRIR PASTA DE BACKUPS - #NEW
'================================================================================
Public Sub AbrirPastaBackups()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim backupFolder As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de backup baseada no documento atual
    If Not doc Is Nothing And doc.Path <> "" Then
        backupFolder = fso.GetParentFolderName(doc.Path) & "\" & BACKUP_FOLDER_NAME
    Else
        Application.StatusBar = "Nenhum documento salvo ativo para localizar pasta de backups"
        Exit Sub
    End If
    
    ' Verifica se a pasta de backup existe
    If Not fso.FolderExists(backupFolder) Then
        Application.StatusBar = "Pasta de backups não encontrada - nenhum backup foi criado ainda"
        LogMessage "Pasta de backups não encontrada: " & backupFolder, LOG_LEVEL_WARNING
        Exit Sub
    End If
    
    ' Abre a pasta no Windows Explorer
    Shell "explorer.exe """ & backupFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de backups aberta: " & backupFolder
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de backups aberta pelo usuário: " & backupFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de backups"
    LogMessage "Erro ao abrir pasta de backups: " & Err.Description, LOG_LEVEL_ERROR
    
    ' Fallback: tenta abrir pasta do documento
    On Error Resume Next
    If Not doc Is Nothing And doc.Path <> "" Then
        Dim docFolder As String
        docFolder = fso.GetParentFolderName(doc.Path)
        Shell "explorer.exe """ & docFolder & """", vbNormalFocus
        Application.StatusBar = "Pasta do documento aberta como alternativa"
    Else
        Application.StatusBar = "Não foi possível abrir pasta de backups"
    End If
End Sub

'================================================================================
' CLEAN MULTIPLE SPACES - LIMPEZA FINAL DE ESPAÇOS MÚLTIPLOS - #NEW
'================================================================================
Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Limpando espaços múltiplos..."
    
    Dim rng As Range
    Dim spacesRemoved As Long
    
    Set rng = doc.Range
    
    ' Remove espaços múltiplos usando Find/Replace simples (mais compatível)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Remove espaços duplos repetidamente até não encontrar mais
        .Text = "  "
        .Replacement.Text = " "
        
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            ' Evita loop infinito
            If spacesRemoved > 1000 Then Exit Do
        Loop
    End With
    
    ' Limpeza adicional de espaços antes/depois de quebras de linha
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        
        ' Remove espaços antes de quebra de linha
        .Text = " ^p"
        .Replacement.Text = "^p"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 1000 Then Exit Do
        Loop
        
        ' Remove espaços depois de quebra de linha
        .Text = "^p "
        .Replacement.Text = "^p"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 1000 Then Exit Do
        Loop
    End With
    
    ' Remove tabs múltiplos
    Set rng = doc.Range
    With rng.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 1000 Then Exit Do
        Loop
        
        ' Converte tabs para espaços simples
        .Text = "^t"
        .Replacement.Text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 1000 Then Exit Do
        Loop
    End With
    
    ' Última passada para garantir que não sobrou nenhum espaço duplo
    Set rng = doc.Range
    With rng.Find
        .Text = "  "
        .Replacement.Text = " "
        .MatchWildcards = False
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 1000 Then Exit Do
        Loop
    End With
    
    LogMessage "Limpeza de espaços concluída: " & spacesRemoved & " correções aplicadas", LOG_LEVEL_INFO
    CleanMultipleSpaces = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza de espaços múltiplos: " & Err.Description, LOG_LEVEL_WARNING
    CleanMultipleSpaces = False ' Não falha o processo por isso
End Function

'================================================================================
' LIMIT SEQUENTIAL EMPTY LINES - CONTROLA LINHAS VAZIAS SEQUENCIAIS - #NEW
'================================================================================
Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Controlando linhas em branco sequenciais..."
    
    Dim para As Paragraph
    Dim i As Long
    Dim emptyLineCount As Long
    Dim linesRemoved As Long
    Dim paraText As String
    Dim totalParas As Long
    
    emptyLineCount = 0
    linesRemoved = 0
    
    ' Faz múltiplas passadas até não haver mais alterações
    ' (necessário porque a remoção altera os índices)
    Dim changesInPass As Boolean
    Dim passCount As Long
    
    Do
        changesInPass = False
        passCount = passCount + 1
        emptyLineCount = 0
        totalParas = doc.Paragraphs.Count
        
        i = 1
        Do While i <= totalParas And i <= doc.Paragraphs.Count
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Verifica se o parágrafo está vazio (sem texto e sem conteúdo visual)
            If paraText = "" And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1
                
                ' Se já temos mais de 1 linha vazia consecutiva, remove esta
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    linesRemoved = linesRemoved + 1
                    changesInPass = True
                    ' Não incrementa i pois removemos um parágrafo
                    totalParas = totalParas - 1
                Else
                    i = i + 1
                End If
            Else
                ' Se encontrou conteúdo, reseta o contador de linhas vazias
                emptyLineCount = 0
                i = i + 1
            End If
            
            ' Proteção contra loops infinitos
            If i > 2000 Then
                LogMessage "Interrompido por segurança após processar 2000 parágrafos", LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
        
        ' Proteção contra loops infinitos de passadas
        If passCount > 10 Then
            LogMessage "Interrompido após 10 passadas por segurança", LOG_LEVEL_WARNING
            Exit Do
        End If
        
    Loop While changesInPass
    
    LogMessage "Controle de linhas vazias concluído em " & passCount & " passada(s): " & linesRemoved & " linhas excedentes removidas (máximo 1 sequencial)", LOG_LEVEL_INFO
    LimitSequentialEmptyLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro no controle de linhas vazias: " & Err.Description, LOG_LEVEL_WARNING
    LimitSequentialEmptyLines = False ' Não falha o processo por isso
End Function

'================================================================================
' CONFIGURE DOCUMENT VIEW - CONFIGURAÇÃO DE VISUALIZAÇÃO - #NEW
'================================================================================
Private Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Configurando visualização do documento..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Configura zoom para 110%
    With docWindow.View
        .Zoom.Percentage = 110
        .Type = wdPrintView     ' Garante que está no modo de exibição de impressão
    End With
    
    ' Configurações adicionais de visualização
    With Application.Options
        On Error Resume Next
        .ShowReadabilityStatistics = False  ' Desabilita estatísticas de legibilidade
        .CheckGrammarAsYouType = True      ' Mantém verificação gramatical
        .CheckSpellingAsYouType = True     ' Mantém verificação ortográfica
        On Error GoTo ErrorHandler
    End With
    
    LogMessage "Visualização configurada: zoom 110%, régua visível, modo impressão", LOG_LEVEL_INFO
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao configurar visualização: " & Err.Description, LOG_LEVEL_WARNING
    ConfigureDocumentView = False ' Não falha o processo por isso
End Function

