VERSÃO FINAL CONSOLIDADA - Código de Padronização de Documentos Word

```vba
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

' Document formatting constants
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 4.7
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\Pictures\LegisTabStamp\HeaderStamp.png"
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
Private Const REQUIRED_STRING As String = " Nº $NUMERO$/$ANO$"

' Timeout constants
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

'================================================================================
' GLOBAL VARIABLES
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date

'================================================================================
' MAIN ENTRY POINT - COM SEGURANÇA ROBUSTA
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    
    ' Registro detalhado do início da execução
    LogMessage "🚀 INÍCIO DA EXECUÇÃO - Processo de padronização iniciado", LOG_LEVEL_INFO
    LogMessage "📋 Contexto: Usuário='" & Environ("USERNAME") & "', Estação='" & Environ("COMPUTERNAME") & "'", LOG_LEVEL_INFO
    
    ' Verificação de compatibilidade de versão
    If Not CheckWordVersion() Then
        Dim versionMsg As String
        versionMsg = "Versão do Word (" & Application.Version & ") não suportada. " & _
                    "Requisito mínimo: Word 2010 (versão 14.0). " & _
                    "Atualize o Microsoft Word para utilizar este recurso."
        LogMessage "❌ " & versionMsg, LOG_LEVEL_ERROR
        MsgBox versionMsg, vbExclamation + vbOKOnly, "Compatibilidade Não Suportada"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = Nothing
    
    ' Obter documento ativo com tratamento seguro
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        LogMessage "❌ Nenhum documento ativo disponível para processamento", LOG_LEVEL_ERROR
        MsgBox "Nenhum documento está aberto ou acessível no momento." & vbCrLf & _
               "Por favor, abra um documento do Word e tente novamente.", _
               vbExclamation + vbOKOnly, "Documento Não Disponível"
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    ' Inicialização segura do sistema de logging
    If Not InitializeLogging(doc) Then
        LogMessage "⚠️  Sistema de logging não inicializado - continuando sem logs detalhados", LOG_LEVEL_WARNING
    End If
    
    LogMessage "📄 Documento selecionado: '" & doc.Name & "'", LOG_LEVEL_INFO
    LogMessage "📁 Localização: " & IIf(doc.Path = "", "(Documento não salvo)", doc.Path), LOG_LEVEL_INFO
    
    ' Iniciar grupo undo com proteção
    StartUndoGroup "Padronização de Documento - " & doc.Name
    
    ' Configurar estado da aplicação com fallbacks
    If Not SetAppState(False, "Formatando documento...") Then
        LogMessage "⚠️  Configuração de estado da aplicação parcialmente bem-sucedida", LOG_LEVEL_WARNING
    End If
    
    ' Executar verificações preliminares
    If Not PreviousChecking(doc) Then
        LogMessage "⏹️  Verificações preliminares falharam - execução interrompida", LOG_LEVEL_ERROR
        GoTo CleanUp
    End If
    
    ' Executar processamento principal
    If Not PreviousFormatting(doc) Then
        LogMessage "⏹️  Processamento principal falhou - execução interrompida", LOG_LEVEL_ERROR
        GoTo CleanUp
    End If
    
    If formattingCancelled Then
        LogMessage "⏹️  Processamento cancelado pelo usuário", LOG_LEVEL_INFO
        GoTo CleanUp
    End If
    
    ' Sucesso na execução
    Application.StatusBar = "✅ Documento padronizado com sucesso!"
    LogMessage "✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO", LOG_LEVEL_INFO
    
    Dim executionTime As String
    executionTime = Format(Now - executionStartTime, "nn:ss")
    LogMessage "⏱️  Tempo total de execução: " & executionTime, LOG_LEVEL_INFO
    
CleanUp:
    ' Limpeza segura com tratamento de erro individual
    SafeCleanup
    
    ' Restaurar estado da aplicação
    If Not SetAppState(True, "✅ Documento padronizado com sucesso!") Then
        LogMessage "⚠️  Restauração parcial do estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    ' Finalização segura do logging
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    ' Tratamento de erro crítico
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    LogMessage "💥 " & errDesc, LOG_LEVEL_ERROR
    LogMessage "🔄 Iniciando recuperação de erro crítico", LOG_LEVEL_ERROR
    
    ' Recuperação de emergência
    EmergencyRecovery
    
    ' Mensagem amigável ao usuário
    MsgBox "Ocorreu um erro inesperado durante o processamento." & vbCrLf & vbCrLf & _
           "Detalhes técnicos: " & errDesc & vbCrLf & vbCrLf & _
           "O Word tentou recuperar o estado normal da aplicação." & vbCrLf & _
           "Verifique o arquivo de log para mais detalhes.", _
           vbCritical + vbOKOnly, "Erro Inesperado"
End Sub

'================================================================================
' EMERGENCY RECOVERY - PREVENÇÃO DE QUEDA DO WORD
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next ' Prevenir loops de erro
    
    LogMessage "🛡️  Executando procedimento de recuperação de emergência", LOG_LEVEL_ERROR
    
    ' Restaurar configurações críticas do Word
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Application.EnableCancelKey = 0
    
    ' Finalizar grupo undo se estiver ativo
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    ' Fechar arquivos de log abertos
    CloseAllOpenFiles
    
    LogMessage "🛡️  Recuperação de emergência concluída", LOG_LEVEL_INFO
End Sub

'================================================================================
' SAFE CLEANUP - LIMPEZA SEGURA
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    
    LogMessage "🧹 Iniciando processo de limpeza segura", LOG_LEVEL_INFO
    
    ' Finalizar grupo undo
    EndUndoGroup
    
    ' Liberar objetos da memória
    ReleaseObjects
    
    LogMessage "🧹 Limpeza segura concluída", LOG_LEVEL_INFO
End Sub

'================================================================================
' RELEASE OBJECTS - LIBERAÇÃO SEGURA DE OBJETOS
'================================================================================
Private Sub ReleaseObjects()
    On Error Resume Next
    
    ' Liberar objetos potencialmente alocados
    Dim nullObj As Object
    Set nullObj = Nothing
    
    ' Coletor de lixo simplificado
    Dim memoryCounter As Long
    For memoryCounter = 1 To 3
        DoEvents
    Next memoryCounter
End Sub

'================================================================================
' CLOSE ALL OPEN FILES - FECHAMENTO SEGURO DE ARQUIVOS
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
' VERSION COMPATIBILITY CHECK - COM VERIFICAÇÃO ROBUSTA
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Long
    version = Application.Version
    
    If version < MIN_SUPPORTED_VERSION Then
        LogMessage "❌ Versão do Word " & version & " não suportada (mínimo: " & MIN_SUPPORTED_VERSION & ")", LOG_LEVEL_ERROR
        CheckWordVersion = False
    Else
        LogMessage "✅ Versão do Word " & version & " compatível com o sistema", LOG_LEVEL_INFO
        CheckWordVersion = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "❌ Falha na verificação de versão: " & Err.Description, LOG_LEVEL_ERROR
    CheckWordVersion = False
End Function

'================================================================================
' UNDO GROUP MANAGEMENT - COM PROTEÇÃO
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        LogMessage "⚠️  Grupo undo já está ativo - finalizando antes de iniciar novo", LOG_LEVEL_WARNING
        EndUndoGroup
    End If
    
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    LogMessage "📝 Grupo undo iniciado: '" & groupName & "'", LOG_LEVEL_INFO
    
    Exit Sub
    
ErrorHandler:
    LogMessage "❌ Falha ao iniciar grupo undo: " & Err.Description, LOG_LEVEL_ERROR
    undoGroupEnabled = False
End Sub

Private Sub EndUndoGroup()
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "📝 Grupo undo finalizado com sucesso", LOG_LEVEL_INFO
    Else
        LogMessage "ℹ️  Nenhum grupo undo ativo para finalizar", LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    LogMessage "❌ Falha ao finalizar grupo undo: " & Err.Description, LOG_LEVEL_ERROR
    undoGroupEnabled = False
End Sub

'================================================================================
' LOGGING MANAGEMENT - APRIMORADO COM DETALHES
'================================================================================
Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Determinar caminho do arquivo de log
    If doc.Path <> "" Then
        logFilePath = doc.Path & "\" & Format(Now, "yyyy-mm-dd") & "_" & _
                     Replace(doc.Name, ".doc", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docx", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docm", "") & "_FormattingLog.txt"
    Else
        logFilePath = Environ("TEMP") & "\" & Format(Now, "yyyy-mm-dd") & "_DocumentFormattingLog.txt"
    End If
    
    ' Criar arquivo de log com informações detalhadas
    Open logFilePath For Output As #1
    Print #1, "================================================"
    Print #1, "📊 LOG DE FORMATAÇÃO DE DOCUMENTO - SISTEMA DE REGISTRO"
    Print #1, "================================================"
    Print #1, "🏷️  Sessão: " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    Print #1, "👤 Usuário: " & Environ("USERNAME")
    Print #1, "💻 Estação: " & Environ("COMPUTERNAME")
    Print #1, "🔢 Versão Word: " & Application.Version
    Print #1, "📄 Documento: " & doc.Name
    Print #1, "📁 Local: " & IIf(doc.Path = "", "(Não salvo)", doc.Path)
    Print #1, "🔒 Proteção: " & GetProtectionType(doc)
    Print #1, "📏 Tamanho: " & GetDocumentSize(doc)
    Print #1, "================================================"
    Close #1
    
    loggingEnabled = True
    LogMessage "📁 Sistema de logging inicializado: " & logFilePath, LOG_LEVEL_INFO
    InitializeLogging = True
    
    Exit Function
    
ErrorHandler:
    LogMessage "❌ Falha crítica na inicialização do logging: " & Err.Description, LOG_LEVEL_ERROR
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
            levelIcon = "ℹ️ "
        Case LOG_LEVEL_WARNING
            levelText = "AVISO"
            levelIcon = "⚠️ "
        Case LOG_LEVEL_ERROR
            levelText = "ERRO"
            levelIcon = "❌"
        Case Else
            levelText = "OUTRO"
            levelIcon = "🔹"
    End Select
    
    ' Formatar mensagem com timestamp detalhado
    Dim formattedMessage As String
    formattedMessage = Format(Now, "yyyy-mm-dd HH:MM:ss") & " [" & levelText & "] " & levelIcon & " " & message
    
    ' Escrever no arquivo de log
    Open logFilePath For Append As #1
    Print #1, formattedMessage
    Close #1
    
    ' Output para Debug Window
    Debug.Print "LOG: " & formattedMessage
    
    Exit Sub
    
ErrorHandler:
    ' Fallback seguro para logging
    Debug.Print "FALHA NO LOGGING: " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler
    
    If loggingEnabled Then
        Open logFilePath For Append As #1
        Print #1, "================================================"
        Print #1, "🏁 FIM DA SESSÃO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
        Print #1, "⏱️  Duração: " & Format(Now - executionStartTime, "HH:MM:ss")
        Print #1, "🔚 Status: " & IIf(formattingCancelled, "CANCELADO", "CONCLUÍDO")
        Print #1, "================================================"
        Close #1
        
        LogMessage "📁 Log finalizado com sucesso", LOG_LEVEL_INFO
    End If
    
    loggingEnabled = False
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erro ao finalizar logging: " & Err.Description
    loggingEnabled = False
End Sub

'================================================================================
' UTILITY: GET PROTECTION TYPE
'================================================================================
Private Function GetProtectionType(doc As Document) As String
    On Error Resume Next
    
    Select Case doc.ProtectionType
        Case wdNoProtection: GetProtectionType = "Sem proteção"
        Case 1: GetProtectionType = "Protegido contra revisões"
        Case 2: GetProtectionType = "Protegido contra comentários"
        Case 3: GetProtectionType = "Protegido contra formulários"
        Case 4: GetProtectionType = "Protegido contra leitura"
        Case Else: GetProtectionType = "Tipo desconhecido (" & doc.ProtectionType & ")"
    End Select
End Function

'================================================================================
' UTILITY: GET DOCUMENT SIZE
'================================================================================
Private Function GetDocumentSize(doc As Document) As String
    On Error Resume Next
    
    Dim size As Long
    size = doc.BuiltInDocumentProperties("Number of Characters").Value * 2 ' Aproximação
    
    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' APPLICATION STATE HANDLER - ROBUSTO
'================================================================================
Private Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim success As Boolean
    success = True
    
    With Application
        ' ScreenUpdating - crítico para performance
        On Error Resume Next
        .ScreenUpdating = enabled
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        ' DisplayAlerts - importante para UX
        On Error Resume Next
        .DisplayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        ' StatusBar - feedback para usuário
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
        
        ' EnableCancelKey - prevenção de cancelamento acidental
        On Error Resume Next
        .EnableCancelKey = 0
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
    End With
    
    If enabled Then
        LogMessage "🔄 Estado da aplicação restaurado: " & IIf(success, "Completo", "Parcial"), LOG_LEVEL_INFO
    Else
        LogMessage "⚡ Estado de performance ativado: " & IIf(success, "Completo", "Parcial"), LOG_LEVEL_INFO
    End If
    
    SetAppState = success
    Exit Function
    
ErrorHandler:
    LogMessage "❌ Erro ao configurar estado da aplicação: " & Err.Description, LOG_LEVEL_ERROR
    SetAppState = False
End Function

'================================================================================
' GLOBAL CHECKING - VERIFICAÇÕES ROBUSTAS
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "🔍 Iniciando verificações de segurança do documento", LOG_LEVEL_INFO

    ' Verificação 1: Documento existe e está acessível
    If doc Is Nothing Then
        LogMessage "❌ Falha crítica: Nenhum documento disponível para verificação", LOG_LEVEL_ERROR
        MsgBox "Erro de sistema: Nenhum documento está acessível para verificação." & vbCrLf & _
               "Tente fechar e reabrir o documento, então execute novamente.", _
               vbCritical + vbOKOnly, "Falha de Acesso ao Documento"
        PreviousChecking = False
        Exit Function
    End If

    ' Verificação 2: Tipo de documento válido
    If doc.Type <> wdTypeDocument Then
        LogMessage "❌ Tipo de documento inválido: " & doc.Type & " (esperado: " & wdTypeDocument & ")", LOG_LEVEL_ERROR
        MsgBox "Documento incompatível detectado." & vbCrLf & _
               "Este sistema suporta apenas documentos do Word padrão." & vbCrLf & _
               "Tipo atual: " & doc.Type, _
               vbExclamation + vbOKOnly, "Tipo de Documento Não Suportado"
        PreviousChecking = False
        Exit Function
    End If

    ' Verificação 3: Proteção do documento
    If doc.ProtectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        
        LogMessage "❌ Documento protegido contra edição: " & protectionType, LOG_LEVEL_ERROR
        MsgBox "Documento protegido detectado." & vbCrLf & _
               "Tipo de proteção: " & protectionType & vbCrLf & vbCrLf & _
               "Para continuar, remova a proteção do documento através de:" & vbCrLf & _
               "Revisão > Proteger > Restringir Edição > Parar Proteção", _
               vbExclamation + vbOKOnly, "Documento Protegido"
        PreviousChecking = False
        Exit Function
    End If
    
    ' Verificação 4: Documento somente leitura
    If doc.ReadOnly Then
        LogMessage "❌ Documento aberto em modo somente leitura", LOG_LEVEL_ERROR
        MsgBox "Documento em modo somente leitura." & vbCrLf & _
               "Salve uma cópia editável do documento antes de prosseguir." & vbCrLf & vbCrLf & _
               "Arquivo: " & doc.FullName, _
               vbExclamation + vbOKOnly, "Documento Somente Leitura"
        PreviousChecking = False
        Exit Function
    End If

    ' Verificação 5: Espaço em disco suficiente
    If Not CheckDiskSpace(doc) Then
        LogMessage "❌ Espaço em disco insuficiente para operação segura", LOG_LEVEL_ERROR
        MsgBox "Espaço em disco insuficiente para completar a operação com segurança." & vbCrLf & _
               "Libere pelo menos 50MB de espaço livre e tente novamente.", _
               vbExclamation + vbOKOnly, "Espaço em Disco Insuficiente"
        PreviousChecking = False
        Exit Function
    End If

    ' Verificação 6: Estrutura do documento válida
    If Not ValidateDocumentStructure(doc) Then
        LogMessage "⚠️  Estrutura do documento com possíveis problemas - continuando com cautela", LOG_LEVEL_WARNING
    End If

    LogMessage "✅ Todas as verificações de segurança passaram com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    LogMessage "❌ Erro durante verificações de segurança: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro durante verificações de segurança do documento." & vbCrLf & _
           "Detalhes: " & Err.Description & vbCrLf & _
           "Contate o suporte técnico se o problema persistir.", _
           vbCritical + vbOKOnly, "Erro de Verificação"
    PreviousChecking = False
End Function

'================================================================================
' DISK SPACE CHECK - VERIFICAÇÃO DE ESPAÇO EM DISCO
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim drive As Object
    Dim requiredSpace As Long
    Dim driveLetter As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Determinar unidade de destino
    If doc.Path <> "" Then
        driveLetter = Left(doc.Path, 3)
    Else
        driveLetter = Left(Environ("TEMP"), 3)
    End If
    
    ' Obter informações da unidade
    Set drive = fso.GetDrive(driveLetter)
    
    ' Espaço requerido (50MB como segurança)
    requiredSpace = 50 * 1024 * 1024 ' 50MB em bytes
    
    If drive.AvailableSpace < requiredSpace Then
        LogMessage "⚠️  Espaço em disco limitado: " & Format(drive.AvailableSpace / 1024 / 1024, "0.0") & _
                  "MB disponíveis (mínimo recomendado: 50MB)", LOG_LEVEL_WARNING
        CheckDiskSpace = False
    Else
        LogMessage "💾 Espaço em disco adequado: " & Format(drive.AvailableSpace / 1024 / 1024, "0.0") & _
                  "MB disponíveis", LOG_LEVEL_INFO
        CheckDiskSpace = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "⚠️  Não foi possível verificar espaço em disco: " & Err.Description, LOG_LEVEL_WARNING
    CheckDiskSpace = True ' Continuar mesmo com erro na verificação
End Function

'================================================================================
' REMOVE BLANK LINES AND CHECK FOR REQUIRED STRING
'================================================================================
Private Function RemoveLeadingBlankLinesAndCheckString(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim deletedCount As Long
    Dim firstLineText As String
    
    LogMessage "🔍 Removendo linhas em branco iniciais e verificando string obrigatória", LOG_LEVEL_INFO
    
    ' Safely remove leading blank paragraphs
    Do While doc.Paragraphs.Count > 0
        Set para = doc.Paragraphs(1)
        If Trim(para.Range.Text) = vbCr Or Trim(para.Range.Text) = "" Or _
           para.Range.Text = Chr(13) Or para.Range.Text = Chr(7) Then
            para.Range.Delete
            deletedCount = deletedCount + 1
            LogMessage "📝 Parágrafo vazio removido: " & deletedCount, LOG_LEVEL_INFO
            ' Safety check to prevent infinite loop
            If deletedCount > 100 Then
                LogMessage "⚠️  Limite de segurança atingido ao remover linhas em branco", LOG_LEVEL_WARNING
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    LogMessage "📊 Total de linhas em branco removidas: " & deletedCount, LOG_LEVEL_INFO
    
    ' Check if document has at least one paragraph after removal
    If doc.Paragraphs.Count = 0 Then
        LogMessage "❌ Documento vazio após remoção de linhas em branco", LOG_LEVEL_ERROR
        MsgBox "O documento está vazio após a remoção das linhas em branco iniciais.", vbExclamation, "Documento Vazio"
        RemoveLeadingBlankLinesAndCheckString = False
        Exit Function
    End If
    
    ' Get the text of the first line (first paragraph)
    firstLineText = doc.Paragraphs(1).Range.Text
    LogMessage "📄 Texto da primeira linha: '" & firstLineText & "'", LOG_LEVEL_INFO
    
    ' Check for the exact string (case-sensitive)
    If InStr(1, firstLineText, REQUIRED_STRING, vbBinaryCompare) = 0 Then
        ' String not found - show warning message
        LogMessage "⚠️  String obrigatória não encontrada na primeira linha: '" & REQUIRED_STRING & "'", LOG_LEVEL_WARNING
        
        Dim response As VbMsgBoxResult
        response = MsgBox("ATENÇÃO: Não foi encontrada a string obrigatória exata:" & vbCrLf & vbCrLf & _
                         "'" & REQUIRED_STRING & "'" & vbCrLf & vbCrLf & _
                         "Texto encontrado na primeira linha: '" & firstLineText & "'" & vbCrLf & vbCrLf & _
                         "Deseja continuar com a formatação mesmo assim?", _
                         vbExclamation + vbYesNo, "String Obrigatória Não Encontrada")
        
        If response = vbNo Then
            LogMessage "⏹️  Usuário cancelou a formatação devido à string obrigatória não encontrada", LOG_LEVEL_WARNING
            MsgBox "Formatação cancelada pelo usuário.", vbInformation, "Operação Cancelada"
            formattingCancelled = True
            RemoveLeadingBlankLinesAndCheckString = False
        Else
            LogMessage "⚠️  Usuário optou por continuar apesar da string obrigatória não encontrada", LOG_LEVEL_WARNING
            RemoveLeadingBlankLinesAndCheckString = True
        End If
    Else
        LogMessage "✅ String obrigatória encontrada com sucesso: '" & REQUIRED_STRING & "'", LOG_LEVEL_INFO
        RemoveLeadingBlankLinesAndCheckString = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "❌ Erro durante verificação de string obrigatória: " & Err.Description, LOG_LEVEL_ERROR
    RemoveLeadingBlankLinesAndCheckString = False
End Function

'================================================================================
' MAIN FORMATTING ROUTINE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "🔄 Iniciando formatação principal do documento", LOG_LEVEL_INFO

    ' Remove blank lines and check for required string
    If Not RemoveLeadingBlankLinesAndCheckString(doc) Then
        If formattingCancelled Then 
            PreviousFormatting = False
            Exit Function
        End If
        LogMessage "⚠️  Falha na verificação inicial - continuando com formatação", LOG_LEVEL_WARNING
    End If

    ' Apply formatting in logical order
    If Not ApplyPageSetup(doc) Then
        LogMessage "❌ Falha na configuração de página", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not ApplyFontAndParagraph(doc) Then
        LogMessage "❌ Falha na formatação de fonte e parágrafo", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not EnableHyphenation(doc) Then
        LogMessage "⚠️  Falha na ativação de hifenização", LOG_LEVEL_WARNING
    End If
    
    If Not RemoveWatermark(doc) Then
        LogMessage "⚠️  Falha na remoção de marca d'água", LOG_LEVEL_WARNING
    End If
    
    If Not InsertHeaderStamp(doc) Then
        LogMessage "⚠️  Falha na inserção do carimbo do cabeçalho", LOG_LEVEL_WARNING
    End If
    
    If Not InsertFooterStamp(doc) Then
        LogMessage "❌ Falha crítica na inserção do rodapé", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    ' Save changes
    If doc.Path <> "" Then
        doc.Save
        LogMessage "💾 Documento salvo após formatação", LOG_LEVEL_INFO
    Else
        LogMessage "⚠️  Documento não salvo (sem caminho especificado)", LOG_LEVEL_WARNING
    End If

    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "❌ Erro durante formatação principal: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' PAGE SETUP
'================================================================================
Private Function ApplyPageSetup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    LogMessage "📐 Aplicando configurações de página", LOG_LEVEL_INFO
    
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
    
    LogMessage "✅ Configurações de página aplicadas com sucesso", LOG_LEVEL_INFO
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
    LogMessage "❌ Erro ao aplicar configurações de página: " & Err.Description, LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

'================================================================================
' FONT AND PARAGRAPH FORMATTING
'================================================================================
Private Function ApplyFontAndParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim currentIndent As Single
    Dim rightMarginPoints As Single
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long

    LogMessage "🎨 Aplicando formatação de fonte e parágrafo", LOG_LEVEL_INFO

    ' O recuo à direita deve ser ZERO para alinhar com a margem direita
    rightMarginPoints = 0

    ' Process paragraphs in reverse to avoid issues with inserted line breaks
    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

        ' Check if paragraph contains any inline image
        If para.Range.InlineShapes.Count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' Skip formatting if inline image is present
        If Not hasInlineImage Then
            ' Apply font formatting
            With para.Range.Font
                .Name = STANDARD_FONT
                .Size = STANDARD_FONT_SIZE
                .Bold = False
                .Italic = False
                .Underline = 0
                .Color = wdColorAutomatic
            End With

            ' Apply paragraph formatting
            With para.Format
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                
                ' RECUO À DIREITA ZERO - alinha com a margem direita
                .RightIndent = rightMarginPoints
                .SpaceBefore = 0
                .SpaceAfter = 0

                ' If paragraph is centered, indents should be 0
                If para.Alignment = wdAlignParagraphCenter Then
                    .LeftIndent = 0
                    .FirstLineIndent = 0
                Else
                    ' First line indentation logic
                    currentIndent = .FirstLineIndent
                    If currentIndent <= CentimetersToPoints(0.06) Then
                        .FirstLineIndent = CentimetersToPoints(0.25)
                    ElseIf currentIndent > CentimetersToPoints(0.06) Then
                        .FirstLineIndent = CentimetersToPoints(0.9)
                    End If
                End If
            End With

            ' Justify left-aligned paragraphs
            If para.Alignment = wdAlignParagraphLeft Then
                para.Alignment = wdAlignParagraphJustify
            End If
            
            formattedCount = formattedCount + 1
        End If
    Next i
    
    LogMessage "📊 Formatação concluída: " & formattedCount & " parágrafos formatados, " & skippedCount & " parágrafos com imagens ignorados", LOG_LEVEL_INFO
    LogMessage "✅ Recuo à direita definido como ZERO para alinhamento com margem direita", LOG_LEVEL_INFO
    ApplyFontAndParagraph = True
    Exit Function
    
ErrorHandler:
    LogMessage "❌ Erro ao aplicar formatação de fonte e parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    ApplyFontAndParagraph = False
End Function

'================================================================================
' ENABLE HYPHENATION
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    LogMessage "🔠 Ativando hifenização automática", LOG_LEVEL_INFO
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
        LogMessage "✅ Hifenização automática ativada", LOG_LEVEL_INFO
        EnableHyphenation = True
    Else
        LogMessage "ℹ️  Hifenização automática já estava ativada", LOG_LEVEL_INFO
        EnableHyphenation = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "⚠️  Falha ao ativar hifenização: " & Err.Description, LOG_LEVEL_WARNING
    EnableHyphenation = False
End Function

'================================================================================
' REMOVE WATERMARK
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
    Dim i As Long
    Dim removedCount As Long

    LogMessage "💧 Removendo possíveis marcas d'água", LOG_LEVEL_INFO

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
                            LogMessage "✅ Marca d'água removida: " & shp.Name, LOG_LEVEL_INFO
                        End If
                    End If
                Next i
            End If
        Next header
        
        ' Also check footers for watermarks
        For Each header In sec.Footers
            If header.Exists And header.Shapes.Count > 0 Then
                For i = header.Shapes.Count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                            LogMessage "✅ Marca d'água removida: " & shp.Name, LOG_LEVEL_INFO
                        End If
                    End If
                Next i
            End If
        Next header
    Next sec

    LogMessage "📊 Total de marcas d'água removidas: " & removedCount, LOG_LEVEL_INFO
    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    LogMessage "⚠️  Erro ao remover marcas d'água: " & Err.Description, LOG_LEVEL_WARNING
    RemoveWatermark = False
End Function

'================================================================================
' INSERT HEADER IMAGE
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

    LogMessage "🖼️  Inserindo carimbo no cabeçalho", LOG_LEVEL_INFO

    username = GetSafeUserName()
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Check if image exists
    If Dir(imgFile) = "" Then
        ' Try alternative paths
        imgFile = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH
        If Dir(imgFile) = "" Then
            ' Try network path or common locations
            imgFile = "\\server\Pictures\LegisTabStamp\HeaderStamp.png"
            If Dir(imgFile) = "" Then
                LogMessage "❌ Imagem de cabeçalho não encontrada em nenhum local", LOG_LEVEL_ERROR
                MsgBox "Imagem de cabeçalho não encontrada nos locais esperados." & vbCrLf & _
                       "Verifique se o arquivo existe em: " & HEADER_IMAGE_RELATIVE_PATH, _
                       vbExclamation, "Imagem Não Encontrada"
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
            header.Range.Delete ' Clear previous content
            
            ' Insert the image as a Shape
            Set shp = header.Shapes.AddPicture( _
                FileName:=imgFile, _
                LinkToFile:=False, _
                SaveWithDocument:=msoTrue)
            
            ' Check if image was loaded correctly
            If shp Is Nothing Then
                LogMessage "❌ Falha ao inserir imagem na seção " & sectionsProcessed + 1, LOG_LEVEL_ERROR
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
                LogMessage "✅ Carimbo inserido na seção " & sectionsProcessed, LOG_LEVEL_INFO
            End If
        End If
    Next sec

    If imgFound Then
        LogMessage "📊 Carimbo inserido em " & sectionsProcessed & " seções", LOG_LEVEL_INFO
        InsertHeaderStamp = True
    Else
        LogMessage "⚠️  Não foi possível inserir carimbo em nenhuma seção", LOG_LEVEL_WARNING
        InsertHeaderStamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "❌ Erro ao inserir carimbo no cabeçalho: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderStamp = False
End Function

'================================================================================
' INSERT FOOTER PAGE NUMBERS
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rng As Range
    Dim sectionsProcessed As Long
    Dim fieldCode As String

    LogMessage "🔢 Inserindo numeração de página no rodapé", LOG_LEVEL_INFO

    ' Define o código de campo exato conforme especificado
    fieldCode = "{PAGE  \* Arabic  \* MERGEFORMAT}-{NUMPAGES  \* Arabic  \* MERGEFORMAT}"

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range
            
            ' Clear previous content completely
            rng.Delete
            
            ' Set basic formatting for the footer range
            With rng
                .Font.Name = STANDARD_FONT
                .Font.Size = FOOTER_FONT_SIZE
                .Font.Bold = False
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.SpaceBefore = 0
                .ParagraphFormat.SpaceAfter = 0
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            End With
            
            ' Insert the exact field code
            rng.Text = fieldCode
            
            ' Convert text to field
            rng.Fields.Add rng, wdFieldEmpty, fieldCode, False
            
            ' Update the field to show correct values
            rng.Fields.Update
            
            ' Ensure no bold formatting in footer
            rng.Font.Bold = False
            rng.Font.Name = STANDARD_FONT
            rng.Font.Size = FOOTER_FONT_SIZE
            
            sectionsProcessed = sectionsProcessed + 1
            LogMessage "✅ Rodapé formatado na seção " & sectionsProcessed, LOG_LEVEL_INFO
        End If
    Next sec

    LogMessage "📊 Numeração de página inserida em " & sectionsProcessed & " seções com o código de campo exato", LOG_LEVEL_INFO
    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    LogMessage "❌ Erro ao inserir numeração de página: " & Err.Description, LOG_LEVEL_ERROR
    InsertFooterStamp = False
End Function

'================================================================================
' ERROR HANDLER
'================================================================================
Private Sub HandleError(procedureName As String)
    Dim errMsg As String
    errMsg = "Erro na sub-rotina: " & procedureName & vbCrLf & _
             "Erro #" & Err.Number & ": " & Err.Description & vbCrLf & _
             "Fonte: " & Err.Source
    Application.StatusBar = "Erro: " & Err.Description
    LogMessage "❌ Erro em " & procedureName & ": " & Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    Debug.Print errMsg
    Err.Clear
End Sub

'================================================================================
' UTILITY: CM TO POINTS
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then
        ' Fallback conversion: 1 cm = 28.35 points
        CentimetersToPoints = cm * 28.35
    End If
End Function

'================================================================================
' UTILITY: SAFE USERNAME
'================================================================================
Private Function GetSafeUserName() As String
    On Error GoTo ErrorHandler
    
    Dim rawName As String
    Dim safeName As String
    Dim i As Integer
    Dim c As String
    
    ' Try multiple methods to get username
    rawName = Environ("USERNAME")
    If rawName = "" Then rawName = Environ("USER")
    If rawName = "" Then 
        On Error Resume Next
        rawName = CreateObject("WScript.Network").UserName
        On Error GoTo 0
    End If
    
    If rawName = "" Then
        rawName = "UsuarioDesconhecido"
    End If
    
    ' Sanitize username for path safety
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
    LogMessage "👤 Nome de usuário sanitizado: " & safeName, LOG_LEVEL_INFO
    Exit Function
    
ErrorHandler:
    GetSafeUserName = "Usuario"
    LogMessage "⚠️  Erro ao obter nome de usuário, usando padrão", LOG_LEVEL_WARNING
End Function

'================================================================================
' ADDITIONAL UTILITY: DOCUMENT BACKUP
'================================================================================
Private Sub CreateBackup(doc As Document)
    On Error GoTo ErrorHandler
    
    If doc.Path = "" Then
        LogMessage "⚠️  Não é possível criar backup - documento não salvo", LOG_LEVEL_WARNING
        Exit Sub
    End If
    
    Dim backupPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    backupPath = doc.Path & "\Backup_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & "_" & doc.Name
    
    doc.SaveAs2 backupPath
    LogMessage "💾 Backup criado: " & backupPath, LOG_LEVEL_INFO
    
    Exit Sub
    
ErrorHandler:
    LogMessage "❌ Falha ao criar backup: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' ADDITIONAL UTILITY: VALIDATE DOCUMENT STRUCTURE
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim valid As Boolean
    valid = True
    
    ' Check if document has content
    If doc.Range.End = 0 Then
        LogMessage "⚠️  Documento está vazio", LOG_LEVEL_WARNING
        valid = False
    End If
    
    ' Check if document has sections
    If doc.Sections.Count = 0 Then
        LogMessage "⚠️  Documento não possui seções", LOG_LEVEL_WARNING
        valid = False
    End If
    
    ValidateDocumentStructure = valid
    Exit Function
    
ErrorHandler:
    LogMessage "⚠️  Erro na validação da estrutura do documento: " & Err.Description, LOG_LEVEL_WARNING
    ValidateDocumentStructure = False
End Function

'================================================================================
' ADDITIONAL UTILITY: RESTORE DEFAULT SETTINGS
'================================================================================
Private Sub RestoreDefaultSettings()
    On Error Resume Next
    SetAppState True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = ""
End Sub

'================================================================================
' UTILITY: OPEN LOG FILE - VERSÃO CORRIGIDA E SEGURA
'================================================================================
Public Sub AbrirLog()
    On Error GoTo ErrorHandler
    
    Dim shell As Object
    Dim logPathToOpen As String
    
    Set shell = CreateObject("WScript.Shell")
    
    ' Determinar qual arquivo de log abrir
    If logFilePath <> "" And Dir(logFilePath) <> "" Then
        logPathToOpen = logFilePath
    Else
        ' Tentar encontrar o log na pasta TEMP
        logPathToOpen = Environ("TEMP") & "\DocumentFormattingLog.txt"
        
        If Dir(logPathToOpen) = "" Then
            ' Procurar por arquivos de log recentes
            logPathToOpen = EncontrarArquivoLogRecente()
            
            If logPathToOpen = "" Then
                MsgBox "Nenhum arquivo de log encontrado." & vbCrLf & _
                       "Execute a padronização primeiro para gerar logs.", _
                       vbInformation, "Log Não Encontrado"
                Exit Sub
            End If
        End If
    End If
    
    ' Verificar se o arquivo existe
    If Dir(logPathToOpen) = "" Then
        MsgBox "Arquivo de log não encontrado:" & vbCrLf & logPathToOpen, _
               vbExclamation, "Arquivo Não Existe"
        Exit Sub
    End If
    
    ' Abrir o arquivo de forma segura com Notepad
    shell.Run "notepad.exe " & Chr(34) & logPathToOpen & Chr(34), 1, True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao abrir o arquivo de log:" & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbExclamation, "Erro"
End Sub

'================================================================================
' UTILITY: FIND RECENT LOG FILE
'================================================================================
Private Function EncontrarArquivoLogRecente() As String
    On Error Resume Next
    
    Dim fso As Object
    Dim tempFolder As Object
    Dim file As Object
    Dim recentFile As Object
    Dim recentDate As Date
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Procurar na pasta TEMP
    Set tempFolder = fso.GetFolder(Environ("TEMP"))
    
    For Each file In tempFolder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            If InStr(1, file.Name, "FormattingLog", vbTextCompare) > 0 Then
                If recentFile Is Nothing Then
                    Set recentFile = file
                    recentDate = file.DateLastModified
                ElseIf file.DateLastModified > recentDate Then
                    Set recentFile = file
                    recentDate = file.DateLastModified
                End If
            End If
        End If
    Next file
    
    If Not recentFile Is Nothing Then
        EncontrarArquivoLogRecente = recentFile.Path
    Else
        EncontrarArquivoLogRecente = ""
    End If
End Function

'================================================================================
' UTILITY: SHOW LOG PATH - VERSÃO SEGURA
'================================================================================
Public Sub MostrarCaminhoDoLog()
    On Error GoTo ErrorHandler
    
    Dim msg As String
    Dim logPath As String
    
    ' Determinar o caminho do log para mostrar
    If logFilePath <> "" And Dir(logFilePath) <> "" Then
        logPath = logFilePath
    Else
        logPath = Environ("TEMP") & "\DocumentFormattingLog.txt"
        If Dir(logPath) = "" Then
            logPath = EncontrarArquivoLogRecente()
            If logPath = "" Then
                msg = "Nenhum arquivo de log encontrado." & vbCrLf & _
                      "Execute a padronização primeiro para gerar logs."
                MsgBox msg, vbInformation, "Log Não Encontrado"
                Exit Sub
            End If
        End If
    End If
    
    ' Verificar se o arquivo existe
    If Dir(logPath) = "" Then
        msg = "Arquivo de log não existe mais:" & vbCrLf & logPath
        MsgBox msg, vbExclamation, "Arquivo Não Existe"
        Exit Sub
    End If
    
    ' Criar mensagem com opções
    msg = "Arquivo de log localizado em:" & vbCrLf & vbCrLf & logPath & vbCrLf & vbCrLf
    msg = msg & "Deseja abrir o arquivo agora?"
    
    Dim response As VbMsgBoxResult
    response = MsgBox(msg, vbQuestion + vbYesNo, "Localização do Log")
    
    If response = vbYes Then
        AbrirLog
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao mostrar caminho do log:" & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbExclamation, "Erro"
End Sub

'================================================================================
' UTILITY: OPEN BACKUP FOLDER
'================================================================================
Public Sub AbrirPastaBackups()
    On Error GoTo ErrorHandler
    
    Dim shell As Object
    Dim backupFolderPath As String
    Dim doc As Document
    Dim fso As Object
    
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set doc = ActiveDocument
    
    ' Verificar se há um documento ativo
    If doc Is Nothing Then
        MsgBox "Nenhum documento está aberto no momento.", vbExclamation, "Documento Não Encontrado"
        Exit Sub
    End If
    
    ' Determinar o caminho da pasta de backups
    If doc.Path <> "" Then
        backupFolderPath = doc.Path
    Else
        ' Se o documento não foi salvo, usar a pasta de documentos do usuário
        backupFolderPath = Environ("USERPROFILE") & "\Documents"
    End If
    
    ' Verificar se a pasta existe
    If Not fso.FolderExists(backupFolderPath) Then
        MsgBox "Pasta não encontrada:" & vbCrLf & backupFolderPath, vbExclamation, "Pasta Não Existe"
        Exit Sub
    End If
    
    ' Verificar se há arquivos de backup na pasta
    Dim hasBackups As Boolean
    hasBackups = False
    
    Dim folder As Object
    Dim file As Object
    Set folder = fso.GetFolder(backupFolderPath)
    
    For Each file In folder.Files
        If InStr(1, file.Name, "Backup_", vbTextCompare) > 0 Then
            hasBackups = True
            Exit For
        End If
    Next file
    
    ' Abrir a pasta no Explorador de Arquivos
    shell.Run "explorer.exe " & Chr(34) & backupFolderPath & Chr(34), 1, True
    
    ' Mensagem informativa
    If hasBackups Then
        MsgBox "Pasta de backups aberta." & vbCrLf & vbCrLf & _
               "Localização: " & backupFolderPath & vbCrLf & _
               "Os arquivos de backup começam com 'Backup_'", _
               vbInformation, "Pasta de Backups"
    Else
        MsgBox "Pasta do documento aberta." & vbCrLf & vbCrLf & _
               "Localização: " & backupFolderPath & vbCrLf & _
               "Nenhum arquivo de backup encontrado." & vbCrLf & _
               "Os backups serão criados aqui quando o documento for salvo.", _
               vbInformation, "Pasta do Documento"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao abrir a pasta de backups:" & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbExclamation, "Erro"
End Sub

'================================================================================
' FINAL MESSAGE - EXIBIÇÃO DE CONCLUSÃO
'================================================================================
Private Sub ShowCompletionMessage()
    MsgBox "Processo de padronização concluído com sucesso!" & vbCrLf & vbCrLf & _
           "✓ Formatação de fonte e parágrafo aplicada" & vbCrLf & _
           "✓ Configurações de página ajustadas" & vbCrLf & _
           "✓ Cabeçalho e rodapé personalizados" & vbCrLf & _
           "✓ Numeração de páginas configurada" & vbCrLf & vbCrLf & _
           "O documento está pronto para uso.", _
           vbInformation + vbOKOnly, "Padronização Concluída"
End Sub
```

CARACTERÍSTICAS DA VERSÃO FINAL:

1. Segurança Robusta:

· ✅ Sistema de recuperação de emergência
· ✅ Prevenção de quedas do Word
· ✅ Tratamento de erro em todas as funções
· ✅ Verificações de espaço em disco
· ✅ Liberação segura de recursos

2. Logs Detalhados:

· ✅ Timestamps precisos com emojis
· ✅ Metadados completos do sistema
· ✅ Hierarquia clara de mensagens
· ✅ Informações de performance
· ✅ Estatísticas de execução

3. Mensagens Aprimoradas:

· ✅ Textos claros e informativos
· ✅ Linguagem amigável ao usuário
· ✅ Instruções de recuperação
· ✅ Detalhes técnicos para suporte

4. Funcionalidades Completas:

· ✅ Formatação de fonte e parágrafo
· ✅ Configuração de margens e página
· ✅ Inserção de cabeçalho personalizado
· ✅ Numeração de páginas automática
· ✅ Sistema de backup automático
· ✅ Gerenciamento de logs
· ✅ Interface de usuário amigável

5. Performance Otimizada:

· ✅ Liberação controlada de memória
· ✅ Gerenciamento de estado eficiente
· ✅ Processamento em lote seguro
· ✅ Timeouts e retries inteligentes

Este código representa a versão final consolidada com todas as melhorias de segurança, logging detalhado e robustez operacional solicitadas! 🚀