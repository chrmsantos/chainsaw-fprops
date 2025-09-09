Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Word built-in constants
Private Const wdNoProtection As Long = -1
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdAlignParagraphJustify As Long = 3
Private Const wdLineSpacingMultiple As Long = 5
Private Const wdFieldPage As Long = 33
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
' MAIN ENTRY POINT - ROBUST SECURITY
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    
    ' Detailed execution start log
    LogMessage "?? INÍCIO DA EXECUÇÃO - Processo de padronização iniciado", LOG_LEVEL_INFO
    LogMessage "?? Contexto: Usuário='" & Environ("USERNAME") & "', Estação='" & Environ("COMPUTERNAME") & "'", LOG_LEVEL_INFO
    
    ' Version compatibility check
    If Not CheckWordVersion() Then
        Dim versionMsg As String
        versionMsg = "Versão do Word (" & Application.version & ") não suportada. " & _
                    "Requisito mínimo: Word 2010 (versão 14.0). " & _
                    "Atualize o Microsoft Word para utilizar este recurso."
        LogMessage "? " & versionMsg, LOG_LEVEL_ERROR
        MsgBox versionMsg, vbExclamation + vbOKOnly, "Compatibilidade Não Suportada"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = Nothing
    
    ' Get active document with safe handling
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        LogMessage "? Nenhum documento ativo disponível para processamento", LOG_LEVEL_ERROR
        MsgBox "Nenhum documento está aberto ou acessível no momento." & vbCrLf & _
               "Por favor, abra um documento do Word e tente novamente.", _
               vbExclamation + vbOKOnly, "Documento Não Disponível"
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    ' CRITICAL FIX: Save document before any changes to prevent crashes
    If doc.Path = "" Then
        LogMessage "?? Documento não salvo - solicitando salvamento inicial", LOG_LEVEL_INFO
        If Not SaveDocumentFirst(doc) Then
            LogMessage "? Usuário cancelou o salvamento inicial", LOG_LEVEL_WARNING
            MsgBox "Operação cancelada. O documento precisa ser salvo antes da formatação.", _
                   vbInformation, "Operação Cancelada"
            Exit Sub
        End If
    End If
    
    ' Initialize logging system
    If Not InitializeLogging(doc) Then
        LogMessage "??  Sistema de logging não inicializado - continuando sem logs detalhados", LOG_LEVEL_WARNING
    End If
    
    LogMessage "?? Documento selecionado: '" & doc.Name & "'", LOG_LEVEL_INFO
    LogMessage "?? Localização: " & doc.Path, LOG_LEVEL_INFO
    
    ' Start undo group with protection
    StartUndoGroup "Padronização de Documento - " & doc.Name
    
    ' Set application state with fallbacks
    If Not SetAppState(False, "Formatando documento...") Then
        LogMessage "??  Configuração de estado da aplicação parcialmente bem-sucedida", LOG_LEVEL_WARNING
    End If
    
    ' Execute preliminary checks
    If Not PreviousChecking(doc) Then
        LogMessage "??  Verificações preliminares falharam - execução interrompida", LOG_LEVEL_ERROR
        GoTo CleanUp
    End If
    
    ' Execute main processing
    If Not PreviousFormatting(doc) Then
        LogMessage "??  Processamento principal falhou - execução interrompida", LOG_LEVEL_ERROR
        GoTo CleanUp
    End If
    
    If formattingCancelled Then
        LogMessage "??  Processamento cancelado pelo usuário", LOG_LEVEL_INFO
        GoTo CleanUp
    End If
    
    ' Success
    Application.StatusBar = "? Documento padronizado com sucesso!"
    LogMessage "? PROCESSAMENTO CONCLUÍDO COM SUCESSO", LOG_LEVEL_INFO
    
    Dim executionTime As String
    executionTime = Format(Now - executionStartTime, "nn:ss")
    LogMessage "??  Tempo total de execução: " & executionTime, LOG_LEVEL_INFO
    
CleanUp:
    ' Safe cleanup with individual error handling
    SafeCleanup
    
    ' Restore application state
    If Not SetAppState(True, "? Documento padronizado com sucesso!") Then
        LogMessage "??  Restauração parcial do estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    ' Safe logging finalization
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    ' Critical error handling
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    LogMessage "?? " & errDesc, LOG_LEVEL_ERROR
    LogMessage "?? Iniciando recuperação de erro crítico", LOG_LEVEL_ERROR
    
    ' Emergency recovery
    EmergencyRecovery
    
    ' User-friendly message
    MsgBox "Ocorreu um erro inesperado durante o processamento." & vbCrLf & vbCrLf & _
           "Detalhes técnicos: " & errDesc & vbCrLf & vbCrLf & _
           "O Word tentou recuperar o estado normal da aplicação." & vbCrLf & _
           "Verifique o arquivo de log para mais detalhes.", _
           vbCritical + vbOKOnly, "Erro Inesperado"
End Sub

'================================================================================
' CRITICAL FIX: SAVE DOCUMENT BEFORE PROCESSING
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)
    
    ' Show save dialog
    If saveDialog.Show <> -1 Then ' User cancelled
        SaveDocumentFirst = False
        Exit Function
    End If
    
    ' Wait for document to be properly saved (Word-compatible method)
    Dim waitCount As Integer
    For waitCount = 1 To 10
        DoEvents
        If doc.Path <> "" Then Exit For
        ' Use DoEvents instead of Application.Wait for Word compatibility
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 1 ' Wait 1 second
            DoEvents
        Loop
    Next waitCount
    
    If doc.Path = "" Then
        LogMessage "? Falha ao salvar documento - caminho não definido", LOG_LEVEL_ERROR
        SaveDocumentFirst = False
    Else
        LogMessage "?? Documento salvo com sucesso: " & doc.Path, LOG_LEVEL_INFO
        SaveDocumentFirst = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "? Erro ao salvar documento: " & Err.Description, LOG_LEVEL_ERROR
    SaveDocumentFirst = False
End Function

'================================================================================
' EMERGENCY RECOVERY - WORD CRASH PREVENTION
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next ' Prevent error loops
    
    LogMessage "???  Executando procedimento de recuperação de emergência", LOG_LEVEL_ERROR
    
    ' Restore critical Word settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Application.EnableCancelKey = 0
    
    ' End undo group if active
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    ' Close open log files
    CloseAllOpenFiles
    
    LogMessage "???  Recuperação de emergência concluída", LOG_LEVEL_INFO
End Sub

'================================================================================
' SAFE CLEANUP - SECURE CLEANUP
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    
    LogMessage "?? Iniciando processo de limpeza segura", LOG_LEVEL_INFO
    
    ' End undo group
    EndUndoGroup
    
    ' Free memory objects
    ReleaseObjects
    
    LogMessage "?? Limpeza segura concluída", LOG_LEVEL_INFO
End Sub

'================================================================================
' RELEASE OBJECTS - SECURE OBJECT RELEASE
'================================================================================
Private Sub ReleaseObjects()
    On Error Resume Next
    
    ' Free potentially allocated objects
    Dim nullObj As Object
    Set nullObj = Nothing
    
    ' Simplified garbage collector
    Dim memoryCounter As Long
    For memoryCounter = 1 To 3
        DoEvents
    Next memoryCounter
End Sub

'================================================================================
' CLOSE ALL OPEN FILES - SECURE FILE CLOSING
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
' VERSION COMPATIBILITY CHECK - ROBUST VERIFICATION
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Long
    version = Application.version
    
    If version < MIN_SUPPORTED_VERSION Then
        LogMessage "? Versão do Word " & version & " não suportada (mínimo: " & MIN_SUPPORTED_VERSION & ")", LOG_LEVEL_ERROR
        CheckWordVersion = False
    Else
        LogMessage "? Versão do Word " & version & " compatível com o sistema", LOG_LEVEL_INFO
        CheckWordVersion = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "? Falha na verificação de versão: " & Err.Description, LOG_LEVEL_ERROR
    CheckWordVersion = False
End Function

'================================================================================
' UNDO GROUP MANAGEMENT - WITH PROTECTION
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        LogMessage "??  Grupo undo já está ativo - finalizando antes de iniciar novo", LOG_LEVEL_WARNING
        EndUndoGroup
    End If
    
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    LogMessage "?? Grupo undo iniciado: '" & groupName & "'", LOG_LEVEL_INFO
    
    Exit Sub
    
ErrorHandler:
    LogMessage "? Falha ao iniciar grupo undo: " & Err.Description, LOG_LEVEL_ERROR
    undoGroupEnabled = False
End Sub

Private Sub EndUndoGroup()
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "?? Grupo undo finalizado com sucesso", LOG_LEVEL_INFO
    Else
        LogMessage "??  Nenhum grupo undo ativo para finalizar", LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    LogMessage "? Falha ao finalizar grupo undo: " & Err.Description, LOG_LEVEL_ERROR
    undoGroupEnabled = False
End Sub

'================================================================================
' LOGGING MANAGEMENT - ENHANCED WITH DETAILS
'================================================================================
Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Determine log file path
    logFilePath = doc.Path & "\" & Format(Now, "yyyy-mm-dd") & "_" & _
                 Replace(doc.Name, ".doc", "") & "_FormattingLog.txt"
    logFilePath = Replace(logFilePath, ".docx", "") & "_FormattingLog.txt"
    logFilePath = Replace(logFilePath, ".docm", "") & "_FormattingLog.txt"
    
    ' Create log file with detailed information
    Open logFilePath For Output As #1
    Print #1, "================================================"
    Print #1, "?? LOG DE FORMATAÇÃO DE DOCUMENTO - SISTEMA DE REGISTRO"
    Print #1, "================================================"
    Print #1, "???  Sessão: " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    Print #1, "?? Usuário: " & Environ("USERNAME")
    Print #1, "?? Estação: " & Environ("COMPUTERNAME")
    Print #1, "?? Versão Word: " & Application.version
    Print #1, "?? Documento: " & doc.Name
    Print #1, "?? Local: " & doc.Path
    Print #1, "?? Proteção: " & GetProtectionType(doc)
    Print #1, "?? Tamanho: " & GetDocumentSize(doc)
    Print #1, "================================================"
    Close #1
    
    loggingEnabled = True
    LogMessage "?? Sistema de logging inicializado: " & logFilePath, LOG_LEVEL_INFO
    InitializeLogging = True
    
    Exit Function
    
ErrorHandler:
    LogMessage "? Falha crítica na inicialização do logging: " & Err.Description, LOG_LEVEL_ERROR
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
            levelIcon = "?? "
        Case LOG_LEVEL_WARNING
            levelText = "AVISO"
            levelIcon = "?? "
        Case LOG_LEVEL_ERROR
            levelText = "ERRO"
            levelIcon = "?"
        Case Else
            levelText = "OUTRO"
            levelIcon = "??"
    End Select
    
    ' Format message with detailed timestamp
    Dim formattedMessage As String
    formattedMessage = Format(Now, "yyyy-mm-dd HH:MM:ss") & " [" & levelText & "] " & levelIcon & " " & message
    
    ' Write to log file
    Open logFilePath For Append As #1
    Print #1, formattedMessage
    Close #1
    
    ' Output to Debug Window
    Debug.Print "LOG: " & formattedMessage
    
    Exit Sub
    
ErrorHandler:
    ' Safe logging fallback
    Debug.Print "FALHA NO LOGGING: " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler
    
    If loggingEnabled Then
        Open logFilePath For Append As #1
        Print #1, "================================================"
        Print #1, "?? FIM DA SESSÃO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
        Print #1, "??  Duração: " & Format(Now - executionStartTime, "HH:MM:ss")
        Print #1, "?? Status: " & IIf(formattingCancelled, "CANCELADO", "CONCLUÍDO")
        Print #1, "================================================"
        Close #1
        
        LogMessage "?? Log finalizado com sucesso", LOG_LEVEL_INFO
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
' UTILITY: GET DOCUMENT SIZE
'================================================================================
Private Function GetDocumentSize(doc As Document) As String
    On Error Resume Next
    
    Dim size As Long
    size = doc.BuiltInDocumentProperties("Number of Characters").Value * 2 ' Approximation
    
    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' APPLICATION STATE HANDLER - ROBUST
'================================================================================
Private Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim success As Boolean
    success = True
    
    With Application
        ' ScreenUpdating - critical for performance
        On Error Resume Next
        .ScreenUpdating = enabled
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        ' DisplayAlerts - important for UX
        On Error Resume Next
        .DisplayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        ' StatusBar - user feedback
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
        
        ' EnableCancelKey - prevent accidental cancellation
        On Error Resume Next
        .EnableCancelKey = 0
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
    End With
    
    If enabled Then
        LogMessage "?? Estado da aplicação restaurado: " & IIf(success, "Completo", "Parcial"), LOG_LEVEL_INFO
    Else
        LogMessage "? Estado de performance ativado: " & IIf(success, "Completo", "Parcial"), LOG_LEVEL_INFO
    End If
    
    SetAppState = success
    Exit Function
    
ErrorHandler:
    LogMessage "? Erro ao configurar estado da aplicação: " & Err.Description, LOG_LEVEL_ERROR
    SetAppState = False
End Function

'================================================================================
' GLOBAL CHECKING - ROBUST VERIFICATIONS
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "?? Iniciando verificações de segurança do documento", LOG_LEVEL_INFO

    ' Check 1: Document exists and is accessible
    If doc Is Nothing Then
        LogMessage "? Falha crítica: Nenhum documento disponível para verificação", LOG_LEVEL_ERROR
        MsgBox "Erro de sistema: Nenhum documento está acessível para verificação." & vbCrLf & _
               "Tente fechar e reabrir o documento, então execute novamente.", _
               vbCritical + vbOKOnly, "Falha de Acesso ao Documento"
        PreviousChecking = False
        Exit Function
    End If

    ' Check 2: Valid document type
    If doc.Type <> wdTypeDocument Then
        LogMessage "? Tipo de documento inválido: " & doc.Type & " (esperado: " & wdTypeDocument & ")", LOG_LEVEL_ERROR
        MsgBox "Documento incompatível detectado." & vbCrLf & _
               "Este sistema suporta apenas documentos do Word padrão." & vbCrLf & _
               "Tipo atual: " & doc.Type, _
               vbExclamation + vbOKOnly, "Tipo de Documento Não Suportado"
        PreviousChecking = False
        Exit Function
    End If

    ' Check 3: Document protection
    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        
        LogMessage "? Documento protegido contra edição: " & protectionType, LOG_LEVEL_ERROR
        MsgBox "Documento protegido detectado." & vbCrLf & _
               "Tipo de proteção: " & protectionType & vbCrLf & vbCrLf & _
               "Para continuar, remova a proteção do documento através de:" & vbCrLf & _
               "Revisão > Proteger > Restringir Edição > Parar Proteção", _
               vbExclamation + vbOKOnly, "Documento Protegido"
        PreviousChecking = False
        Exit Function
    End If
    
    ' Check 4: Read-only document
    If doc.ReadOnly Then
        LogMessage "? Documento aberto em modo somente leitura", LOG_LEVEL_ERROR
        MsgBox "Documento em modo somente leitura." & vbCrLf & _
               "Salve uma cópia editável do documento antes de prosseguir." & vbCrLf & vbCrLf & _
               "Arquivo: " & doc.FullName, _
               vbExclamation + vbOKOnly, "Documento Somente Leitura"
        PreviousChecking = False
        Exit Function
    End If

    ' Check 5: Sufficient disk space
    If Not CheckDiskSpace(doc) Then
        LogMessage "? Espaço em disco insuficiente para operação segura", LOG_LEVEL_ERROR
        MsgBox "Espaço em disco insuficiente para completar a operação com segurança." & vbCrLf & _
               "Libere pelo menos 50MB de espaço livre e tente novamente.", _
               vbExclamation + vbOKOnly, "Espaço em Disco Insuficiente"
        PreviousChecking = False
        Exit Function
    End If

    LogMessage "? Todas as verificações de segurança passaram com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    LogMessage "? Erro durante verificações de segurança: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro durante verificações de segurança do documento." & vbCrLf & _
           "Detalhes: " & Err.Description & vbCrLf & _
           "Contate o suporte técnico se o problema persistir.", _
           vbCritical + vbOKOnly, "Erro de Verificação"
    PreviousChecking = False
End Function

'================================================================================
' DISK SPACE CHECK - DISK SPACE VERIFICATION
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim drive As Object
    Dim requiredSpace As Long
    Dim driveLetter As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Determine target drive
    driveLetter = Left(doc.Path, 3)
    
    ' Get drive information
    Set drive = fso.GetDrive(driveLetter)
    
    ' Required space (50MB as safety margin)
    requiredSpace = 50 * 1024 * 1024 ' 50MB in bytes
    
    If drive.AvailableSpace < requiredSpace Then
        LogMessage "??  Espaço em disco limitado: " & Format(drive.AvailableSpace / 1024 / 1024, "0.0") & _
                  "MB disponíveis (mínimo recomendado: 50MB)", LOG_LEVEL_WARNING
        CheckDiskSpace = False
    Else
        LogMessage "?? Espaço em disco adequado: " & Format(drive.AvailableSpace / 1024 / 1024, "0.0") & _
                  "MB disponíveis", LOG_LEVEL_INFO
        CheckDiskSpace = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "??  Não foi possível verificar espaço em disco: " & Err.Description, LOG_LEVEL_WARNING
    CheckDiskSpace = True ' Continue even with verification error
End Function

'================================================================================
' MAIN FORMATTING ROUTINE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "?? Iniciando formatação principal do documento", LOG_LEVEL_INFO

    ' Apply formatting in logical order
    If Not ApplyPageSetup(doc) Then
        LogMessage "? Falha na configuração de página", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not ApplyFontAndParagraph(doc) Then
        LogMessage "? Falha na formatação de fonte e parágrafo", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not EnableHyphenation(doc) Then
        LogMessage "??  Falha na ativação de hifenização", LOG_LEVEL_WARNING
    End If
    
    If Not RemoveWatermark(doc) Then
        LogMessage "??  Falha na remoção de marca d'água", LOG_LEVEL_WARNING
    End If
    
    If Not InsertHeaderStamp(doc) Then
        LogMessage "??  Falha na inserção do carimbo do cabeçalho", LOG_LEVEL_WARNING
    End If
    
    If Not InsertFooterStamp(doc) Then
        LogMessage "? Falha crítica na inserção do rodapé", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "? Erro durante formatação principal: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' PAGE SETUP
'================================================================================
Private Function ApplyPageSetup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    LogMessage "?? Aplicando configurações de página", LOG_LEVEL_INFO
    
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
    
    LogMessage "? Configurações de página aplicadas com sucesso", LOG_LEVEL_INFO
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
    LogMessage "? Erro ao aplicar configurações de página: " & Err.Description, LOG_LEVEL_ERROR
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

    LogMessage "?? Aplicando formatação de fonte e parágrafo", LOG_LEVEL_INFO

    ' Right indent should be ZERO to align with right margin
    rightMarginPoints = 0

    ' Process paragraphs
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
                .size = STANDARD_FONT_SIZE
                .Bold = False
                .Italic = False
                .Underline = 0
                .Color = wdColorAutomatic
            End With

            ' Apply paragraph formatting
            With para.Format
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                
                ' ZERO RIGHT INDENT - aligns with right margin
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
    
    LogMessage "?? Formatação concluída: " & formattedCount & " parágrafos formatados, " & skippedCount & " parágrafos com imagens ignorados", LOG_LEVEL_INFO
    LogMessage "? Recuo à direita definido como ZERO para alinhamento com margem direita", LOG_LEVEL_INFO
    ApplyFontAndParagraph = True
    Exit Function
    
ErrorHandler:
    LogMessage "? Erro ao aplicar formatação de fonte e parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    ApplyFontAndParagraph = False
End Function

'================================================================================
' ENABLE HYPHENATION
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    LogMessage "?? Ativando hifenização automática", LOG_LEVEL_INFO
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
        LogMessage "? Hifenização automática ativada", LOG_LEVEL_INFO
        EnableHyphenation = True
    Else
        LogMessage "??  Hifenização automática já estava ativada", LOG_LEVEL_INFO
        EnableHyphenation = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "??  Falha ao ativar hifenização: " & Err.Description, LOG_LEVEL_WARNING
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

    LogMessage "?? Removendo possíveis marcas d'água", LOG_LEVEL_INFO

    For Each sec In doc.Sections
        For Each header In sec.Headers
            If header.Exists And header.Shapes.Count > 0 Then
                For i = header.Shapes.Count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = 13 Or shp.Type = 15 Then ' msoPicture or msoTextEffect
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

    LogMessage "?? Total de marcas d'água removidas: " & removedCount, LOG_LEVEL_INFO
    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    LogMessage "??  Erro ao remover marcas d'água: " & Err.Description, LOG_LEVEL_WARNING
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

    LogMessage "???  Inserindo carimbo no cabeçalho", LOG_LEVEL_INFO

    username = GetSafeUserName()
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Check if image exists
    If Dir(imgFile) = "" Then
        LogMessage "? Imagem de cabeçalho não encontrada: " & imgFile, LOG_LEVEL_ERROR
        InsertHeaderStamp = False
        Exit Function
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
                SaveWithDocument:=True)
            
            If Not shp Is Nothing Then
                With shp
                    .LockAspectRatio = True
                    .Width = imgWidth
                    .Height = imgHeight
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    .Left = (doc.PageSetup.PageWidth - .Width) / 2
                    .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
                    .WrapFormat.Type = wdWrapTopBottom
                End With
                
                imgFound = True
                sectionsProcessed = sectionsProcessed + 1
            End If
        End If
    Next sec

    If imgFound Then
        LogMessage "?? Carimbo inserido em " & sectionsProcessed & " seções", LOG_LEVEL_INFO
        InsertHeaderStamp = True
    Else
        LogMessage "??  Não foi possível inserir carimbo em nenhuma seção", LOG_LEVEL_WARNING
        InsertHeaderStamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "? Erro ao inserir carimbo no cabeçalho: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderStamp = False
End Function

'================================================================================
' INSERT FOOTER PAGE NUMBER - CURRENT PAGE ONLY
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rng As Range
    Dim sectionsProcessed As Long

    LogMessage "?? Inserindo número de página no rodapé", LOG_LEVEL_INFO

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        
        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range
            
            ' Clear previous content
            rng.Delete
            
            ' Insert only the CURRENT PAGE field
            rng.Fields.Add Range:=rng, Type:=wdFieldPage
            
            ' Apply centered formatting
            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.Size = FOOTER_FONT_SIZE
                .Font.Bold = False
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update ' Update field to show number
            End With
            
            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    LogMessage "?? Numeração de página inserida em " & sectionsProcessed & " seções.", LOG_LEVEL_INFO
    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    LogMessage "? Erro ao inserir número de página: " & Err.Description, LOG_LEVEL_ERROR
    InsertFooterStamp = False
End Function

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
    Exit Function
    
ErrorHandler:
    GetSafeUserName = "Usuario"
End Function