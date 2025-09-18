' =============================================================================
' std-legis-docs - Padronização e automação avançada de documentos Word em VBA
' Versão: 2.0.1-stable | Data: 2025-09-11
' Autor: Christian Martin dos Santos | github.com/chrmsantos/std-legis-docs
' =============================================================================
' Solução open source para formatação, segurança e interface.
' aprimorada de documentos institucionais no Microsoft Word.
' Licença: Apache 2.0 modificada com cláusula 10 (restrição comercial), ver LICENSE
' =============================================================================
' Funcionalidades principais:
' - Segurança e robustez (recuperação, tratamento de erros)
' - Interface aprimorada (mensagens, status dinâmico)
' - Padronização e formatação:
'     • Margens automáticas (superior, inferior, esquerda, direita)
'     • Fonte padrão Arial, tamanho 12, espaçamento entrelinhas 1,4
'     • Recuo de parágrafo e primeira linha conforme regras institucionais
'     • Cabeçalho institucional com imagem personalizada centralizada
'     • Numeração automática de páginas centralizada no rodapé
'     • Remoção de marcas d'água e linhas em branco iniciais
'     • Verificação de estrutura mínima
'     • Hifenização automática e alinhamento justificado
' - Performance e utilitários (processamento seguro, undo customizado)
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

'================================================================================
' GLOBAL VARIABLES
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date

'================================================================================
' MAIN ENTRY POINT - #STABLE
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    
    If Not CheckWordVersion() Then
        Dim versionMsg As String
        versionMsg = "Versão do Word (" & Application.version & ") não suportada. " & _
                    "Requisito mínimo: Word 2010 (versão 14.0). " & _
                    "Atualize o Microsoft Word para utilizar este recurso."
        MsgBox versionMsg, vbExclamation + vbOKOnly, "Compatibilidade Não Suportada"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = Nothing
    
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Nenhum documento está aberto ou acessível no momento." & vbCrLf & _
               "Por favor, abra um documento do Word e tente novamente.", _
               vbExclamation + vbOKOnly, "Documento Não Disponível"
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    If Not InitializeLogging(doc) Then
    End If
    
    StartUndoGroup "Padronização de Documento - " & doc.Name
    
    If Not SetAppState(False, "Formatando documento...") Then
    End If
    
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If
    
    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            MsgBox "Operação cancelada. O documento precisa ser salvo antes da formatação.", _
                   vbInformation, "Operação Cancelada"
            Exit Sub
        End If
    End If
    
    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    Application.StatusBar = "Documento padronizado com sucesso!"

CleanUp:
    SafeCleanup
    
    If Not SetAppState(True, "Documento padronizado com sucesso!") Then
    End If
    
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    EmergencyRecovery
    
    MsgBox "Ocorreu um erro inesperado durante o processamento." & vbCrLf & vbCrLf & _
           "Detalhes técnicos: " & errDesc & vbCrLf & vbCrLf & _
           "O Word tentou recuperar o estado normal da aplicação.", _
           vbCritical + vbOKOnly, "Erro Inesperado"
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
        MsgBox "Erro de sistema: Nenhum documento está acessível para verificação." & vbCrLf & _
               "Tente fechar e reabrir o documento, então execute novamente.", _
               vbCritical + vbOKOnly, "Falha de Acesso ao Documento"
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        MsgBox "Documento incompatível detectado." & vbCrLf & _
               "Este sistema suporta apenas documentos do Word padrão." & vbCrLf & _
               "Tipo atual: " & doc.Type, _
               vbExclamation + vbOKOnly, "Tipo de Documento Não Suportado"
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        
        MsgBox "Documento protegido detectado." & vbCrLf & _
               "Tipo de proteção: " & protectionType & vbCrLf & vbCrLf & _
               "Para continuar, remova a proteção do documento através de:" & vbCrLf & _
               "Revisão > Proteger > Restringir Edição > Parar Proteção", _
               vbExclamation + vbOKOnly, "Documento Protegido"
        PreviousChecking = False
        Exit Function
    End If
    
    If doc.ReadOnly Then
        MsgBox "Documento em modo somente leitura." & vbCrLf & _
               "Salve uma cópia editável do documento antes de prosseguir." & vbCrLf & vbCrLf & _
               "Arquivo: " & doc.FullName, _
               vbExclamation + vbOKOnly, "Documento Somente Leitura"
        PreviousChecking = False
        Exit Function
    End If

    If Not CheckDiskSpace(doc) Then
        MsgBox "Espaço em disco insuficiente para completar a operação com segurança." & vbCrLf & _
               "Libere pelo menos 50MB de espaço livre e tente novamente.", _
               vbExclamation + vbOKOnly, "Espaço em Disco Insuficiente"
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
    End If

    PreviousChecking = True
    Exit Function

ErrorHandler:
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
    
    If doc.Path <> "" Then
        driveLetter = Left(doc.Path, 3)
    Else
        driveLetter = Left(Environ("TEMP"), 3)
    End If
    
    Set drive = fso.GetDrive(driveLetter)
    
    requiredSpace = 50 * 1024 * 1024
    
    If drive.AvailableSpace < requiredSpace Then
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If
    
    Exit Function
    
ErrorHandler:
    CheckDiskSpace = True
End Function

'================================================================================
' MAIN FORMATTING ROUTINE - #STABLE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

     If Not ApplyPageSetup(doc) Then
        PreviousFormatting = False
        Exit Function
    End If

    If Not ApplyStdFont(doc) Then
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not ApplyStdParagraphs(doc) Then
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not EnableHyphenation(doc) Then
    End If
    
    If Not RemoveWatermark(doc) Then
    End If
    
    If Not InsertHeaderStamp(doc) Then
    End If
    
    If Not InsertFooterStamp(doc) Then
        PreviousFormatting = False
        Exit Function
    End If
    
    PreviousFormatting = True
    Exit Function

ErrorHandler:
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
    
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
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

    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

        If para.Range.InlineShapes.Count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        If Not hasInlineImage Then
            With para.Range.Font
                .Name = STANDARD_FONT
                .Size = STANDARD_FONT_SIZE
                .Underline = wdUnderlineNone
                .Color = wdColorAutomatic
            End With
            
            formattedCount = formattedCount + 1
        End If
    Next i
    
    ApplyStdFont = True
    Exit Function

ErrorHandler:
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

        If Not hasInlineImage Then
            Do While InStr(para.Range.Text, "  ") > 0
                para.Range.Text = Replace(para.Range.Text, "  ", " ")
            Loop

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
    
    ApplyStdParagraphs = True
    Exit Function

ErrorHandler:
    ApplyStdParagraphs = False
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
        EnableHyphenation = True
    Else
        EnableHyphenation = True
    End If
    
    Exit Function
    
ErrorHandler:
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

    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
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

    If Dir(imgFile) = "" Then
        imgFile = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH
        If Dir(imgFile) = "" Then
            imgFile = "\\strqnapmain\Dir. Legislativa\Christian" & HEADER_IMAGE_RELATIVE_PATH
            If Dir(imgFile) = "" Then
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
            header.Range.Delete
            
            Set shp = header.Shapes.AddPicture( _
                FileName:=imgFile, _
                LinkToFile:=False, _
                SaveWithDocument:=msoTrue)
            
            If shp Is Nothing Then
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
        InsertHeaderStamp = True
    Else
        InsertHeaderStamp = False
    End If

    Exit Function

ErrorHandler:
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
                .Font.Bold = False
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
            
            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    InsertFooterStamp = False
End Function

'================================================================================
' ERROR HANDLER - #STABLE
'================================================================================
Private Sub HandleError(procedureName As String)
    Dim errMsg As String
    errMsg = "Erro na sub-rotina: " & procedureName & vbCrLf & _
             "Erro #" & Err.Number & ": " & Err.Description & vbCrLf & _
             "Fonte: " & Err.Source
    Application.StatusBar = "Erro: " & Err.Description
    Err.Clear
End Sub

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
' ADDITIONAL UTILITY: VALIDATE DOCUMENT STRUCTURE - #STABLE
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim valid As Boolean
    valid = True
    
    If doc.Range.End = 0 Then
        valid = False
    End If
    
    If doc.Sections.Count = 0 Then
        valid = False
    End If
    
    ValidateDocumentStructure = valid
    Exit Function
    
ErrorHandler:
    ValidateDocumentStructure = False
End Function

'================================================================================
' ADDITIONAL UTILITY: RESTORE DEFAULT SETTINGS - #STABLE
'================================================================================
Private Sub RestoreDefaultSettings()
    On Error Resume Next
    SetAppState True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = ""
End Sub


'================================================================================
' CRITICAL FIX: SAVE DOCUMENT BEFORE PROCESSING
' TO PREVENT CRASHES ON NEW NON SAVED DOCUMENTS - #STABLE
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
        SaveDocumentFirst = False
        Exit Function
    End If

    Dim waitCount As Integer
    For waitCount = 1 To 10
        DoEvents
        If doc.Path <> "" Then Exit For
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
    Next waitCount

    If doc.Path = "" Then
        SaveDocumentFirst = False
    Else
        SaveDocumentFirst = True
    End If

    Exit Function

ErrorHandler:
    SaveDocumentFirst = False
End Function

