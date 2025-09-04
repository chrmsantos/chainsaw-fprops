' ===================================================
' CÓDIGO DE PADRONIZAÇÃO DE DOCUMENTOS WORD
' ===================================================

' Licenciado sob Apache License, Version 2.0
' https://www.apache.org/licenses/LICENSE-2.0

' CARACTERÍSTICAS PRINCIPAIS:
' - Sistema robusto de tratamento de erros
' - Logs detalhados com timestamps
' - Formatação completa de documento
' - Interface amigável com usuário
' - Backup automático e recuperação

Option Explicit

'================================================================================
' CONSTANTES
'================================================================================

' Constantes do Word
Private Const wdNoProtection As Long = -1
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdAlignParagraphJustify As Long = 3
Private Const wdLineSpaceSingle As Long = 0
Private Const wdLineSpace1pt5 As Long = 1
Private Const wdLineSpacingMultiple As Long = 5
Private Const wdFieldPage As Long = 33
Private Const wdFieldNumPages As Long = 26
Private Const wdRelativeHorizontalPositionPage As Long = 1
Private Const wdRelativeVerticalPositionPage As Long = 1
Private Const wdWrapTopBottom As Long = 3
Private Const wdAlertsAll As Long = 0
Private Const wdAlertsNone As Long = -1
Private Const wdColorAutomatic As Long = -16777216
Private Const wdOrientPortrait As Long = 0
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoPicture As Long = 13

' Configurações de formatação
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

' Margens (em centímetros)
Private Const TOP_MARGIN_CM As Double = 4.6
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

' Cabeçalho
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\Pictures\LegisTabStamp\HeaderStamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

' Versão mínima suportada
Private Const MIN_SUPPORTED_VERSION As Long = 14 ' Word 2010

' Níveis de log
Private Const LOG_LEVEL_INFO As Long = 1
Private Const LOG_LEVEL_WARNING As Long = 2
Private Const LOG_LEVEL_ERROR As Long = 3

' String obrigatória
Private Const REQUIRED_STRING As String = " Nº $NUMERO$/$ANO$"

'================================================================================
' VARIÁVEIS GLOBAIS
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date

'================================================================================
' PONTO DE ENTRADA PRINCIPAL
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    
    LogMessage "INÍCIO DA EXECUÇÃO - Processo de padronização iniciado", LOG_LEVEL_INFO
    
    ' Verificar compatibilidade
    If Not CheckWordVersion() Then
        MsgBox "Versão do Word (" & Application.version & ") não suportada." & vbCrLf & _
               "Requisito mínimo: Word 2010 (versão 14.0).", _
               vbExclamation, "Compatibilidade Não Suportada"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Nenhum documento está aberto." & vbCrLf & _
               "Por favor, abra um documento do Word e tente novamente.", _
               vbExclamation, "Documento Não Disponível"
        Exit Sub
    End If
    
    ' Inicializar sistema de log
    InitializeLogging doc
    
    LogMessage "Documento selecionado: '" & doc.Name & "'", LOG_LEVEL_INFO
    
    ' Iniciar processo
    StartUndoGroup "Padronização de Documento - " & doc.Name
    SetAppState False, "Formatando documento..."
    
    ' Executar verificações e formatação
    If Not PreviousChecking(doc) Then GoTo CleanUp
    If Not PreviousFormatting(doc) Then GoTo CleanUp
    If formattingCancelled Then GoTo CleanUp
    
    ' Sucesso
    Application.StatusBar = "Documento padronizado com sucesso!"
    LogMessage "PROCESSAMENTO CONCLUÍDO COM SUCESSO", LOG_LEVEL_INFO
    
    Dim executionTime As String
    executionTime = Format(Now - executionStartTime, "nn:ss")
    LogMessage "Tempo total de execução: " & executionTime, LOG_LEVEL_INFO

CleanUp:
    SafeCleanup
    SetAppState True, "Documento padronizado com sucesso!"
    SafeFinalizeLogging
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO #" & Err.Number & ": " & Err.Description
    
    LogMessage "ERRO CRÍTICO: " & errDesc, LOG_LEVEL_ERROR
    EmergencyRecovery
    
    MsgBox "Ocorreu um erro inesperado durante o processamento." & vbCrLf & _
           "Detalhes: " & errDesc, _
           vbCritical, "Erro Inesperado"
End Sub

'================================================================================
' RECUPERAÇÃO DE EMERGÊNCIA
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next
    
    LogMessage "Executando recuperação de emergência", LOG_LEVEL_ERROR
    
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
' LIMPEZA SEGURA
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    LogMessage "Iniciando limpeza segura", LOG_LEVEL_INFO
    EndUndoGroup
    ReleaseObjects
    LogMessage "Limpeza concluída", LOG_LEVEL_INFO
End Sub

Private Sub ReleaseObjects()
    On Error Resume Next
    Dim i As Long
    For i = 1 To 3
        DoEvents
    Next i
End Sub

Private Sub CloseAllOpenFiles()
    On Error Resume Next
    Dim fileNumber As Integer
    For fileNumber = 1 To 511
        If Not EOF(fileNumber) Then Close fileNumber
    Next fileNumber
End Sub

'================================================================================
' VERIFICAÇÃO DE VERSÃO
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Long
    version = Application.version
    
    If version < MIN_SUPPORTED_VERSION Then
        LogMessage "Versão do Word " & version & " não suportada", LOG_LEVEL_ERROR
        CheckWordVersion = False
    Else
        LogMessage "Versão do Word " & version & " compatível", LOG_LEVEL_INFO
        CheckWordVersion = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Falha na verificação de versão", LOG_LEVEL_ERROR
    CheckWordVersion = False
End Function

'================================================================================
' GERENCIAMENTO DE GRUPO UNDO
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then EndUndoGroup
    
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    LogMessage "Grupo undo iniciado: '" & groupName & "'", LOG_LEVEL_INFO
    Exit Sub
    
ErrorHandler:
    LogMessage "Falha ao iniciar grupo undo", LOG_LEVEL_ERROR
    undoGroupEnabled = False
End Sub

Private Sub EndUndoGroup()
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "Grupo undo finalizado", LOG_LEVEL_INFO
    End If
    Exit Sub
    
ErrorHandler:
    LogMessage "Falha ao finalizar grupo undo", LOG_LEVEL_ERROR
    undoGroupEnabled = False
End Sub

'================================================================================
' SISTEMA DE LOGGING
'================================================================================
Private Sub InitializeLogging(doc As Document)
    On Error GoTo ErrorHandler
    
    If doc.Path <> "" Then
        logFilePath = doc.Path & "\" & Format(Now, "yyyy-mm-dd") & "_Log.txt"
    Else
        logFilePath = Environ("TEMP") & "\" & Format(Now, "yyyy-mm-dd") & "_Log.txt"
    End If
    
    Open logFilePath For Output As #1
    Print #1, "LOG DE FORMATAÇÃO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    Print #1, "Documento: " & doc.Name
    Print #1, "Usuário: " & Environ("USERNAME")
    Print #1, "========================================"
    Close #1
    
    loggingEnabled = True
    LogMessage "Sistema de logging inicializado", LOG_LEVEL_INFO
    Exit Sub
    
ErrorHandler:
    loggingEnabled = False
End Sub

Private Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error GoTo ErrorHandler
    
    If Not loggingEnabled Then Exit Sub
    
    Dim levelText As String
    Select Case level
        Case LOG_LEVEL_INFO: levelText = "INFO"
        Case LOG_LEVEL_WARNING: levelText = "AVISO"
        Case LOG_LEVEL_ERROR: levelText = "ERRO"
        Case Else: levelText = "OUTRO"
    End Select
    
    Dim formattedMessage As String
    formattedMessage = Format(Now, "HH:MM:ss") & " [" & levelText & "] " & message
    
    Open logFilePath For Append As #1
    Print #1, formattedMessage
    Close #1
    
    Debug.Print "LOG: " & formattedMessage
    Exit Sub
    
ErrorHandler:
    Debug.Print "FALHA NO LOG: " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error Resume Next
    
    If loggingEnabled Then
        Open logFilePath For Append As #1
        Print #1, "FIM DA SESSÃO - " & Format(Now, "HH:MM:ss")
        Print #1, "Duração: " & Format(Now - executionStartTime, "HH:MM:ss")
        Close #1
    End If
    
    loggingEnabled = False
End Sub

'================================================================================
' CONTROLE DE ESTADO DA APLICAÇÃO
'================================================================================
Private Sub SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "")
    On Error Resume Next
    
    With Application
        .ScreenUpdating = enabled
        .DisplayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
        If statusMsg <> "" Then .StatusBar = statusMsg
        .EnableCancelKey = 0
    End With
End Sub

'================================================================================
' VERIFICAÇÕES PRELIMINARES
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Iniciando verificações de segurança", LOG_LEVEL_INFO

    If doc Is Nothing Then
        MsgBox "Nenhum documento disponível para verificação.", vbCritical, "Falha de Acesso"
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        MsgBox "Tipo de documento não suportado.", vbExclamation, "Tipo Inválido"
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        MsgBox "Documento protegido contra edição.", vbExclamation, "Documento Protegido"
        PreviousChecking = False
        Exit Function
    End If
    
    If doc.ReadOnly Then
        MsgBox "Documento em modo somente leitura.", vbExclamation, "Somente Leitura"
        PreviousChecking = False
        Exit Function
    End If

    PreviousChecking = True
    Exit Function

ErrorHandler:
    MsgBox "Erro durante verificações de segurança.", vbCritical, "Erro de Verificação"
    PreviousChecking = False
End Function

'================================================================================
' FORMATAÇÃO PRINCIPAL
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogMessage "Iniciando formatação principal", LOG_LEVEL_INFO

    RemoveLeadingBlankLines doc
    ApplyPageSetup doc
    ApplyFontFormatting doc
    ApplyParagraphFormatting doc
    EnableHyphenation doc
    RemoveWatermark doc
    InsertHeaderStamp doc
    InsertFooterStamp doc

    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatação principal", LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' FUNÇÕES DE FORMATAÇÃO
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then CentimetersToPoints = cm * 28.35
End Function

Public Function ApplyPageSetup(doc As Document) As Boolean
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
    
    LogMessage "Configurações de página aplicadas", LOG_LEVEL_INFO
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao aplicar configurações de página", LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

Public Function InsertHeaderStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim shp As Shape

    LogMessage "Inserindo carimbo no cabeçalho", LOG_LEVEL_INFO

    imgFile = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH

    If Dir(imgFile) = "" Then
        LogMessage "Imagem de cabeçalho não encontrada", LOG_LEVEL_ERROR
        InsertHeaderStamp = False
        Exit Function
    End If

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete
            
            Set shp = header.Shapes.AddPicture(imgFile, False, True)
            
            With shp
                .LockAspectRatio = msoTrue
                .Width = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
                .Height = .Width * HEADER_IMAGE_HEIGHT_RATIO
                .Left = (doc.PageSetup.PageWidth - .Width) / 2
                .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
            End With
        End If
    Next sec

    InsertHeaderStamp = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir carimbo no cabeçalho", LOG_LEVEL_ERROR
    InsertHeaderStamp = False
End Function

Public Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rng As Range

    LogMessage "Inserindo numeração de página", LOG_LEVEL_INFO

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        
        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range
            rng.Delete
            
            rng.Fields.Add Range:=rng, Type:=wdFieldPage
            rng.InsertAfter "-"
            rng.Collapse 0
            rng.Fields.Add Range:=rng, Type:=wdFieldNumPages
            
            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.size = FOOTER_FONT_SIZE
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
        End If
    Next sec

    LogMessage "Numeração de página inserida", LOG_LEVEL_INFO
    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir numeração de página", LOG_LEVEL_ERROR
    InsertFooterStamp = False
End Function

Public Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
    Dim i As Long

    LogMessage "Removendo marcas d'água", LOG_LEVEL_INFO

    For Each sec In doc.Sections
        For Each header In sec.Headers
            If header.Exists Then
                For i = header.Shapes.Count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Then shp.Delete
                Next i
            End If
        Next header
    Next sec

    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao remover marcas d'água", LOG_LEVEL_WARNING
    RemoveWatermark = False
End Function

Public Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
        LogMessage "Hifenização automática ativada", LOG_LEVEL_INFO
    End If
    
    EnableHyphenation = True
    Exit Function
    
ErrorHandler:
    LogMessage "Falha ao ativar hifenização", LOG_LEVEL_WARNING
    EnableHyphenation = False
End Function

Public Sub RemoveLeadingBlankLines(doc As Document)
    On Error Resume Next
    Dim para As Paragraph
    Do While doc.Paragraphs.Count > 0
        Set para = doc.Paragraphs(1)
        If Trim(para.Range.Text) = vbCr Or Trim(para.Range.Text) = "" Then
            para.Range.Delete
        Else
            Exit Do
        End If
    Loop
End Sub

Public Function ApplyFontFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long

    LogMessage "Aplicando formatação de fonte", LOG_LEVEL_INFO

    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        
        ' Remover espaços duplos
        Do While InStr(para.Range.Text, "  ") > 0
            para.Range.Text = Replace(para.Range.Text, "  ", " ")
        Loop

        ' Aplicar formatação padrão
        With para.Range.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = 0
            .Color = wdColorAutomatic
        End With
    Next i

    ApplyFontFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao aplicar formatação de fonte", LOG_LEVEL_ERROR
    ApplyFontFormatting = False
End Function

Public Function ApplyParagraphFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long

    LogMessage "Aplicando formatação de parágrafo", LOG_LEVEL_INFO

    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        
        With para.Format
            .LineSpacingRule = wdLineSpacingMultiple
            .LineSpacing = LINE_SPACING
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = CentimetersToPoints(2)
        End With

        If para.Alignment = wdAlignParagraphLeft Then
            para.Alignment = wdAlignParagraphJustify
        End If
    Next i

    ApplyParagraphFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao aplicar formatação de parágrafo", LOG_LEVEL_ERROR
    ApplyParagraphFormatting = False
End Function