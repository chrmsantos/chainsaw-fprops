' =============================================================================
' chainsaw-fprops - Padronização e automação avançada de documentos Word em VBA
' Versão: 2.0.1-stable | Data: 2025-09-11 (revisado)
' Autor: Christian Martin dos Santos | github.com/chrmsantos/chainsaw-fprops
' Revisado: simplificação, correções e otimizações (logging reduzido)
' =============================================================================

Option Explicit

'---------------------------
' CONSTANTS (mantidas)
'---------------------------
Private Const wdNoProtection As Long = -1
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdAlignParagraphJustify As Long = 3
Private Const wdLineSpacingMultiple As Long = 5
Private Const wdStatisticPages As Long = 2
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoPicture As Long = 13
Private Const msoTextEffect As Long = 15
Private Const wdCollapseEnd As Long = 0
Private Const wdFieldPage As Long = 33
Private Const wdFieldNumPages As Long = 26
Private Const wdRelativeHorizontalPositionPage As Long = 1
Private Const wdRelativeVerticalPositionPage As Long = 1
Private Const wdWrapTopBottom As Long = 3
Private Const wdAlertsAll As Long = 0
Private Const wdAlertsNone As Long = -1
Private Const wdColorAutomatic As Long = -16777216
Private Const wdOrientPortrait As Long = 0
Private Const wdUnderlineNone As Long = 0

Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

Private Const TOP_MARGIN_CM As Double = 4.6
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

Private Const HEADER_IMAGE_RELATIVE_PATH As String = 
    "\\strqnapmain\Dir. Legislativa\Christian\chainsaw-fprops\stamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

Private Const MIN_SUPPORTED_VERSION As Double = 14# ' Word 2010
Private Const LOG_LEVEL_INFO As Long = 1
Private Const LOG_LEVEL_WARNING As Long = 2
Private Const LOG_LEVEL_ERROR As Long = 3

Private Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

'---------------------------
' GLOBAL STATE
'---------------------------
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date
Private fnum As Integer ' file handle for open log (kept open for session)

'================================================================================
' MAIN ENTRY POINT
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler

    executionStartTime = Now
    formattingCancelled = False

    Dim doc As Document
    Set doc = Nothing

    ' Active document safe-get
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo CriticalErrorHandler

    If doc Is Nothing Then
        MsgBox "Nenhum documento está aberto. Abra um documento e tente novamente.", vbExclamation, "Documento Não Disponível"
        Exit Sub
    End If

    ' Initialize logging (non-fatal)
    If Not InitializeLogging(doc) Then
        ' proceed without logging
    End If

    ' Version check
    If Not CheckWordVersion() Then
        MsgBox "Versão do Word não suportada. Requisitos mínimos: Word 2010 (14.0).", vbExclamation, "Compatibilidade"
        GoTo CleanUp
    End If

    ' Start undo grouping
    StartUndoGroup "Padronização de Documento - " & doc.Name

    ' Set application state for performance; non-fatal
    SetAppState False, "Formatando documento..."

    ' Pre-checks
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If

    ' Ensure saved
    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            MsgBox "Operação cancelada. O documento precisa ser salvo antes da formatação.", vbInformation, "Operação Cancelada"
            GoTo CleanUp
        End If
    End If

    ' Main formatting
    If Not PreviousFormatting(doc) Then
        ShowCompletionMessage False
        GoTo CleanUp
    End If

    If formattingCancelled Then
        ShowCompletionMessage False
        GoTo CleanUp
    End If

    Application.StatusBar = "Documento padronizado com sucesso!"
    LogMessage "Processamento concluído", LOG_LEVEL_INFO

    ShowCompletionMessage True

CleanUp:
    ' Safe cleanup
    SafeCleanup
    SetAppState True, ""
    SafeFinalizeLogging
    Exit Sub

CriticalErrorHandler:
    LogMessage "Erro crítico #" & Err.Number & ": " & Err.Description, LOG_LEVEL_ERROR
    EmergencyRecovery
    MsgBox "Ocorreu um erro inesperado: " & Err.Description & vbCrLf & "Verifique o log.", vbCritical, "Erro"
End Sub

'================================================================================
' RECOVERY & CLEANUP
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Application.EnableCancelKey = 0

    If undoGroupEnabled Then
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If

    ' Close only our open log handle (fallback will attempt to close others)
    If fnum <> 0 Then
        Close #fnum
        fnum = 0
    End If
End Sub

Private Sub SafeCleanup()
    On Error Resume Next
    EndUndoGroup
    ReleaseObjects
End Sub

Private Sub ReleaseObjects()
    On Error Resume Next
    Dim i As Long
    For i = 1 To 3
        DoEvents
    Next i
End Sub

' Close all potential open files (conservative)
Private Sub CloseAllOpenFiles()
    On Error Resume Next
    If fnum <> 0 Then
        Close #fnum
        fnum = 0
    End If
    Dim fil As Integer
    For fil = 1 To 255
        On Error Resume Next
        Close fil
    Next fil
End Sub

'================================================================================
' VERSION CHECK
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    Dim v As Double
    v = Val(Application.Version)
    If v < MIN_SUPPORTED_VERSION Then
        LogMessage "Versão do Word não suportada: " & Application.Version, LOG_LEVEL_ERROR
        CheckWordVersion = False
    Else
        CheckWordVersion = True
    End If
    Exit Function
ErrorHandler:
    LogMessage "Falha ao verificar versão do Word: " & Err.Description, LOG_LEVEL_WARNING
    CheckWordVersion = True ' conservador: continue if unknown
End Function

'================================================================================
' UNDO GROUP
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error GoTo ErrorHandler
    If undoGroupEnabled Then EndUndoGroup
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = (Err.Number = 0)
    Err.Clear
    Exit Sub
ErrorHandler:
    undoGroupEnabled = False
End Sub

Private Sub EndUndoGroup()
    On Error Resume Next
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
End Sub

'================================================================================
' LOGGING (simplificado e otimizado)
' - Mantém arquivo aberto durante a sessão (fnum)
' - Registra apenas WARN/ERROR e INFOS relevantes
'================================================================================
Private Sub WriteLogLine(ByVal line As String)
    On Error GoTo Fallback
    If fnum <> 0 Then
        Print #fnum, line
        Exit Sub
    End If
Fallback:
    On Error Resume Next
    If logFilePath = "" Then Exit Sub
    Dim tf As Integer: tf = FreeFile
    Open logFilePath For Append As #tf
    Print #tf, line
    Close #tf
End Sub

Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim baseName As String
    baseName = doc.Name
    baseName = Replace(baseName, ".docx", "")
    baseName = Replace(baseName, ".docm", "")
    baseName = Replace(baseName, ".doc", "")
    baseName = Replace(baseName, ":", "-")
    baseName = Replace(baseName, "/", "-")
    baseName = Replace(baseName, "\", "-")
    baseName = Replace(baseName, "?", "")
    baseName = Replace(baseName, "*", "")
    baseName = Trim(baseName)
    If doc.Path <> "" Then
        logFilePath = doc.Path & "\" & Format(Now, "yyyy-mm-dd") & "_" & baseName & "_FormattingLog.txt"
    Else
        logFilePath = Environ("TEMP") & "\" & Format(Now, "yyyy-mm-dd") & "_" & baseName & "_FormattingLog.txt"
    End If
    fnum = FreeFile
    Open logFilePath For Append As #fnum
    WriteLogLine "LOG DE FORMATAÇÃO"
    WriteLogLine "Sessão: " & Format(Now, "yyyy-mm-dd HH:MM:ss") & " | Usuário: " & Environ("USERNAME")
    WriteLogLine "Documento: " & doc.Name & IIf(doc.Path = "", " (Não salvo)", "")
    WriteLogLine String(40, "-")
    loggingEnabled = True
    executionStartTime = Now
    LogMessage "Logging iniciado", LOG_LEVEL_INFO
    InitializeLogging = True
    Exit Function
ErrorHandler:
    loggingEnabled = False
    If fnum <> 0 Then Close #fnum: fnum = 0
    InitializeLogging = False
End Function

Private Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error Resume Next
    ' If logging not enabled, allow ERROR fallback to file path if present
    If Not loggingEnabled Then
        If level = LOG_LEVEL_ERROR And logFilePath <> "" Then
            Dim errEntry As String
            errEntry = Format(Now, "yyyy-mm-dd HH:MM:ss") & " [ERRO] " & message
            Dim tf As Integer: tf = FreeFile
            Open logFilePath For Append As #tf
            Print #tf, errEntry
            Close #tf
        End If
        Exit Sub
    End If

    Dim mustLog As Boolean: mustLog = False
    If level >= LOG_LEVEL_WARNING Then
        mustLog = True
    Else
        ' INFO: log only high-value info keywords
        Dim low As String: low = LCase$(message)
        If InStr(low, "iniciado") > 0 Or InStr(low, "concluido") > 0 Or InStr(low, "erro") > 0 Or InStr(low, "falha") > 0 Or InStr(low, "cancel") > 0 Then
            mustLog = True
        End If
    End If

    If Not mustLog Then Exit Sub

    Dim label As String
    Select Case level
        Case LOG_LEVEL_INFO: label = "INFO"
        Case LOG_LEVEL_WARNING: label = "AVISO"
        Case LOG_LEVEL_ERROR: label = "ERRO"
        Case Else: label = "LOG"
    End Select

    WriteLogLine Format(Now, "yyyy-mm-dd HH:MM:ss") & " [" & label & "] " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error Resume Next
    If Not loggingEnabled Then Exit Sub
    Dim duration As Date: duration = Now - executionStartTime
    WriteLogLine String(30, "-")
    WriteLogLine "FIM DA SESSÃO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    WriteLogLine "Duração aproximada: " & Format(duration, "hh:nn:ss")
    WriteLogLine "Status: " & IIf(formattingCancelled, "CANCELADO", "CONCLUÍDO")
    If fnum <> 0 Then Close #fnum: fnum = 0
    loggingEnabled = False
End Sub

'================================================================================
' UTILITIES
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

Private Function GetDocumentSize(doc As Document) As String
    On Error Resume Next
    If doc.Path <> "" Then
        GetDocumentSize = Format(FileLen(doc.FullName) / 1024, "0.0") & " KB"
    Else
        Dim chars As Long: chars = 0
        chars = doc.Characters.Count
        GetDocumentSize = Format(chars / 1000, "0.0") & " KB (aprox.)"
    End If
End Function

Private Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then CentimetersToPoints = cm * 28.35
End Function

Private Function GetSafeUserName() As String
    On Error GoTo ErrorHandler
    Dim rawName As String, safeName As String
    rawName = Environ("USERNAME")
    If rawName = "" Then rawName = Environ("USER")
    If rawName = "" Then
        On Error Resume Next
        rawName = CreateObject("WScript.Network").UserName
        On Error GoTo ErrorHandler
    End If
    If rawName = "" Then rawName = "UsuarioDesconhecido"
    Dim i As Long, c As String, code As Long
    For i = 1 To Len(rawName)
        c = Mid$(rawName, i, 1)
        code = Asc(c)
        ' Digits (48-57), uppercase (65-90), lowercase (97-122), underscore or hyphen allowed.
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or c = "_" Or c = "-" Then
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
' PRE-CHECKS
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then
        MsgBox "Nenhum documento disponível.", vbCritical, "Erro"
        PreviousChecking = False: Exit Function
    End If
    If doc.Type <> wdTypeDocument Then
        MsgBox "Documento incompatível.", vbExclamation, "Tipo inválido"
        PreviousChecking = False: Exit Function
    End If
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "Documento protegido. Remova a proteção antes de prosseguir.", vbExclamation, "Protegido"
        PreviousChecking = False: Exit Function
    End If
    If doc.ReadOnly Then
        MsgBox "Documento em somente leitura. Salve uma cópia editável.", vbExclamation, "Somente leitura"
        PreviousChecking = False: Exit Function
    End If
    If Not CheckDiskSpace(doc) Then
        MsgBox "Espaço em disco insuficiente. Libere ~50MB e tente novamente.", vbExclamation, "Espaço insuficiente"
        PreviousChecking = False: Exit Function
    End If
    ' Validate structure but allow continuation
    If Not ValidateDocumentStructure(doc) Then
        ' continue but warn
    End If
    PreviousChecking = True
    Exit Function
ErrorHandler:
    MsgBox "Erro durante verificações: " & Err.Description, vbCritical, "Erro"
    PreviousChecking = False
End Function

Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim fso As Object, drive As Object, driveLetter As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If doc.Path <> "" Then driveLetter = Left(doc.Path, 3) Else driveLetter = Left(Environ("TEMP"), 3)
    Set drive = fso.GetDrive(driveLetter)
    If drive.AvailableSpace < 50 * 1024 * 1024 Then
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If
    Exit Function
ErrorHandler:
    CheckDiskSpace = True ' be permissive on failure to check
End Function

Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim valid As Boolean: valid = True
    If doc.Range.End = 0 Then valid = False
    If doc.Sections.Count = 0 Then valid = False
    ValidateDocumentStructure = valid
    Exit Function
ErrorHandler:
    ValidateDocumentStructure = False
End Function

'================================================================================
' MAIN FORMATTING PIPELINE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Remove leading blanks and check required string
    If Not RemoveLeadingBlankLinesAndCheckString(doc) Then
        PreviousFormatting = False: Exit Function
    End If

    DataAtualExtensoNoFinal doc ' informs user when missing

    If Not ApplyPageSetup(doc) Then GoTo ErrExit
    If Not ApplyStdFont(doc) Then GoTo ErrExit
    If Not ApplyStdParagraphs(doc) Then GoTo ErrExit
    EnableHyphenation doc ' non-fatal
    RemoveWatermark doc ' non-fatal
    InsertHeaderStamp doc ' non-fatal
    If Not InsertFooterStamp(doc) Then GoTo ErrExit

    PreviousFormatting = True
    Exit Function

ErrExit:
    PreviousFormatting = False
    Exit Function
ErrorHandler:
    LogMessage "Erro na formatação: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' PAGE SETUP
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
        .Orientation = wdOrientPortrait
    End With
    ApplyPageSetup = True: Exit Function
ErrorHandler:
    LogMessage "Erro ApplyPageSetup: " & Err.Description, LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

'================================================================================
' FONT FORMATTING (optimized: avoid Range.Text rewriting; use Font on Range)
'================================================================================
Private Function ApplyStdFont(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If para.Range.InlineShapes.Count = 0 Then
            With para.Range.Font
                .Name = STANDARD_FONT
                .Size = STANDARD_FONT_SIZE
                .Color = wdColorAutomatic
            End With
        End If
    Next para
    ApplyStdFont = True: Exit Function
ErrorHandler:
    LogMessage "Erro ApplyStdFont: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFont = False
End Function

'================================================================================
' PARAGRAPH FORMATTING (simplified & safer)
'================================================================================
Private Function ApplyStdParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If para.Range.InlineShapes.Count = 0 Then
            With para.Range.ParagraphFormat
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                .SpaceBefore = 0
                .SpaceAfter = 0
                If para.Alignment <> wdAlignParagraphCenter Then
                    .FirstLineIndent = CentimetersToPoints(1.5)
                    .LeftIndent = 0
                Else
                    .FirstLineIndent = 0
                    .LeftIndent = 0
                End If
            End With
            If para.Alignment = wdAlignParagraphLeft Then para.Alignment = wdAlignParagraphJustify
            ' normalize double spaces once
            If InStr(para.Range.Text, "  ") > 0 Then
                para.Range.Text = Replace(para.Range.Text, "  ", " ")
            End If
        End If
    Next para
    ApplyStdParagraphs = True: Exit Function
ErrorHandler:
    LogMessage "Erro ApplyStdParagraphs: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdParagraphs = False
End Function

'================================================================================
' HYPHENATION
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
    End If
    EnableHyphenation = True: Exit Function
ErrorHandler:
    EnableHyphenation = False
End Function

'================================================================================
' WATERMARK REMOVAL (simplified)
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim sec As Section, hf As HeaderFooter, shp As Shape, i As Long
    For Each sec In doc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                For i = hf.Shapes.Count To 1 Step -1
                    Set shp = hf.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                        End If
                    End If
                Next i
            End If
        Next hf
        For Each hf In sec.Footers
            If hf.Exists Then
                For i = hf.Shapes.Count To 1 Step -1
                    Set shp = hf.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                        End If
                    End If
                Next i
            End If
        Next hf
    Next sec
    RemoveWatermark = True: Exit Function
ErrorHandler:
    RemoveWatermark = False
End Function

'================================================================================
' HEADER IMAGE INSERTION (keeps behavior, improved path detection)
'================================================================================
Private Function InsertHeaderStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim sec As Section, header As HeaderFooter, imgFile As String
    imgFile = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH
    If Dir(imgFile) = "" Then imgFile = "C:\Users\" & GetSafeUserName() & HEADER_IMAGE_RELATIVE_PATH
    If Dir(imgFile) = "" Then imgFile = "\\server\Pictures\LegisTabStamp\HeaderStamp.png"
    If Dir(imgFile) = "" Then
        InsertHeaderStamp = False: Exit Function
    End If
    Dim imgWidth As Single: imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    Dim imgHeight As Single: imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO
    Dim shp As Shape
    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete
            Set shp = header.Shapes.AddPicture(FileName:=imgFile, LinkToFile:=False, SaveWithDocument:=msoTrue)
            If Not shp Is Nothing Then
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
            End If
        End If
    Next sec
    InsertHeaderStamp = True: Exit Function
ErrorHandler:
    InsertHeaderStamp = False
End Function

'================================================================================
' FOOTER: PAGE NUMBERS (corrected insertion order)
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim sec As Section, footer As HeaderFooter, rng As Range
    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range
            rng.Delete
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            ' Insert current page field
            doc.Fields.Add Range:=rng, Type:=wdFieldPage
            rng.Collapse Direction:=wdCollapseEnd
            rng.InsertAfter " de "
            rng.Collapse Direction:=wdCollapseEnd
            doc.Fields.Add Range:=rng, Type:=wdFieldNumPages
            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.Size = FOOTER_FONT_SIZE
                .Font.Bold = False
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
        End If
    Next sec
    InsertFooterStamp = True: Exit Function
ErrorHandler:
    InsertFooterStamp = False
End Function

'================================================================================
' REMOVE LEADING BLANK LINES AND CHECK REQUIRED STRING (safer & simpler)
'================================================================================
Private Function RemoveLeadingBlankLinesAndCheckString(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim loopCount As Long: loopCount = 0
    Do While doc.Paragraphs.Count > 0
        If Trim(doc.Paragraphs(1).Range.Text) = "" Then
            doc.Paragraphs(1).Range.Delete
            loopCount = loopCount + 1
            If loopCount > 50 Then Exit Do
        Else
            Exit Do
        End If
    Loop
    If doc.Paragraphs.Count = 0 Then
        MsgBox "Documento vazio após remoção de linhas em branco.", vbExclamation, "Documento Vazio"
        RemoveLeadingBlankLinesAndCheckString = False: Exit Function
    End If
    Dim firstLineText As String: firstLineText = doc.Paragraphs(1).Range.Text
    If InStr(1, firstLineText, REQUIRED_STRING, vbBinaryCompare) = 0 Then
        MsgBox "ATENÇÃO: A variável " & REQUIRED_STRING & " NÃO foi encontrada na primeira linha.", vbExclamation, "String Obrigatória"
        RemoveLeadingBlankLinesAndCheckString = True ' permitimos continuar, but user warned
    Else
        RemoveLeadingBlankLinesAndCheckString = True
    End If
    Exit Function
ErrorHandler:
    RemoveLeadingBlankLinesAndCheckString = False
End Function

'================================================================================
' DATE CHECK (simplified)
'================================================================================
Private Function DataAtualExtensoNoFinal(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim dataHoje As Date: dataHoje = Date
    Dim v1 As String, v2 As String
    v1 = LCase$(Format(dataHoje, "d 'de' mmmm 'de' yyyy"))
    v2 = LCase$(Format(dataHoje, "dd 'de' mmmm 'de' yyyy"))
    Dim i As Long, textoPara As String, found As Boolean
    For i = 1 To doc.Paragraphs.Count
        textoPara = LCase$(Trim(doc.Paragraphs(i).Range.Text))
        If Right$(textoPara, Len(v1)) = v1 Or Right$(textoPara, Len(v2)) = v2 Then
            found = True: Exit For
        End If
    Next i
    If Not found Then
        MsgBox "ATENÇÃO: A data atual por extenso NÃO foi encontrada ao final de nenhum parágrafo.", vbExclamation, "Data Não Localizada"
        DataAtualExtensoNoFinal = False
    Else
        DataAtualExtensoNoFinal = True
    End If
    Exit Function
ErrorHandler:
    DataAtualExtensoNoFinal = False
End Function

'================================================================================
' SAVE DOCUMENT FIRST (improved wait)
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc.Path <> "" Then
        SaveDocumentFirst = True: Exit Function
    End If
    If Application.Dialogs(wdDialogFileSaveAs).Show <> -1 Then
        SaveDocumentFirst = False: Exit Function
    End If
    Dim t0 As Single: t0 = Timer
    Do While doc.Path = "" And Timer - t0 < 10
        DoEvents
    Loop
    If doc.Path = "" Then SaveDocumentFirst = False Else SaveDocumentFirst = True
    Exit Function
ErrorHandler:
    SaveDocumentFirst = False
End Function

'================================================================================
' BACKUP (keeps behavior)
'================================================================================
Private Sub CreateBackup(doc As Document)
    On Error GoTo ErrorHandler
    If doc.Path = "" Then Exit Sub
    Dim backupPath As String
    backupPath = doc.Path & "\Backup_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & "_" & doc.Name
    doc.SaveAs2 backupPath
    Exit Sub
ErrorHandler:
    ' silent fail
End Sub

'================================================================================
' MISC UTILITIES (Open log, find recent log, show path, backups folder)
'================================================================================
Public Sub AbrirLog()
    On Error GoTo ErrorHandler
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim pathToOpen As String
    If logFilePath <> "" And Dir(logFilePath) <> "" Then
        pathToOpen = logFilePath
    Else
        pathToOpen = EncontrarArquivoLogRecente()
        If pathToOpen = "" Then MsgBox "Nenhum log encontrado.", vbInformation: Exit Sub
    End If
    If Dir(pathToOpen) = "" Then MsgBox "Arquivo de log não encontrado: " & pathToOpen, vbExclamation: Exit Sub
    shell.Run "notepad.exe " & Chr(34) & pathToOpen & Chr(34), 1, True
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao abrir log: " & Err.Description, vbExclamation
End Sub

Private Function EncontrarArquivoLogRecente() As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempFolder As Object: Set tempFolder = fso.GetFolder(Environ("TEMP"))
    Dim file As Object, recentFile As Object, recentDate As Date
    For Each file In tempFolder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "txt" Then
            If InStr(1, file.Name, "FormattingLog", vbTextCompare) > 0 Then
                If recentFile Is Nothing Then
                    Set recentFile = file: recentDate = file.DateLastModified
                ElseIf file.DateLastModified > recentDate Then
                    Set recentFile = file: recentDate = file.DateLastModified
                End If
            End If
        End If
    Next file
    If Not recentFile Is Nothing Then EncontrarArquivoLogRecente = recentFile.Path Else EncontrarArquivoLogRecente = ""
End Function

Public Sub MostrarCaminhoDoLog()
    On Error GoTo ErrorHandler
    Dim path As String
    If logFilePath <> "" And Dir(logFilePath) <> "" Then
        path = logFilePath
    Else
        path = EncontrarArquivoLogRecente()
        If path = "" Then MsgBox "Nenhum arquivo de log encontrado.", vbInformation: Exit Sub
    End If
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Arquivo de log: " & vbCrLf & path & vbCrLf & vbCrLf & "Abrir agora?", vbQuestion + vbYesNo, "Log")
    If resp = vbYes Then AbrirLog
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao localizar log: " & Err.Description, vbExclamation
End Sub

Public Sub AbrirPastaBackups()
    On Error GoTo ErrorHandler
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then MsgBox "Nenhum documento aberto.", vbExclamation: Exit Sub
    Dim folderPath As String
    If doc.Path <> "" Then folderPath = doc.Path Else folderPath = Environ("USERPROFILE") & "\Documents"
    If Not fso.FolderExists(folderPath) Then MsgBox "Pasta não encontrada: " & folderPath, vbExclamation: Exit Sub
    shell.Run "explorer.exe " & Chr(34) & folderPath & Chr(34), 1, True
    Exit Sub
ErrorHandler:
    MsgBox "Erro ao abrir pasta de backups: " & Err.Description, vbExclamation
End Sub

'================================================================================
' FINISHING HELPERS
'================================================================================
Private Sub ShowCompletionMessage(Optional ByVal sucesso As Boolean = True)
    Dim msg As String, resp As VbMsgBoxResult
    If sucesso Then
        msg = "Padronização concluída com sucesso." & vbCrLf & "Deseja abrir o LOG?"
        resp = MsgBox(msg, vbInformation + vbYesNo, "Concluído")
        If resp = vbYes Then AbrirLog
    Else
        msg = "Padronização parcial/erro. Deseja abrir o LOG?"
        resp = MsgBox(msg, vbExclamation + vbYesNo, "Parcial")
        If resp = vbYes Then AbrirLog
    End If
End Sub

'================================================================================
' TEXT TRANSFORMS (examples preserved, logs reduced)
'================================================================================
Private Function RemoveLeadingLinesAtEnd(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim i As Long
    For i = doc.Paragraphs.Count To 1 Step -1
        If Trim(doc.Paragraphs(i).Range.Text) = "" Then
            doc.Paragraphs(i).Range.Delete
        Else
            Exit For
        End If
    Next i
    RemoveLeadingLinesAtEnd = True: Exit Function
ErrorHandler:
    RemoveLeadingLinesAtEnd = False
End Function

Private Function ReplaceSugereWithIndicaInEmenta(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, i As Long, paraText As String
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = LCase$(Trim(para.Range.Text))
        If InStr(paraText, "ementa:") > 0 And InStr(paraText, "indicação") > 0 Then
            If InStr(paraText, "sugere") > 0 Then para.Range.Text = Replace(para.Range.Text, "sugere", "indica", , , vbTextCompare)
        End If
    Next i
    ReplaceSugereWithIndicaInEmenta = True: Exit Function
ErrorHandler:
    ReplaceSugereWithIndicaInEmenta = False
End Function

Private Function FormatConsideringParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, i As Long
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        If LCase$(Trim(para.Range.Text)) Like "considerando*" Then
            If para.Range.InlineShapes.Count = 0 Then
                With para.Range.Font
                    .Name = STANDARD_FONT
                    .Size = STANDARD_FONT_SIZE
                    .Bold = True
                    .AllCaps = True
                End With
            End If
        End If
    Next i
    FormatConsideringParagraphs = True: Exit Function
ErrorHandler:
    FormatConsideringParagraphs = False
End Function

'================================================================================
' APPLICATION STATE (performance helper)
' - SetAppState False: desativa ScreenUpdating e silencia alertas, mostra status.
' - SetAppState True: restaura estados padrão.
'================================================================================
Private Sub SetAppState(ByVal enable As Boolean, Optional ByVal statusMsg As String = "")
    On Error Resume Next
    If enable Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = wdAlertsAll
        Application.StatusBar = False
    Else
        Application.ScreenUpdating = False
        Application.DisplayAlerts = wdAlertsNone
        If Len(Trim$(statusMsg)) > 0 Then
            Application.StatusBar = statusMsg
        Else
            Application.StatusBar = "Processando..."
        End If
    End If
End Sub

