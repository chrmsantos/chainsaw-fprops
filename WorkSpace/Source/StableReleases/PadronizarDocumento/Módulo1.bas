Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Word built-in constants (define if not referenced from Word object library)
Private Const wdNoProtection As Long = -1
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdAlignParagraphJustify As Long = 3
Private Const wdLineSpace1pt5 As Long = 1
Private Const wdLineSpacingMultiple As Long = 5
Private Const msoTrue As Long = -1
Private Const msoPicture As Long = 13
Private Const msoTextEffect As Long = 15
Private Const wdCollapseEnd As Long = 0
Private Const wdFieldPage As Long = 33
Private Const wdFieldNumPages As Long = 26
Private Const wdRelativeHorizontalPositionPage As Long = 1
Private Const wdRelativeVerticalPositionPage As Long = 1
Private Const wdWrapTopBottom As Long = 3

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

'================================================================================
' GLOBAL VARIABLES
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String

'================================================================================
' MAIN ENTRY POINT
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo ErrHandler
    
    ' Check version compatibility
    If Not CheckWordVersion() Then
        MsgBox "Este código é compatível apenas com Word 2010 ou superior.", vbExclamation, "Versão Não Suportada"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Initialize logging
    InitializeLogging doc
    
    LogMessage "Iniciando padronização do documento: " & doc.Name, LOG_LEVEL_INFO
    
    ' Start undo group
    StartUndoGroup "Padronizar Documento"
    
    SetAppState False, "Formatting document..."
    
    If Not PreviousChecking(doc) Then GoTo CleanUp
    
    PreviousFormatting doc
    
    Application.StatusBar = "Document standardized successfully!"
    LogMessage "Documento padronizado com sucesso!", LOG_LEVEL_INFO
    
CleanUp:
    ' End undo group
    EndUndoGroup
    
    ' Save and close log
    FinalizeLogging
    
    SetAppState True
    Exit Sub
    
ErrHandler:
    LogMessage "Erro na rotina principal: " & Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    Application.StatusBar = "Error: " & Err.Description
    MsgBox "An error occurred while standardizing the document:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume CleanUp
End Sub

'================================================================================
' VERSION COMPATIBILITY CHECK
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Long
    version = Application.Version
    
    If version < MIN_SUPPORTED_VERSION Then
        LogMessage "Versão do Word não suportada: " & version & " (mínima: " & MIN_SUPPORTED_VERSION & ")", LOG_LEVEL_ERROR
        CheckWordVersion = False
    Else
        LogMessage "Versão do Word compatível: " & version, LOG_LEVEL_INFO
        CheckWordVersion = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao verificar versão do Word: " & Err.Description, LOG_LEVEL_ERROR
    CheckWordVersion = False
End Function

'================================================================================
' UNDO GROUP MANAGEMENT
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    LogMessage "Iniciando grupo undo: " & groupName, LOG_LEVEL_INFO
End Sub

Private Sub EndUndoGroup()
    On Error Resume Next
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "Grupo undo finalizado", LOG_LEVEL_INFO
    End If
End Sub

'================================================================================
' LOGGING MANAGEMENT
'================================================================================
Private Sub InitializeLogging(doc As Document)
    On Error GoTo ErrorHandler
    
    ' Determine log file path
    If doc.Path <> "" Then
        logFilePath = doc.Path & "\" & Replace(doc.Name, ".docx", "") & "_FormattingLog.txt"
    Else
        logFilePath = Environ("TEMP") & "\DocumentFormattingLog.txt"
    End If
    
    ' Write initial log entry
    Open logFilePath For Output As #1
    Print #1, "================================================"
    Print #1, "LOG DE FORMATAÇÃO DE DOCUMENTO"
    Print #1, "Data/Hora: " & Now
    Print #1, "Documento: " & doc.FullName
    Print #1, "Usuário: " & Environ("USERNAME")
    Print #1, "Versão do Word: " & Application.Version
    Print #1, "================================================"
    Close #1
    
    loggingEnabled = True
    Exit Sub
    
ErrorHandler:
    loggingEnabled = False
    Debug.Print "Não foi possível inicializar logging: " & Err.Description
End Sub

Private Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error Resume Next
    
    If Not loggingEnabled Then Exit Sub
    
    Dim levelText As String
    Select Case level
        Case LOG_LEVEL_INFO
            levelText = "INFO"
        Case LOG_LEVEL_WARNING
            levelText = "AVISO"
        Case LOG_LEVEL_ERROR
            levelText = "ERRO"
        Case Else
            levelText = "OUTRO"
    End Select
    
    Open logFilePath For Append As #1
    Print #1, Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & levelText & "] " & message
    Close #1
    
    ' Also output to Immediate Window for debugging
    Debug.Print "LOG: " & levelText & " - " & message
End Sub

Private Sub FinalizeLogging()
    On Error Resume Next
    
    If loggingEnabled Then
        Open logFilePath For Append As #1
        Print #1, "================================================"
        Print #1, "FIM DO LOG"
        Print #1, "================================================"
        Close #1
    End If
    
    loggingEnabled = False
End Sub

'================================================================================
' APPLICATION STATE HANDLER
'================================================================================
Private Sub SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "")
    With Application
        .ScreenUpdating = enabled
        .DisplayAlerts = enabled
        .EnableEvents = enabled
        If statusMsg <> "" Then
            .StatusBar = statusMsg
        End If
    End With
End Sub

'================================================================================
' GLOBAL CHECKING
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        LogMessage "Nenhum documento ativo encontrado", LOG_LEVEL_ERROR
        MsgBox "No active document found. Please open a document to format.", vbExclamation, "Inactive Document"
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        LogMessage "Tipo de documento inválido: " & doc.Type, LOG_LEVEL_ERROR
        MsgBox "The active document is not a Word document. Please open a Word document to format.", vbExclamation, "Invalid Document Type"
        Exit Function
    End If

    ' Security: Check if document is protected
    If doc.ProtectionType <> wdNoProtection Then
        LogMessage "Documento protegido não pode ser formatado", LOG_LEVEL_ERROR
        MsgBox "The document is protected. Please unprotect it before formatting.", vbExclamation, "Protected Document"
        Exit Function
    End If
    
    ' Check if document is read-only
    If doc.ReadOnly Then
        LogMessage "Documento somente leitura não pode ser formatado", LOG_LEVEL_ERROR
        MsgBox "The document is read-only. Please save a copy before formatting.", vbExclamation, "Read-Only Document"
        Exit Function
    End If

    PreviousChecking = True
    LogMessage "Verificações preliminares passaram com sucesso", LOG_LEVEL_INFO
    Exit Function

ErrorHandler:
    HandleError "PreviousChecking"
    PreviousChecking = False
End Function

'================================================================================
' REMOVE BLANK LINES AT THE BEGINNING OF THE DOCUMENT
'================================================================================
Private Sub RemoveLeadingBlankLines(doc As Document)
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim deletedCount As Long
    
    LogMessage "Removendo linhas em branco iniciais", LOG_LEVEL_INFO
    
    ' Safely remove leading blank paragraphs
    Do While doc.Paragraphs.Count > 0
        Set para = doc.Paragraphs(1)
        If Trim(para.Range.Text) = vbCr Or Trim(para.Range.Text) = "" Then
            para.Range.Delete
            deletedCount = deletedCount + 1
            LogMessage "Parágrafo vazio removido: " & deletedCount, LOG_LEVEL_INFO
            ' Safety check to prevent infinite loop
            If deletedCount > 100 Then Exit Do
        Else
            Exit Do
        End If
    Loop
    
    LogMessage "Total de linhas em branco removidas: " & deletedCount, LOG_LEVEL_INFO
    Exit Sub
    
ErrorHandler:
    HandleError "RemoveLeadingBlankLines"
End Sub

'================================================================================
' MAIN FORMATTING ROUTINE
'================================================================================
Private Sub PreviousFormatting(doc As Document)
    On Error GoTo ErrorHandler

    LogMessage "Iniciando formatação principal do documento", LOG_LEVEL_INFO

    ' Remove blank lines at the beginning of the document
    RemoveLeadingBlankLines doc

    ' Apply formatting in logical order
    ApplyPageSetup doc
    ApplyFontAndParagraph doc
    EnableHyphenation doc
    RemoveWatermark doc
    InsertHeaderStamp doc
    InsertFooterStamp doc
    
    ' Save changes
    doc.Save
    LogMessage "Documento salvo após formatação", LOG_LEVEL_INFO

    Exit Sub

ErrorHandler:
    HandleError "PreviousFormatting"
End Sub

'================================================================================
' PAGE SETUP
'================================================================================
Private Sub ApplyPageSetup(doc As Document)
    On Error GoTo ErrorHandler
    
    LogMessage "Aplicando configurações de página", LOG_LEVEL_INFO
    
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
        .Gutter = 0
    End With
    
    LogMessage "Configurações de página aplicadas com sucesso", LOG_LEVEL_INFO
    Exit Sub
    
ErrorHandler:
    HandleError "ApplyPageSetup"
End Sub

'================================================================================
' FONT AND PARAGRAPH FORMATTING
'================================================================================
Private Sub ApplyFontAndParagraph(doc As Document)
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim currentIndent As Single
    Dim rightMarginPoints As Single
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long

    LogMessage "Aplicando formatação de fonte e parágrafo", LOG_LEVEL_INFO

    ' Calculate right indent based on document right margin
    rightMarginPoints = doc.PageSetup.RightMargin

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
            End With

            ' Apply paragraph formatting
            With para.Format
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                .RightIndent = rightMarginPoints

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

                ' Replace spacing before/after with line breaks
                If .SpaceBefore > 0 Then
                    .SpaceBefore = 0
                    If i > 1 Then
                        doc.Paragraphs(i - 1).Range.InsertAfter vbCr
                    End If
                End If
                
                If .SpaceAfter > 0 Then
                    .SpaceAfter = 0
                    para.Range.InsertAfter vbCr
                End If
            End With

            ' Justify left-aligned paragraphs
            If para.Alignment = wdAlignParagraphLeft Then
                para.Alignment = wdAlignParagraphJustify
            End If
            
            formattedCount = formattedCount + 1
        End If
    Next i
    
    LogMessage "Formatação concluída: " & formattedCount & " parágrafos formatados, " & skippedCount & " parágrafos com imagens ignorados", LOG_LEVEL_INFO
    Exit Sub
    
ErrorHandler:
    HandleError "ApplyFontAndParagraph"
End Sub

'================================================================================
' ENABLE HYPHENATION
'================================================================================
Private Sub EnableHyphenation(doc As Document)
    On Error GoTo ErrorHandler
    
    LogMessage "Ativando hifenização automática", LOG_LEVEL_INFO
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63) ' Standard zone
        LogMessage "Hifenização automática ativada", LOG_LEVEL_INFO
    Else
        LogMessage "Hifenização automática já estava ativada", LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError "EnableHyphenation"
End Sub

'================================================================================
' REMOVE WATERMARK
'================================================================================
Private Sub RemoveWatermark(doc As Document)
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
    Dim i As Long
    Dim removedCount As Long

    LogMessage "Removendo possíveis marcas d'água", LOG_LEVEL_INFO

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
                        End If
                    End If
                Next i
            End If
        Next header
    Next sec

    LogMessage "Removidas " & removedCount & " marcas d'água", LOG_LEVEL_INFO
    Exit Sub
    
ErrorHandler:
    HandleError "RemoveWatermark"
End Sub

'================================================================================
' INSERT HEADER IMAGE
'================================================================================
Private Sub InsertHeaderStamp(doc As Document)
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

    LogMessage "Inserindo carimbo no cabeçalho", LOG_LEVEL_INFO

    username = GetSafeUserName()
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Check if image exists
    If Dir(imgFile) = "" Then
        ' Try alternative paths
        imgFile = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH
        If Dir(imgFile) = "" Then
            LogMessage "Imagem de cabeçalho não encontrada em: C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH & " ou " & imgFile, LOG_LEVEL_ERROR
            MsgBox "Header image not found at:" & vbCrLf & _
                   "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH & vbCrLf & _
                   "or at: " & imgFile, vbExclamation, "Image Missing"
            Exit Sub
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
                SaveWithDocument:=msoTrue, _
                Left:=0, _
                Top:=0, _
                Width:=imgWidth, _
                Height:=imgHeight)

            With shp
                .LockAspectRatio = msoTrue
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
    Next sec

    If imgFound Then
        LogMessage "Carimbo inserido em " & sectionsProcessed & " seções", LOG_LEVEL_INFO
    Else
        LogMessage "Não foi possível inserir carimbo em nenhuma seção", LOG_LEVEL_WARNING
        MsgBox "Could not insert header image in any section.", vbExclamation, "Insertion Failed"
    End If

    Exit Sub

ErrorHandler:
    HandleError "InsertHeaderStamp"
End Sub

'================================================================================
' INSERT FOOTER PAGE NUMBERS
'================================================================================
Private Sub InsertFooterStamp(doc As Document)
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rng As Range
    Dim sectionsProcessed As Long
    Dim pageField As Field
    Dim totalPagesField As Field
    Dim totalPages As Long

    LogMessage "Inserindo numeração de página no rodapé", LOG_LEVEL_INFO
    
    ' Calculate total pages once for the entire document
    totalPages = doc.ComputeStatistics(wdStatisticPages)
    LogMessage "Total de páginas do documento: " & totalPages, LOG_LEVEL_INFO

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
            
            ' Insert current page number field
            Set pageField = rng.Fields.Add(Range:=rng, Type:=wdFieldPage)
            
            ' Insert hyphen
            rng.Collapse Direction:=wdCollapseEnd
            rng.Text = "-"
            
            ' Insert total pages as text (not field) to show document total
            rng.Collapse Direction:=wdCollapseEnd
            rng.Text = CStr(totalPages)
            
            ' Remove any extra paragraphs or line breaks
            If rng.Paragraphs.Count > 1 Then
                Do While rng.Paragraphs.Count > 1
                    rng.Paragraphs.Last.Range.Delete
                Loop
            End If
            
            ' Ensure no bold formatting anywhere in footer
            For Each rng In footer.Range.StoryRanges
                rng.Font.Bold = False
            Next rng
            
            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    LogMessage "Numeração de página inserida em " & sectionsProcessed & " seções. Formato: [página atual]-[total: " & totalPages & "]", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    HandleError "InsertFooterStamp"
End Sub

'================================================================================
' ERROR HANDLER
'================================================================================
Private Sub HandleError(procedureName As String)
    Dim errMsg As String
    errMsg = "Error in subroutine: " & procedureName & vbCrLf & _
             "Error #" & Err.Number & ": " & Err.Description & vbCrLf & _
             "Source: " & Err.Source
    Application.StatusBar = "Error: " & Err.Description
    LogMessage "Erro em " & procedureName & ": " & Err.Number & " - " & Err.Description, LOG_LEVEL_ERROR
    Err.Clear
End Sub

'================================================================================
' UTILITY: CM TO POINTS
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    CentimetersToPoints = Application.CentimetersToPoints(cm)
End Function

'================================================================================
' UTILITY: SAFE USERNAME
'================================================================================
Private Function GetSafeUserName() As String
    Dim rawName As String
    Dim safeName As String
    Dim i As Integer
    Dim c As String
    
    ' Try multiple methods to get username
    On Error Resume Next
    rawName = Environ("USERNAME")
    If rawName = "" Then rawName = Environ("USER")
    If rawName = "" Then rawName = CreateObject("WScript.Network").UserName
    On Error GoTo 0
    
    ' Sanitize username for path safety
    For i = 1 To Len(rawName)
        c = Mid(rawName, i, 1)
        If c Like "[A-Za-z0-9_\-]" Then
            safeName = safeName & c
        End If
    Next i
    
    GetSafeUserName = safeName
    LogMessage "Nome de usuário sanitizado: " & safeName, LOG_LEVEL_INFO
End Function