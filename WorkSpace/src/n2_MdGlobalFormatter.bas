Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Constants for Word operations
Private Const wdFindContinue As Long = 1 ' Continue search after the first match
Private Const wdReplaceOne As Long = 1 ' Replace only one occurrence
Private Const wdLineSpaceSingle As Long = 1.5 ' Standard Line spacing
Private Const STANDARD_FONT As String = "Arial" ' Standard font for the document
Private Const STANDARD_FONT_SIZE As Long = 12 ' Standard font size
Private Const LINE_SPACING As Long = 12 ' Line spacing in points

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 4 ' Top margin in cm
Private Const BOTTOM_MARGIN_CM As Double = 3   ' Bottom margin in cm
Private Const LEFT_MARGIN_CM As Double = 3     ' Left margin in cm
Private Const RIGHT_MARGIN_CM As Double = 3    ' Right margin in cm
Private Const HEADER_DISTANCE_CM As Double = 0.5 ' Distance from header to content in cm
Private Const FOOTER_DISTANCE_CM As Double = 1 ' Distance from footer to content in cm

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\Documents\OneDrive - Personal\Profissional" _
    "\c_2025\c_PubGithubRepos\c_LegisToolbox\Workspace\img\DefaultHeader.png" ' Relative path to the header image
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 20 ' Maximum width of the header image in cm
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.1 ' Top margin for the header image in cm
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.18 ' Height-to-width ratio for the header image

'================================================================================
' Main module for formatting
'================================================================================

' Entry point for macro button: applies formatting to the active document
Public Sub PageFormatterMain()
    On Error GoTo ErrorHandler

    ' Otimização de desempenho: desabilita atualizações de tela e alertas
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatando documento..."
    End With


    ' Apply formatting steps to the active document
    BasicFormatting ActiveDocument

    ' Remove any existing watermark shapes
    RemoveWatermark ActiveDocument

    ' Insere a imagem de cabeçalho padrão
    InsertHeaderImage ActiveDocument
    

    ' Restaura o estado da aplicação
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With

    Exit Sub

ErrorHandler:
    ' Garante restauração do estado mesmo em caso de erro
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    HandleError "Main"
End Sub

'================================================================================
' RemoveWatermark
' Purpose: Removes watermark shapes from all sections if present.
'================================================================================
Private Sub RemoveWatermark(doc As Document)
    On Error Resume Next

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
    Dim i As Long

    For Each sec In doc.Sections
        For Each header In sec.Headers
            For i = header.Shapes.Count To 1 Step -1
                Set shp = header.Shapes(i)
                If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                    If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Then
                        shp.Delete
                    End If
                End If
            Next i
        Next header
    Next sec

    On Error GoTo 0
End Sub

'================================================================================
' HandleError
' Purpose: Handles errors by displaying an error message and logging it to the
' debug console.
'================================================================================
Private Sub HandleError(procedureName As String)
    Dim errMsg As String ' Variable to hold the error message

    ' Build the error message
    errMsg = "Erro na sub-rotina: " & procedureName & vbCrLf & _
             "Erro #" & Err.Number & ": " & Err.Description

    ' Display the error message to the user
    MsgBox errMsg, vbCritical, "Erro de Formatação"

    ' Log the error message to the debug console
    Debug.Print errMsg

    ' Clear the error
    Err.Clear
End Sub

'================================================================================
' CentimetersToPoints
' Purpose: Converts a value in centimeters to points.
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    CentimetersToPoints = Application.CentimetersToPoints(cm)
End Function

'================================================================================
' BasicFormatting
' Purpose: Applies standard formatting to the document, including font, margins,
' and paragraph formatting.
'================================================================================
Private Sub BasicFormatting(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling

    ' Verifica se o documento está protegido
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "O documento está protegido. Por favor, desproteja-o antes de continuar.", _
               vbExclamation, "Documento Protegido"
        Exit Sub
    End If

    ' Set page layout and margins
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
    End With

    ' Apply font formatting to the entire document content
    With doc.Content.Font
        .Name = STANDARD_FONT
        .Size = STANDARD_FONT_SIZE
    End With

    Exit Sub

ErrorHandler:
    HandleError "BasicFormatting"
End Sub

'================================================================================
' InsertHeaderImage
' Purpose: Inserts a standard header image into the document's headers.
'================================================================================
Private Sub InsertHeaderImage(doc As Document)
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single

    ' Monta o caminho completo da imagem do cabeçalho
    username = Environ("USERNAME")
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Verifica se a imagem existe
    If Dir(imgFile) = "" Then
        MsgBox "Header image not found at: " & vbCrLf & imgFile, vbExclamation, "Image Missing"
        Exit Sub
    End If

    ' Calcula as dimensões da imagem
    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    ' Para cada seção, insere a imagem no cabeçalho
    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        header.LinkToPrevious = False

        ' Remove todo o conteúdo anterior do cabeçalho
        header.Range.Delete

        ' Limpa a formatação da fonte do cabeçalho
        With header.Range
            .Font.Reset
            .Font.Name = STANDARD_FONT
            .Font.Size = STANDARD_FONT_SIZE
            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
            .ParagraphFormat.LineSpacing = 1.5 * 12 ' 1.5 linhas (18 pontos)
        End With

        ' Adiciona a imagem e ajusta suas propriedades
        With header.Shapes.AddPicture( _
            fileName:=imgFile, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0, _
            Width:=imgWidth, _
            Height:=imgHeight)

            .WrapFormat.Type = wdWrapTight ' Quebra de texto do tipo "justa"
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .Left = wdShapeCenter
            .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
            .LockAspectRatio = msoTrue
        End With
    Next sec

    Exit Sub

ErrorHandler:
    HandleError "InsertHeaderImage"
End Sub


