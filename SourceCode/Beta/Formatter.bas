Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Constants for Word operations
Private Const wdFindContinue As Long = 1 ' Continue search after the first match
Private Const wdReplaceOne As Long = 1 ' Replace only one occurrence
Private Const wdLineSpaceSingle As Long = 0 ' Single line spacing
Private Const STANDARD_FONT As String = "Arial" ' Standard font for the document
Private Const STANDARD_FONT_SIZE As Long = 12 ' Standard font size
Private Const LINE_SPACING As Long = 12 ' Line spacing in points

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 3.8 ' Top margin in cm
Private Const BOTTOM_MARGIN_CM As Double = 2.5   ' Bottom margin in cm
Private Const LEFT_MARGIN_CM As Double = 2.5     ' Left margin in cm
Private Const RIGHT_MARGIN_CM As Double = 2.5    ' Right margin in cm
Private Const HEADER_DISTANCE_CM As Double = 0.5 ' Distance from header to content in cm
Private Const FOOTER_DISTANCE_CM As Double = 0.5 ' Distance from footer to content in cm

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\SetStandardFormat\Personalization\StandardHeader.png" ' Relative path to the header image
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 19 ' Maximum width of the header image in cm
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.27 ' Top margin for the header image in cm
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.175 ' Height-to-width ratio for the header image

'================================================================================
' Main module for formatting
'================================================================================

' Entry point for macro button: applies formatting to the active document
Public Sub Formatter()
    On Error GoTo ErrorHandler

    ' Otimização de desempenho: desabilita atualizações de tela e alertas
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatando documento..."
    End With

    ' Apply formatting steps to the active document
    ApplyStandardFormatting ActiveDocument
    InsertStandardHeaderImage ActiveDocument

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
' ApplyStandardFormatting
' Purpose: Applies standard formatting to the document, including font, margins,
' and paragraph formatting.
'================================================================================
Private Sub ApplyStandardFormatting(doc As Document)
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
    HandleError "ApplyStandardFormatting"
End Sub

'================================================================================
' InsertStandardHeaderImage
' Purpose: Inserts a standard header image into the document's headers.
'================================================================================
Private Sub InsertStandardHeaderImage(doc As Document)
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
    HandleError "InsertStandardHeaderImage"
End Sub


