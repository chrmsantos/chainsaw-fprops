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
Private Const TOP_MARGIN_CM As Double = 4.5 ' Top margin in cm
Private Const BOTTOM_MARGIN_CM As Double = 2   ' Bottom margin in cm
Private Const LEFT_MARGIN_CM As Double = 3     ' Left margin in cm
Private Const RIGHT_MARGIN_CM As Double = 3    ' Right margin in cm
Private Const HEADER_DISTANCE_CM As Double = 0.7 ' Distance from header to content in cm
Private Const FOOTER_DISTANCE_CM As Double = 0.7 ' Distance from footer to content in cm

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\RevisorDeProposituras\Personalizations\DefaultHeader.png" ' Relative path to the header image
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 17 ' Maximum width of the header image in cm
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.27 ' Top margin for the header image in cm
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.21 ' Height-to-width ratio for the header image

'================================================================================
' Main module for formatting
'================================================================================
Public Sub Main_SDF(doc As Document)
    On Error GoTo ErrorHandler

    ' Otimização de desempenho: desabilita atualizações de tela e alertas
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatando documento..."
    End With

    ' Setting format steps
    ApplyStandardFormatting doc ' Apply standard formatting
    RemoveBlankLines doc ' Remove blank lines
    EnsureBlankLineBelowTextParagraphs doc ' Ensure blank line below text paragraphs
    InsertStandardHeaderImage doc ' Insert standard header image
    FormatSpecificLines doc ' Format specific lines

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
    HandleError "Main_SDF"
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
' RemoveBlankLines
' Purpose: Removes all blank lines (empty paragraphs) from the document.
'================================================================================
Private Sub RemoveBlankLines(doc As Document)
    On Error GoTo ErrorHandler

    Dim i As Long
    For i = doc.Paragraphs.Count To 1 Step -1
        If Len(Trim(doc.Paragraphs(i).Range.Text)) = 0 Then
            doc.Paragraphs(i).Range.Delete
        End If
    Next i

    Exit Sub

ErrorHandler:
    HandleError "RemoveBlankLines"
End Sub

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
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With

    ' Apply paragraph formatting to the entire document
    With doc.Content.ParagraphFormat
        .SpaceAfter = 12
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = 1.15 * 12
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(2.5)
    End With

    ' Format the first paragraph of the document
    If doc.Paragraphs.Count > 0 Then
        Dim firstPara As Paragraph
        Set firstPara = doc.Paragraphs(1)
        With firstPara.Range
            .Font.Bold = True
            .Font.AllCaps = True
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.RightIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
        End With
    End If

    ' Format the second paragraph of the document
    If doc.Paragraphs.Count > 1 Then
        Dim secondPara As Paragraph
        Set secondPara = doc.Paragraphs(2)
        With secondPara.Range.ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = CentimetersToPoints(9)
            .RightIndent = 0
        End With
    End If

    ' Apply font formatting to headers and footers
    Dim sec As Section
    Dim hdrFtr As HeaderFooter
    For Each sec In doc.Sections
        ' Format headers
        For Each hdrFtr In sec.Headers
            If Len(Trim(hdrFtr.Range.Text)) > 0 Then
                With hdrFtr.Range.Font
                    .Name = STANDARD_FONT
                    .Size = STANDARD_FONT_SIZE
                    .Bold = False
                    .Italic = False
                    .Underline = wdUnderlineNone
                End With
            End If
        Next hdrFtr

        ' Format footers
        For Each hdrFtr In sec.Footers
            If Len(Trim(hdrFtr.Range.Text)) > 0 Then
                With hdrFtr.Range.Font
                    .Name = STANDARD_FONT
                    .Size = STANDARD_FONT_SIZE
                    .Bold = False
                    .Italic = False
                    .Underline = wdUnderlineNone
                End With
            End If
        Next hdrFtr
    Next sec

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

    username = Environ("USERNAME")
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    If Dir(imgFile) = "" Then
        MsgBox "Header image not found at: " & vbCrLf & imgFile, vbExclamation, "Image Missing"
        Exit Sub
    End If

    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        header.LinkToPrevious = False
        header.Range.Delete

        With header.Shapes.AddPicture( _
            fileName:=imgFile, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0, _
            Width:=imgWidth, _
            Height:=imgHeight)

            .WrapFormat.Type = wdWrapTopBottom
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

'================================================================================
' EnsureBlankLineBelowTextParagraphs
' Purpose: Ensures that every paragraph with text has a blank line below it.
'================================================================================
Private Sub EnsureBlankLineBelowTextParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    Dim i As Long
    For i = doc.Paragraphs.Count To 1 Step -1
        Dim para As Paragraph
        Set para = doc.Paragraphs(i)
        Dim paraText As String
        paraText = Trim(para.Range.Text)
        If Len(paraText) > 0 Then
            If i = doc.Paragraphs.Count Then
                ' Último parágrafo: não faz nada
            Else
                Dim nextParaText As String
                nextParaText = Trim(doc.Paragraphs(i + 1).Range.Text)
                If Len(nextParaText) > 0 Then
                    doc.Paragraphs(i).Range.InsertAfter vbCr
                End If
            End If
        End If
    Next i

    Exit Sub

ErrorHandler:
    HandleError "EnsureBlankLineBelowTextParagraphs"
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

    maxParagraphs = 1000

    paraIndex = 1
    For Each para In doc.Paragraphs
        If paraIndex > maxParagraphs Then Exit For

        paraText = Trim(Replace(para.Range.Text, vbCr, ""))

        ' Formata a primeira linha do texto
        If paraIndex = 1 Then
            With para.Range
                .Font.Bold = True
                .Font.Underline = wdUnderlineSingle
                .Font.AllCaps = True
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.LeftIndent = 0
                .ParagraphFormat.RightIndent = 0
                .ParagraphFormat.FirstLineIndent = 0
            End With
        End If

        ' Formata a linha com "justificativa" ou "justificativas"
        If LCase(paraText) = "justificativa" Or LCase(paraText) = "justificativas" Then
            With para.Range
                .Font.Bold = True
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.LeftIndent = 0
                .ParagraphFormat.RightIndent = 0
                .ParagraphFormat.FirstLineIndent = 0
            End With
        End If

        ' Formata a linha com "anexo" ou "anexos"
        If LCase(paraText) = "anexo" Or LCase(paraText) = "anexos" Then
            With para.Range
                .Font.Bold = True
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.LeftIndent = 0
                .ParagraphFormat.RightIndent = 0
                .ParagraphFormat.FirstLineIndent = 0
            End With
        End If

        paraIndex = paraIndex + 1
    Next para

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao formatar linhas específicas: " & Err.Description, vbCritical, "Erro"
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: RemoveExtraPageBreaks
' Finalidade: Remove quebras de página extras no documento.
'--------------------------------------------------------------------------------
Private Function RemoveExtraPageBreaks(doc As Document) As Integer
    On Error Resume Next

    Dim editCount As Integer: editCount = 0

    With doc.Content.Find
        .Text = "^m^m"
        .Replacement.Text = "^m"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceAll)
            editCount = editCount + 1
        Loop
    End With

    RemoveExtraPageBreaks = editCount
End Function

