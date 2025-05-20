Public Sub Main_SDF(doc As Document)
    On Error GoTo ErrorHandler

    ' Setting format steps
    ApplyStandardFormatting doc ' Apply standard formatting
    EnsureBlankLineBelowTextParagraphs doc ' Ensure blank line below text paragraphs
    InsertStandardHeaderImage doc ' Insert standard header image
    FormatSpecificLines doc ' Format specific lines

    Exit Sub

ErrorHandler:
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
Private Function CentimetersToPoints(cm As Double) As Single
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
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With

    ' Apply paragraph formatting to the entire document
    With doc.Content.ParagraphFormat
        .SpaceAfter = 12 ' Add 12 points of space after each paragraph
        .LineSpacingRule = wdLineSpaceMultiple ' Set line spacing rule to multiple
        .LineSpacing = 1.15 * 12 ' Set line spacing to 1.15
        .Alignment = wdAlignParagraphJustify ' Justify alignment
        .FirstLineIndent = CentimetersToPoints(2.5) ' Set first line indent to 2.5 cm
    End With

    ' Format the first paragraph of the document
    If doc.Paragraphs.Count > 0 Then
        Dim firstPara As Paragraph
        Set firstPara = doc.Paragraphs(1)
        With firstPara.Range
            .Font.Bold = True ' Apply bold
            .Font.AllCaps = True ' Convert text to uppercase
            .ParagraphFormat.Alignment = wdAlignParagraphCenter ' Center alignment
            .ParagraphFormat.LeftIndent = 0 ' Remove left indent
            .ParagraphFormat.RightIndent = 0 ' Remove right indent
            .ParagraphFormat.FirstLineIndent = 0 ' Remove first line indent
        End With
    End If

    ' Format the second paragraph of the document
    If doc.Paragraphs.Count > 1 Then
        Dim secondPara As Paragraph
        Set secondPara = doc.Paragraphs(2)
        With secondPara.Range.ParagraphFormat
            .LeftIndent = 0 ' Remove any left indent
            .FirstLineIndent = CentimetersToPoints(9) ' Set first line indent to 9 cm
            .RightIndent = 0 ' Ensure no right indent
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
    
    Exit Sub ' Exit the function
    
ErrorHandler:
    ' Handle errors
    HandleError "ApplyStandardFormatting"
End Sub

'================================================================================
' InsertStandardHeaderImage
' Purpose: Inserts a standard header image into the document's headers.
'================================================================================
Private Sub InsertStandardHeaderImage(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling
    
    Dim sec As Section ' Variable to hold each section
    Dim header As HeaderFooter ' Variable to hold the primary header
    Dim imgFile As String ' Path to the header image
    Dim username As String ' Current username
    Dim imgWidth As Single ' Width of the image in points
    Dim imgHeight As Single ' Height of the image in points
    
    ' Get the current username from the environment variable
    username = Environ("USERNAME")
    
    ' Build the full path to the header image
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH
    
    ' Check if the image file exists
    If Dir(imgFile) = "" Then
        MsgBox "Header image not found at: " & vbCrLf & imgFile, vbExclamation, "Image Missing"
        Exit Sub
    End If
    
    ' Calculate proportional dimensions in points
    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO
    
    ' Loop through all sections and insert the header image
    For Each sec In doc.Sections
        ' Modify the primary header
        Set header = sec.Headers(wdHeaderFooterPrimary)
        
        ' Clear existing header content
        header.LinkToPrevious = False
        header.Range.Delete
        
        ' Insert and format the image with proportional sizing
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
            .LockAspectRatio = msoTrue ' Maintain aspect ratio
        End With
    Next sec
    
    Exit Sub ' Exit the function
    
ErrorHandler:
    ' Handle errors
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
                    ' Insere linha em branco somente se o próximo não for em branco
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

    maxParagraphs = 1000 ' Limite de segurança

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

        ' Formata a linha com "anexo" ou "anexos" (apenas se for a palavra isolada)
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

    ' Remove quebras de página extras
    With doc.Content.Find
        .Text = "^m^m" ' Duas quebras de página consecutivas
        .Replacement.Text = "^m" ' Uma quebra de página
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceAll)
            editCount = editCount + 1
        Loop
    End With

    RemoveExtraPageBreaks = editCount
End Function

