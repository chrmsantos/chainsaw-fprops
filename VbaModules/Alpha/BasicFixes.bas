Option Explicit

'================================================================================
' DOCUMENT FORMATTING TOOL
'================================================================================
' Description: Standardizes document formatting to formal specifications
' Compatibility: Microsoft Word 2010 and later versions
' Author: [Your Name]
' Version: 1.5
' Last Modified: [Date]
'================================================================================

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
Private Const BOTTOM_MARGIN_CM As Double = 3# ' Bottom margin in cm
Private Const LEFT_MARGIN_CM As Double = 3# ' Left margin in cm
Private Const RIGHT_MARGIN_CM As Double = 3# ' Right margin in cm
Private Const HEADER_DISTANCE_CM As Double = 0.7 ' Distance from header to content in cm
Private Const FOOTER_DISTANCE_CM As Double = 0.7 ' Distance from footer to content in cm

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\RevisorDeProposituras\Personalizations\DefaultHeader.png" ' Relative path to the header image
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Single = 17 ' Maximum width of the header image in cm
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Single = 0.27 ' Top margin for the header image in cm
Private Const HEADER_IMAGE_HEIGHT_RATIO As Single = 0.21 ' Height-to-width ratio for the header image

'================================================================================
' MAIN PROCEDURE: BasicFixes
'================================================================================
' Purpose: Orchestrates the document formatting process by calling various helper
' functions to apply standard formatting, clean up spacing, and insert headers.
'================================================================================
Public Sub BasicFixes()
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Verifica se a versão do Word é 2007 ou superior
    If Application.Version < 12 Then
        MsgBox "Este script requer o Microsoft Word 2007 ou superior.", vbExclamation, "Versão Incompatível"
        Exit Sub
    End If
    
    ' Validate document state
    If Not IsDocumentValid() Then Exit Sub ' Exit if the document is invalid
    
    Dim doc As Document ' Variable to hold the active document
    Set doc = ActiveDocument
    
    ' Verifica se o documento está protegido
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "O documento está protegido. Por favor, desproteja-o antes de continuar.", _
               vbExclamation, "Documento Protegido"
        Exit Sub
    End If
    
    ' Optimize performance by disabling screen updates
    With Application
        .ScreenUpdating = False
        .StatusBar = "Formatting document..."
    End With
    
    ' Execute formatting steps
    ResetBasicFormatting doc ' Reset basic formatting
    RemoveBlankLines doc ' Remove all blank lines
    RemoveLeadingBlankLines doc ' Remove leading blank lines
    CleanDocumentSpacing doc ' Clean up document spacing
    ApplyStandardFormatting doc ' Apply standard formatting
    RemoveAllWatermarks doc ' Remove watermarks
    InsertStandardHeaderImage doc ' Insert standard header image
    EnsureBlankLineBelowTextParagraphs doc ' Ensure blank line below text paragraphs
    
    ' Restore application state
    With Application
        .ScreenUpdating = True
        .StatusBar = False
    End With
    
    ' Notify the user of completion
    MsgBox "Document formatting completed successfully.", _
           vbInformation, "Formatting Complete"
    
    ' Limpeza de variáveis
    Set doc = Nothing
    Exit Sub ' Exit the procedure
    
ErrorHandler:
    ' Handle errors and restore application state
    HandleError "BasicFixes"
    With Application
        .ScreenUpdating = True
        .StatusBar = False
    End With
    Set doc = Nothing
End Sub

'================================================================================
' DOCUMENT VALIDATION FUNCTIONS
'================================================================================

'================================================================================
' IsDocumentValid
' Purpose: Validates the state of the active document to ensure it is suitable
' for formatting.
'================================================================================
Private Function IsDocumentValid() As Boolean
    ' Check if any document is open
    If Documents.Count = 0 Then
        MsgBox "No document is currently open.", vbExclamation, "Document Required"
        Exit Function
    End If
    
    ' Check if the active window contains a valid Word document
    If Not TypeOf ActiveDocument Is Document Then
        MsgBox "The active window does not contain a valid Word document.", _
               vbExclamation, "Invalid Document Type"
        Exit Function
    End If
    
    ' Check if the document contains any text
    If Len(Trim(ActiveDocument.Content.text)) = 0 Then
        MsgBox "The document contains no text to format.", _
               vbExclamation, "Empty Document"
        Exit Function
    End If
    
    IsDocumentValid = True ' Return True if all checks pass
End Function

'================================================================================
' FORMATTING FUNCTIONS
'================================================================================

'================================================================================
' RemoveLeadingBlankLines
' Purpose: Removes blank paragraphs at the beginning of the document.
'================================================================================
Private Sub RemoveLeadingBlankLines(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling
    
    Dim firstPara As Paragraph ' Variable to hold the first paragraph
    
    ' Check if the document contains any paragraphs
    If doc.Paragraphs.Count = 0 Then Exit Sub ' Exit if no paragraphs exist
    
    ' Loop through and remove blank paragraphs at the beginning
    Set firstPara = doc.Paragraphs(1)
    Do While Len(Trim(firstPara.Range.Text)) = 0 ' Check if the paragraph is blank
        firstPara.Range.Delete ' Delete the blank paragraph
        If doc.Paragraphs.Count = 0 Then Exit Do ' Exit if no more paragraphs exist
        Set firstPara = doc.Paragraphs(1) ' Update the first paragraph
    Loop
    
    Exit Sub ' Exit the function
    
ErrorHandler:
    ' Handle errors
    HandleError "RemoveLeadingBlankLines"
End Sub

'================================================================================
' CleanDocumentSpacing
' Purpose: Cleans up unnecessary spaces and paragraph breaks in the document.
'================================================================================
Private Sub CleanDocumentSpacing(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling
    
    Dim searchRange As Range ' Variable to hold the search range
    
    ' Check if the document is protected
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "Document is protected. Please unprotect it before formatting.", _
               vbExclamation, "Document Protected"
        Exit Sub
    End If
    
    Set searchRange = doc.Content ' Set the search range to the entire document content
    
    ' Replace multiple spaces with a single space
    With searchRange.Find
        .ClearFormatting
        .text = "  " ' Two spaces
        .Replacement.text = " " ' Single space
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Replace multiple paragraph breaks with a single break
    With searchRange.Find
        .text = "^p^p" ' Two paragraph marks
        .Replacement.text = "^p" ' Single paragraph mark
        .Execute Replace:=wdReplaceAll
    End With
    
    Exit Sub ' Exit the function
    
ErrorHandler:
    ' Handle errors
    HandleError "CleanDocumentSpacing"
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
' ResetBasicFormatting
' Purpose: Resets all direct formatting in the document to its default state.
'================================================================================
Private Sub ResetBasicFormatting(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Reset all direct formatting
    doc.Content.Font.Reset
    doc.Content.ParagraphFormat.Reset
    
    Exit Sub ' Exit the function
    
ErrorHandler:
    ' Handle errors
    HandleError "ResetBasicFormatting"
End Sub

'================================================================================
' RemoveAllWatermarks
' Purpose: Removes all watermarks from the document by deleting shapes in headers.
'================================================================================
Private Sub RemoveAllWatermarks(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling
    
    Dim sec As section ' Variable to hold each section
    Dim hdr As HeaderFooter ' Variable to hold each header/footer
    Dim shp As shape ' Variable to hold each shape
    
    ' Loop through all sections and headers
    For Each sec In doc.Sections
        For Each hdr In sec.Headers
            ' Remove all shapes in headers
            For Each shp In hdr.Shapes
                shp.Delete ' Delete the shape
            Next shp
        Next hdr
    Next sec
    
    Exit Sub ' Exit the function
    
ErrorHandler:
    ' Handle errors
    HandleError "RemoveAllWatermarks"
End Sub

'================================================================================
' InsertStandardHeaderImage
' Purpose: Inserts a standard header image into the document's headers.
'================================================================================
Private Sub InsertStandardHeaderImage(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling
    
    Dim sec As section ' Variable to hold each section
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
' ApplySpellingAndGrammarCorrections
' Purpose: Automatically applies all spelling and grammar corrections suggested
' by Word's native proofing tools without opening dialog boxes.
'================================================================================
Private Sub ApplySpellingAndGrammarCorrections(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling

    Dim errorRange As Range
    Dim suggestions As SpellingSuggestions
    Dim correctedText As String

    ' Loop through all spelling errors and correct them
    For Each errorRange In doc.SpellingErrors
        Set suggestions = Application.GetSpellingSuggestions(errorRange.Text) ' Get suggestions
        If suggestions.Count > 0 Then
            correctedText = suggestions(1).Name ' Use the first suggestion
            errorRange.Text = correctedText ' Replace the error with the suggestion
        End If
    Next errorRange

    ' Loop through all grammar errors and correct them
    For Each errorRange In doc.GrammaticalErrors
        ' Grammar errors do not have a direct suggestion API, so handle them differently if needed
        ' For now, we skip grammar corrections
    Next errorRange

    Exit Sub ' Exit the function

ErrorHandler:
    ' Handle errors
    HandleError "ApplySpellingAndGrammarCorrections"
End Sub

'================================================================================
' EnsureBlankLineBelowTextParagraphs
' Purpose: Ensures that every paragraph with text has a blank line below it.
'================================================================================
Private Sub EnsureBlankLineBelowTextParagraphs(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling

    Dim para As Paragraph
    Dim paraRange As Range

    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        Set paraRange = para.Range

        ' Check if the paragraph contains text (ignoring empty or whitespace-only paragraphs)
        If Len(Trim(paraRange.Text)) > 0 Then
            ' Check if the next character is not a paragraph mark (indicating no blank line)
            If Not paraRange.Characters.Last.Text = vbCr Then
                paraRange.InsertAfter vbCr ' Insert a blank line
            End If
        End If
    Next para

    Exit Sub ' Exit the function

ErrorHandler:
    ' Handle errors
    HandleError "EnsureBlankLineBelowTextParagraphs"
End Sub

'================================================================================
' RemoveBlankLines
' Purpose: Removes all blank lines (empty paragraphs) from the document.
'================================================================================
Private Sub RemoveBlankLines(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling

    Dim para As Paragraph
    Dim paraRange As Range

    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        Set paraRange = para.Range

        ' Check if the paragraph is empty or contains only whitespace
        If Len(Trim(paraRange.Text)) = 0 Then
            paraRange.Delete ' Delete the blank paragraph
        End If
    Next para

    Exit Sub ' Exit the function

ErrorHandler:
    ' Handle errors
    HandleError "RemoveBlankLines"
End Sub

'================================================================================
' HELPER FUNCTIONS
'================================================================================

'================================================================================
' CentimetersToPoints
' Purpose: Converts a value in centimeters to points.
'================================================================================
Private Function CentimetersToPoints(cm As Double) As Single
    CentimetersToPoints = Application.CentimetersToPoints(cm)
End Function

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

