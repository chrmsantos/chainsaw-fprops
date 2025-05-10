Option Explicit

'================================================================================
' DOCUMENT FORMATTING TOOL
'================================================================================
' Description: Standardizes document formatting to formal specifications
' Compatibility: Microsoft Word 2010 and later versions
' Author: [Your Name]
' Version: 1.4
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
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\Documents\Configurations\DefaultHeader.png" ' Relative path to the header image
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Single = 14.8 ' Maximum width of the header image in cm
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Single = 0.27 ' Top margin for the header image in cm
Private Const HEADER_IMAGE_HEIGHT_RATIO As Single = 0.22 ' Height-to-width ratio for the header image

'================================================================================
' MAIN PROCEDURE: FormatDocumentStandard
'================================================================================
' Purpose: Orchestrates the document formatting process by calling various helper
' functions to apply standard formatting, clean up spacing, and insert headers.
'================================================================================
Public Sub FormatDocumentStandard()
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Validate document state
    If Not IsDocumentValid() Then Exit Sub ' Exit if the document is invalid
    
    Dim doc As Document ' Variable to hold the active document
    Set doc = ActiveDocument
    
    ' Optimize performance by disabling screen updates
    With Application
        .ScreenUpdating = False
        .StatusBar = "Formatting document..."
    End With
    
    ' Execute formatting steps
    ResetBasicFormatting doc ' Reset basic formatting
    RemoveLeadingBlankLines doc ' Remove leading blank lines
    CleanDocumentSpacing doc ' Clean up document spacing
    ApplyStandardFormatting doc ' Apply standard formatting
    RemoveAllWatermarks doc ' Remove watermarks
    InsertStandardHeaderImage doc ' Insert standard header image
    
    ' Restore application state
    With Application
        .ScreenUpdating = True
        .StatusBar = False
    End With
    
    ' Notify the user of completion
    MsgBox "Document formatting completed successfully.", _
           vbInformation, "Formatting Complete"
    
    Exit Sub ' Exit the procedure
    
ErrorHandler:
    ' Handle errors and restore application state
    HandleError "FormatDocumentStandard"
    With Application
        .ScreenUpdating = True
        .StatusBar = False
    End With
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
    If Len(Trim(ActiveDocument.Content.Text)) = 0 Then
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
        .Text = "  " ' Two spaces
        .Replacement.Text = " " ' Single space
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Replace multiple paragraph breaks with a single break
    With searchRange.Find
        .Text = "^p^p" ' Two paragraph marks
        .Replacement.Text = "^p" ' Single paragraph mark
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

    ' Apply paragraph formatting to the main document content
    Dim para As Paragraph
    Dim paraIndex As Long: paraIndex = 1
    For Each para In doc.Paragraphs
        With para.Range.ParagraphFormat
            ' Apply standard formatting to all paragraphs except the third
            If paraIndex <> 3 Then
                .LeftIndent = 0
                .RightIndent = 0
                .FirstLineIndent = CentimetersToPoints(2.5) ' First line indent
                .Alignment = wdAlignParagraphJustify ' Justified alignment
            End If
            
            ' Apply specific formatting to the third paragraph
            If paraIndex = 3 Then
                .LeftIndent = CentimetersToPoints(9) ' Left indent of 9 cm
                .FirstLineIndent = 0 ' No first line indent
                .Alignment = wdAlignParagraphLeft ' Left alignment
            End If
            
            .SpaceBefore = 0
            .SpaceAfter = LINE_SPACING
            .LineSpacingRule = wdLineSpaceSingle
        End With
        
        paraIndex = paraIndex + 1 ' Increment the paragraph index
    Next para

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
    
    Dim sec As Section ' Variable to hold each section
    Dim hdr As HeaderFooter ' Variable to hold each header/footer
    Dim shp As Shape ' Variable to hold each shape
    
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
    errMsg = "Error in " & procedureName & ":" & vbCrLf & _
              "Error #" & Err.Number & vbCrLf & _
              Err.Description
    MsgBox errMsg, vbCritical, "Formatting Error" ' Display the error message
    Debug.Print errMsg ' Log the error message to the debug console
End Sub