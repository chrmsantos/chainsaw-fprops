Option Explicit

'================================================================================
' DOCUMENT FORMATTING TOOL
'================================================================================
' Description: Standardizes document formatting to formal specifications
' Compatibility: Microsoft Word 2010 and later versions
' Author: [Your Name]
' Version: 1.2
' Last Modified: [Date]
'================================================================================

' Constants for Word operations
Private Const wdFindContinue As Long = 1
Private Const wdReplaceOne As Long = 1
Private Const wdLineSpaceSingle As Long = 0
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const LINE_SPACING As Long = 12 ' Points

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 4.5
Private Const BOTTOM_MARGIN_CM As Double = 3.0
Private Const LEFT_MARGIN_CM As Double = 3.0
Private Const RIGHT_MARGIN_CM As Double = 3.0
Private Const HEADER_DISTANCE_CM As Double = 0.7
Private Const FOOTER_DISTANCE_CM As Double = 0.7

' Header image constants
Private Const HEADER_IMAGE_PATH As String = "C:\Users\YourUsername\Documents\Configurations\DefaultHeader.png"
Private Const HEADER_IMAGE_WIDTH As Single = 468 ' 6.5 inches in points
Private Const HEADER_IMAGE_HEIGHT As Single = 72 ' 1 inch in points

'================================================================================
' MAIN PROCEDURE: FormatDocumentStandard
' Purpose: Orchestrates the document formatting process
'================================================================================
Public Sub FormatDocumentStandard()
    On Error GoTo ErrorHandler
    
    ' Validate document state
    If Not IsDocumentValid() Then Exit Sub
    
    Dim doc As Document
    Set doc = ActiveDocument
    Dim editCount As Long
    
    ' Optimize performance
    With Application
        .ScreenUpdating = False
        .StatusBar = "Formatting document..."
    End With
    
    ' Execute formatting steps
    editCount = editCount + RemoveLeadingBlankLines(doc)
    editCount = editCount + CleanDocumentSpacing(doc)
    editCount = editCount + ApplyStandardFormatting(doc)
    editCount = editCount + ResetBasicFormatting(doc)
    editCount = editCount + RemoveAllWatermarks(doc)
    editCount = editCount + InsertStandardHeaderImage(doc)
    
    ' Restore application state
    With Application
        .ScreenUpdating = True
        .StatusBar = False
    End With
    
    ' Completion notification
    MsgBox "Document formatting completed successfully." & vbCrLf & _
           "Total modifications: " & editCount, _
           vbInformation, "Formatting Complete"
    
    Exit Sub
    
ErrorHandler:
    HandleError "FormatDocumentStandard"
    With Application
        .ScreenUpdating = True
        .StatusBar = False
    End With
End Sub

'================================================================================
' DOCUMENT VALIDATION FUNCTIONS
'================================================================================

Private Function IsDocumentValid() As Boolean
    If Documents.Count = 0 Then
        MsgBox "No document is currently open.", vbExclamation, "Document Required"
        Exit Function
    End If
    
    If Not TypeOf ActiveDocument Is Document Then
        MsgBox "The active window does not contain a valid Word document.", _
               vbExclamation, "Invalid Document Type"
        Exit Function
    End If
    
    If Len(Trim(ActiveDocument.Content.Text)) = 0 Then
        MsgBox "The document contains no text to format.", _
               vbExclamation, "Empty Document"
        Exit Function
    End If
    
    IsDocumentValid = True
End Function

'================================================================================
' FORMATTING FUNCTIONS
'================================================================================

Private Function RemoveLeadingBlankLines(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    Dim edits As Long: edits = 0
    Dim firstPara As Paragraph
    
    ' Check for paragraphs
    If doc.Paragraphs.Count = 0 Then Exit Function
    
    ' Remove leading blank paragraphs
    Set firstPara = doc.Paragraphs(1)
    Do While Len(Trim(firstPara.Range.Text)) = 0
        firstPara.Range.Delete
        edits = edits + 1
        If doc.Paragraphs.Count = 0 Then Exit Do
        Set firstPara = doc.Paragraphs(1)
    Loop
    
    RemoveLeadingBlankLines = edits
    Exit Function
    
ErrorHandler:
    HandleError "RemoveLeadingBlankLines"
    RemoveLeadingBlankLines = edits
End Function

Private Function CleanDocumentSpacing(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    Dim edits As Long: edits = 0
    Dim searchRange As Range
    
    ' Check if document is protected
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "Document is protected. Please unprotect it before formatting.", _
               vbExclamation, "Document Protected"
        Exit Function
    End If
    
    Set searchRange = doc.Content
    
    ' First pass: Replace multiple spaces with single space
    With searchRange.Find
        .ClearFormatting
        .Text = "  "  ' Two spaces
        .Replacement.Text = " "  ' Single space
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
        edits = edits + 1
    End With
    
    ' Second pass: Replace remaining multiple spaces
    With searchRange.Find
        .Text = "  "  ' Two spaces
        .Execute Replace:=wdReplaceAll
        edits = edits + 1
    End With
    
    ' Handle page breaks
    With searchRange.Find
        .ClearFormatting
        .Text = "^m^m"  ' Two manual page breaks
        .Replacement.Text = "^m"  ' Single page break
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
        edits = edits + 1
    End With
    
    ' Handle paragraph breaks
    With searchRange.Find
        .Text = "^p^p"  ' Two paragraph marks
        .Replacement.Text = "^p"  ' Single paragraph mark
        .Execute Replace:=wdReplaceAll
        edits = edits + 1
    End With
    
    CleanDocumentSpacing = edits
    Exit Function
    
ErrorHandler:
    HandleError "CleanDocumentSpacing"
    CleanDocumentSpacing = edits
End Function

Private Function ApplyStandardFormatting(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    Dim edits As Long: edits = 0
    
    ' Set page layout
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
    End With
    
    ' Apply paragraph formatting
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        With para.Range
            ' Font formatting
            With .Font
                .Name = STANDARD_FONT
                .Size = STANDARD_FONT_SIZE
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .LeftIndent = 0
                .RightIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = LINE_SPACING
                .LineSpacingRule = wdLineSpaceSingle
                .Alignment = wdAlignParagraphLeft
                .FirstLineIndent = 0
            End With
        End With
        edits = edits + 1
    Next para
    
    ApplyStandardFormatting = edits
    Exit Function
    
ErrorHandler:
    HandleError "ApplyStandardFormatting"
    ApplyStandardFormatting = edits
End Function

Private Function ResetBasicFormatting(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    ' Reset all direct formatting
    doc.Content.Font.Reset
    doc.Content.ParagraphFormat.Reset
    
    ResetBasicFormatting = 1 ' Count as one operation
    Exit Function
    
ErrorHandler:
    HandleError "ResetBasicFormatting"
    ResetBasicFormatting = 0
End Function

Private Function RemoveAllWatermarks(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    Dim edits As Long: edits = 0
    Dim sec As Section
    Dim hdr As HeaderFooter
    Dim shp As Shape
    
    ' Process all sections and headers
    For Each sec In doc.Sections
        For Each hdr In sec.Headers
            ' Remove all shapes in headers
            For Each shp In hdr.Shapes
                shp.Delete
                edits = edits + 1
            Next shp
        Next hdr
    Next sec
    
    RemoveAllWatermarks = edits
    Exit Function
    
ErrorHandler:
    HandleError "RemoveAllWatermarks"
    RemoveAllWatermarks = edits
End Function

'================================================================================
' HEADER IMAGE FUNCTION
'================================================================================

Private Function InsertStandardHeaderImage(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    Dim edits As Long: edits = 0
    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    
    ' Check if image file exists
    imgFile = HEADER_IMAGE_PATH
    If Dir(imgFile) = "" Then
        MsgBox "Header image not found at: " & vbCrLf & imgFile, vbExclamation, "Image Missing"
        Exit Function
    End If
    
    ' Process each section
    For Each sec In doc.Sections
        ' Modify primary headers
        Set header = sec.Headers(wdHeaderFooterPrimary)
        
        ' Clear existing header content
        header.LinkToPrevious = False
        header.Range.Delete
        
        ' Insert and format the image
        With header.Shapes.AddPicture( _
            FileName:=imgFile, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0, _
            Width:=HEADER_IMAGE_WIDTH, _
            Height:=HEADER_IMAGE_HEIGHT)
            
            .WrapFormat.Type = wdWrapTopBottom
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .Left = wdShapeCenter
            .Top = 0
        End With
        
        edits = edits + 1
    Next sec
    
    InsertStandardHeaderImage = edits
    Exit Function
    
ErrorHandler:
    HandleError "InsertStandardHeaderImage"
    InsertStandardHeaderImage = edits
End Function

'================================================================================
' HELPER FUNCTIONS
'================================================================================

Private Function CentimetersToPoints(cm As Double) As Single
    CentimetersToPoints = Application.CentimetersToPoints(cm)
End Function

Private Sub HandleError(procedureName As String)
    Dim errMsg As String
    errMsg = "Error in " & procedureName & ":" & vbCrLf & _
              "Error #" & Err.Number & vbCrLf & _
              Err.Description
    MsgBox errMsg, vbCritical, "Formatting Error"
    Debug.Print errMsg
End Sub