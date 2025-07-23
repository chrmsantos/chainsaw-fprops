Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Constants for Word operations
Private Const wdFindContinue As Long = 1 ' Continue search after the first match
Private Const wdReplaceOne As Long = 1 ' Replace only one occurrence
Private Const wdLineSpaceSingle As Double = 1.5 ' Standard line spacing (should be Double)
Private Const STANDARD_FONT As String = "Arial" ' Standard font for the document
Private Const STANDARD_FONT_SIZE As Long = 12 ' Standard font size
Private Const LINE_SPACING As Long = 12 ' Line spacing in points

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 5 ' Top margin in cm
Private Const BOTTOM_MARGIN_CM As Double = 3 ' Bottom margin in cm
Private Const LEFT_MARGIN_CM As Double = 3 ' Left margin in cm
Private Const RIGHT_MARGIN_CM As Double = 3 ' Right margin in cm
Private Const HEADER_DISTANCE_CM As Double = 0.5 ' Distance from header to content in cm
Private Const FOOTER_DISTANCE_CM As Double = 1.5 ' Distance from footer to content in cm

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\Documents\HeaderStamp.png" ' Relative path to the header image
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 22 ' Maximum width of the header image in cm
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7 ' Top margin for the header image in cm
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19 ' Height-to-width ratio for the header image

'================================================================================
' Main module for formatting
'================================================================================

' Entry point for macro button: applies formatting to the active document
Public Sub GlobalFormatting()
    ' Set up error handling for the procedure
    On Error GoTo ErrorHandler

    ' Save the document if there are unsaved changes
    If ActiveDocument.Saved = False Then
        ActiveDocument.Save
    End If

    ' Temporarily disable screen updating and alerts for performance and user experience
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatting document..."
    End With

    ' Apply basic formatting (margins, font, etc.)
    BasicFormatting ActiveDocument

    ' Remove any watermark shapes from all headers in all sections
    RemoveWatermark ActiveDocument

    ' Insert the standard header image in all sections
    InsertHeaderStamp ActiveDocument

    ' Insert the footer stamp (page numbers) in all sections
    InsertFooterStamp ActiveDocument

    ' Restore application state (reenable screen updating and alerts)
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With

    Exit Sub

ErrorHandler:
    ' Always restore application state, even if an error occurs
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    ' Call the error handler subroutine with a label for this procedure
    HandleError "Main"
End Sub

'================================================================================
' RemoveWatermark
' Purpose: Removes watermark shapes from all sections if present.
'================================================================================
Private Sub RemoveWatermark(doc As Document)
    ' Ignore errors in this routine
    On Error Resume Next

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
    Dim i As Long

    ' Loop through all sections in the document
    For Each sec In doc.Sections
        ' Loop through all headers in the section
        For Each header In sec.Headers
            ' Loop backwards through all shapes in the header (safe for deletion)
            For i = header.Shapes.Count To 1 Step -1
                Set shp = header.Shapes(i)
                ' Check if the shape is a picture or text effect
                If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                    ' Check if the shape name contains "Watermark"
                    If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Then
                        shp.Delete ' Delete the watermark shape
                    End If
                End If
            Next i
        Next header
    Next sec

    ' Restore normal error handling
    On Error GoTo 0
End Sub

'================================================================================
' HandleError
' Purpose: Handles errors by displaying an error message and logging it to the debug console.
'================================================================================
Public Sub HandleError(procedureName As String)
    Dim errMsg As String ' Variable to hold the error message

    ' Build a detailed error message with procedure name, error number, and description
    errMsg = "Error in subroutine: " & procedureName & vbCrLf & _
             "Error #" & Err.Number & ": " & Err.Description

    ' Show the error message to the user
    MsgBox errMsg, vbCritical, "Formatting Error"

    ' Output the error message to the Immediate window for debugging
    Debug.Print errMsg

    ' Clear the error object
    Err.Clear
End Sub

'================================================================================
' CentimetersToPoints
' Purpose: Converts a value in centimeters to points.
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    ' Use Word's built-in conversion function
    CentimetersToPoints = Application.CentimetersToPoints(cm)
End Function

'================================================================================
' BasicFormatting
' Purpose: Applies standard formatting to the document, including font, margins, and paragraph formatting.
'================================================================================
Private Sub BasicFormatting(doc As Document)
    On Error GoTo ErrorHandler ' Enable error handling

    ' Check if the document is protected and cannot be modified
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "The document is protected. Please unprotect it before continuing.", _
               vbExclamation, "Protected Document"
        Exit Sub
    End If

    ' Set page layout and margins using constants and conversion function
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
        '.Reset ' Uncomment to reset font formatting to default
        .Name = STANDARD_FONT
        .Size = STANDARD_FONT_SIZE
    End With

    ' Optionally, set paragraph formatting (uncomment if needed)
    'With doc.Content.ParagraphFormat
    '    .LineSpacingRule = wdLineSpaceMultiple
    '    .LineSpacing = wdLineSpaceSingle * LINE_SPACING
    'End With

    Exit Sub

ErrorHandler:
    HandleError "BasicFormatting"
End Sub

'================================================================================
' InsertHeaderStamp
' Purpose: Inserts a standard header image into the document's headers.
'================================================================================
Private Sub InsertHeaderStamp(doc As Document)
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single

    ' Build the full path to the header image using the current user's folder
    username = Environ("USERNAME")
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Check if the image file exists
    If Dir(imgFile) = "" Then
        MsgBox "Header image not found at: " & vbCrLf & imgFile, vbExclamation, "Image Missing"
        Exit Sub
    End If

    ' Calculate the image dimensions in points
    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    ' Loop through all sections and insert the image in the primary header
    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        header.LinkToPrevious = False ' Unlink header from previous section

        ' Remove all existing content in the header
        header.Range.Delete

        ' Set font and paragraph formatting for the header range
        With header.Range
            .Font.Reset
            .Font.Name = STANDARD_FONT
            .Font.Size = STANDARD_FONT_SIZE
            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
            .ParagraphFormat.LineSpacing = wdLineSpaceSingle * LINE_SPACING ' 1.5 lines (18 points)
        End With

        ' Add the image and adjust its properties
        With header.Shapes.AddPicture( _
            fileName:=imgFile, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0, _
            Width:=imgWidth, _
            Height:=imgHeight)

            .WrapFormat.Type = wdWrapTight ' Set text wrapping style to "tight"
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .Left = wdShapeCenter ' Center the image horizontally
            .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM) ' Set top margin
            .LockAspectRatio = msoTrue ' Maintain aspect ratio
        End With
    Next sec

    Exit Sub

ErrorHandler:
    HandleError "InsertHeaderStamp"
End Sub

'================================================================================
' InsertFooterStamp
' Purpose: Inserts centered automatic page numbers in the footer in the format "1-1" (where both are the page number), with numbers in bold.
'================================================================================
Private Sub InsertFooterStamp(doc As Document)
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rng As Range
    Dim fld1 As Field
    Dim fld2 As Field

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        footer.LinkToPrevious = False

        ' Clear existing footer content
        footer.Range.Text = ""

        ' Set range to the footer and center it
        Set rng = footer.Range
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter

        ' Insert first page number field and make it bold
        Set fld1 = rng.Fields.Add(Range:=rng, Type:=wdFieldPage)
        fld1.Result.Font.Bold = True

        ' Move to end and insert hyphen (not bold)
        rng.Collapse Direction:=wdCollapseEnd
        rng.InsertAfter "-"
        rng.Collapse Direction:=wdCollapseEnd

        ' Insert second page number field and make it bold
        Set fld2 = rng.Fields.Add(Range:=rng, Type:=wdFieldPage)
        fld2.Result.Font.Bold = True
    Next sec

    Exit Sub

ErrorHandler:
    HandleError "InsertFooterStamp"
End Sub