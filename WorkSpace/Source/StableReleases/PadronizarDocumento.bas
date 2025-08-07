Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Word built-in constants (define if not referenced from Word object library)
Private Const wdNoProtection As Long = 0
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdLineSpace1pt5 As Long = 4
Private Const wdLineSpaceMultiple As Long = 5
Private Const wdWrapTight As Long = 1
Private Const wdRelativeHorizontalPositionPage As Long = 1
Private Const wdRelativeVerticalPositionPage As Long = 1
Private Const wdShapeCenter As Long = -999999 ' Center constant for shapes
Private Const msoTrue As Long = -1
Private Const msoPicture As Long = 13
Private Const msoTextEffect As Long = 15
Private Const wdCollapseEnd As Long = 0
Private Const wdFieldPage As Long = 33

' Document formatting constants
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const LINE_SPACING As Long = 10

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 5
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\Documents\HeaderStamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

'================================================================================
' Entry point for the "Standardize Document" button
'================================================================================
Public Sub BtnMAIN()
    On Error GoTo ErrHandler

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatting document..."
    End With

    Call GlobalChecking
    Call GlobalFormatting

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With

    Exit Sub

ErrHandler:
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    MsgBox "An error occurred while standardizing the document:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

'================================================================================
' Checks initial conditions before running other routines.
'================================================================================
Sub GlobalChecking()
    On Error GoTo ErrorHandler

    If ActiveDocument Is Nothing Then
        MsgBox "No active document found. Please open a document to format.", _
               vbExclamation, "Inactive Document"
        Exit Sub
    End If

    If ActiveDocument.Type <> wdTypeDocument Then
        MsgBox "The active document is not a Word document. Please open a Word document to format.", _
               vbExclamation, "Invalid Document Type"
        Exit Sub
    End If

    Exit Sub

ErrorHandler:
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    HandleError "GlobalChecking"
End Sub

'================================================================================
' Main formatting routine
'================================================================================
Public Sub GlobalFormatting()
    On Error GoTo ErrorHandler

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatting document..."
    End With

    BasicFormatting ActiveDocument
    RemoveWatermark ActiveDocument
    InsertHeaderStamp ActiveDocument
    InsertFooterStamp ActiveDocument

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With

    Exit Sub

ErrorHandler:
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    HandleError "Main"
End Sub

'================================================================================
' Removes watermark shapes from all sections if present.
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
' Handles errors by displaying an error message and logging it to the debug console.
'================================================================================
Public Sub HandleError(procedureName As String)
    Dim errMsg As String
    errMsg = "Error in subroutine: " & procedureName & vbCrLf & _
             "Error #" & Err.Number & ": " & Err.Description
    MsgBox errMsg, vbCritical, "Formatting Error"
    Debug.Print errMsg
    Err.Clear
End Sub

'================================================================================
' Converts a value in centimeters to points.
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    CentimetersToPoints = Application.CentimetersToPoints(cm)
End Function

'================================================================================
' Applies standard formatting to the document, including font, margins, and paragraph formatting.
'================================================================================
Private Sub BasicFormatting(doc As Document)
    On Error GoTo ErrorHandler

    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
    End With

    With doc.Content.Font
        .Name = STANDARD_FONT
        .Size = STANDARD_FONT_SIZE
    End With

    ' Enable automatic hyphenation for the document
    doc.AutoHyphenation = True

    With doc.Content.ParagraphFormat
        .LineSpacingRule = wdLineSpace1pt5
        .LineSpacing = 18
    End With

    Exit Sub

ErrorHandler:
    HandleError "BasicFormatting"
End Sub

'================================================================================
' Inserts a standard header image into the document's headers.
'================================================================================
Private Sub InsertHeaderStamp(doc As Document)
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

        With header.Range
            .Font.Reset
            .Font.Name = STANDARD_FONT
            .Font.Size = STANDARD_FONT_SIZE
            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
            .ParagraphFormat.LineSpacing = 18
        End With

        With header.Shapes.AddPicture( _
            fileName:=imgFile, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=wdShapeCenter, _
            Top:=CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM), _
            Width:=imgWidth, _
            Height:=imgHeight)
            .WrapFormat.Type = wdWrapTight
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .LockAspectRatio = msoTrue
        End With
    Next sec

    Exit Sub

ErrorHandler:
    HandleError "InsertHeaderStamp"
End Sub

'================================================================================
' Inserts centered automatic page numbers in the footer in the format "1-1" (both bold).
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
        Set rng = footer.Range

        rng.Text = ""
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter

        ' Insert first page number (bold)
        rng.Font.Name = STANDARD_FONT
        rng.Font.Size = STANDARD_FONT_SIZE - 3
        rng.Font.Bold = True
        Set fld1 = rng.Fields.Add(Range:=rng, Type:=wdFieldPage)
        rng.Collapse Direction:=wdCollapseEnd

        ' Insert hyphen (not bold)
        rng.Font.Bold = False
        rng.InsertAfter "-"
        rng.Collapse Direction:=wdCollapseEnd

        ' Insert second page number (bold)
        rng.Font.Bold = True
        Set fld2 = rng.Fields.Add(Range:=rng, Type:=wdFieldPage)
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next sec

    Exit Sub

ErrorHandler:
    HandleError "InsertFooterStamp"
End Sub

