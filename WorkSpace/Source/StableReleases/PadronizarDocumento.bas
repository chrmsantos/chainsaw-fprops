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
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Long = 13

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 5
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

'================================================================================
' MAIN ENTRY POINT
'================================================================================
Public Sub BtnMAIN()
    On Error GoTo ErrHandler

    SetAppState False, "Formatting document..."

    If Not GlobalChecking Then GoTo CleanUp

    With ActiveDocument
        GlobalFormatting .Application.ActiveDocument
    End With

    Application.StatusBar = "Document standardized successfully!"
    Exit Sub

ErrHandler:
    Application.StatusBar = "Error: " & Err.Description
    MsgBox "An error occurred while standardizing the document:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
CleanUp:
    SetAppState True, ""   ' Do NOT clear the status bar here
    Exit Sub
End Sub

'================================================================================
' APPLICATION STATE HANDLER
'================================================================================
Private Sub SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "")
    With Application
        .ScreenUpdating = enabled
        .DisplayAlerts = enabled
        If Not enabled And statusMsg <> "" Then
            .StatusBar = statusMsg ' Show the custom message
        End If
        ' Do NOT clear the status bar when enabled
    End With
End Sub

'================================================================================
' GLOBAL CHECKING
'================================================================================
Private Function GlobalChecking() As Boolean
    On Error GoTo ErrorHandler

    If ActiveDocument Is Nothing Then
        MsgBox "No active document found. Please open a document to format.", vbExclamation, "Inactive Document"
        Exit Function
    End If

    If ActiveDocument.Type <> wdTypeDocument Then
        MsgBox "The active document is not a Word document. Please open a Word document to format.", vbExclamation, "Invalid Document Type"
        Exit Function
    End If

    ' Security: Check if document is protected
    If ActiveDocument.ProtectionType <> 0 And ActiveDocument.ProtectionType <> -1 Then
        MsgBox "The document is protected. Please unprotect it before formatting.", vbExclamation, "Protected Document"
        Exit Function
    End If

    ' Security: Check for macros in the document (warn user)
    If HasMacros(ActiveDocument) Then
        If MsgBox("Warning: This document contains macros. Continue formatting?", vbExclamation + vbYesNo, "Macro Warning") = vbNo Then
            Exit Function
        End If
    End If

    GlobalChecking = True
    Exit Function

ErrorHandler:
    HandleError "GlobalChecking"
    GlobalChecking = False
End Function

'================================================================================
' MAIN FORMATTING ROUTINE
'================================================================================
Private Sub GlobalFormatting(doc As Document)
    On Error GoTo ErrorHandler

    ' Performance: Use With block for doc to minimize property lookups
    With doc
        ApplyPageSetup .Application.ActiveDocument
        ApplyFontAndParagraph .Application.ActiveDocument
        EnableHyphenation .Application.ActiveDocument
        RemoveWatermark .Application.ActiveDocument
        InsertHeaderStamp .Application.ActiveDocument
        InsertFooterStamp .Application.ActiveDocument
    End With

    Exit Sub

ErrorHandler:
    HandleError "GlobalFormatting"
End Sub

'================================================================================
' PAGE SETUP
'================================================================================
Private Sub ApplyPageSetup(doc As Document)
    ' Performance: Only set if different to avoid unnecessary redraws
    With doc.PageSetup
        If .TopMargin <> CentimetersToPoints(TOP_MARGIN_CM) Then .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        If .BottomMargin <> CentimetersToPoints(BOTTOM_MARGIN_CM) Then .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        If .LeftMargin <> CentimetersToPoints(LEFT_MARGIN_CM) Then .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        If .RightMargin <> CentimetersToPoints(RIGHT_MARGIN_CM) Then .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        If .HeaderDistance <> CentimetersToPoints(HEADER_DISTANCE_CM) Then .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        If .FooterDistance <> CentimetersToPoints(FOOTER_DISTANCE_CM) Then .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
    End With
End Sub

'================================================================================
' FONT AND PARAGRAPH FORMATTING
'================================================================================
Private Sub ApplyFontAndParagraph(doc As Document)
    Dim para As Paragraph
    Dim hasInlineImage As Boolean

    For Each para In doc.Paragraphs
        hasInlineImage = False

        ' Check if paragraph contains any inline image
        If para.Range.InlineShapes.Count > 0 Then
            hasInlineImage = True
        End If

        ' Skip formatting if inline image is present
        If Not hasInlineImage Then
            ' Apply font formatting
            With para.Range.Font
                .Name = STANDARD_FONT
                .Size = STANDARD_FONT_SIZE
            End With

            ' Apply paragraph formatting
            With para.Format
                .LineSpacingRule = wdLineSpace1pt5
                .LineSpacing = LINE_SPACING
            End With

            ' Justify left-aligned paragraphs
            If para.Alignment = wdAlignParagraphLeft Then
                para.Alignment = wdAlignParagraphJustify
            End If
        End If
    Next para
End Sub


'================================================================================
' ENABLE HYPHENATION
'================================================================================
Private Sub EnableHyphenation(doc As Document)
    On Error Resume Next
    If Not doc.AutoHyphenation Then doc.AutoHyphenation = True
    On Error GoTo 0
End Sub

'================================================================================
' REMOVE WATERMARK
'================================================================================
Private Sub RemoveWatermark(doc As Document)
    On Error Resume Next

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
    Dim i As Long

    For Each sec In doc.Sections
        For Each header In sec.Headers
            If header.Shapes.Count > 0 Then
                For i = header.Shapes.Count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If (shp.Type = msoPicture Or shp.Type = msoTextEffect) And InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Then
                        shp.Delete
                    End If
                Next i
            End If
        Next header
    Next sec

    On Error GoTo 0
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

    username = GetSafeUserName()
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Security: Validate image path (prevent directory traversal)
    If InStr(imgFile, "..") > 0 Or InStr(imgFile, ":") > 2 Then
        MsgBox "Invalid image path detected. Aborting header image insertion.", vbCritical, "Security Alert"
        Exit Sub
    End If

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
            .ParagraphFormat.LineSpacing = LINE_SPACING
        End With

        ' Performance: Only add image if not already present
        If header.Shapes.Count = 0 Then
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
        End If
    Next sec

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
        rng.Font.Size = FOOTER_FONT_SIZE
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

'================================================================================
' ERROR HANDLER
'================================================================================
Public Sub HandleError(procedureName As String)
    Dim errMsg As String
    errMsg = "Error in subroutine: " & procedureName & vbCrLf & _
             "Error #" & Err.Number & ": " & Err.Description
    Application.StatusBar = "Error: " & Err.Description
    MsgBox errMsg, vbCritical, "Formatting Error"
    Debug.Print errMsg
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
    ' Only allow alphanumeric and underscore in username for path safety
    Dim rawName As String, c As String, i As Integer
    rawName = Environ("USERNAME")
    For i = 1 To Len(rawName)
        c = Mid(rawName, i, 1)
        If c Like "[A-Za-z0-9_]" Then
            GetSafeUserName = GetSafeUserName & c
        End If
    Next i
End Function

'================================================================================
' UTILITY: CHECK FOR MACROS
'================================================================================
Private Function HasMacros(doc As Document) As Boolean
    ' Checks for VBA project in the document (Word 2010+)
    On Error Resume Next
    HasMacros = (doc.HasVBProject)
    On Error GoTo 0
End Function

