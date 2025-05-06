Option Explicit

'================================================================================
' MAIN FORMATTING MODULE
'================================================================================
' Description: Corrects the document formatting to a formal standard.
' Compatible with Microsoft Office Word 2010 and later versions.
'================================================================================

Const wdFindContinue As Integer = 1
Const wdReplaceAll As Integer = 2
Const wdLineSpaceSingle As Integer = 0

'--------------------------------------------------------------------------------
' MAIN SUBROUTINE: FixDocumentFormatting
' Purpose: Manages the processing flow and corrects the formatting of the active document.
'--------------------------------------------------------------------------------
Sub FixDocumentFormatting()
    On Error GoTo ErrorHandler

    ' Validate if there is an active document
    If Documents.Count = 0 Then
        MsgBox "No document is open. Please open a document before running the formatting correction.", _
               vbExclamation, "Document Not Found"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    ' Validate if the active document is a Word document
    If Not TypeOf doc Is Document Then
        MsgBox "The active document is not a Word document.", vbExclamation, "Invalid Document"
        Exit Sub
    End If

    ' Validate if the document is empty
    If Trim(doc.Content.Text) = "" Then
        MsgBox "The document is empty. Please add content before running the formatting correction.", _
               vbExclamation, "Empty Document"
        Exit Sub
    End If

    Dim editCount As Integer: editCount = 0

    ' Disable screen updates and events to improve performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Processing flow
    editCount = editCount + RemoveBlankLinesBeforeTitle(doc)
    editCount = editCount + RemoveExtraSpacesAndPageBreaks(doc)
    editCount = editCount + FormatDocument(doc)
    editCount = editCount + ClearDocumentFormatting(doc)
    editCount = editCount + RemoveDocumentWatermarks(doc)

    ' Re-enable screen updates and events
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Completion message
    MsgBox "Formatting correction completed successfully!" & vbCrLf & _
           "Number of edits made: " & editCount, vbInformation, "Formatting Corrected"
    Exit Sub

ErrorHandler:
    ' Error handling
    MsgBox "Error in FixDocumentFormatting: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Set doc = Nothing
End Sub

'--------------------------------------------------------------------------------
' FUNCTION: RemoveBlankLinesBeforeTitle
' Purpose: Removes all blank lines before the document title.
'--------------------------------------------------------------------------------
Private Function RemoveBlankLinesBeforeTitle(doc As Document) As Integer
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        MsgBox "Invalid document parameter.", vbCritical, "Error"
        Exit Function
    End If

    Dim editCount As Integer: editCount = 0
    Dim firstPara As Range

    ' Validate if the document has at least one paragraph
    If doc.Paragraphs.Count = 0 Then Exit Function

    ' Loop to remove blank lines at the beginning of the document
    Set firstPara = doc.Paragraphs(1).Range
    Do While Trim(firstPara.Text) = ""
        firstPara.Delete
        editCount = editCount + 1
        Set firstPara = doc.Paragraphs(1).Range
    Loop

    RemoveBlankLinesBeforeTitle = editCount
    Exit Function

ErrorHandler:
    MsgBox "Error in RemoveBlankLinesBeforeTitle: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
    RemoveBlankLinesBeforeTitle = editCount
    Set firstPara = Nothing
End Function

'--------------------------------------------------------------------------------
' FUNCTION: RemoveExtraSpacesAndPageBreaks
' Purpose: Removes duplicate spaces and extra page breaks in a single execution.
'--------------------------------------------------------------------------------
Private Function RemoveExtraSpacesAndPageBreaks(doc As Document) As Integer
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        MsgBox "Invalid document parameter.", vbCritical, "Error"
        Exit Function
    End If

    Dim editCount As Integer: editCount = 0
    Dim result As Boolean

    ' Remove duplicate spaces and extra page breaks
    With doc.Content.Find
        .Text = "[  ]{2,}|^m^m" ' Duplicate spaces or consecutive page breaks
        .Replacement.Text = " " ' Replace with a single space or a single page break
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        Do
            result = .Execute(Replace:=wdReplaceOne)
            If result Then editCount = editCount + 1
        Loop While result
    End With

    RemoveExtraSpacesAndPageBreaks = editCount
    Exit Function

ErrorHandler:
    MsgBox "Error in RemoveExtraSpacesAndPageBreaks: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
    RemoveExtraSpacesAndPageBreaks = editCount
End Function

'--------------------------------------------------------------------------------
' FUNCTION: FormatDocument
' Purpose: Adjusts the formatting of the active document to a predefined standard.
'--------------------------------------------------------------------------------
Private Function FormatDocument(doc As Document) As Integer
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        MsgBox "Invalid document parameter.", vbCritical, "Error"
        Exit Function
    End If

    Dim editCount As Integer: editCount = 0

    ' Set page layout margins and spacing
    With doc.PageSetup
        .TopMargin = Application.CentimetersToPoints(4.5)
        .BottomMargin = Application.CentimetersToPoints(3)
        .LeftMargin = Application.CentimetersToPoints(3)
        .RightMargin = Application.CentimetersToPoints(3)
        .HeaderDistance = Application.CentimetersToPoints(0.7)
        .FooterDistance = Application.CentimetersToPoints(0.7)
    End With

    ' Format each paragraph
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        With para.Range.Font
            .Name = "Arial"
            .Size = 12
        End With
        With para.Format
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
        End With
        editCount = editCount + 1
    Next para

    FormatDocument = editCount
    Exit Function

ErrorHandler:
    MsgBox "Error in FormatDocument: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
    FormatDocument = editCount
End Function

'--------------------------------------------------------------------------------
' FUNCTION: ClearDocumentFormatting
' Purpose: Removes all formatting from the active document.
'--------------------------------------------------------------------------------
Private Function ClearDocumentFormatting(doc As Document) As Integer
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        MsgBox "Invalid document parameter.", vbCritical, "Error"
        Exit Function
    End If

    doc.Content.Font.Reset
    doc.Content.ParagraphFormat.Reset

    ClearDocumentFormatting = 1 ' Considered as one edit made
    Exit Function

ErrorHandler:
    MsgBox "Error in ClearDocumentFormatting: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
    ClearDocumentFormatting = 0
End Function

'--------------------------------------------------------------------------------
' FUNCTION: RemoveDocumentWatermarks
' Purpose: Removes all watermarks from the active document.
'--------------------------------------------------------------------------------
Private Function RemoveDocumentWatermarks(doc As Document) As Integer
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        MsgBox "Invalid document parameter.", vbCritical, "Error"
        Exit Function
    End If

    Dim editCount As Integer: editCount = 0
    Dim section As Section
    Dim header As HeaderFooter
    Dim shape As Shape

    ' Loop through all sections to remove watermarks
    For Each section In doc.Sections
        ' Process headers
        For Each header In section.Headers
            ' Remove text effect shapes (watermarks)
            For Each shape In header.Shapes
                If shape.Type = msoTextEffect Then
                    shape.Delete
                    editCount = editCount + 1
                End If
            Next shape
        Next header
    Next section

    RemoveDocumentWatermarks = editCount
    Exit Function

ErrorHandler:
    MsgBox "Error in RemoveDocumentWatermarks: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
    RemoveDocumentWatermarks = editCount
End Function