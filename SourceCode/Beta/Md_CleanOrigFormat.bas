Public Sub Main_COF(doc As Document)
    On Error GoTo ErrorHandler

    ' Otimização de desempenho: desabilita atualizações de tela e alertas
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Limpando formatação do documento..."
    End With

    ' Cleaning format steps
    ResetBasicFormatting doc ' Reset basic formatting
    RemoveAllWatermarks doc ' Remove watermarks
    RemoveLeadingBlankLines doc ' Remove leading blank lines
    CleanDocumentSpacing doc ' Clean up document spacing
    RemoveExtraPageBreaks doc ' Remove extra page breaks

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
    HandleError "Main_COF"
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

