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
Private Const BOTTOM_MARGIN_CM As Double = 2# ' Bottom margin in cm
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
' MAIN PROCEDURE: Main_PR
'================================================================================
' Purpose: Orchestrates the document formatting process by calling various helper
' functions to apply standard formatting, clean up spacing, and insert headers.
'================================================================================
Public Sub Main_PR()
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

    ' === AJUSTA O ZOOM PARA 110% ===
    On Error Resume Next
    ActiveWindow.View.Zoom.Percentage = 110
    On Error GoTo ErrorHandler

    ' === ATIVAR CONTROLAR ALTERAÇÕES ===
    doc.TrackRevisions = True

    ' === BACKUP ANTES DE QUALQUER ALTERAÇÃO ===
    Dim backupPath As String
    backupPath = Main_OPB(doc) ' Md_OrigPropsBackup.CreateBackup deve estar disponível no projeto
    
    ' Verifica se o documento está protegido
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "O documento está protegido. Por favor, desproteja-o antes de continuar.", _
               vbExclamation, "Documento Protegido"
        Exit Sub
    End If
    
    ' Optimize performance by disabling screen updates and alerts
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatting document..."
    End With
    
    ' Limpa os metadados do documento
    ClearDocumentMetadata doc ' Clear document metadata

    ' Formatting the document
    Main_COF doc ' Call the format cleaner module
    Main_SDF doc ' Call the set default format module
    
    ' Calling the text replacement subroutine
    Main_BNATF doc ' Call the text replacement module

    ' Restore application state
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    
    ' Mensagem de conclusão
    Dim docPath As String: docPath = doc.FullName
    ShowCompletionMessage backupPath, docPath
    
    ' Limpeza de variáveis
    Set doc = Nothing
    Exit Sub ' Exit the procedure
    
ErrorHandler:
    ' Handle errors and restore application state
    HandleError "Main_PR"
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    Set doc = Nothing
End Sub

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

'--------------------------------------------------------------------------------
' SUBROTINA: ClearDocumentMetadata
' Remove os metadados do documento ativo.
'--------------------------------------------------------------------------------
Private Sub ClearDocumentMetadata(doc As Document)
    On Error Resume Next

    Dim prop As DocumentProperty
    doc.BuiltInDocumentProperties("Title") = ""
    doc.BuiltInDocumentProperties("Subject") = ""
    doc.BuiltInDocumentProperties("Keywords") = ""
    doc.BuiltInDocumentProperties("Comments") = ""
    doc.BuiltInDocumentProperties("Author") = "Anônimo"
    doc.BuiltInDocumentProperties("Last Author") = "Anônimo"
    doc.BuiltInDocumentProperties("Manager") = ""
    doc.BuiltInDocumentProperties("Company") = ""

    For Each prop In doc.CustomDocumentProperties
        prop.Delete
    Next prop
End Sub

'--------------------------------------------------------------------------------
' SUBROTINA: ShowCompletionMessage
' Exibe uma mensagem de conclusão com informações sobre o backup e o processamento.
'--------------------------------------------------------------------------------
Private Sub ShowCompletionMessage(backupPath As String, docPath As String, editCount As Integer, executionTime As Double)
    MsgBox "Retificação concluída com sucesso!" & vbCrLf & vbCrLf & _
           "Backup criado em: " & backupPath & vbCrLf & _
           vbInformation, "Retificação Completa"
End Sub


