'================================================================================
' GlobalChecking
' Purpose: Checks initial conditions before running other routines.
'================================================================================
Sub GlobalChecking()
    ' Check if there is an active document
    ' This prevents errors if the user tries to run formatting without any document open
    If ActiveDocument Is Nothing Then
        MsgBox "No active document found. Please open a document to format.", _
               vbExclamation, "Inactive Document"
        Exit Sub
    End If

    ' Check if the document is protected
    ' Formatting routines require the document to be unprotected for changes
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "The document is protected. Please unprotect it before continuing.", _
               vbExclamation, "Protected Document"
        Exit Sub
    End If

    ' Check if the document contains any content
    ' Prevents running formatting on an empty document
    If Trim(ActiveDocument.Content.Text) = "" Then
        MsgBox "The document is empty. Please add content before formatting.", _
               vbExclamation, "Empty Document"
        Exit Sub
    End If

    ' Check if the document is a Word document
    ' Ensures the macro is only run on valid Word documents
    If ActiveDocument.Type <> wdTypeDocument Then
        MsgBox "The active document is not a Word document. Please open a Word document to format.", _
               vbExclamation, "Invalid Document Type"
        Exit Sub
    End If

    ' If all checks pass, exit the subroutine normally
    Exit Sub

ErrorHandler:
    ' Always restore application state, even if an error occurs
    ' This prevents the application from remaining in a disabled state after an error
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    ' Call the error handler routine to display/log the error
    HandleError "GlobalChecking"
End Sub
