'--------------------------------------------------------------------------------
' SUBROUTINE: SaveAndExit
' Purpose: Saves all open documents and exits Microsoft Word after confirmation.
' Features:
'   - Saves all modified documents
'   - Provides option to cancel operation
'   - Handles multiple documents
'--------------------------------------------------------------------------------
Sub SaveAndExit()
    On Error GoTo ErrorHandler
    
    ' Check if there are any open documents
    If Documents.Count = 0 Then
        MsgBox "No documents are open to save.", vbInformation, "No Open Documents"
        Application.Quit
        Exit Sub
    End If
    
    ' Ask for user confirmation
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you want to save all documents and exit Word?", _
                     vbQuestion + vbYesNoCancel, "Exit Word")
    
    If response = vbCancel Then Exit Sub
    
    Dim doc As Document
    Dim unsavedDocs As Integer: unsavedDocs = 0
    
    ' Save all modified documents if user confirmed
    If response = vbYes Then
        For Each doc In Documents
            If Not doc.Saved Then
                doc.Save
                unsavedDocs = unsavedDocs + 1
            End If
        Next doc
    End If
    
    ' Provide feedback and exit
    If response = vbYes Then
        If unsavedDocs > 0 Then
            MsgBox unsavedDocs & " document(s) saved successfully.", _
                  vbInformation, "Documents Saved"
        Else
            MsgBox "No unsaved documents found.", vbInformation, "No Changes"
        End If
    End If
    
    ' Close all documents and exit Word
    Application.Quit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in SaveAndExit:" & vbCrLf & _
           "Error #" & Err.Number & vbCrLf & _
           Err.Description, vbCritical, "Save and Exit Error"
End Sub

'--------------------------------------------------------------------------------
' SUBROUTINE: SaveAndCloseActive
' Purpose: Saves and closes the active document with advanced options.
' Features:
'   - Handles save confirmation
'   - Preserves other open documents
'   - Provides detailed feedback
'--------------------------------------------------------------------------------
Sub SaveAndCloseActive()
    On Error GoTo ErrorHandler
    
    ' Check if there is an active document
    If Documents.Count = 0 Then
        MsgBox "No document is currently open.", vbExclamation, "No Document"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = ActiveDocument
    Dim docName As String
    docName = doc.Name
    
    ' Check if document needs saving
    If Not doc.Saved Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Do you want to save changes to '" & docName & "'?", _
                         vbQuestion + vbYesNoCancel, "Save Changes")
        
        Select Case response
            Case vbYes
                doc.Save
                MsgBox "Document '" & docName & "' saved successfully.", _
                      vbInformation, "Document Saved"
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    ' Close the document
    doc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' Handle remaining documents
    If Documents.Count > 0 Then
        ' Optional: Activate next available document
        Documents(1).Activate
    Else
        ' Option 1: Leave Word open with new blank document
        Documents.Add
        
        ' Option 2: Quit Word (uncomment if preferred)
        ' Application.Quit
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in SaveAndCloseActive:" & vbCrLf & _
           "Error #" & Err.Number & vbCrLf & _
           Err.Description, vbCritical, "Save and Close Error"
End Sub