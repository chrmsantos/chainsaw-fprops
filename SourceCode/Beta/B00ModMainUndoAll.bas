'--------------------------------------------------------------------------------
' PROCEDURE: UndoAllChanges
' Purpose: Reverts all changes made to the active document with improved error handling,
'          performance, and user feedback.
' Parameters: None
' Returns: Nothing
'--------------------------------------------------------------------------------
Sub UndoAllChanges()
    On Error GoTo ErrorHandler
    
    ' Validate active document exists
    If Not IsDocumentAvailable() Then Exit Sub
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Check if there are changes to undo
    If Not HasChangesToUndo(doc) Then Exit Sub
    
    ' Perform undo operations with safety limits
    UndoChanges doc
    
    ' Notify user of completion
    ShowCompletionMessage
    
ExitProcedure:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    HandleError "UndoAllChanges"
    Resume ExitProcedure
End Sub

Private Function IsDocumentAvailable() As Boolean
    If Documents.Count = 0 Then
        MsgBox "No document is open to undo changes.", vbExclamation, "Document not found"
        IsDocumentAvailable = False
        Exit Function
    End If
    
    IsDocumentAvailable = True
End Function

Private Function HasChangesToUndo(doc As Document) As Boolean
    If doc.Saved Then
        MsgBox "No changes made since last save.", vbInformation, "Nothing to undo"
        HasChangesToUndo = False
        
        ' Small optimization - close document reference if we're exiting here anyway 
        Set doc = Nothing
        
        Exit Function
    End If
    
    HasChangesToUndo = True
End Function

Private Sub UndoChanges(doc As Document)
    Const MAX_UNDO_OPERATIONS As Long = 1000 ' Safety limit for undo operations
    
    Application.ScreenUpdating = False
    
    Dim undoCount As Long: undoCount = 0
    
     ' More efficient loop structure with explicit condition check first 
     While doc.Undo And undoCount < MAX_UNDO_OPERATIONS 
         undoCount = undoCount + 1 
     Wend 
    
End Sub

Private Sub ShowCompletionMessage()
     MsgBox "All changes since file opening or last save were successfully reverted.", _
            vbInformation, "Changes Reverted" 
End Sub 

Private Sub HandleError(procedureName As String)
     Application.ScreenUpdating = True 
     
     Dim errorMsg As String 
     errorMsg = "Error in procedure '" & procedureName & "'" & vbCrLf & _ 
                "Error #" & Err.Number & ": " & Err.Description 

     MsgBox errorMsg, vbCritical, "Operation Failed" 
End Sub 
