'--------------------------------------------------------------------------------
' SUBROUTINE: SaveAndExit
' Purpose: Saves the active document and exits Microsoft Word.
'--------------------------------------------------------------------------------
Sub SaveAndExit()
    On Error GoTo ErrorHandler

    ' Check if there is an active document
    If Documents.Count = 0 Then
        MsgBox "No document is open to save and exit.", vbExclamation, "Document Not Found"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Save the active document
    If Not doc.Saved Then
        doc.Save
        MsgBox "Document saved successfully.", vbInformation, "Save Completed"
    Else
        MsgBox "No changes detected. The document is already saved.", vbInformation, "No Changes"
    End If

    ' Close the document and exit Word
    doc.Close SaveChanges:=wdDoNotSaveChanges
    Application.Quit

    Exit Sub

ErrorHandler:
    ' Error handling
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in Save and Exit"
End Sub

'--------------------------------------------------------------------------------
' SUBROUTINE: SaveAndCloseActive
' Purpose: Saves and closes the active document. If other documents are open,
'          Microsoft Word is minimized instead of exiting.
'--------------------------------------------------------------------------------
Sub SaveAndCloseActive()
    On Error GoTo ErrorHandler

    ' Check if there is an active document
    If Documents.Count = 0 Then
        MsgBox "No document is open to save and close.", vbExclamation, "Document Not Found"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Save the active document
    If Not doc.Saved Then
        doc.Save
        MsgBox "Document saved successfully.", vbInformation, "Save Completed"
    Else
        MsgBox "No changes detected. The document is already saved.", vbInformation, "No Changes"
    End If

    ' Close the active document
    doc.Close SaveChanges:=wdDoNotSaveChanges

    ' Check if there are other open documents
    If Documents.Count > 0 Then
        ' Minimize Microsoft Word
        Application.WindowState = wdWindowStateMinimize
    Else
        ' Exit Microsoft Word if no other documents are open
        Application.Quit
    End If

    Exit Sub

ErrorHandler:
    ' Error handling
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in Save and Close Active"
End Sub