' Entry point and orchestration module for the "Standardize Document" button

Public Sub BtnMAIN()
    On Error GoTo ErrHandler

    ' Performance optimization: disable screen updating and alerts during processing
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatting document..."
    End With

    ' Check prerequisites before formatting (active document, protection, etc.)
    Call GlobalChecking

    ' Run the global formatting routine (margins, font, header, watermark, etc.)
    Call GlobalFormatting

    ' Restore application state after formatting
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With

    Exit Sub

ErrHandler:
    ' Restore application state in case of error
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    ' Show detailed error message to the user
    MsgBox "An error occurred while standardizing the document:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

