'--------------------------------------------------------------------------------
' Botão: Salvar, fechar e minimizar o Word
'--------------------------------------------------------------------------------
' SUBROTINA: SaveMinAndExit
' Finalidade: Salva o documento ativo, fecha-o e minimiza a janela do Microsoft Word.
'--------------------------------------------------------------------------------
Public Sub Main_SCM()
    On Error GoTo ErrorHandler

    ' Verifica se há um documento ativo
    If Documents.Count = 0 Then
        MsgBox "Nenhum documento está aberto para salvar e fechar.", vbExclamation, "Documento não encontrado"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = Nothing

    ' Tenta obter o documento ativo
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Não foi possível acessar o documento ativo.", vbCritical, "Erro"
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' Salva o documento ativo
    If Not doc.Saved Then
        On Error Resume Next
        doc.Save
        If Err.Number <> 0 Then
            MsgBox "Erro ao salvar o documento: " & Err.Description, vbCritical, "Erro ao Salvar"
            Err.Clear
            GoTo Finalize
        Else
            MsgBox "Documento salvo com sucesso.", vbInformation, "Salvamento Concluído"
        End If
        On Error GoTo ErrorHandler
    Else
        MsgBox "Nenhuma alteração foi detectada. O documento já está salvo.", vbInformation, "Sem Alterações"
    End If

    ' Fecha o documento ativo
    On Error Resume Next
    doc.Close SaveChanges:=wdDoNotSaveChanges
    If Err.Number <> 0 Then
        MsgBox "Erro ao fechar o documento: " & Err.Description, vbCritical, "Erro ao Fechar"
        Err.Clear
        GoTo Finalize
    End If
    On Error GoTo ErrorHandler

    ' Minimiza a janela do Microsoft Word
    On Error Resume Next
    Application.WindowState = wdWindowStateMinimize
    If Err.Number <> 0 Then
        MsgBox "Erro ao minimizar o Word: " & Err.Description, vbExclamation, "Erro ao Minimizar"
        Err.Clear
    End If
    On Error GoTo ErrorHandler

Finalize:
    Set doc = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro ao Salvar, Fechar e Minimizar"
    Resume Finalize
End Sub

