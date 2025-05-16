'--------------------------------------------------------------------------------
' Botão: Salvar, fechar e minimizar o Word
'--------------------------------------------------------------------------------
' SUBROTINA: SaveAndExit
' Finalidade: Salva o documento ativo, fecha-o e minimiza a janela do Microsoft Word.
'--------------------------------------------------------------------------------
Sub SaveAndExit()
    On Error GoTo ErrorHandler

    ' Verifica se há um documento ativo
    If Documents.Count = 0 Then
        MsgBox "Nenhum documento está aberto para salvar e fechar.", vbExclamation, "Documento não encontrado"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Salva o documento ativo
    If Not doc.Saved Then
        doc.Save
        MsgBox "Documento salvo com sucesso.", vbInformation, "Salvamento Concluído"
    Else
        MsgBox "Nenhuma alteração foi detectada. O documento já está salvo.", vbInformation, "Sem Alterações"
    End If

    ' Fecha o documento ativo
    doc.Close SaveChanges:=wdDoNotSaveChanges

    ' Minimiza a janela do Microsoft Word
    Application.WindowState = wdWindowStateMinimize

    Exit Sub

ErrorHandler:
    ' Tratamento de erros
    MsgBox "Erro " & Err.Number & ": " & Err.description, vbCritical, "Erro ao Salvar, Fechar e Minimizar"
End Sub
