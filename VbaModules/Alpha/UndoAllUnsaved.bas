' SUBROTINA: UndoAllChanges
' Finalidade: Desfaz todas as alterações realizadas no documento ativo.
'--------------------------------------------------------------------------------
Sub UndoAllChanges()
    On Error GoTo ErrorHandler

    ' Verifica se há um documento ativo
    If Documents.Count = 0 Then
        MsgBox "Nenhum documento está aberto para desfazer alterações.", vbExclamation, "Documento não encontrado"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Verifica se há alterações a serem desfeitas
    If doc.Saved Then
        MsgBox "Nenhuma alteração foi realizada desde o último salvamento.", vbInformation, "Nada a desfazer"
        Exit Sub
    End If

    ' Desativa a atualização da tela para melhorar a performance
    Application.ScreenUpdating = False

    ' Desfaz todas as alterações realizadas no documento
    While doc.Undo
        ' Continua desfazendo até que não haja mais alterações
    Wend

    ' Restaura a atualização da tela
    Application.ScreenUpdating = True

    ' Mensagem de conclusão
    MsgBox "Todas as alterações realizadas desde a abertura do arquivo ou do último salvamento foram desfeitas com sucesso.", _
           vbInformation, "Alterações Desfeitas"
    Exit Sub

ErrorHandler:
    ' Tratamento de erros
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro ao Desfazer Alterações"
End Sub
