    Sub GlobalPreRunnerMain()
    ' GlobalPreRunner: Verifica condições iniciais antes de executar outras rotinas
    
    ' Otimização de desempenho: desabilita atualizações de tela e alertas
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatando documento..."
    End With

    ' Verifica se há um documento ativo
    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento ativo encontrado. Por favor, abra um documento para formatar.", _
               vbExclamation, "Documento Inativo"
        Exit Sub
    End If
    ' Verifica se o documento está protegido
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "O documento está protegido. Por favor, desproteja-o antes de continuar.", _
               vbExclamation, "Documento Protegido"
        Exit Sub
    End If
    ' Verifica se o documento contém conteúdo
    If ActiveDocument.Content.Text = "" Then
        MsgBox "O documento está vazio. Por favor, adicione conteúdo antes de formatar.", _
               vbExclamation, "Documento Vazio"
        Exit Sub
    End If
    ' Verifica se o documento é do tipo Word
    If ActiveDocument.Type <> wdTypeDocument Then
        MsgBox "O documento ativo não é um documento do Word. Por favor, abra um documento do Word para formatar.", _
               vbExclamation, "Tipo de Documento Inválido"
        Exit Sub
    End If
    ' Verifica se o documento está salvo
    If ActiveDocument.Saved = False Then
        Dim response As VbMsgBoxResult
        response = MsgBox("O documento não foi salvo. Deseja salvar antes de formatar?", _
                          vbYesNo + vbQuestion, "Salvar Documento")
        If response = vbYes Then
            ActiveDocument.Save
        Else
            MsgBox "Por favor, salve o documento antes de continuar.", vbExclamation, "Documento Não Salvo"
            Exit Sub
        End If
    End If
    ' Verifica se o documento está em modo de exibição de layout de impressão
    If ActiveWindow.View.Type <> wdPrintView Then
        MsgBox "Por favor, mude o modo de exibição para 'Layout de Impressão' antes de formatar.", _
               vbExclamation, "Modo de Exibição Inválido"
        Exit Sub
    End If

    
    
    

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
    HandleError "Main"
End Sub
