' Módulo exclusivamente de entry point e orquestração do fluxo do botão Padronizar Documento

Public Sub BtnMAIN()
    On Error GoTo ErrHandler

    ' Otimização de desempenho: desabilita atualizações de tela e alertas
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Formatando documento..."
    End With

    ' Verificação de pré-requisitos
    Call GlobalChecking
    
    ' Rotina de formatação global
    Call GlobalFormatter

    ' Restaura o estado da aplicação
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With

    Exit Sub

ErrHandler:
    MsgBox "Ocorreu um erro ao padronizar o documento:" & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub

