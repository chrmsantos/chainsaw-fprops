Public Sub Main_SuppAndDoc()
    ' Abre o link de suporte/documentação no navegador padrão
    On Error GoTo ErrorHandler
    Dim supportUrl As String
    supportUrl = "https://github.com/chrmsantos/RevisorDeProposituras" ' Altere para o link desejado

    Dim ret As Variant
    ret = Shell("cmd /c start " & supportUrl, vbHide)
    If ret = 0 Then
        MsgBox "Não foi possível abrir o navegador. Verifique as configurações do sistema.", vbExclamation, "Erro ao Abrir Link"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao tentar abrir o link de suporte:" & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub