Public Sub Main_SuppAndDoc()
    ' Abre o link de suporte/documentação no navegador padrão
    On Error GoTo ErrorHandler
    Dim supportUrl As String
    supportUrl = "https://github.com/chrmsantos/RevisorDeProposituras" ' Altere para o link desejado

    If supportUrl <> "" Then
        ActiveDocument.FollowHyperlink Address:=supportUrl, NewWindow:=True
    Else
        MsgBox "URL de suporte não definida.", vbExclamation, "Atenção"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao tentar abrir o link de suporte:" & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub