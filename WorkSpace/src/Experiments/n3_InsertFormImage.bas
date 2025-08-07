
'MÓDULO EM VERSÃO ALPHA

' Salvar imagem de um controle em arquivo temporário
Private Sub InserirImagemDoUserForm()
    Dim imgPath As String
    imgPath = Environ("TEMP") & "\tempImg.bmp"

    ' Salva a imagem do controle Image para um arquivo
    SavePicture UserForm1.Image1.Picture, imgPath

    ' Abre o Word e insere a imagem
    Dim wdApp As Object
    Dim wdDoc As Object
    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Add

    wdApp.Visible = True

    wdDoc.Content.InlineShapes.AddPicture FileName:=imgPath, _
        LinkToFile:=False, SaveWithDocument:=True

    Kill imgPath ' Apaga imagem temporária

    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
