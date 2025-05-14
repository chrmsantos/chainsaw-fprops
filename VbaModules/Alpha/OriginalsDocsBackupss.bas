Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: CreateBackup
' Cria um backup do documento ativo, garantindo que o diretório seja criado se não existir.
'--------------------------------------------------------------------------------
Public Function CreateBackup(doc As Document) As String
    On Error GoTo ErrorHandler

    Dim backupFolder As String
    Dim backupPath As String
    Dim docName As String

    ' Define o nome do documento e o caminho do backup
    docName = IIf(doc.Path = "", "Documento1.docx", doc.Name)
    backupFolder = Environ("USERPROFILE") & "\RevisorDeProposituras\BackupsPropositurasOriginais\" & Format(Date, "yyyy-mm-dd") & "\"

    ' Garante que o diretório do backup seja criado
    If Not CreateDirectoryIfNeeded(backupFolder) Then
        MsgBox "Não foi possível criar o diretório de backup: " & backupFolder & vbCrLf & _
               "A execução continuará sem o backup.", vbExclamation, "Aviso de Backup"
        CreateBackup = "" ' Retorna vazio, mas permite continuar
        Exit Function
    End If

    ' Define o caminho completo do backup
    backupPath = backupFolder & SanitizeFileName(docName)

    ' Salva o documento no caminho do backup
    doc.SaveAs FileName:=backupPath, FileFormat:=wdFormatDocumentDefault ' Compatível com Word 2007 e posterior
    CreateBackup = backupPath
    Exit Function

ErrorHandler:
    MsgBox "Erro ao criar o backup: " & Err.Description & vbCrLf & _
           "A execução continuará sem o backup.", vbExclamation, "Aviso de Backup"
    CreateBackup = "" ' Retorna vazio, mas permite continuar
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: CreateDirectoryIfNeeded
' Cria o diretório especificado, incluindo subdiretórios, se necessário.
'--------------------------------------------------------------------------------
Public Function CreateDirectoryIfNeeded(folderPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim parentFolder As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Verifica se o diretório já existe
    If Not fso.FolderExists(folderPath) Then
        ' Cria diretórios recursivamente
        parentFolder = fso.GetParentFolderName(folderPath)
        If parentFolder <> "" And Not fso.FolderExists(parentFolder) Then
            CreateDirectoryIfNeeded parentFolder
        End If
        fso.CreateFolder(folderPath)
    End If

    ' Verifica novamente se o diretório foi criado com sucesso
    CreateDirectoryIfNeeded = fso.FolderExists(folderPath)
    Exit Function

ErrorHandler:
    MsgBox "Erro ao criar o diretório: " & folderPath & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro no Backup"
    CreateDirectoryIfNeeded = False
End Function

'--------------------------------------------------------------------------------
' FUNÇÃO: SanitizeFileName
' Remove caracteres inválidos do nome do arquivo.
'--------------------------------------------------------------------------------
Public Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String: invalidChars = "\/:*?""<>|"
    Dim i As Integer
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i
    SanitizeFileName = fileName
End Function