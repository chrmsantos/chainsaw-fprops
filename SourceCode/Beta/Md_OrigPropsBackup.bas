Option Explicit

'--------------------------------------------------------------------------------
' FUNÇÃO: Main_OPB
' Cria um backup do documento ativo, garantindo que o diretório seja criado se não existir.
' Retorna o caminho do backup criado ou "" em caso de erro.
'--------------------------------------------------------------------------------
Public Function Main_OPB(doc As Document) As String
    On Error GoTo ErrorHandler
    
    Const MAX_PATH_LENGTH As Integer = 260
    
    Dim backupFolder As String
    Dim backupPath As String
    Dim docName As String
    
    ' Define o nome do documento e o caminho do backup
    If doc.Path = "" Then
        docName = "Documento1.docx"
    Else
        docName = doc.Name
    End If
    
    ' Sanitize and ensure filename isn't too long
    docName = SanitizeFileName(docName)
    If Len(docName) > 50 Then  ' Arbitrary safe limit for filename part
        docName = Left(docName, 45) & "..." & Right(docName, 5)
    End If
    
    backupFolder = Environ("USERPROFILE") & "\RevisorDeProposituras\BkpsPropsOriginais\" & Format(Date, "yyyy-mm-dd") & "\"
    
    ' Validate total path length won't exceed limits when combined with filename
    If Len(backupFolder) + Len(docName) > MAX_PATH_LENGTH Then
        MsgBox "O caminho do backup é muito longo para ser salvo.", vbExclamation, "Aviso de Backup"
        Main_OPB = ""
        Exit Function
    End If
    
    ' Garante que o diretório do backup seja criado
    If Not CreateDirectoryIfNeeded(backupFolder) Then
        MsgBox "Não foi possível criar o diretório de backup: " & backupFolder & vbCrLf & _
               "A execução continuará sem o backup.", vbExclamation, "Aviso de Backup"
        Main_OPB = ""
        Exit Function
    End If
    
    ' Define o caminho completo do backup com verificação final de duplicados 
    backupPath = GetUniqueFileName(backupFolder, docName)
    
    ' Salva o documento no caminho do backup com tratamento específico para erros de salvamento
    On Error Resume Next
    doc.SaveAs FileName:=backupPath, FileFormat:=doc.SaveFormat  ' Preserve original format by default
    
Select Case Err.Number 
Case 0     ' Success - no error
        
Case Else  ' Any other error        
        MsgBox "Erro ao salvar o backup: " & Err.Description, vbExclamation, "Aviso de Backup"
        Main_OPB = ""
        Err.Clear        
Exit Function        
End Select        
On Error GoTo ErrorHandler
    
Main_OPB = backupPath    
Exit Function

ErrorHandler:    
MsgBox "Erro ao criar o backup: " & Err.Description & vbCrLf & _       
"A execução continuará sem o backup.", vbExclamation, "Aviso de Backup"    
Main_OPB = ""    
End Function


'-------------------------------------------------------------------------------- 
' FUNÇÃO: GetUniqueFileName  
' Gera um nome de arquivo único adicionando um sufixo numérico se necessário  
'-------------------------------------------------------------------------------- 
Private Function GetUniqueFileName(folderPath As String, baseName As String) As String
    Dim fso As Object
    Dim counter As Integer
    Dim testPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Remove qualquer barra invertida final do caminho
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    testPath = folderPath & "\" & baseName

    If Not fso.FileExists(testPath) Then
        GetUniqueFileName = testPath
        Set fso = Nothing
        Exit Function
    End If

    counter = 1

    Do While True
        testPath = folderPath & "\" & fso.GetBaseName(baseName) & "_" & counter
        If Right(baseName, 5) Like "*.[a-zA-Z]??" Then
            testPath = testPath & "." & fso.GetExtensionName(baseName)
        End If
        If Not fso.FileExists(testPath) Then Exit Do
        counter = counter + 1
        If counter > 1000 Then Exit Do  ' Prevents infinite loops in extreme cases
    Loop

    GetUniqueFileName = testPath
    Set fso = Nothing
End Function


'-------------------------------------------------------------------------------- 
' FUNÇÃO: SanitizeFileName  
' Remove caracteres inválidos e problemas comuns em nomes de arquivo  
'-------------------------------------------------------------------------------- 
Public Function SanitizeFileName(fileName As String) As String   
Dim invalidChars As String   
Dim i As Integer   
Dim result As String    

invalidChars = "\/:*?""<>|"     
result = fileName    

For i=1 To Len(invalidChars)
result=Replace(result,Mid(invalidChars,i ,1),"_")
Next i    

result=Trim(result)
While InStr(result,"..")>0     
result=Replace(result,"..",".")
Wend    

SanitizeFileName=result   
End Function


'--------------------------------------------------------------------------------
' FUNÇÃO: CreateDirectoryIfNeeded
' Cria o diretório especificado (e subdiretórios) se não existirem.
' Retorna True se o diretório existir ou for criado com sucesso, False caso contrário.
'--------------------------------------------------------------------------------
Public Function CreateDirectoryIfNeeded(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim fso As Object
    Dim folders() As String
    Dim currentPath As String
    Dim i As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Trim(folderPath)
    If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath) - 1)
    
    If fso.FolderExists(folderPath) Then
        CreateDirectoryIfNeeded = True
        Set fso = Nothing
        Exit Function
    End If

    folders = Split(folderPath, "\")
    currentPath = folders(0)
    If InStr(currentPath, ":") = 0 Then currentPath = folders(0) & "\"
    For i = 1 To UBound(folders)
        If currentPath <> "" And Right(currentPath, 1) <> "\" Then currentPath = currentPath & "\"
        currentPath = currentPath & folders(i)
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder currentPath
        End If
    Next i

    CreateDirectoryIfNeeded = fso.FolderExists(folderPath)
    Set fso = Nothing
    Exit Function

ErrorHandler:
    CreateDirectoryIfNeeded = False
    Set fso = Nothing
End Function


