Option Explicit ' Garante que todas as variáveis sejam declaradas explicitamente

' =================================================================================================
' ||                                                                                               ||
' ||  INSTRUÇÕES IMPORTANTES ANTES DE EXECUTAR (MANTIDAS DO SEU CÓDIGO ORIGINAL):                   ||
' ||                                                                                               ||
' ||  1. Habilitar a Referência "Microsoft XML, v6.0":                                             ||
' ||     - No editor VBA (Alt + F11), vá em Ferramentas > Referências...                             ||
' ||     - Marque a caixa de seleção para "Microsoft XML, v6.0" (ou a versão mais recente).          ||
' ||                                                                                               ||
' ||  2. Importar o Módulo JsonConverter:                                                          ||
' ||     - Baixe o arquivo 'JsonConverter.bas' do GitHub: https://github.com/VBA-tools/VBA-JSON    ||
' ||     - No editor VBA, vá em Arquivo > Importar Arquivo... e selecione o arquivo baixado.        ||
' ||     - Isso adicionará um módulo chamado 'JsonConverter' ao seu projeto VBA.                     ||
' ||                                                                                               ||
' ||  3. Inserir sua Chave de API:                                                                 ||
' ||     - Obtenha sua chave de API no Google AI Studio: https://aistudio.google.com/app/apikey     ||
' ||     - Substitua "SUA_CHAVE_API_AQUI" na variável 'apiKey' abaixo pela sua chave real.          ||
' ||     - Mantenha sua chave de API em segredo e não a compartilhe.                               ||
' ||                                                                                               ||
' =================================================================================================

Sub GeminiAPIconn()
    ' Declaração de todas as variáveis
    Dim http As Object
    Dim url As String
    Dim apiKey As String
    Dim prompt As String
    Dim jsonResponse As String
    Dim jsonParsed As Object
    Dim respostaTexto As String
    Dim jsonBody As String
    Dim jsonTemplate As String
    Dim escapedPrompt As String

    ' --- CONFIGURAÇÃO ---
    ' 1. Insira sua chave de API aqui.
    apiKey = "AIzaSyD3kBMb-NDZZa2XNVo7e7z0_8N1VHBJOAs" ' <--- SUBSTITUA PELA SUA CHAVE DE API REAL

    ' 2. URL da API do Google Gemini.
    ' --- ATUALIZAÇÃO FINAL (CORREÇÃO DO ERRO 404) ---
    ' Voltamos para o endpoint 'v1beta' que é o correto para 'generateContent'
    ' e usamos o modelo padrão e estável 'gemini-1.0-pro'.
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" & apiKey
    
    ' 3. Prompt (pergunta) a ser enviado para o Gemini.
    prompt = "Qual é a capital dos EUA?"

    ' Verifica se a chave de API foi inserida.
    If apiKey = "" Then
        MsgBox "ERRO: Por favor, insira sua chave de API na variável 'apiKey' antes de executar o código.", vbCritical, "Chave de API Faltando"
        Exit Sub
    End If

    ' --- PREPARAÇÃO DA REQUISIÇÃO ---
    ' Cria o objeto de requisição HTTP.
    ' Usamos late binding para não depender da versão exata do MSXML.
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then
        Set http = CreateObject("MSXML2.XMLHTTP")
    End If
    On Error GoTo 0
    
    If http Is Nothing Then
        MsgBox "Não foi possível criar o objeto HTTP. Verifique se a referência 'Microsoft XML, v6.0' está habilitada.", vbCritical, "Erro de Dependência"
        Exit Sub
    End If


    ' --- CORREÇÃO PRINCIPAL: Montagem do Corpo JSON ---
    ' Escapa quaisquer caracteres especiais no prompt para garantir que o JSON seja válido.
    escapedPrompt = EscapeJsonString(prompt)

    ' Em vez de construir o JSON manualmente com concatenação (que é frágil),
    ' usamos um modelo e a função Replace. É mais seguro e fácil de ler.
    jsonTemplate = "{""contents"":[{""parts"":[{""text"":""__PROMPT__""}]}]}"
    jsonBody = Replace(jsonTemplate, "__PROMPT__", escapedPrompt)
    
    ' Para depuração: visualize o JSON que será enviado no Painel de Verificação Imediata (Ctrl+G).
    Debug.Print "JSON Body Enviado: " & jsonBody

    ' --- ENVIO DA REQUISIÇÃO HTTP ---
    ' Abre a conexão POST de forma síncrona ('False').
    http.Open "POST", url, False
    ' Define o cabeçalho obrigatório para a API.
    http.setRequestHeader "Content-Type", "application/json"
    ' Envia a requisição com o corpo JSON.
    http.send jsonBody

    ' --- PROCESSAMENTO DA RESPOSTA ---
    ' Verifica se a requisição foi bem-sucedida (status 200 = OK).
    If http.Status = 200 Then
        jsonResponse = http.responseText
        ' Para depuração: visualize a resposta JSON bruta.
        Debug.Print "Status: " & http.Status
        Debug.Print "StatusText: " & http.StatusText
        Debug.Print "Resposta JSON bruta: " & jsonResponse

        ' Analisa (parse) a string de resposta JSON em um objeto VBA.
        On Error Resume Next
        Set jsonParsed = JsonConverter.ParseJson(jsonResponse)
        On Error GoTo 0

        If jsonParsed Is Nothing Then
            respostaTexto = "Falha ao analisar a resposta JSON do servidor. A resposta pode estar malformada."
        Else
            ' Extrai o texto da resposta de forma segura, verificando cada nível do objeto.
            ' Isso evita erros caso a estrutura da resposta mude ou um erro seja retornado.
            If jsonParsed.Exists("candidates") And jsonParsed("candidates").Count > 0 Then
                If jsonParsed("candidates")(1)("content")("parts").Count > 0 Then
                    respostaTexto = jsonParsed("candidates")(1)("content")("parts")(1)("text")
                Else
                    respostaTexto = "A estrutura da resposta não continha 'parts'."
                End If
            ' A API também pode retornar um erro estruturado em JSON.
            ElseIf jsonParsed.Exists("error") Then
                respostaTexto = "A API retornou um erro: " & vbCrLf & jsonParsed("error")("message")
            Else
                respostaTexto = "A resposta JSON não continha 'candidates' nem uma mensagem de 'error' reconhecível."
            End If
        End If

        ' Exibe o resultado final.
        MsgBox respostaTexto, vbInformation, "Resposta da API Gemini"

    Else ' Se a requisição falhar (status diferente de 200)
        ' Monta uma mensagem de erro detalhada.
        Dim erroMsg As String
        erroMsg = "Erro na requisição: " & http.Status & " - " & http.StatusText & vbCrLf & vbCrLf & _
                  "Resposta do Servidor:" & vbCrLf & http.responseText
        MsgBox erroMsg, vbCritical, "Erro na API Google Gemini"
    End If

    ' Libera os objetos da memória para evitar vazamentos.
    Set http = Nothing
    Set jsonParsed = Nothing
End Sub


' --------------------------------------------------------------------------
' FUNÇÃO AUXILIAR PARA ESCAPAR CARACTERES ESPECIAIS EM STRINGS JSON
' (Essencial para evitar o erro "Invalid JSON payload")
' --------------------------------------------------------------------------
Private Function EscapeJsonString(ByVal textToEscape As String) As String
    Dim result As String
    result = textToEscape
    ' A ordem das substituições é importante. A barra invertida deve ser a primeira.
    result = Replace(result, "\", "\\") ' Escapa a própria barra invertida
    result = Replace(result, """", "\""") ' Escapa aspas duplas
    result = Replace(result, vbCr, "\r")   ' Escapa retorno de carro
    result = Replace(result, vbLf, "\n")   ' Escapa nova linha
    result = Replace(result, vbTab, "\t")  ' Escapa tabulação
    ' Outros caracteres de controle que podem quebrar o JSON
    result = Replace(result, vbBack, "\b") ' Escapa backspace
    result = Replace(result, vbFormFeed, "\f") ' Escapa form feed
    
    EscapeJsonString = result
End Function


