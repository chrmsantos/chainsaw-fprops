Sub ChamarGoogleGeminiAPI()
    Dim http As Object
    Dim url As String
    Dim apiKey As String
    Dim prompt As String
    Dim jsonResponse As String
    Dim jsonParsed As Object
    Dim respostaTexto As String
    Dim jsonBody As String

    ' Certifique-se de que você tem a biblioteca JSON instalada no seu projeto VBA (VBA-JSON)
    ' E a referência para Microsoft XML, v6.0 ou superior.

    ' Sua chave de API - insira aqui.
    ' Obtenha sua chave de API no Google AI Studio: https://aistudio.google.com/app/apikey
    apiKey = "SUA_CHAVE_API_AQUI" ' <--- SUBSTITUA PELA SUA CHAVE DE API REAL

    ' URL da API do Google Gemini. Para o modelo 'gemini-pro'.
    ' O endpoint mudou, agora é models/MODEL_ID:generateContent
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=" & apiKey

    ' Prompt a ser enviado
    prompt = "Cumprimente a Câmara Municipal de Santa Bárbara dOeste de forma formal e educada."

    ' Cria objeto XMLHTTP para requisição
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Monta o corpo da requisição JSON para o modelo Gemini
    ' A estrutura do JSON para o Gemini é diferente, usando "contents" e "parts".
    jsonBody = "{""contents"":[{""parts"":[{""text"":""" & Replace(prompt, """", "\""") & ""!"}]}]}"

    ' Abre a conexão POST
    http.Open "POST", url, False ' O terceiro parâmetro 'False' indica que a requisição é síncrona.

    ' Define os headers
    http.setRequestHeader "Content-Type", "application/json"
    ' A chave de API agora é passada na URL, então o header Authorization não é mais necessário para este método.

    ' Envia a requisição
    http.send jsonBody

    ' Verifica se a requisição foi bem sucedida
    If http.Status = 200 Then
        jsonResponse = http.responseText
        Debug.Print "Resposta JSON bruta: " & jsonResponse ' Para depuração

        ' Parse do JSON - usando JsonConverter (necessário importar módulo JSON)
        ' Instale o JsonConverter de https://github.com/VBA-tools/VBA-JSON
        Set jsonParsed = JsonConverter.ParseJson(jsonResponse)

        ' Acessa a resposta de texto do Gemini. A estrutura é aninhada.
        ' A resposta geralmente vem em "candidates" -> "content" -> "parts" -> "text"
        If Not jsonParsed("candidates") Is Nothing Then
            If Not jsonParsed("candidates")(1)("content") Is Nothing Then
                If Not jsonParsed("candidates")(1)("content")("parts") Is Nothing Then
                    If Not jsonParsed("candidates")(1)("content")("parts")(1)("text") Is Nothing Then
                        respostaTexto = jsonParsed("candidates")(1)("content")("parts")(1)("text")
                    Else
                        respostaTexto = "Não foi possível extrair o texto da resposta."
                    End If
                End If
            End If
        End If

        ' Insere a resposta no documento Word
        ' Certifique-se de que há um documento Word aberto e a seleção está onde você quer inserir.
        On Error Resume Next ' Para evitar erro se não houver seleção ou documento aberto
        Selection.TypeText respostaTexto
        On Error GoTo 0 ' Retorna o tratamento de erro normal
        MsgBox "Resposta da API Gemini inserida no documento: " & respostaTexto, vbInformation, "Sucesso!"

    Else
        ' Se a requisição falhar, exibe uma mensagem de erro
        MsgBox "Erro na requisição: " & http.Status & " - " & http.statusText & vbCrLf & _
               "Resposta: " & http.responseText, vbCritical, "Erro na API Gemini"
    End If

    Set http = Nothing
    Set jsonParsed = Nothing
End Sub