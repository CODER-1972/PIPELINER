Attribute VB_Name = "M05_OpenAI_API_request_enviar"
Option Explicit

Private Const OPENAI_ENDPOINT As String = "https://api.openai.com/v1/responses"

' ============================================================
' JSON helpers (escape / unescape / parsing simples)
' ============================================================

Private Function JsonEscapar(ByVal s As String) As String
    JsonEscapar = Json_EscapeString(CStr(s))
End Function



Private Function EncontrarFimStringJson(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long, escaped As Boolean
    escaped = False

    For i = startPos To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If escaped Then
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            EncontrarFimStringJson = i
            Exit Function
        End If
    Next i

    EncontrarFimStringJson = 0
End Function

Private Function DesescaparJson(ByVal s As String) As String
    Dim i As Long, result As String
    result = ""
    i = 1

    Do While i <= Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch <> "\" Then
            result = result & ch
            i = i + 1
        Else
            If i = Len(s) Then Exit Do

            Dim nxt As String
            nxt = Mid$(s, i + 1, 1)

            Select Case nxt
                Case """": result = result & """": i = i + 2
                Case "\": result = result & "\": i = i + 2
                Case "/": result = result & "/": i = i + 2
                Case "n": result = result & vbLf: i = i + 2
                Case "r": result = result & vbCr: i = i + 2
                Case "t": result = result & vbTab: i = i + 2

                Case "u"
                    If i + 5 <= Len(s) Then
                        Dim hex4 As String
                        hex4 = Mid$(s, i + 2, 4)

                        If IsHex4(hex4) Then
                            result = result & ChrW$(CLng("&H" & hex4))
                            i = i + 6
                        Else
                            result = result & "\u"
                            i = i + 2
                        End If
                    Else
                        result = result & "\u"
                        i = i + 2
                    End If

                Case Else
                    result = result & nxt
                    i = i + 2
            End Select
        End If
    Loop

    DesescaparJson = result
End Function

Private Function IsHex4(ByVal texto As String) As Boolean
    Dim i As Long
    If Len(texto) <> 4 Then IsHex4 = False: Exit Function

    For i = 1 To 4
        Dim c As String
        c = Mid$(texto, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", c, vbBinaryCompare) = 0 Then
            IsHex4 = False
            Exit Function
        End If
    Next i

    IsHex4 = True
End Function

Private Function ExtrairTextoOutputText(ByVal json As String) As String
    Dim pos As Long: pos = 1
    Dim saida As String: saida = ""

    Do
        Dim posOut As Long
        posOut = InStr(pos, json, "output_text")
        If posOut = 0 Then Exit Do

        Dim posTextKey As Long
        posTextKey = InStr(posOut, json, """text""")
        If posTextKey = 0 Then Exit Do

        Dim posDoisPontos As Long
        posDoisPontos = InStr(posTextKey, json, ":")
        If posDoisPontos = 0 Then Exit Do

        Dim posAspaInicio As Long
        posAspaInicio = InStr(posDoisPontos, json, """")
        If posAspaInicio = 0 Then Exit Do

        Dim posAspaFim As Long
        posAspaFim = EncontrarFimStringJson(json, posAspaInicio + 1)
        If posAspaFim = 0 Then Exit Do

        Dim raw As String
        raw = Mid$(json, posAspaInicio + 1, posAspaFim - posAspaInicio - 1)

        If saida <> "" Then saida = saida & vbLf
        saida = saida & DesescaparJson(raw)

        pos = posAspaFim + 1
    Loop

    ExtrairTextoOutputText = Trim$(saida)
End Function

Private Function ExtrairCampoJsonSimples(ByVal json As String, ByVal chave As String) As String
    Dim p As Long
    p = InStr(1, json, chave)
    If p = 0 Then Exit Function

    Dim pAspa1 As Long
    pAspa1 = InStr(p + Len(chave), json, """")
    If pAspa1 = 0 Then Exit Function

    Dim pAspa2 As Long
    pAspa2 = EncontrarFimStringJson(json, pAspa1 + 1)
    If pAspa2 = 0 Then Exit Function

    ExtrairCampoJsonSimples = Mid$(json, pAspa1 + 1, pAspa2 - pAspa1 - 1)
End Function

' ============================================================
' Normalizacao e validacao do input JSON literal
'   - garante que input e um array JSON quando fornecido
' ============================================================

Private Function Json_LastNonWhitespaceChar(ByVal s As String) As String
    Dim i As Long
    For i = Len(s) To 1 Step -1
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then
            Json_LastNonWhitespaceChar = ch
            Exit Function
        End If
    Next i
    Json_LastNonWhitespaceChar = ""
End Function

Private Function NormalizarInputJsonLiteral(ByVal s As String) As String
    Dim t As String
    t = Trim$(CStr(s))

    If t = "" Then
        NormalizarInputJsonLiteral = ""
        Exit Function
    End If

    ' Remover BOM (U+FEFF) se existir
    On Error Resume Next
    If Len(t) > 0 Then
        If AscW(Left$(t, 1)) = &HFEFF Then t = Mid$(t, 2)
    End If
    On Error GoTo 0

    t = Trim$(t)

    ' Se vier como string JSON (entre aspas), retirar aspas e desescapar
    If Len(t) >= 2 Then
        If Left$(t, 1) = """" And Right$(t, 1) = """" Then
            t = Mid$(t, 2, Len(t) - 2)
            t = DesescaparJson(t)
            t = Trim$(t)
        End If
    End If

    ' Se vier como objeto único { ... }, embrulhar em array [ ... ]
    If Left$(t, 1) = "{" And Json_LastNonWhitespaceChar(t) = "}" Then
        t = "[" & t & "]"
    End If

    NormalizarInputJsonLiteral = Trim$(t)
End Function

Private Function ExtraFragment_TemTools(ByVal extraFragmentSemInput As String) As Boolean
    Dim t As String
    t = Trim$(CStr(extraFragmentSemInput))

    If t = "" Then
        ExtraFragment_TemTools = False
        Exit Function
    End If

    ' Apenas considerar tools explícitas.
    ' NOTA: tool_choice sozinho NÃO deve desactivar a auto-injecção de tools,
    ' porque a API exige que o tool escolhido exista em "tools".
    If InStr(1, t, """tools""", vbTextCompare) > 0 Then
        ExtraFragment_TemTools = True
        Exit Function
    End If

    ExtraFragment_TemTools = False
End Function


Private Function Modos_Contem(ByVal modos As String, ByVal token As String) As Boolean
    Dim m As String
    m = LCase$(Trim$(CStr(modos)))
    token = LCase$(Trim$(CStr(token)))
    If m = "" Or token = "" Then
        Modos_Contem = False
        Exit Function
    End If
    ' Aceita listas simples: "Web search", "Code Interpreter", "Web search + Code Interpreter", "web_search;code interpreter", etc.
    Modos_Contem = (InStr(1, m, token, vbTextCompare) > 0)
End Function

' ============================================================
' API: OpenAI_Executar  (/v1/responses)
' ============================================================

Public Function OpenAI_Executar( _
    ByVal apiKey As String, _
    ByVal modelo As String, _
    ByVal textoPrompt As String, _
    ByVal temperatura As Double, _
    ByVal maxOutputTokens As Long, _
    ByVal modos As String, _
    ByVal storage As Boolean, _
    ByVal inputJsonLiteralOpcional As String, _
    ByVal extraFragmentSemInput As String, _
    Optional ByVal promptIdForDebug As String = "" _
) As ApiResultado


    Dim resultado As ApiResultado
    Dim dbgPromptId As String
    dbgPromptId = IIf(Trim$(promptIdForDebug) <> "", promptIdForDebug, "M05")

    On Error GoTo TrataErro

    If Trim$(apiKey) = "" Then
        resultado.Erro = "API key vazia."
        OpenAI_Executar = resultado
        Exit Function
    End If

    Dim tempStr As String
    tempStr = Replace(CStr(temperatura), ",", ".")

    Dim storeStr As String
    If storage Then storeStr = "true" Else storeStr = "false"

    Dim usarInputArray As Boolean
    usarInputArray = (Trim$(CStr(inputJsonLiteralOpcional)) <> "")

    Dim inputLit As String
    inputLit = ""

    If usarInputArray Then
        inputLit = NormalizarInputJsonLiteral(inputJsonLiteralOpcional)

        ' GARANTIA: input tem de ser array JSON (para suportar input_file/input_image)
        If inputLit = "" Or Left$(inputLit, 1) <> "[" Or Json_LastNonWhitespaceChar(inputLit) <> "]" Then
            resultado.Erro = "INPUT override fornecido mas nao e um JSON array valido. " & _
                             "Comeca por: [" & Left$(Trim$(inputLit), 60) & "]"
            OpenAI_Executar = resultado
            Exit Function
        End If
    End If

    ' -------------------------
    ' Diagnostico do INPUT (antes de construir payload final)
    ' -------------------------
    Dim hasInputFile As Boolean, hasInputImage As Boolean
    Dim hasFileData As Boolean, hasFileId As Boolean

    hasInputFile = False: hasInputImage = False
    hasFileData = False: hasFileId = False

    If usarInputArray Then
        hasInputFile = (InStr(1, inputLit, """type"":""input_file""", vbTextCompare) > 0)
        hasInputImage = (InStr(1, inputLit, """type"":""input_image""", vbTextCompare) > 0)
        hasFileData = (InStr(1, inputLit, """file_data"":""data:", vbTextCompare) > 0)
        hasFileId = (InStr(1, inputLit, """file_id"":""file-", vbTextCompare) > 0)

        ' GARANTIA: input_file tem de ter file_data OU file_id
        If hasInputFile Then
            If (Not hasFileData) And (Not hasFileId) Then
                resultado.Erro = "INPUT array contem input_file mas nao tem file_data nem file_id."
                OpenAI_Executar = resultado
                Exit Function
            End If
        End If
    End If

    ' -------------------------
    ' Construir payload JSON (/v1/responses)
    ' -------------------------
    Dim json As String
    json = "{""model"":""" & JsonEscapar(modelo) & ""","

    If usarInputArray Then
        ' GARANTIA: passa o array diretamente no campo input (NAO como string)
        json = json & """input"":" & inputLit & ","
    Else
        json = json & """input"":""" & JsonEscapar(textoPrompt) & ""","
    End If

    json = json & """temperature"":" & tempStr & _
                  ",""max_output_tokens"":" & CStr(maxOutputTokens) & _
                  ",""store"":" & storeStr

    ' -------------------------
    ' Tools: web_search
    '  - Regra: nao auto-adicionar web_search quando ha ficheiro/imagem
    '  - Se quiseres web_search com ficheiros, define tools explicitamente em extraFragmentSemInput
    ' -------------------------
    Dim modosWebSearch As Boolean
    modosWebSearch = Modos_Contem(modos, "Web search")

    Dim modosCodeInterpreter As Boolean
    modosCodeInterpreter = Modos_Contem(modos, "Code Interpreter")

    Dim extraTemTools As Boolean
    extraTemTools = ExtraFragment_TemTools(extraFragmentSemInput)

    Dim autoAddWebSearch As Boolean
    autoAddWebSearch = False

    If modosWebSearch Then
        If extraTemTools Then
            autoAddWebSearch = False
        ElseIf hasInputFile Or hasInputImage Then
            autoAddWebSearch = False
        Else
            autoAddWebSearch = True
        End If
    End If

    Dim autoAddCodeInterpreter As Boolean
    autoAddCodeInterpreter = False

    If modosCodeInterpreter Then
        If extraTemTools Then
            autoAddCodeInterpreter = False
        Else
            autoAddCodeInterpreter = True
        End If
    End If

    If autoAddWebSearch Or autoAddCodeInterpreter Then
        Dim toolsFrag As String
        toolsFrag = """tools"":["
        Dim firstTool As Boolean
        firstTool = True

        If autoAddWebSearch Then
            toolsFrag = toolsFrag & "{""type"":""web_search""}"
            firstTool = False
        End If

        If autoAddCodeInterpreter Then
            If Not firstTool Then toolsFrag = toolsFrag & ","
            toolsFrag = toolsFrag & "{""type"":""code_interpreter"",""container"":{""type"":""auto""}}"
        End If

        toolsFrag = toolsFrag & "]"
        json = json & "," & toolsFrag
    End If

    ' merge do extra (sem input)
    If Trim$(extraFragmentSemInput) <> "" Then
        json = json & "," & extraFragmentSemInput
    End If

    json = json & "}"

    ' -------------------------
    ' Log diagnostico (sem despejar base64)
    ' -------------------------
    Dim toolMsg As String
    If modosWebSearch Then
        If autoAddWebSearch Then
            toolMsg = "web_search=ADICIONADO_AUTO"
        ElseIf extraTemTools Then
            toolMsg = "web_search=NAO_AUTO (tools no extra)"
        ElseIf hasInputFile Or hasInputImage Then
            toolMsg = "web_search=NAO_AUTO (ha anexos)"
        Else
            toolMsg = "web_search=NAO_AUTO"
        End If
    Else
        toolMsg = "web_search=OFF"
    End If

        Dim ciMsg As String
    If modosCodeInterpreter Then
        If autoAddCodeInterpreter Then
            ciMsg = "code_interpreter=ADICIONADO_AUTO"
        ElseIf extraTemTools Then
            ciMsg = "code_interpreter=NAO_AUTO (tools no extra)"
        Else
            ciMsg = "code_interpreter=NAO_AUTO"
        End If
    Else
        ciMsg = "code_interpreter=OFF"
    End If
    toolMsg = toolMsg & " | " & ciMsg

On Error Resume Next
    Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_PAYLOAD_CHECK", _
        "endpoint=" & OPENAI_ENDPOINT & _
        " | model=" & modelo & _
        " | input_is_array=" & IIf(usarInputArray, "SIM", "NAO") & _
        " | has_input_file=" & IIf(hasInputFile, "SIM", "NAO") & _
        " | has_file_data=" & IIf(hasFileData, "SIM", "NAO") & _
        " | has_file_id=" & IIf(hasFileId, "SIM", "NAO") & _
        " | has_input_image=" & IIf(hasInputImage, "SIM", "NAO") & _
        " | " & toolMsg & _
        " | payload_len=" & CStr(Len(json)), _
        "")
    On Error GoTo TrataErro

    ' -------------------------
    ' HTTP POST
    ' -------------------------
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Mais tolerante para payloads com base64 (INLINE_BASE64)
    http.SetTimeouts 15000, 15000, 60000, 120000

    http.Open "POST", OPENAI_ENDPOINT, False
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"

    http.Send json

    resultado.httpStatus = CLng(http.status)

    Dim resposta As String
    resposta = CStr(http.ResponseText)

    ' Guardar JSON bruto para auditoria (sempre)
    resultado.rawResponseJson = resposta

    If resultado.httpStatus < 200 Or resultado.httpStatus >= 300 Then
        resultado.Erro = "HTTP " & resultado.httpStatus & " - " & resposta
        OpenAI_Executar = resultado
        Exit Function
    End If

    resultado.responseId = ExtrairCampoJsonSimples(resposta, """id"":")
    resultado.outputText = ExtrairTextoOutputText(resposta)

    OpenAI_Executar = resultado
    Exit Function

TrataErro:
    resultado.Erro = "Erro VBA: " & Err.Description
    OpenAI_Executar = resultado
End Function


