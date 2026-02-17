Attribute VB_Name = "M05_OpenAI_API_request_enviar"
Option Explicit

Private Const PAYLOAD_DEBUG_OFF As String = "off"
Private Const PAYLOAD_DEBUG_BASIC As String = "basic"
Private Const PAYLOAD_DEBUG_VERBOSE As String = "verbose"

' =============================================================================
' Módulo: M05_OpenAI_API_request_enviar
' Propósito:
' - Construir e enviar pedidos para a Responses API com tratamento de retries/erros.
' - Extrair campos úteis da resposta JSON para consumo da orquestração.
'
' Atualizações:
' - 2026-02-17 | Codex | Guard rails de payload para invalid_json + debug por nível
'   - Adiciona pré-validação JSON antes de http.Send e bloqueia envio em caso de payload inválido.
'   - Introduz DEBUG_PAYLOAD_LEVEL (OFF|BASIC|VERBOSE), hash e slices contextuais no DEBUG.
' - 2026-02-16 | Codex | Dump opcional do payload final para troubleshooting local
'   - Adiciona escrita do JSON final em C:\Temp\payload.json antes do envio HTTP.
'   - Regista INFO/ALERTA no DEBUG sem expor segredos.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - OpenAI_Executar (Function): rotina pública do módulo.
' =============================================================================

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
    Dim payloadDebugLevel As String
    payloadDebugLevel = M05_GetPayloadDebugLevel()

    Dim payloadErr As String
    If Not M05_ValidatePayloadBeforeSend(json, payloadErr) Then
        resultado.Erro = "Payload JSON inválido (pré-validação): " & payloadErr
        Call Debug_Registar(0, dbgPromptId, "ERRO", "", "M05_PAYLOAD_INVALID", _
            "Request bloqueado antes de http.Send. motivo=" & payloadErr, _
            "Corrija a montagem do payload (text.format/schema/extraFragment) e repita.")
        If payloadDebugLevel <> PAYLOAD_DEBUG_OFF Then
            Call M05_DumpPayloadForDebug(json, dbgPromptId, payloadDebugLevel)
        End If
        OpenAI_Executar = resultado
        Exit Function
    End If

    If payloadDebugLevel <> PAYLOAD_DEBUG_OFF Then
        Call M05_DumpPayloadForDebug(json, dbgPromptId, payloadDebugLevel)
    End If

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
    ' HTTP POST com retry 5xx
    ' -------------------------
    Const MAX_RETRIES_5XX As Long = 3
    Dim waitsSec(1 To 3) As Long
    waitsSec(1) = 2
    waitsSec(2) = 6
    waitsSec(3) = 14

    Dim attempt As Long
    Dim httpStatus As Long
    Dim resposta As String
    Dim reqId As String

    attempt = 0

    Do
        attempt = attempt + 1
        reqId = ""

        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

        ' Mais tolerante para payloads com base64 (INLINE_BASE64)
        http.SetTimeouts 15000, 15000, 60000, 120000

        http.Open "POST", OPENAI_ENDPOINT, False
        http.SetRequestHeader "Authorization", "Bearer " & apiKey
        http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        http.SetRequestHeader "Accept", "application/json"

        http.Send json

        httpStatus = CLng(http.Status)
        resposta = CStr(http.ResponseText)

        ' Guardar JSON bruto para auditoria (sempre)
        resultado.rawResponseJson = resposta
        resultado.httpStatus = httpStatus

        If httpStatus >= 200 And httpStatus < 300 Then
            resultado.responseId = ExtrairCampoJsonSimples(resposta, """id"":")
            resultado.outputText = ExtrairTextoOutputText(resposta)
            OpenAI_Executar = resultado
            Exit Function
        End If

        If httpStatus >= 500 And httpStatus <= 599 Then
            reqId = ExtrairCampoJsonSimples(resposta, """request_id"":")
            If reqId = "" Then
                reqId = M05_ExtrairReqIdDeTexto(resposta)
            End If

            If attempt <= MAX_RETRIES_5XX Then
                Dim waitS As Long
                waitS = waitsSec(attempt)

                On Error Resume Next
                Call Debug_Registar(0, dbgPromptId, "ALERTA", "", "API_RETRY_5XX", _
                    "HTTP " & httpStatus & " (server_error) | attempt=" & attempt & "/" & MAX_RETRIES_5XX & _
                    " | wait=" & waitS & "s" & IIf(reqId <> "", " | req_id=" & reqId, ""), _
                    "Sugestão: erro transitório do servidor. O PIPELINER vai repetir automaticamente.")
                On Error GoTo TrataErro

                Call M05_SleepSeconds(waitS)
            Else
                resultado.Erro = "HTTP " & httpStatus & " - " & resposta
                On Error Resume Next
                Call Debug_Registar(0, dbgPromptId, "ERRO", "", "API", _
                    "HTTP " & httpStatus & " (server_error) | retries_esgotados=" & MAX_RETRIES_5XX & _
                    IIf(reqId <> "", " | req_id=" & reqId, ""), _
                    "Sugestão: tente mais tarde. Se persistir, contacte suporte com o req_id.")
                On Error GoTo TrataErro

                OpenAI_Executar = resultado
                Exit Function
            End If
        Else
            resultado.Erro = "HTTP " & httpStatus & " - " & resposta
            OpenAI_Executar = resultado
            Exit Function
        End If
    Loop

TrataErro:
    resultado.Erro = "Erro VBA: " & Err.Description
    OpenAI_Executar = resultado
End Function

Private Sub M05_SleepSeconds(ByVal seconds As Long)
    On Error GoTo EH

    If seconds <= 0 Then Exit Sub
    Application.Wait Now + TimeSerial(0, 0, seconds)
    Exit Sub

EH:
    ' Falha silenciosa: não bloquear o pipeline por causa do wait
End Sub

Private Function M05_ExtrairReqIdDeTexto(ByVal s As String) As String
    On Error GoTo EH

    Dim p As Long
    p = InStr(1, s, "req_", vbTextCompare)
    If p = 0 Then
        M05_ExtrairReqIdDeTexto = ""
        Exit Function
    End If

    Dim i As Long
    Dim ch As String
    Dim tok As String
    tok = ""

    For i = p To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            tok = tok & ch
        Else
            Exit For
        End If
    Next i

    If Left$(tok, 4) = "req_" Then
        M05_ExtrairReqIdDeTexto = tok
    Else
        M05_ExtrairReqIdDeTexto = ""
    End If
    Exit Function

EH:
    M05_ExtrairReqIdDeTexto = ""
End Function



Private Sub M05_DumpPayloadForDebug(ByVal payloadJson As String, ByVal dbgPromptId As String, ByVal debugLevel As String)
    On Error GoTo Falha

    Dim targetPath As String
    targetPath = "C:\Temp\payload.json"

    Dim folderPath As String
    folderPath = Left$(targetPath, InStrRev(targetPath, "\") - 1)

    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    Dim ff As Integer
    ff = FreeFile
    Open targetPath For Output As #ff
    Print #ff, payloadJson
    Close #ff

    Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_PAYLOAD_DUMP", _
        "Payload final gravado em " & targetPath & " | len=" & CStr(Len(payloadJson)) & " | hash=" & M05_PayloadHash(payloadJson), _
        "Use este ficheiro para validar text.format.schema e outros fragmentos antes de novo envio.")

    If debugLevel = PAYLOAD_DEBUG_VERBOSE Then
        Call M05_LogPayloadSlices(payloadJson, dbgPromptId)
    End If
    Exit Sub

Falha:
    On Error Resume Next
    If ff > 0 Then Close #ff
    Call Debug_Registar(0, dbgPromptId, "ALERTA", "", "M05_PAYLOAD_DUMP_FAIL", _
        "Não foi possível gravar payload em C:\Temp\payload.json: " & Err.Description, _
        "Verifique permissões locais e existência da pasta C:\Temp.")
End Sub


Private Function M05_GetPayloadDebugLevel() As String
    On Error GoTo Fallback

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 1 To lr
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "debug_payload_level" Then
            Dim raw As String
            raw = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))

            Select Case raw
                Case PAYLOAD_DEBUG_OFF, PAYLOAD_DEBUG_BASIC, PAYLOAD_DEBUG_VERBOSE
                    M05_GetPayloadDebugLevel = raw
                    Exit Function
            End Select
            Exit For
        End If
    Next r

Fallback:
    M05_GetPayloadDebugLevel = PAYLOAD_DEBUG_BASIC
End Function

Private Function M05_ValidatePayloadBeforeSend(ByVal payloadJson As String, ByRef outErr As String) As Boolean
    outErr = ""

    If Trim$(payloadJson) = "" Then
        outErr = "payload vazio"
        M05_ValidatePayloadBeforeSend = False
        Exit Function
    End If

    Dim syntaxErr As String
    If Not M05_ValidateJsonSyntaxBasic(payloadJson, syntaxErr) Then
        outErr = syntaxErr
        M05_ValidatePayloadBeforeSend = False
        Exit Function
    End If

    If InStr(1, payloadJson, """strict"":true}}}", vbTextCompare) > 0 Then
        outErr = "detetado padrão de fecho extra em text.format (strict=true}}})"
        M05_ValidatePayloadBeforeSend = False
        Exit Function
    End If

    M05_ValidatePayloadBeforeSend = True
End Function

Private Function M05_ValidateJsonSyntaxBasic(ByVal s As String, ByRef outErr As String) As Boolean
    On Error GoTo Falha

    Dim i As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim objDepth As Long
    Dim arrDepth As Long

    inString = False
    escaped = False
    objDepth = 0
    arrDepth = 0

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            End If
        Else
            Select Case ch
                Case """"
                    inString = True
                Case "{"
                    objDepth = objDepth + 1
                Case "}"
                    objDepth = objDepth - 1
                    If objDepth < 0 Then
                        outErr = "fecho de objeto sem abertura na posição " & CStr(i)
                        M05_ValidateJsonSyntaxBasic = False
                        Exit Function
                    End If
                Case "["
                    arrDepth = arrDepth + 1
                Case "]"
                    arrDepth = arrDepth - 1
                    If arrDepth < 0 Then
                        outErr = "fecho de array sem abertura na posição " & CStr(i)
                        M05_ValidateJsonSyntaxBasic = False
                        Exit Function
                    End If
            End Select
        End If
    Next i

    If inString Then
        outErr = "aspas não terminadas"
        M05_ValidateJsonSyntaxBasic = False
        Exit Function
    End If

    If objDepth <> 0 Or arrDepth <> 0 Then
        outErr = "delimitadores desbalanceados (obj=" & CStr(objDepth) & ", arr=" & CStr(arrDepth) & ")"
        M05_ValidateJsonSyntaxBasic = False
        Exit Function
    End If

    M05_ValidateJsonSyntaxBasic = True
    Exit Function

Falha:
    outErr = "falha na validação sintática: " & Err.Description
    M05_ValidateJsonSyntaxBasic = False
End Function

Private Sub M05_LogPayloadSlices(ByVal payloadJson As String, ByVal dbgPromptId As String)
    On Error GoTo Falha

    Call M05_LogSliceNear(payloadJson, """schema""", "M05_PAYLOAD_SCHEMA_SLICE", dbgPromptId)
    Call M05_LogSliceNear(payloadJson, """strict""", "M05_PAYLOAD_STRICT_SLICE", dbgPromptId)
    Call M05_LogSliceNear(payloadJson, """format""", "M05_PAYLOAD_FORMAT_SLICE", dbgPromptId)

    Dim tailLen As Long
    tailLen = 200
    If Len(payloadJson) < tailLen Then tailLen = Len(payloadJson)
    If tailLen > 0 Then
        Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_PAYLOAD_TAIL", _
            "tail_last_" & CStr(tailLen) & "=" & Mid$(payloadJson, Len(payloadJson) - tailLen + 1), _
            "Inspecione os últimos carateres para detetar fechos/vírgulas indevidos.")
    End If
    Exit Sub

Falha:
    Call Debug_Registar(0, dbgPromptId, "ALERTA", "", "M05_PAYLOAD_SLICE_FAIL", _
        "Não foi possível gerar slices do payload: " & Err.Description, _
        "Verifique o formato do payload em C:\Temp\payload.json.")
End Sub

Private Sub M05_LogSliceNear(ByVal payloadJson As String, ByVal token As String, ByVal debugTag As String, ByVal dbgPromptId As String)
    Dim p As Long
    p = InStr(1, payloadJson, token, vbTextCompare)
    If p = 0 Then Exit Sub

    Dim startPos As Long
    Dim stopPos As Long
    startPos = p - 90
    If startPos < 1 Then startPos = 1
    stopPos = p + 220
    If stopPos > Len(payloadJson) Then stopPos = Len(payloadJson)

    Call Debug_Registar(0, dbgPromptId, "INFO", "", debugTag, _
        Mid$(payloadJson, startPos, stopPos - startPos + 1), _
        "Slice de payload para troubleshooting local.")
End Sub

Private Function M05_PayloadHash(ByVal payloadJson As String) As String
    Dim h As Double
    Dim i As Long

    h = 5381#
    For i = 1 To Len(payloadJson)
        h = ((h * 33#) + AscW(Mid$(payloadJson, i, 1))) Mod 2147483647#
    Next i

    M05_PayloadHash = CStr(CLng(h))
End Function
