Attribute VB_Name = "M05_OpenAI_API_request_enviar"
Option Explicit

' =============================================================================
' Módulo: M05_OpenAI_API_request_enviar
' Propósito:
' - Construir e enviar pedidos para a Responses API com tratamento de retries/erros.
' - Extrair campos úteis da resposta JSON para consumo da orquestração.
'
' Atualizações:
' - 2026-03-01 | Codex | Evita auto-injecao indevida de Code Interpreter com anexos
'   - Quando ha input_file/input_image e nao existe intencao explicita de CI no extra, suprime auto-add de code_interpreter.
'   - Regista alerta M05_CI_AUTO_SUPPRESS para troubleshooting quando Modos inclui Code Interpreter nestas condicoes.
' - 2026-02-28 | Codex | Diagnostico detalhado para context_length_exceeded (HTTP 400)
'   - Regista evento dedicado `API_CONTEXT_LENGTH_EXCEEDED` com metrica de tamanho (payload/input/prompt), anexos e faixas de risco.
'   - Emite mensagem didatica no DEBUG com ação recomendada (reduzir input, anexos text_embed ou max_output_tokens).
' - 2026-02-28 | Codex | Diagnostico operativo com contexto de rede e retry dirigido
'   - Regista started_at/failed_at absolutos, contexto de rede do host e decisao automatica no evento M05_TIMEOUT_DECISION.
'   - Adiciona retry unico para timeout em stage=Send com novo socket e registo de retry_outcome.
' - 2026-02-28 | Codex | Aumenta precisao de causa provavel em timeout
'   - Acrescenta hint de causa com confianca usando fase HTTP, tamanho de payload/resposta e estado parcial.
'   - Regista contexto extra (attempt, payload_len, http_status_partial, response_len, Err.Source) no erro final.
' - 2026-02-28 | Codex | Diagnostico enriquecido para timeouts HTTP
'   - Classifica timeout provavel (resolve/connect/send/receive/outro) com base no tempo decorrido e limites configurados.
'   - Regista tempo decorrido ate falha e valores efetivos de HTTP_TIMEOUT_*_MS no DEBUG (M05_HTTP_TIMEOUT_ERROR).
' - 2026-02-28 | Codex | Preserva detalhes de erro no handler de timeout
'   - Guarda Err.Number/Err.Description antes de logging para evitar perda de mensagem no resultado final.
'   - Inclui diagnostico de timeout (tipo, elapsed e HTTP_TIMEOUT_*_MS) tambem em resultado.Erro no Seguimento.
' - 2026-02-27 | Codex | Diagnóstico com fingerprint e distinção transporte vs contrato
'   - Adiciona fingerprint textual (FP=...) em M05_PAYLOAD_CHECK/M05_HTTP_TIMEOUTS/M05_HTTP_RESULT.
'   - Torna mensagens de M05 mais explicativas para correlacionar facilmente com eventos M10 do mesmo passo.
' - 2026-02-26 | Codex | Torna timeouts HTTP configuráveis via folha Config
'   - Adiciona leitura tolerante das chaves HTTP_TIMEOUT_*_MS (com fallback seguro para defaults atuais).
'   - Regista no DEBUG os timeouts efetivos aplicados e alerta quando houver valor inválido fora de intervalo.
' - 2026-02-17 | Codex | Remove gating de web_search por anexos
'   - Garante que `Modos=Web search` injeta sempre `tools:[{"type":"web_search"}]` quando não existem tools explícitas no extra.
'   - Elimina a dependência de configuração por anexos para evitar execuções sem acesso web neste cenário.
' - 2026-02-17 | Codex | Corrige literal de troubleshooting do preflight para compilar em VBA
'   - Substitui montagem ambígua de escapes na mensagem do DEBUG por literal seguro com aspas duplicadas.
'   - Mantém orientações de escapes JSON sem depender de barras invertidas como pseudo-escape de aspas no VBA.
' - 2026-02-17 | Codex | Preflight estrutural de JSON para reduzir 400 invalid_json
'   - Adiciona verificação de aspas/chaves/arrays e deteção de vírgula final inválida (`,}`/`,]`).
'   - Regista diagnóstico acionável no DEBUG antes do HTTP quando o payload não fecha estruturalmente.
' - 2026-02-17 | Codex | Correção de sintaxe VBA em validação de JSON preflight
'   - Corrige literais com aspas duplas em Select Case e comparações de string para evitar erro de compilação.
'   - Mantém validação de escapes JSON com mensagem de diagnóstico preservada no DEBUG.
' - 2026-02-17 | Codex | Validação preventiva para escape inválido com backslash no JSON
'   - Adiciona deteção de sequências de escape inválidas (ex.: \x) em strings JSON no preflight.
'   - Bloqueia envio com erro acionável no DEBUG e indica escapes válidos após \ (" \\ / b f n r t uXXXX).
' - 2026-02-17 | Codex | Melhoria das sugestões de escaping no preflight de JSON
'   - Detalha escape recomendado por carácter de controlo detectado (ex.: \n, \r, \t, \u00XX).
'   - Expande mensagem de troubleshooting no DEBUG para reduzir tentativa/erro em invalid_json.
' - 2026-02-17 | Codex | Preflight de JSON para diagnosticar invalid_json antes do HTTP
'   - Adiciona validação leve de controlo bruto em strings JSON (CR/LF/TAB não escapados).
'   - Em caso de falha, bloqueia envio e regista snippet/contexto no DEBUG para correção rápida.
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
Private Const HTTP_TIMEOUT_RESOLVE_MS_DEFAULT As Long = 15000
Private Const HTTP_TIMEOUT_CONNECT_MS_DEFAULT As Long = 15000
Private Const HTTP_TIMEOUT_SEND_MS_DEFAULT As Long = 60000
Private Const HTTP_TIMEOUT_RECEIVE_MS_DEFAULT As Long = 120000
Private Const HTTP_TIMEOUT_MIN_MS As Long = 1000
Private Const HTTP_TIMEOUT_MAX_MS As Long = 900000
' ============================================================
' JSON helpers (escape / unescape / parsing simples)
' ============================================================

Private Function JsonEscapar(ByVal s As String) As String
    JsonEscapar = Json_EscapeString(CStr(s))
End Function



Private Function M05_ValidateUtf8Roundtrip(ByVal textIn As String, ByRef outDiag As String) As Boolean
    On Error GoTo Falha
    outDiag = ""

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText CStr(textIn)
    stm.Position = 0
    Dim roundtrip As String
    roundtrip = stm.ReadText(-1)
    stm.Close

    If Len(roundtrip) <> Len(textIn) Then
        outDiag = "len_original=" & CStr(Len(textIn)) & " len_roundtrip=" & CStr(Len(roundtrip))
        M05_ValidateUtf8Roundtrip = False
        Exit Function
    End If

    If StrComp(roundtrip, textIn, vbBinaryCompare) <> 0 Then
        outDiag = "binary_diff_detected"
        M05_ValidateUtf8Roundtrip = False
        Exit Function
    End If

    M05_ValidateUtf8Roundtrip = True
    Exit Function
Falha:
    outDiag = "exception=" & CStr(Err.Number) & " desc=" & Err.Description
    M05_ValidateUtf8Roundtrip = False
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

Private Function M05_CountOccurrences(ByVal haystack As String, ByVal needle As String) As Long
    Dim pos As Long
    Dim nextPos As Long

    If Len(haystack) = 0 Or Len(needle) = 0 Then
        M05_CountOccurrences = 0
        Exit Function
    End If

    pos = 1
    Do
        nextPos = InStr(pos, haystack, needle, vbTextCompare)
        If nextPos = 0 Then Exit Do
        M05_CountOccurrences = M05_CountOccurrences + 1
        pos = nextPos + Len(needle)
    Loop
End Function

Private Function M05_LengthBand(ByVal valueLen As Long) As String
    If valueLen <= 20000 Then
        M05_LengthBand = "baixo"
    ElseIf valueLen <= 120000 Then
        M05_LengthBand = "medio"
    ElseIf valueLen <= 300000 Then
        M05_LengthBand = "alto"
    Else
        M05_LengthBand = "muito_alto"
    End If
End Function

Private Function M05_BuildContextLengthDetail( _
    ByVal modelo As String, _
    ByVal payloadLen As Long, _
    ByVal promptLen As Long, _
    ByVal inputLitLen As Long, _
    ByVal maxOutputTokens As Long, _
    ByVal hasInputFile As Boolean, _
    ByVal hasInputImage As Boolean, _
    ByVal hasFileData As Boolean, _
    ByVal hasFileId As Boolean, _
    ByVal inputLit As String _
) As String
    Dim inputTextCount As Long
    Dim inputFileCount As Long
    Dim inputImageCount As Long
    Dim contextPart As String

    inputTextCount = M05_CountOccurrences(inputLit, """type"":""input_text""")
    inputFileCount = M05_CountOccurrences(inputLit, """type"":""input_file""")
    inputImageCount = M05_CountOccurrences(inputLit, """type"":""input_image""")

    contextPart = "model=" & modelo & _
        " | payload_len=" & CStr(payloadLen) & "(" & M05_LengthBand(payloadLen) & ")" & _
        " | prompt_len=" & CStr(promptLen) & "(" & M05_LengthBand(promptLen) & ")" & _
        " | input_array_len=" & CStr(inputLitLen) & "(" & M05_LengthBand(inputLitLen) & ")" & _
        " | max_output_tokens=" & CStr(maxOutputTokens) & _
        " | input_text_items=" & CStr(inputTextCount) & _
        " | input_file_items=" & CStr(inputFileCount) & _
        " | input_image_items=" & CStr(inputImageCount) & _
        " | has_input_file=" & IIf(hasInputFile, "SIM", "NAO") & _
        " | has_input_image=" & IIf(hasInputImage, "SIM", "NAO") & _
        " | has_file_data=" & IIf(hasFileData, "SIM", "NAO") & _
        " | has_file_id=" & IIf(hasFileId, "SIM", "NAO")

    M05_BuildContextLengthDetail = contextPart
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

Private Function M05_EscapeHintForControlChar(ByVal charCode As Long) As String
    Select Case charCode
        Case 10: M05_EscapeHintForControlChar = "\n"
        Case 13: M05_EscapeHintForControlChar = "\r"
        Case 9: M05_EscapeHintForControlChar = "\t"
        Case Else: M05_EscapeHintForControlChar = "\u" & Right$("000" & Hex$(charCode), 4)
    End Select
End Function

Private Function M05_IsHexChar(ByVal ch As String) As Boolean
    M05_IsHexChar = (InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) > 0)
End Function

Private Function M05_JsonEscapeIsValid(ByVal jsonText As String, ByVal slashPos As Long, ByRef outDetail As String) As Boolean
    Dim nxt As String
    Dim j As Long

    outDetail = ""

    If slashPos >= Len(jsonText) Then
        outDetail = "escape=barra_final @pos=" & CStr(slashPos)
        M05_JsonEscapeIsValid = False
        Exit Function
    End If

    nxt = Mid$(jsonText, slashPos + 1, 1)

    Select Case nxt
        Case """", "\", "/", "b", "f", "n", "r", "t"
            M05_JsonEscapeIsValid = True
            Exit Function
        Case "u"
            If slashPos + 5 > Len(jsonText) Then
                outDetail = "escape=unicode_incompleto @pos=" & CStr(slashPos)
                M05_JsonEscapeIsValid = False
                Exit Function
            End If

            For j = slashPos + 2 To slashPos + 5
                If Not M05_IsHexChar(Mid$(jsonText, j, 1)) Then
                    outDetail = "escape=unicode_invalido @pos=" & CStr(slashPos)
                    M05_JsonEscapeIsValid = False
                    Exit Function
                End If
            Next j

            M05_JsonEscapeIsValid = True
            Exit Function
        Case Else
            outDetail = "escape_invalido=" & Chr$(34) & nxt & Chr$(34) & " @pos=" & CStr(slashPos)
            M05_JsonEscapeIsValid = False
            Exit Function
    End Select
End Function

Private Function M05_JsonHasRawControlInString(ByVal jsonText As String, ByRef outDetail As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim escapeHint As String

    inString = False
    escaped = False
    outDetail = ""

    For i = 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)
        code = AscW(ch)

        If inString Then
            If escaped Then
                Dim escapeDetail As String
                If Not M05_JsonEscapeIsValid(jsonText, i - 1, escapeDetail) Then
                    outDetail = escapeDetail & " | escapes_validos=" & Chr$(92) & Chr$(34) & " " & Chr$(92) & Chr$(92) & " / b f n r t uXXXX"
                    M05_JsonHasRawControlInString = True
                    Exit Function
                End If
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            ElseIf code >= 0 And code <= 31 Then
                escapeHint = M05_EscapeHintForControlChar(code)
                outDetail = "char_code=" & CStr(code) & " @pos=" & CStr(i) & " | escape_sugerido=" & escapeHint
                M05_JsonHasRawControlInString = True
                Exit Function
            End If
        Else
            If ch = """" Then inString = True
        End If
    Next i

    If escaped Then
        outDetail = "escape=barra_final @pos=" & CStr(Len(jsonText)) & " | escapes_validos=" & Chr$(92) & Chr$(34) & " " & Chr$(92) & Chr$(92) & " / b f n r t uXXXX"
        M05_JsonHasRawControlInString = True
        Exit Function
    End If

    M05_JsonHasRawControlInString = False
End Function

Private Function M05_JsonStructuralPreflight(ByVal jsonText As String, ByRef outDetail As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim token As String
    Dim stack As String

    inString = False
    escaped = False
    token = ""
    stack = ""

    For i = 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)

        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
                token = "VALUE"
            End If
        Else
            Select Case ch
                Case """"
                    inString = True
                    token = "VALUE"

                Case "{"
                    stack = stack & "}"
                    token = "OPEN"

                Case "["
                    stack = stack & "]"
                    token = "OPEN"

                Case "}", "]"
                    If Len(stack) = 0 Then
                        outDetail = "fecho_sem_abertura @pos=" & CStr(i) & " char=" & ch
                        M05_JsonStructuralPreflight = False
                        Exit Function
                    End If

                    If Right$(stack, 1) <> ch Then
                        outDetail = "fecho_incompativel @pos=" & CStr(i) & " esperado=" & Right$(stack, 1) & " recebido=" & ch
                        M05_JsonStructuralPreflight = False
                        Exit Function
                    End If

                    If token = "COMMA" Then
                        outDetail = "virgula_final_invalida @pos=" & CStr(i) & " sequencia=," & ch
                        M05_JsonStructuralPreflight = False
                        Exit Function
                    End If

                    stack = Left$(stack, Len(stack) - 1)
                    token = "CLOSE"

                Case ","
                    token = "COMMA"

                Case " ", vbTab, vbCr, vbLf, ":"
                    ' Ignorar whitespace e separador de objeto.

                Case Else
                    token = "VALUE"
            End Select
        End If
    Next i

    If inString Then
        outDetail = "string_nao_fechada"
        M05_JsonStructuralPreflight = False
        Exit Function
    End If

    If Len(stack) > 0 Then
        outDetail = "estrutura_nao_fechada esperado=" & Right$(stack, 1)
        M05_JsonStructuralPreflight = False
        Exit Function
    End If

    M05_JsonStructuralPreflight = True
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



Private Function ExtraFragment_RequestsCodeInterpreter(ByVal extraFragmentSemInput As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(CStr(extraFragmentSemInput)))

    If t = "" Then
        ExtraFragment_RequestsCodeInterpreter = False
        Exit Function
    End If

    If InStr(1, t, """process_mode"":""code_interpreter""", vbTextCompare) > 0 Then
        ExtraFragment_RequestsCodeInterpreter = True
        Exit Function
    End If

    If InStr(1, t, """tool_choice"":""code_interpreter""", vbTextCompare) > 0 Then
        ExtraFragment_RequestsCodeInterpreter = True
        Exit Function
    End If

    If InStr(1, t, """tool_choice"":{""type"":""code_interpreter""", vbTextCompare) > 0 Then
        ExtraFragment_RequestsCodeInterpreter = True
        Exit Function
    End If

    ExtraFragment_RequestsCodeInterpreter = False
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
    Optional ByVal promptIdForDebug As String = "", _
    Optional ByVal debugFingerprintSeed As String = "" _
) As ApiResultado

    Dim resultado As ApiResultado
    Dim dbgPromptId As String
    dbgPromptId = IIf(Trim$(promptIdForDebug) <> "", promptIdForDebug, "M05")

    Dim fpBase As String
    fpBase = M05_BuildFingerprint(debugFingerprintSeed, dbgPromptId, "[pendente]", modelo, "[pendente]", "[pendente]")

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
    '  - Regra: se Modos incluir Web search, auto-adicionar sempre (inclusive com anexos).
    '  - Excecao: quando o extra ja traz "tools", evita duplicar chave no JSON final.
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
        Else
            autoAddWebSearch = True
        End If
    End If

    Dim ciExplicitInExtra As Boolean
    ciExplicitInExtra = ExtraFragment_RequestsCodeInterpreter(extraFragmentSemInput)

    Dim autoAddCodeInterpreter As Boolean
    autoAddCodeInterpreter = False

    If modosCodeInterpreter Then
        If extraTemTools Then
            autoAddCodeInterpreter = False
        ElseIf (hasInputFile Or hasInputImage) And (Not ciExplicitInExtra) Then
            autoAddCodeInterpreter = False
            On Error Resume Next
            Call Debug_Registar(0, dbgPromptId, "ALERTA", "", "M05_CI_AUTO_SUPPRESS", _
                "Modos inclui Code Interpreter, mas auto-add foi suprimido: ha anexos input_file/input_image e o extra nao pede CI explicitamente.", _
                "Se este passo precisar mesmo de Code Interpreter, definir process_mode: code_interpreter no Config extra (ou tools explicitas).")
            On Error GoTo TrataErro
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

    Dim utf8Diag As String
    If Not M05_ValidateUtf8Roundtrip(json, utf8Diag) Then
        On Error Resume Next
        Call Debug_Registar(0, dbgPromptId, "ERRO", "", "M05_UTF8_ROUNDTRIP", _
            "Payload bloqueado: roundtrip UTF-8 falhou (" & utf8Diag & ")", _
            "Rever origem de texto/codificacao dos INPUTS e evitar colagens com codificacao inconsistente.")
        On Error GoTo TrataErro

        resultado.Erro = "Payload invalido (utf8_roundtrip): " & utf8Diag
        OpenAI_Executar = resultado
        Exit Function
    Else
        On Error Resume Next
        Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_UTF8_ROUNDTRIP", _
            "roundtrip UTF-8 OK | payload_len=" & CStr(Len(json)), _
            "Gate de codificacao concluido antes do envio HTTP.")
        On Error GoTo TrataErro
    End If

    Dim preflightDetail As String
    If M05_JsonHasRawControlInString(json, preflightDetail) Then
        On Error Resume Next
        Call Debug_Registar(0, dbgPromptId, "ERRO", "", "M05_JSON_PREFLIGHT", _
            "Payload bloqueado antes do envio: possivel controlo nao escapado dentro de string JSON (" & preflightDetail & ")", _
            "Revise input_json/extraFragment. Escapes validos em JSON incluem \n, \r, \t, \u00XX e, apos \\, apenas " & Chr$(92) & Chr$(34) & ", " & Chr$(92) & Chr$(92) & ", \/, \b, \f, \n, \r, \t e \uXXXX.")
        On Error GoTo TrataErro

        resultado.Erro = "Payload invalido (preflight): controlo nao escapado em string JSON. " & preflightDetail
        OpenAI_Executar = resultado
        Exit Function
    End If

    If Not M05_JsonStructuralPreflight(json, preflightDetail) Then
        On Error Resume Next
        Call Debug_Registar(0, dbgPromptId, "ERRO", "", "M05_JSON_PREFLIGHT", _
            "Payload bloqueado antes do envio: estrutura JSON invalida (" & preflightDetail & ")", _
            "Revise fusao de fragments (Config extra/File Output) e valide C:\Temp\payload.json num validador JSON.")
        On Error GoTo TrataErro

        resultado.Erro = "Payload invalido (preflight): estrutura JSON invalida. " & preflightDetail
        OpenAI_Executar = resultado
        Exit Function
    End If

    Call M05_DumpPayloadForDebug(json, dbgPromptId)

    ' -------------------------
    ' Log diagnostico (sem despejar base64)
    ' -------------------------
    Dim toolMsg As String
    If modosWebSearch Then
        If autoAddWebSearch Then
            toolMsg = "web_search=ADICIONADO_AUTO"
        ElseIf extraTemTools Then
            toolMsg = "web_search=NAO_AUTO (tools no extra)"
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
        "FP=" & fpBase & _
        " | endpoint=" & OPENAI_ENDPOINT & _
        " | model=" & modelo & _
        " | input_is_array=" & IIf(usarInputArray, "SIM", "NAO") & _
        " | has_input_file=" & IIf(hasInputFile, "SIM", "NAO") & _
        " | has_file_data=" & IIf(hasFileData, "SIM", "NAO") & _
        " | has_file_id=" & IIf(hasFileId, "SIM", "NAO") & _
        " | has_input_image=" & IIf(hasInputImage, "SIM", "NAO") & _
        " | " & toolMsg & _
        " | payload_len=" & CStr(Len(json)), _
        "Pedido preparado e enviado com sucesso técnico; validação de contrato de output será feita após resposta (correlacionar com M10_* do mesmo FP).")
    On Error GoTo TrataErro

    ' -------------------------
    ' HTTP POST com retry 5xx
    ' -------------------------
    Const MAX_RETRIES_5XX As Long = 3
    Const SEND_TIMEOUT_RETRY_WAIT_S As Long = 3
    Dim waitsSec(1 To 3) As Long
    waitsSec(1) = 2
    waitsSec(2) = 6
    waitsSec(3) = 14

    Dim attempt As Long
    Dim httpStatus As Long
    Dim resposta As String
    Dim reqId As String
    Dim timeoutResolveMs As Long
    Dim timeoutConnectMs As Long
    Dim timeoutSendMs As Long
    Dim timeoutReceiveMs As Long
    Dim attemptStartTick As Double
    Dim timeoutElapsedMs As Long
    Dim httpStage As String
    Dim httpLastDllError As Long
    Dim requestPayloadLen As Long
    Dim httpStatusSnapshot As Long
    Dim responseLenSnapshot As Long
    Dim attemptStartedAt As Date
    Dim timeoutFailedAt As Date
    Dim sendTimeoutRetryUsed As Boolean
    Dim sendTimeoutRetryOutcome As String

    timeoutResolveMs = M05_GetHttpTimeoutMs("HTTP_TIMEOUT_RESOLVE_MS", HTTP_TIMEOUT_RESOLVE_MS_DEFAULT, dbgPromptId)
    timeoutConnectMs = M05_GetHttpTimeoutMs("HTTP_TIMEOUT_CONNECT_MS", HTTP_TIMEOUT_CONNECT_MS_DEFAULT, dbgPromptId)
    timeoutSendMs = M05_GetHttpTimeoutMs("HTTP_TIMEOUT_SEND_MS", HTTP_TIMEOUT_SEND_MS_DEFAULT, dbgPromptId)
    timeoutReceiveMs = M05_GetHttpTimeoutMs("HTTP_TIMEOUT_RECEIVE_MS", HTTP_TIMEOUT_RECEIVE_MS_DEFAULT, dbgPromptId)

    On Error Resume Next
    Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_HTTP_TIMEOUTS", _
        "FP=" & fpBase & _
        " | resolve_ms=" & timeoutResolveMs & _
        " | connect_ms=" & timeoutConnectMs & _
        " | send_ms=" & timeoutSendMs & _
        " | receive_ms=" & timeoutReceiveMs, _
        "Time-outs aplicados nesta execucao; util para distinguir lentidao de erro logico. Se houver timeout, aumentar HTTP_TIMEOUT_RECEIVE_MS e repetir 1 vez.")
    On Error GoTo TrataErro

    requestPayloadLen = Len(json)
    attempt = 0
    sendTimeoutRetryUsed = False
    sendTimeoutRetryOutcome = "not_used"

    Do
        attempt = attempt + 1
        reqId = ""
        attemptStartTick = Timer
        attemptStartedAt = Now
        httpStage = "init"
        httpStatusSnapshot = 0
        responseLenSnapshot = 0

        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

        httpStage = "set_timeouts"
        http.SetTimeouts timeoutResolveMs, timeoutConnectMs, timeoutSendMs, timeoutReceiveMs

        httpStage = "open"
        http.Open "POST", OPENAI_ENDPOINT, False
        http.SetRequestHeader "Authorization", "Bearer " & apiKey
        http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        http.SetRequestHeader "Accept", "application/json"

        httpStage = "send"
        On Error Resume Next
        Err.Clear
        http.Send json
        Dim sendErrNumber As Long
        Dim sendErrDescription As String
        Dim sendErrSource As String
        sendErrNumber = Err.Number
        sendErrDescription = Err.Description
        sendErrSource = Err.Source
        On Error GoTo TrataErro

        If sendErrNumber <> 0 Then
            If M05_IsTimeoutError(sendErrNumber, sendErrDescription) And (Not sendTimeoutRetryUsed) Then
                sendTimeoutRetryUsed = True
                sendTimeoutRetryOutcome = "retry_scheduled"
                timeoutFailedAt = Now

                On Error Resume Next
                Call Debug_Registar(0, dbgPromptId, "ALERTA", "", "M05_TIMEOUT_DECISION", _
                    "FP=" & fpBase & _
                    " | decision=retry_send_once" & _
                    " | reason=timeout_stage_send" & _
                    " | started_at=" & M05_FormatIsoTimestamp(attemptStartedAt) & _
                    " | failed_at=" & M05_FormatIsoTimestamp(timeoutFailedAt) & _
                    " | backoff_s=" & CStr(SEND_TIMEOUT_RETRY_WAIT_S) & _
                    " | retry_outcome=" & sendTimeoutRetryOutcome, _
                    "Timeout em Send na primeira tentativa; o motor agenda 1 retry automatico com novo socket antes de falhar.")
                On Error GoTo TrataErro

                Call M05_SleepSeconds(SEND_TIMEOUT_RETRY_WAIT_S)
                GoTo ProximaTentativa
            Else
                If M05_IsTimeoutError(sendErrNumber, sendErrDescription) Then
                    sendTimeoutRetryOutcome = "retry_failed"
                End If
                Err.Raise sendErrNumber, sendErrSource, sendErrDescription
            End If
        End If

        httpStage = "status"
        httpStatus = CLng(http.Status)
        httpStatusSnapshot = httpStatus
        httpStage = "response_text"
        resposta = CStr(http.ResponseText)
        responseLenSnapshot = Len(resposta)

        resultado.rawResponseJson = resposta
        resultado.httpStatus = httpStatus

        On Error Resume Next
        Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_HTTP_RESULT", _
            "FP=" & M05_BuildFingerprint(debugFingerprintSeed, dbgPromptId, ExtrairCampoJsonSimples(resposta, """id"":"), modelo, IIf(httpStatus >= 200 And httpStatus < 300, "SIM", "NAO"), "[pendente]") & _
            " | http_status=" & CStr(httpStatus), _
            IIf(httpStatus >= 200 And httpStatus < 300, _
                "Transporte HTTP concluido com sucesso (2xx). O contrato funcional de output deve ser confirmado pelos eventos M10.", _
                "Transporte HTTP falhou; validar conectividade, payload e timeouts antes de analisar contrato de output."))
        On Error GoTo TrataErro

        If httpStatus >= 200 And httpStatus < 300 Then
            If sendTimeoutRetryUsed Then
                sendTimeoutRetryOutcome = "retry_success"
                On Error Resume Next
                Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_TIMEOUT_DECISION", _
                    "FP=" & fpBase & _
                    " | decision=retry_send_once" & _
                    " | retry_outcome=" & sendTimeoutRetryOutcome, _
                    "Retry automatico de timeout em Send teve sucesso nesta execucao.")
                On Error GoTo TrataErro
            End If

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
                    "Sugestao: erro transitorio do servidor. O PIPELINER vai repetir automaticamente.")
                On Error GoTo TrataErro

                Call M05_SleepSeconds(waitS)
            Else
                resultado.Erro = "HTTP " & httpStatus & " - " & resposta
                On Error Resume Next
                Call Debug_Registar(0, dbgPromptId, "ERRO", "", "API", _
                    "HTTP " & httpStatus & " (server_error) | retries_esgotados=" & MAX_RETRIES_5XX & _
                    IIf(reqId <> "", " | req_id=" & reqId, ""), _
                    "Sugestao: tente mais tarde. Se persistir, contacte suporte com o req_id.")
                On Error GoTo TrataErro

                OpenAI_Executar = resultado
                Exit Function
            End If
        Else
            resultado.Erro = "HTTP " & httpStatus & " - " & resposta

            If httpStatus = 400 Then
                Dim apiErrCode As String
                Dim apiErrType As String
                Dim apiErrParam As String
                Dim apiErrMessage As String
                Dim contextLengthDetail As String

                apiErrCode = LCase$(Trim$(ExtrairCampoJsonSimples(resposta, """code"":")))
                apiErrType = Trim$(ExtrairCampoJsonSimples(resposta, """type"":"))
                apiErrParam = Trim$(ExtrairCampoJsonSimples(resposta, """param"":"))
                apiErrMessage = Trim$(ExtrairCampoJsonSimples(resposta, """message"":"))

                If apiErrCode = "context_length_exceeded" Then
                    contextLengthDetail = M05_BuildContextLengthDetail(modelo, requestPayloadLen, Len(textoPrompt), Len(inputLit), maxOutputTokens, hasInputFile, hasInputImage, hasFileData, hasFileId, inputLit)
                    On Error Resume Next
                    Call Debug_Registar(0, dbgPromptId, "ERRO", "", "API_CONTEXT_LENGTH_EXCEEDED", _
                        "FP=" & fpBase & " | http_status=400 | error_code=" & apiErrCode & " | error_type=" & apiErrType & " | param=" & apiErrParam & " | api_message=" & apiErrMessage & " | " & contextLengthDetail, _
                        "A entrada excedeu a janela de contexto do modelo; reduzir texto de INPUTS/OUTPUTS, anexos em text_embed ou max_output_tokens antes de repetir.")
                    Call Debug_Registar(0, dbgPromptId, "INFO", "", "API_CONTEXT_LENGTH_ACTION", _
                        Diag_Format("M05_CONTEXT", "context window excedida", "Pedido recusado com HTTP 400", "Resumir input e limitar tamanho de anexos text_embed", "Se houver anexos, prefira PDF selecionado ou divida em passos menores."), _
                        "Checklist: validar REQ_INPUT_JSON len, payload_len em M05_PAYLOAD_CHECK e FILES_TEXT_EMBED_MAX_CHARS.")
                    On Error GoTo TrataErro
                End If
            End If

            OpenAI_Executar = resultado
            Exit Function
        End If

ProximaTentativa:
    Loop

TrataErro:
    Dim errNumber As Long
    Dim errDescription As String
    Dim timeoutType As String
    Dim timeoutDiag As String
    Dim timeoutCauseHint As String
    Dim errSource As String
    Dim timeoutMetric As String
    Dim timeoutDecision As String
    Dim hostNetworkContext As String

    errNumber = Err.Number
    errDescription = Trim$(Err.Description)
    errSource = Trim$(Err.Source)
    httpLastDllError = Err.LastDllError
    timeoutFailedAt = Now

    timeoutElapsedMs = M05_ElapsedMsFromTick(attemptStartTick)
    If M05_IsTimeoutError(errNumber, errDescription) Then
        timeoutType = M05_ClassifyTimeoutType(httpStage, timeoutElapsedMs, timeoutResolveMs, timeoutConnectMs, timeoutSendMs, timeoutReceiveMs)
        timeoutCauseHint = M05_BuildTimeoutCauseHint(httpStage, timeoutElapsedMs, timeoutSendMs, timeoutReceiveMs, requestPayloadLen, responseLenSnapshot, httpStatusSnapshot)
        timeoutMetric = M05_RegisterTimeoutMetric(dbgPromptId, modelo)
        hostNetworkContext = M05_GetHostNetworkContext()
        timeoutDecision = M05_BuildTimeoutDecision(httpStage, timeoutSendMs, timeoutReceiveMs, sendTimeoutRetryUsed, sendTimeoutRetryOutcome)

        timeoutDiag = timeoutType & _
            " | stage=" & M05_TimeoutStageLabel(httpStage) & _
            " | started_at=" & M05_FormatIsoTimestamp(attemptStartedAt) & _
            " | failed_at=" & M05_FormatIsoTimestamp(timeoutFailedAt) & _
            " | elapsed_ms=" & CStr(timeoutElapsedMs) & _
            " | attempt=" & CStr(attempt) & _
            " | payload_len=" & CStr(requestPayloadLen) & _
            " | http_status_partial=" & CStr(httpStatusSnapshot) & _
            " | response_len=" & CStr(responseLenSnapshot) & _
            " | HTTP_TIMEOUT_RESOLVE_MS=" & CStr(timeoutResolveMs) & _
            " | HTTP_TIMEOUT_CONNECT_MS=" & CStr(timeoutConnectMs) & _
            " | HTTP_TIMEOUT_SEND_MS=" & CStr(timeoutSendMs) & _
            " | HTTP_TIMEOUT_RECEIVE_MS=" & CStr(timeoutReceiveMs)
        If timeoutCauseHint <> "" Then timeoutDiag = timeoutDiag & " | " & timeoutCauseHint
        If timeoutMetric <> "" Then timeoutDiag = timeoutDiag & " | " & timeoutMetric
        If hostNetworkContext <> "" Then timeoutDiag = timeoutDiag & " | " & hostNetworkContext
        If timeoutDecision <> "" Then timeoutDiag = timeoutDiag & " | " & timeoutDecision

        On Error Resume Next
        Call Debug_Registar(0, dbgPromptId, "ERRO", "", "M05_HTTP_TIMEOUT_ERROR", _
            "FP=" & fpBase & " | " & timeoutDiag, _
            "Timeout detetado na chamada a /v1/responses; rever limite indicado e latencia/rede antes de repetir.")
        Call Debug_Registar(0, dbgPromptId, "INFO", "", "M05_TIMEOUT_DECISION", _
            "FP=" & fpBase & " | " & timeoutDecision, _
            "Decisao sugerida automaticamente com base na fase do timeout e nos limites atuais.")
        On Error GoTo 0
    End If

    If errDescription = "" Then
        errDescription = "(sem descricao VBA)"
    End If

    resultado.Erro = "Erro VBA: " & errDescription
    If timeoutDiag <> "" Then
        resultado.Erro = resultado.Erro & " | " & timeoutDiag
    End If

    If errNumber <> 0 Then
        resultado.Erro = resultado.Erro & " | Err.Number=" & CStr(errNumber)
    End If
    If httpLastDllError <> 0 Then
        resultado.Erro = resultado.Erro & " | LastDllError=" & CStr(httpLastDllError)
    End If
    If errSource <> "" Then
        resultado.Erro = resultado.Erro & " | Err.Source=" & errSource
    End If

    OpenAI_Executar = resultado
End Function

Private Function M05_ElapsedMsFromTick(ByVal startTick As Double) As Long
    On Error GoTo EH

    If startTick <= 0 Then
        M05_ElapsedMsFromTick = 0
        Exit Function
    End If

    Dim nowTick As Double
    Dim elapsedSec As Double
    nowTick = Timer

    If nowTick >= startTick Then
        elapsedSec = nowTick - startTick
    Else
        elapsedSec = (86400# - startTick) + nowTick
    End If

    M05_ElapsedMsFromTick = CLng(elapsedSec * 1000#)
    Exit Function

EH:
    M05_ElapsedMsFromTick = 0
End Function

Private Function M05_IsTimeoutError(ByVal errNumber As Long, ByVal errDescription As String) As Boolean
    Dim d As String
    d = LCase$(Trim$(CStr(errDescription)))

    If InStr(1, d, "timed out", vbTextCompare) > 0 Then
        M05_IsTimeoutError = True
        Exit Function
    End If

    If InStr(1, d, "tempo limite", vbTextCompare) > 0 Then
        M05_IsTimeoutError = True
        Exit Function
    End If

    ' WinHTTP ERROR_WINHTTP_TIMEOUT
    If errNumber = -2147012894 Then
        M05_IsTimeoutError = True
        Exit Function
    End If

    M05_IsTimeoutError = False
End Function

Private Function M05_ClassifyTimeoutType( _
    ByVal httpStage As String, _
    ByVal elapsedMs As Long, _
    ByVal resolveMs As Long, _
    ByVal connectMs As Long, _
    ByVal sendMs As Long, _
    ByVal receiveMs As Long _
) As String
    Const MARGIN_MS As Long = 1500

    Dim tResolve As Long
    Dim tConnect As Long
    Dim tSend As Long
    Dim tReceive As Long

    tResolve = resolveMs
    tConnect = resolveMs + connectMs
    tSend = resolveMs + connectMs + sendMs
    tReceive = resolveMs + connectMs + sendMs + receiveMs

    Select Case LCase$(Trim$(httpStage))
        Case "open", "set_timeouts"
            M05_ClassifyTimeoutType = "Timeout de ligacao TCP/TLS (ms) para /v1/responses"
            Exit Function
        Case "send"
            M05_ClassifyTimeoutType = "Timeout de envio do request (ms) para /v1/responses"
            Exit Function
        Case "status", "response_text"
            M05_ClassifyTimeoutType = "Timeout de espera da resposta (ms) para /v1/responses"
            Exit Function
    End Select

    If elapsedMs > 0 And elapsedMs <= (tResolve + MARGIN_MS) Then
        M05_ClassifyTimeoutType = "Timeout DNS/resolve (ms) para /v1/responses"
    ElseIf elapsedMs > 0 And elapsedMs <= (tConnect + MARGIN_MS) Then
        M05_ClassifyTimeoutType = "Timeout de ligacao TCP/TLS (ms) para /v1/responses"
    ElseIf elapsedMs > 0 And elapsedMs <= (tSend + MARGIN_MS) Then
        M05_ClassifyTimeoutType = "Timeout de envio do request (ms) para /v1/responses"
    ElseIf elapsedMs > 0 And elapsedMs <= (tReceive + MARGIN_MS) Then
        M05_ClassifyTimeoutType = "Timeout de espera da resposta (ms) para /v1/responses"
    Else
        M05_ClassifyTimeoutType = "Outro tipo de Timeout"
    End If
End Function

Private Function M05_BuildTimeoutCauseHint( _
    ByVal httpStage As String, _
    ByVal elapsedMs As Long, _
    ByVal sendMs As Long, _
    ByVal receiveMs As Long, _
    ByVal payloadLen As Long, _
    ByVal responseLen As Long, _
    ByVal httpStatusPartial As Long _
) As String
    Dim stageNorm As String
    Dim causeHint As String
    Dim confidence As String
    Dim actionHint As String

    stageNorm = LCase$(Trim$(httpStage))
    causeHint = "indeterminada"
    confidence = "baixa"
    actionHint = "correlacionar com M05_PAYLOAD_CHECK e repetir com payload reduzido"

    Select Case stageNorm
        Case "open", "set_timeouts"
            causeHint = "rede_dns_proxy_tls"
            confidence = "media"
            actionHint = "validar DNS/proxy/TLS/firewall e conectividade para api.openai.com"

        Case "send"
            If payloadLen > 250000 Then
                causeHint = "upload_lento_payload_grande"
                confidence = "alta"
                actionHint = "reduzir payload/anexos ou aumentar HTTP_TIMEOUT_SEND_MS"
            Else
                causeHint = "uplink_lento_ou_interrupcao_envio"
                confidence = "media"
                actionHint = "validar rede e considerar aumentar HTTP_TIMEOUT_SEND_MS"
            End If

        Case "status"
            causeHint = "servidor_lento_ate_headers"
            confidence = "media"
            actionHint = "aumentar HTTP_TIMEOUT_RECEIVE_MS e verificar latencia do endpoint"

        Case "response_text"
            If httpStatusPartial >= 200 And httpStatusPartial < 300 And responseLen = 0 Then
                causeHint = "stream_resposta_interrompido"
                confidence = "media"
                actionHint = "repetir 1x e validar estabilidade de rede/proxy"
            ElseIf responseLen > 200000 Then
                causeHint = "resposta_grande_lenta"
                confidence = "alta"
                actionHint = "reduzir output esperado ou aumentar HTTP_TIMEOUT_RECEIVE_MS"
            Else
                causeHint = "espera_resposta_excedida"
                confidence = "media"
                actionHint = "aumentar HTTP_TIMEOUT_RECEIVE_MS"
            End If

        Case Else
            If elapsedMs > 0 And elapsedMs <= sendMs Then
                causeHint = "tempo_gasto_antes_rececao"
                confidence = "baixa"
            ElseIf elapsedMs > sendMs And elapsedMs <= (sendMs + receiveMs) Then
                causeHint = "tempo_gasto_em_espera_resposta"
                confidence = "baixa"
            End If
    End Select

    M05_BuildTimeoutCauseHint = "cause_hint=" & causeHint & " | confidence=" & confidence & " | action=" & actionHint
End Function

Private Function M05_BuildTimeoutDecision( _
    ByVal httpStage As String, _
    ByVal sendMs As Long, _
    ByVal receiveMs As Long, _
    ByVal retryUsed As Boolean, _
    ByVal retryOutcome As String _
) As String
    Dim decision As String
    Dim st As String

    st = LCase$(Trim$(httpStage))
    decision = "decision=investigar"

    Select Case st
        Case "send"
            decision = "decision=aumentar_send_para_" & CStr(sendMs + 180000) & "ms"
        Case "status", "response_text"
            decision = "decision=aumentar_receive_para_" & CStr(receiveMs + 180000) & "ms"
        Case "open", "set_timeouts"
            decision = "decision=validar_proxy_dns_tls"
    End Select

    M05_BuildTimeoutDecision = decision & " | retry_outcome=" & IIf(Trim$(retryOutcome) = "", "n/d", retryOutcome) & " | retry_used=" & IIf(retryUsed, "SIM", "NAO")
End Function

Private Function M05_RegisterTimeoutMetric(ByVal promptId As String, ByVal modelName As String) As String
    On Error GoTo EH

    Static timeoutByPromptModel As Object
    Static timeoutGlobal As Long

    If timeoutByPromptModel Is Nothing Then
        Set timeoutByPromptModel = CreateObject("Scripting.Dictionary")
    End If

    Dim k As String
    Dim c As Long

    k = Trim$(promptId) & "|" & Trim$(modelName)
    If timeoutByPromptModel.Exists(k) Then
        c = CLng(timeoutByPromptModel(k))
    Else
        c = 0
    End If

    c = c + 1
    timeoutByPromptModel(k) = c
    timeoutGlobal = timeoutGlobal + 1

    M05_RegisterTimeoutMetric = "timeout_count_prompt_model=" & CStr(c) & " | timeout_count_global=" & CStr(timeoutGlobal)
    Exit Function

EH:
    M05_RegisterTimeoutMetric = ""
End Function

Private Function M05_GetHostNetworkContext() As String
    On Error GoTo EH

    Dim hostName As String
    Dim ipMasked As String
    Dim proxySummary As String
    Dim vpnFlag As String

    hostName = Trim$(Environ$("COMPUTERNAME"))
    If hostName = "" Then hostName = "[n/d]"

    ipMasked = M05_GetLocalIPv4Masked()
    If ipMasked = "" Then ipMasked = "[n/d]"

    proxySummary = M05_GetWinHttpProxySummary()
    If proxySummary = "" Then proxySummary = "[n/d]"

    vpnFlag = M05_GetVpnFlagHeuristic()
    If vpnFlag = "" Then vpnFlag = "UNKNOWN"

    M05_GetHostNetworkContext = "host=" & hostName & " | ip_masked=" & ipMasked & " | winhttp_proxy=" & proxySummary & " | vpn_flag=" & vpnFlag
    Exit Function

EH:
    M05_GetHostNetworkContext = ""
End Function

Private Function M05_GetVpnFlagHeuristic() As String
    On Error GoTo EH

    Dim s As String
    s = LCase$(M05_RunCommandCapture("cmd /c ipconfig"))

    If InStr(1, s, "ppp adapter", vbTextCompare) > 0 Or _
       InStr(1, s, "vpn", vbTextCompare) > 0 Or _
       InStr(1, s, "tap-", vbTextCompare) > 0 Then
        M05_GetVpnFlagHeuristic = "LIKELY_ON"
    ElseIf Trim$(s) <> "" Then
        M05_GetVpnFlagHeuristic = "LIKELY_OFF"
    Else
        M05_GetVpnFlagHeuristic = "UNKNOWN"
    End If
    Exit Function

EH:
    M05_GetVpnFlagHeuristic = "UNKNOWN"
End Function

Private Function M05_GetWinHttpProxySummary() As String
    On Error GoTo EH

    Dim s As String
    Dim line As Variant
    s = M05_RunCommandCapture("cmd /c netsh winhttp show proxy")
    If Trim$(s) = "" Then GoTo EH

    If InStr(1, s, "Direct access (no proxy server)", vbTextCompare) > 0 Then
        M05_GetWinHttpProxySummary = "DIRECT"
        Exit Function
    End If

    For Each line In Split(s, vbCrLf)
        If InStr(1, CStr(line), "Proxy Server", vbTextCompare) > 0 Then
            M05_GetWinHttpProxySummary = Trim$(Replace(CStr(line), "Proxy Server(s) :", ""))
            If M05_GetWinHttpProxySummary = "" Then M05_GetWinHttpProxySummary = "SET"
            Exit Function
        End If
    Next line

    M05_GetWinHttpProxySummary = "SET"
    Exit Function

EH:
    M05_GetWinHttpProxySummary = ""
End Function

Private Function M05_GetLocalIPv4Masked() As String
    On Error GoTo EH

    Dim s As String
    Dim line As Variant
    Dim t As String
    Dim ip As String

    s = M05_RunCommandCapture("cmd /c ipconfig")
    For Each line In Split(s, vbCrLf)
        t = Trim$(CStr(line))
        If InStr(1, t, "IPv4", vbTextCompare) > 0 Then
            ip = Trim$(Split(t, ":")(1))
            M05_GetLocalIPv4Masked = M05_MaskIPv4(ip)
            Exit Function
        End If
    Next line

EH:
    M05_GetLocalIPv4Masked = ""
End Function

Private Function M05_MaskIPv4(ByVal ip As String) As String
    On Error GoTo EH

    Dim p() As String
    p = Split(Trim$(ip), ".")
    If UBound(p) <> 3 Then GoTo EH

    M05_MaskIPv4 = p(0) & "." & p(1) & "." & p(2) & ".x"
    Exit Function

EH:
    M05_MaskIPv4 = ""
End Function

Private Function M05_RunCommandCapture(ByVal commandLine As String) As String
    On Error GoTo EH

    Dim sh As Object
    Dim ex As Object

    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec(commandLine)
    M05_RunCommandCapture = ex.StdOut.ReadAll
    Exit Function

EH:
    M05_RunCommandCapture = ""
End Function

Private Function M05_FormatIsoTimestamp(ByVal dt As Date) As String
    On Error GoTo EH

    If dt <= 0 Then
        M05_FormatIsoTimestamp = "[n/d]"
    Else
        M05_FormatIsoTimestamp = Format$(dt, "yyyy-mm-dd") & "T" & Format$(dt, "hh:nn:ss")
    End If
    Exit Function

EH:
    M05_FormatIsoTimestamp = "[n/d]"
End Function

Private Function M05_TimeoutStageLabel(ByVal httpStage As String) As String
    Select Case LCase$(Trim$(httpStage))
        Case "set_timeouts": M05_TimeoutStageLabel = "SetTimeouts"
        Case "open": M05_TimeoutStageLabel = "Open"
        Case "send": M05_TimeoutStageLabel = "Send"
        Case "status": M05_TimeoutStageLabel = "Status"
        Case "response_text": M05_TimeoutStageLabel = "ResponseText"
        Case Else: M05_TimeoutStageLabel = IIf(Trim$(httpStage) = "", "[n/d]", Trim$(httpStage))
    End Select
End Function


Private Function M05_BuildFingerprint( _
    ByVal seed As String, _
    ByVal promptId As String, _
    ByVal responseId As String, _
    ByVal modelName As String, _
    ByVal okHttp As String, _
    ByVal modeTxt As String _
) As String
    Dim s As String
    s = Trim$(seed)
    If s = "" Then
        s = "pipeline=[n/d]|step=[n/d]|prompt=" & Trim$(promptId)
    End If

    If Trim$(responseId) = "" Then responseId = "[pendente]"
    If Trim$(modelName) = "" Then modelName = "[n/d]"
    If Trim$(okHttp) = "" Then okHttp = "[pendente]"
    If Trim$(modeTxt) = "" Then modeTxt = "[pendente]"

    M05_BuildFingerprint = s & "|resp=" & Trim$(responseId) & "|model=" & Trim$(modelName) & "|ok_http=" & Trim$(okHttp) & "|mode=" & Trim$(modeTxt)
End Function

Private Function M05_GetHttpTimeoutMs(ByVal keyName As String, ByVal defaultValue As Long, ByVal dbgPromptId As String) As Long
    On Error GoTo EH

    Dim raw As String
    raw = Trim$(M05_Config_GetByKey(keyName, ""))

    If raw = "" Then
        M05_GetHttpTimeoutMs = defaultValue
        Exit Function
    End If

    If Not IsNumeric(raw) Then GoTo InvalidValue

    Dim n As Double
    n = CDbl(raw)

    If n < HTTP_TIMEOUT_MIN_MS Or n > HTTP_TIMEOUT_MAX_MS Then GoTo InvalidValue

    M05_GetHttpTimeoutMs = CLng(n)
    Exit Function

InvalidValue:
    On Error Resume Next
    Call Debug_Registar(0, dbgPromptId, "ALERTA", "", "M05_HTTP_TIMEOUT_INVALID", _
        "Valor inválido em Config para " & keyName & "=[" & raw & "]; a usar default=" & defaultValue & "ms.", _
        "Use inteiro entre " & HTTP_TIMEOUT_MIN_MS & " e " & HTTP_TIMEOUT_MAX_MS & " (ms).")
    On Error GoTo EH

    M05_GetHttpTimeoutMs = defaultValue
    Exit Function

EH:
    M05_GetHttpTimeoutMs = defaultValue
End Function

Private Function M05_Config_GetByKey(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo EH

    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim k As String

    Set ws = ThisWorkbook.Worksheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 1 To lastRow
        k = UCase$(Trim$(CStr(ws.Cells(i, 1).Value)))
        If k = UCase$(Trim$(keyName)) Then
            M05_Config_GetByKey = Trim$(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i

    M05_Config_GetByKey = defaultValue
    Exit Function

EH:
    M05_Config_GetByKey = defaultValue
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



Private Sub M05_DumpPayloadForDebug(ByVal payloadJson As String, ByVal dbgPromptId As String)
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
        "Payload final gravado em " & targetPath & " | len=" & CStr(Len(payloadJson)), _
        "Use este ficheiro para validar text.format.schema e outros fragmentos antes de novo envio.")
    Exit Sub

Falha:
    On Error Resume Next
    If ff > 0 Then Close #ff
    Call Debug_Registar(0, dbgPromptId, "ALERTA", "", "M05_PAYLOAD_DUMP_FAIL", _
        "Não foi possível gravar payload em C:\Temp\payload.json: " & Err.Description, _
        "Verifique permissões locais e existência da pasta C:\Temp.")
End Sub
