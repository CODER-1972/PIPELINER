Attribute VB_Name = "M12_SelfTests"
Option Explicit

' =============================================================================
' Módulo: M12_SelfTests
' Propósito:
' - Executar self-tests idempotentes de componentes críticos do motor VBA.
' - Registar resultados PASS/FAIL/ALERTA no DEBUG para diagnóstico rápido.
'
' Atualizações:
' - 2026-02-16 | Codex | Novos self-tests para resolução segura de OPENAI_API_KEY
'   - Substitui teste de presença simples por cenários de precedência ENV -> Config!B1.
'   - Valida alertas/erros de migração sem exposição de segredo em logs.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - SelfTest_RunAll (Sub): rotina pública do módulo.
' =============================================================================

' =============================================================================
' M12_SelfTests
' - Testes internos idempotentes (não alteram Config permanentemente)
' - Escrevem resultados na folha DEBUG via Debug_Registar (M02_Logger_DEBUG_e_Seguimento)
' - Não dependem de funções inexistentes no M09 (evita "Sub or Function not defined")
'
' Como usar:
'   - Corre SelfTest_RunAll a partir do Editor VBA (F5) ou via Macro.
'   - Resultados ficam na folha DEBUG com Prompt ID = "SELFTEST".
' =============================================================================

Private Const SELFTEST_PROMPT_ID As String = "SELFTEST"

' Severidades (alinhadas com o DEBUG do PIPELINER)
Private Const SEV_INFO As String = "INFO"
Private Const SEV_ALERTA As String = "ALERTA"
Private Const SEV_ERRO As String = "ERRO"

' Prefixo para identificar linhas do SelfTest (para limpeza idempotente)
Private Const SELFTEST_PARAM_PREFIX As String = "SELFTEST_"

' =============================================================================
' Entry point
' =============================================================================

Public Sub SelfTest_RunAll()
    On Error GoTo EH

    ' Idempotência: remove apenas linhas antigas do SELFTEST (sem tocar noutros logs)
    SelfTest_ClearPreviousDebugRows

    SelfTest_Log SEV_INFO, "SELFTEST_RUN", "Início dos testes internos.", "OK"

    ' 1) Sanitização de filename (ASCII_SAFE) - teste local
    SelfTest_SanitizeFilename

    ' 2) Multipart em bytes - teste local (estrutura boundary/CRLF/fecho)
    SelfTest_MultipartBuild_Local

    ' 3) Disponibilidade de engines COM (WinHTTP / MSXML)
    SelfTest_EnginesAvailability

    ' 4) Resolução de OPENAI_API_KEY (precedência e diagnósticos)
    SelfTest_ConfigApiKeyResolution

    ' 5) Esquema mínimo da FILES_MANAGEMENT para output chain
    SelfTest_Schema_FilesManagement

    ' 6) Fluxo register/resolve output->input
    SelfTest_OutputRegister_And_Resolve

    SelfTest_Log SEV_INFO, "SELFTEST_RUN", "Fim dos testes internos.", "OK"
    Exit Sub

EH:
    SelfTest_Log SEV_ERRO, "SELFTEST_RUN", "Exceção no SelfTest_RunAll: " & Err.Number & " - " & Err.Description, "Verifique o código do M12 e o logger (M02)."
End Sub

' =============================================================================
' Teste 1: Sanitização ASCII_SAFE
' =============================================================================

Private Sub SelfTest_SanitizeFilename()
    On Error GoTo EH

    Dim inName As String
    inName = "MODELO – RELATÓRIO – Comunicação CPSA – 2026-02-08_1517.docx"

    Dim outName As String
    outName = SanitizeFilename_AsciiSafe(inName)

    Dim ok As Boolean
    ok = True

    ' Checks mínimos
    If Right$(LCase$(outName), 5) <> ".docx" Then ok = False
    If InStr(1, outName, " ", vbBinaryCompare) > 0 Then ok = False
    If InStr(1, outName, ChrW(8211), vbBinaryCompare) > 0 Then ok = False ' – en dash
    If InStr(1, outName, ChrW(8212), vbBinaryCompare) > 0 Then ok = False ' — em dash
    If InStr(1, UCase$(outName), "RELATORIO", vbBinaryCompare) = 0 Then ok = False
    If Not IsAsciiOnly(outName) Then ok = False

    If ok Then
        SelfTest_Log SEV_INFO, "SELFTEST_FILENAME", "FILENAME_SANITIZED PASS: " & outName, "OK"
    Else
        SelfTest_Log SEV_ALERTA, "SELFTEST_FILENAME", "FILENAME_SANITIZED FAIL: " & outName, "Rever regras ASCII_SAFE (acentos, espaços, extensão, caracteres especiais)."
    End If

    Exit Sub

EH:
    SelfTest_Log SEV_ERRO, "SELFTEST_FILENAME", "Exceção no SelfTest_SanitizeFilename: " & Err.Number & " - " & Err.Description, "Rever função SanitizeFilename_AsciiSafe no M12."
End Sub

' =============================================================================
' Teste 2: Multipart build (LOCAL) em Byte()
' =============================================================================

Private Sub SelfTest_MultipartBuild_Local()
    On Error GoTo EH

    Dim boundary As String
    boundary = BuildBoundary_SelfTest()

    Dim purpose As String: purpose = "user_data"
    Dim fileName As String: fileName = "arquivo_teste.bin"
    Dim contentType As String: contentType = "application/octet-stream"

    Dim fileBytes() As Byte
    fileBytes = ToBytes_Ansi("ABC123") ' bytes fictícios (não lê ficheiro em disco)

    Dim body() As Byte
    body = BuildMultipartBody_Bytes(boundary, purpose, fileName, contentType, fileBytes)

    Dim ok As Boolean
    ok = True

    ' Validações estruturais
    If ByteLen(body) <= ByteLen(fileBytes) Then ok = False

    Dim startNeedle() As Byte
    startNeedle = ToBytes_Ansi("--" & boundary & vbCrLf)
    If BytesIndexOf(body, startNeedle) <> 0 Then ok = False

    Dim endNeedle() As Byte
    endNeedle = ToBytes_Ansi(vbCrLf & "--" & boundary & "--" & vbCrLf)
    If BytesLastIndexOf(body, endNeedle) < 0 Then ok = False

    ' Deve conter o bloco "purpose"
    Dim needlePurpose() As Byte
    needlePurpose = ToBytes_Ansi("name=""purpose""" & vbCrLf & vbCrLf & purpose)
    If BytesIndexOf(body, needlePurpose) < 0 Then ok = False

    ' Deve conter o bloco "file" e o filename
    Dim needleFile() As Byte
    needleFile = ToBytes_Ansi("name=""file""; filename=""" & fileName & """")
    If BytesIndexOf(body, needleFile) < 0 Then ok = False

    ' Deve conter os bytes do ficheiro fictício (ABC123)
    If BytesIndexOf(body, fileBytes) < 0 Then ok = False

    If ok Then
        SelfTest_Log SEV_INFO, "SELFTEST_MULTIPART", "MULTIPART_BUILD PASS (len=" & CStr(ByteLen(body)) & "; boundary=" & boundary & ")", "OK"
    Else
        SelfTest_Log SEV_ALERTA, "SELFTEST_MULTIPART", "MULTIPART_BUILD FAIL (len=" & CStr(ByteLen(body)) & "; boundary=" & boundary & ")", "Rever CRLF/boundary/fecho --boundary-- e concatenação de bytes."
    End If

    Exit Sub

EH:
    SelfTest_Log SEV_ERRO, "SELFTEST_MULTIPART", "Exceção no SelfTest_MultipartBuild_Local: " & Err.Number & " - " & Err.Description, "Rever BuildMultipartBody_Bytes/BytesIndexOf no M12."
End Sub

' =============================================================================
' Teste 3: Engines COM disponíveis (WinHTTP / MSXML)
' =============================================================================

Private Sub SelfTest_EnginesAvailability()
    On Error GoTo EH

    Dim okWinHttp As Boolean, okMsxml As Boolean
    Dim errWinHttp As String, errMsxml As String

    okWinHttp = TryCreateObject("WinHttp.WinHttpRequest.5.1", errWinHttp)
    okMsxml = TryCreateObject("MSXML2.ServerXMLHTTP.6.0", errMsxml)

    If okWinHttp Then
        SelfTest_Log SEV_INFO, "SELFTEST_ENGINE", "WINHTTP disponível (CreateObject OK).", "OK"
    Else
        SelfTest_Log SEV_ALERTA, "SELFTEST_ENGINE", "WINHTTP indisponível: " & errWinHttp, "Pode falhar upload. Verificar instalação/políticas do Windows."
    End If

    If okMsxml Then
        SelfTest_Log SEV_INFO, "SELFTEST_ENGINE", "MSXML disponível (CreateObject OK).", "OK"
    Else
        SelfTest_Log SEV_ALERTA, "SELFTEST_ENGINE", "MSXML indisponível: " & errMsxml, "Fallback de engine pode não funcionar. Verificar MSXML6."
    End If

    Exit Sub

EH:
    SelfTest_Log SEV_ERRO, "SELFTEST_ENGINE", "Exceção no SelfTest_EnginesAvailability: " & Err.Number & " - " & Err.Description, "Rever TryCreateObject no M12."
End Sub

' =============================================================================
' Teste 4: resolução de OPENAI_API_KEY (ENV -> Config!B1)
' =============================================================================

Private Sub SelfTest_ConfigApiKeyResolution()
    On Error GoTo EH

    Dim apiKey As String, src As String, warnTxt As String, errTxt As String
    Dim ok As Boolean

    ' Cenário A: ENV presente, Config literal também presente -> usa ENV + alerta
    ok = Config_SelfTest_ResolveOpenAIApiKey("env-secret", "cfg-secret", apiKey, src, warnTxt, errTxt)
    If ok And src = "ENV" And apiKey = "env-secret" And warnTxt <> "" And errTxt = "" Then
        SelfTest_Log SEV_INFO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION ENV precedence: PASS", "OK"
    Else
        SelfTest_Log SEV_ERRO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION ENV precedence: FAIL", "Esperado source=ENV com alerta de key literal em Config!B1."
    End If

    ' Cenário B: ENV vazio, Config literal válida -> fallback com alerta
    ok = Config_SelfTest_ResolveOpenAIApiKey("", "cfg-secret", apiKey, src, warnTxt, errTxt)
    If ok And src = "CONFIG_B1" And apiKey = "cfg-secret" And warnTxt <> "" And errTxt = "" Then
        SelfTest_Log SEV_INFO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION Config fallback: PASS", "OK"
    Else
        SelfTest_Log SEV_ERRO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION Config fallback: FAIL", "Esperado source=CONFIG_B1 com alerta de migração para ambiente."
    End If

    ' Cenário C: Config diretiva Environ e ENV vazio -> erro
    ok = Config_SelfTest_ResolveOpenAIApiKey("", "(Environ(""OPENAI_API_KEY""))", apiKey, src, warnTxt, errTxt)
    If (Not ok) And errTxt <> "" Then
        SelfTest_Log SEV_INFO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION Environ directive sem ENV: PASS", "OK"
    Else
        SelfTest_Log SEV_ERRO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION Environ directive sem ENV: FAIL", "Esperado erro com instrução para definir OPENAI_API_KEY no ambiente."
    End If

    ' Cenário D: sem ENV e sem Config válida -> erro
    ok = Config_SelfTest_ResolveOpenAIApiKey("", "", apiKey, src, warnTxt, errTxt)
    If (Not ok) And errTxt <> "" Then
        SelfTest_Log SEV_INFO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION sem fontes: PASS", "OK"
    Else
        SelfTest_Log SEV_ERRO, "SELFTEST_CONFIG", "APIKEY_RESOLUTION sem fontes: FAIL", "Esperado erro quando ENV e Config!B1 estão vazios/inválidos."
    End If

    Exit Sub
EH:
    SelfTest_Log SEV_ERRO, "SELFTEST_CONFIG", "Exceção no SelfTest_ConfigApiKeyResolution: " & Err.Number & " - " & Err.Description, "Verificar M14_ConfigApiKey e cenários de precedência."
End Sub


Private Sub SelfTest_Schema_FilesManagement()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("FILES_MANAGEMENT")

    Dim must As Variant
    must = Array("Timestamp", "File name", "Full path", "notes")

    Dim i As Long
    For i = LBound(must) To UBound(must)
        If SelfTest_FindHeader(ws, CStr(must(i))) = 0 Then
            SelfTest_Log SEV_ERRO, "SELFTEST_SCHEMA", "Header em falta na FILES_MANAGEMENT: " & CStr(must(i)), "Adicionar cabeçalho em linha 1 sem renomear colunas existentes."
            Exit Sub
        End If
    Next i

    SelfTest_Log SEV_INFO, "SELFTEST_SCHEMA", "Schema FILES_MANAGEMENT mínimo: PASS", "OK"
    Exit Sub
EH:
    SelfTest_Log SEV_ERRO, "SELFTEST_SCHEMA", "Exceção no SelfTest_Schema_FilesManagement: " & Err.Number & " - " & Err.Description, "Verificar folha FILES_MANAGEMENT e headers."
End Sub

Private Sub SelfTest_OutputRegister_And_Resolve()
    On Error GoTo EH

    Dim runId As String
    runId = "SELFTEST_RUN_" & Format$(Now, "yyyymmdd_hhnnss")

    Call Files_SetRunToken(runId)

    Dim tempFolder As String
    tempFolder = Environ$("TEMP")
    If Trim$(tempFolder) = "" Then tempFolder = ThisWorkbook.Path

    Dim f1 As String
    f1 = tempFolder & "\pipeliner_selftest_output1.txt"

    Dim ff As Integer
    ff = FreeFile
    Open f1 For Output As #ff
    Print #ff, "SELFTEST OUTPUT 1"
    Close #ff

    Call Files_LogEventOutput("SELFTEST_PIPE", "SELFTEST/PROMPT/A", tempFolder, f1, "output(selftest)", "DL", "selftest=1", "", runId, 1, 0, "OUTPUT")

    Dim rp As String, rn As String, st As String, cand As String
    Call SelfTest_InvokeResolve("@LAST_OUTPUT", 2, rp, rn, st, cand)

    If st = "OK" And Len(Trim$(rp)) > 0 Then
        SelfTest_Log SEV_INFO, "SELFTEST_OUTPUT_CHAIN", "@LAST_OUTPUT resolve: PASS -> " & rn, "OK"
    Else
        SelfTest_Log SEV_ERRO, "SELFTEST_OUTPUT_CHAIN", "@LAST_OUTPUT resolve: FAIL | status=" & st & " | " & cand, "Confirmar registo de output e parser de tokens @LAST_OUTPUT/@OUTPUT(...)."
    End If

    Call SelfTest_InvokeResolve("@OUTPUT(step_n=1,index=0)", 2, rp, rn, st, cand)
    If st = "OK" Then
        SelfTest_Log SEV_INFO, "SELFTEST_OUTPUT_CHAIN", "@OUTPUT(...) resolve: PASS -> " & rn, "OK"
    Else
        SelfTest_Log SEV_ERRO, "SELFTEST_OUTPUT_CHAIN", "@OUTPUT(...) resolve: FAIL | status=" & st & " | " & cand, "Confirmar filtros prompt_id/step_n/filename/index."
    End If

    On Error Resume Next
    Kill f1
    On Error GoTo 0
    Exit Sub
EH:
    SelfTest_Log SEV_ERRO, "SELFTEST_OUTPUT_CHAIN", "Exceção no SelfTest_OutputRegister_And_Resolve: " & Err.Number & " - " & Err.Description, "Verificar M09 (Files_LogEventOutput / resolução de tokens)."
End Sub

Private Sub SelfTest_InvokeResolve(ByVal token As String, ByVal stepN As Long, ByRef resolvedPath As String, ByRef resolvedName As String, ByRef status As String, ByRef candidatos As String)
    On Error GoTo EH
    Application.Run "Files_ResolverOutputToken", "SELFTEST_PIPE", "SELFTEST", stepN, token, resolvedPath, resolvedName, status, candidatos
    Exit Sub
EH:
    status = "NOT_FOUND"
    candidatos = "invoke-fail: " & Err.Description
End Sub

Private Function SelfTest_FindHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long
    On Error GoTo Fim
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            SelfTest_FindHeader = c
            Exit Function
        End If
    Next c
Fim:
End Function

' =============================================================================
' Logging (compatível com o PIPELINER)
' =============================================================================

Private Sub SelfTest_Log(ByVal severidade As String, ByVal parametro As String, ByVal problema As String, ByVal sugestao As String)
    On Error Resume Next
    ' Usa o logger do PIPELINER (M02_Logger_DEBUG_e_Seguimento)
    Debug_Registar 0, SELFTEST_PROMPT_ID, severidade, Empty, parametro, problema, sugestao
End Sub

' =============================================================================
' Idempotência: limpar linhas antigas do SELFTEST na folha DEBUG
' (Só remove linhas cujo Prompt ID seja SELFTEST e Parametro comece por SELFTEST_)
' =============================================================================

Private Sub SelfTest_ClearPreviousDebugRows()
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DEBUG")

    Dim colPrompt As Long, colParam As Long
    colPrompt = FindHeaderColumn(ws, "Prompt ID")
    colParam = FindHeaderColumn(ws, "Parâmetro") ' ou "Parametro" (normalização trata)

    If colPrompt = 0 Or colParam = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rowS.Count, colPrompt).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    For r = lastRow To 2 Step -1
        Dim pid As String, par As String
        pid = CStr(ws.Cells(r, colPrompt).value)
        par = CStr(ws.Cells(r, colParam).value)

        If StrComp(Trim$(pid), SELFTEST_PROMPT_ID, vbTextCompare) = 0 Then
            If Left$(Trim$(par), Len(SELFTEST_PARAM_PREFIX)) = SELFTEST_PARAM_PREFIX Then
                ws.rowS(r).Delete
            End If
        End If
    Next r

    Exit Sub

FailSoft:
    ' não falhar o pipeline por causa de limpeza de logs
End Sub

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim target As String
    target = NormalizeHeader(headerName)

    Dim c As Long
    For c = 1 To lastCol
        Dim h As String
        h = NormalizeHeader(CStr(ws.Cells(1, c).value))
        If h = target Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    s = LCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = RemoveDiacriticsPT(s)
    NormalizeHeader = s
End Function

' =============================================================================
' Helpers: sanitização ASCII_SAFE
' =============================================================================

Private Function SanitizeFilename_AsciiSafe(ByVal fileName As String) As String
    Dim base As String, ext As String
    SplitNameExt fileName, base, ext

    Dim s As String
    s = base

    ' Remover diacríticos (PT)
    s = RemoveDiacriticsPT(s)

    ' Normalizar travessões
    s = Replace(s, ChrW(8211), "-") ' –
    s = Replace(s, ChrW(8212), "-") ' —

    ' Espaços para hífen
    s = Replace(s, " ", "-")

    ' Remover/normalizar caracteres problemáticos
    s = SanitizeForbiddenChars(s)

    ' Colapsar hífens repetidos
    s = CollapseRepeats(s, "-")

    ' Limpar extremos
    s = TrimChars(s, "-_.")

    ' Limite simples (preserva extensão)
    If Len(s) > 160 Then s = Left$(s, 160)

    If ext <> "" Then
        SanitizeFilename_AsciiSafe = s & "." & ext
    Else
        SanitizeFilename_AsciiSafe = s
    End If
End Function

Private Sub SplitNameExt(ByVal fileName As String, ByRef outBase As String, ByRef outExt As String)
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 And p < Len(fileName) Then
        outBase = Left$(fileName, p - 1)
        outExt = Mid$(fileName, p + 1)
    Else
        outBase = fileName
        outExt = ""
    End If
End Sub

Private Function IsAsciiOnly(ByVal s As String) As Boolean
    Dim i As Long, code As Long
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code < 0 Or code > 127 Then
            IsAsciiOnly = False
            Exit Function
        End If
    Next i
    IsAsciiOnly = True
End Function

Private Function RemoveDiacriticsPT(ByVal s As String) As String
    ' Usa ChrW para evitar problemas de encoding ao importar/exportar .bas
    ' a/A
    s = Replace(s, ChrW(225), "a")
    s = Replace(s, ChrW(224), "a")
    s = Replace(s, ChrW(226), "a")
    s = Replace(s, ChrW(227), "a")
    s = Replace(s, ChrW(228), "a")
    s = Replace(s, ChrW(193), "A")
    s = Replace(s, ChrW(192), "A")
    s = Replace(s, ChrW(194), "A")
    s = Replace(s, ChrW(195), "A")
    s = Replace(s, ChrW(196), "A")

    ' e/E
    s = Replace(s, ChrW(233), "e")
    s = Replace(s, ChrW(232), "e")
    s = Replace(s, ChrW(234), "e")
    s = Replace(s, ChrW(235), "e")
    s = Replace(s, ChrW(201), "E")
    s = Replace(s, ChrW(200), "E")
    s = Replace(s, ChrW(202), "E")
    s = Replace(s, ChrW(203), "E")

    ' i/I
    s = Replace(s, ChrW(237), "i")
    s = Replace(s, ChrW(236), "i")
    s = Replace(s, ChrW(238), "i")
    s = Replace(s, ChrW(239), "i")
    s = Replace(s, ChrW(205), "I")
    s = Replace(s, ChrW(204), "I")
    s = Replace(s, ChrW(206), "I")
    s = Replace(s, ChrW(207), "I")

    ' o/O
    s = Replace(s, ChrW(243), "o")
    s = Replace(s, ChrW(242), "o")
    s = Replace(s, ChrW(244), "o")
    s = Replace(s, ChrW(245), "o")
    s = Replace(s, ChrW(246), "o")
    s = Replace(s, ChrW(211), "O")
    s = Replace(s, ChrW(210), "O")
    s = Replace(s, ChrW(212), "O")
    s = Replace(s, ChrW(213), "O")
    s = Replace(s, ChrW(214), "O")

    ' u/U
    s = Replace(s, ChrW(250), "u")
    s = Replace(s, ChrW(249), "u")
    s = Replace(s, ChrW(251), "u")
    s = Replace(s, ChrW(252), "u")
    s = Replace(s, ChrW(218), "U")
    s = Replace(s, ChrW(217), "U")
    s = Replace(s, ChrW(219), "U")
    s = Replace(s, ChrW(220), "U")

    ' c/C
    s = Replace(s, ChrW(231), "c")
    s = Replace(s, ChrW(199), "C")

    ' n/N (não PT puro, mas aparece)
    s = Replace(s, ChrW(241), "n")
    s = Replace(s, ChrW(209), "N")

    RemoveDiacriticsPT = s
End Function

Private Function SanitizeForbiddenChars(ByVal s As String) As String
    Dim forb As Variant, i As Long
    forb = Array(":", "*", "?", """", "<", ">", "|", "\", "/", vbTab, vbCr, vbLf)
    For i = LBound(forb) To UBound(forb)
        s = Replace$(s, CStr(forb(i)), "_")
    Next i
    SanitizeForbiddenChars = s
End Function

Private Function CollapseRepeats(ByVal s As String, ByVal token As String) As String
    Dim doubleToken As String
    doubleToken = token & token
    Do While InStr(1, s, doubleToken, vbBinaryCompare) > 0
        s = Replace$(s, doubleToken, token)
    Loop
    CollapseRepeats = s
End Function

Private Function TrimChars(ByVal s As String, ByVal chars As String) As String
    Do While Len(s) > 0 And InStr(1, chars, Left$(s, 1), vbBinaryCompare) > 0
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0 And InStr(1, chars, Right$(s, 1), vbBinaryCompare) > 0
        s = Left$(s, Len(s) - 1)
    Loop
    TrimChars = s
End Function

' =============================================================================
' Helpers: multipart LOCAL em bytes
' =============================================================================

Private Function BuildBoundary_SelfTest() As String
    Randomize
    BuildBoundary_SelfTest = "----SELFTEST_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 1000000))
End Function

Private Function BuildMultipartBody_Bytes( _
    ByVal boundary As String, _
    ByVal purpose As String, _
    ByVal fileName As String, _
    ByVal contentType As String, _
    ByRef fileBytes() As Byte _
) As Byte()

    Dim pre1 As String, pre2 As String, post As String

    pre1 = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""purpose""" & vbCrLf & vbCrLf & _
           purpose & vbCrLf

    pre2 = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf & _
           "Content-Type: " & contentType & vbCrLf & vbCrLf

    post = vbCrLf & "--" & boundary & "--" & vbCrLf

    Dim b1() As Byte, b2() As Byte, b4() As Byte
    b1 = ToBytes_Ansi(pre1)
    b2 = ToBytes_Ansi(pre2)
    b4 = ToBytes_Ansi(post)

    BuildMultipartBody_Bytes = BytesConcat4(b1, b2, fileBytes, b4)
End Function

Private Function ToBytes_Ansi(ByVal s As String) As Byte()
    Dim b() As Byte
    b = StrConv(s, vbFromUnicode)
    ToBytes_Ansi = b
End Function

Private Function BytesConcat4(ByRef a() As Byte, ByRef b() As Byte, ByRef c() As Byte, ByRef d() As Byte) As Byte()
    Dim la As Long, lb As Long, lc As Long, ld As Long
    la = ByteLen(a): lb = ByteLen(b): lc = ByteLen(c): ld = ByteLen(d)

    Dim total As Long
    total = la + lb + lc + ld
    If total <= 0 Then
        ReDim BytesConcat4(0 To 0) As Byte
        Exit Function
    End If

    Dim out() As Byte
    ReDim out(0 To total - 1) As Byte

    Dim pos As Long
    pos = 0
    CopyBytes out, pos, a: pos = pos + la
    CopyBytes out, pos, b: pos = pos + lb
    CopyBytes out, pos, c: pos = pos + lc
    CopyBytes out, pos, d

    BytesConcat4 = out
End Function

Private Sub CopyBytes(ByRef dest() As Byte, ByVal destPos As Long, ByRef src() As Byte)
    Dim i As Long, ls As Long
    ls = ByteLen(src)
    If ls <= 0 Then Exit Sub
    For i = 0 To ls - 1
        dest(destPos + i) = src(i)
    Next i
End Sub

Private Function ByteLen(ByRef b() As Byte) As Long
    On Error GoTo EH
    ByteLen = (UBound(b) - LBound(b) + 1)
    Exit Function
EH:
    ByteLen = 0
End Function

Private Function BytesIndexOf(ByRef hay() As Byte, ByRef needle() As Byte) As Long
    Dim lh As Long, ln As Long
    lh = ByteLen(hay)
    ln = ByteLen(needle)

    If ln = 0 Or lh = 0 Or ln > lh Then
        BytesIndexOf = -1
        Exit Function
    End If

    Dim i As Long, j As Long
    For i = 0 To lh - ln
        For j = 0 To ln - 1
            If hay(i + j) <> needle(j) Then GoTo NextI
        Next j
        BytesIndexOf = i
        Exit Function
NextI:
    Next i

    BytesIndexOf = -1
End Function

Private Function BytesLastIndexOf(ByRef hay() As Byte, ByRef needle() As Byte) As Long
    Dim lh As Long, ln As Long
    lh = ByteLen(hay)
    ln = ByteLen(needle)

    If ln = 0 Or lh = 0 Or ln > lh Then
        BytesLastIndexOf = -1
        Exit Function
    End If

    Dim i As Long, j As Long
    For i = lh - ln To 0 Step -1
        For j = 0 To ln - 1
            If hay(i + j) <> needle(j) Then GoTo NextI
        Next j
        BytesLastIndexOf = i
        Exit Function
NextI:
    Next i

    BytesLastIndexOf = -1
End Function

' =============================================================================
' Helpers: COM
' =============================================================================

Private Function TryCreateObject(ByVal progId As String, ByRef outErr As String) As Boolean
    On Error GoTo EH
    Dim o As Object
    Set o = CreateObject(progId)
    Set o = Nothing
    outErr = ""
    TryCreateObject = True
    Exit Function
EH:
    outErr = "CreateObject falhou (" & progId & "): " & Err.Number & " - " & Err.Description
    TryCreateObject = False
End Function


