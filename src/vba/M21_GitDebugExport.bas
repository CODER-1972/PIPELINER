Attribute VB_Name = "M21_GitDebugExport"
Option Explicit

' =============================================================================
' MÃ³dulo: M21_GitDebugExport
' PropÃ³sito:
' - Orquestrar exportaÃ§Ã£o opcional dos logs DEBUG/Seguimento/CatÃ¡logo para GitHub.
' - Manter a assinatura pÃºblica estÃ¡vel para chamadas de outros mÃ³dulos.
' - Produzir bundle por run com observabilidade e manifesto de execuÃ§Ã£o.
'
' AtualizaÃ§Ãµes:
' - 2026-03-05 | Codex | Hardening e observabilidade (PR2)
'   - Adiciona telemetria por etapa e sanitizaÃ§Ã£o bÃ¡sica de dados sensÃ­veis.
'   - Adiciona `manifest.json` com estado por artefacto e parÃ¢metros de execuÃ§Ã£o.
'   - Integra timeout/retry/backoff configurÃ¡veis do mÃ³dulo HTTP.
' - 2026-03-05 | Codex | CorreÃ§Ã£o funcional crÃ­tica e estrutura por pasta em logs
'   - Passa a exportar trÃªs ficheiros (catalogo/debug/seguimento) por execuÃ§Ã£o.
'   - Cria pasta lÃ³gica em `logs/` no formato `YYYY-MM-SS - HHMM - [pipeline]`.
'   - Adiciona leitura de SHA atual no GitHub para suportar update idempotente em `/contents`.
' - 2026-03-04 | Codex | RefatoraÃ§Ã£o para facade de alto nÃ­vel
'   - Preserva entry point pÃºblico PipelineGitDebug_ExportIfEnabled.
'   - Move leitura de config, HTTP, blob/base64 e logging para mÃ³dulos M22-M26.
'
' FunÃ§Ãµes e procedimentos:
' - PipelineGitDebug_ExportIfEnabled(Optional pipelineIndex As Long = 0) (Sub)
'   - Entry point pÃºblico; avalia enable/config e executa exportaÃ§Ã£o para GitHub.
' - M21_ExportRunBundle(cfg As Object, pipelineIndex As Long, pipelineName As String) As Boolean
'   - Publica bundle (catÃ¡logo, DEBUG, Seguimento e manifesto) em `logs/<run>/`.
' - M21_CatalogosDosPromptsDoDebug() As String
'   - Resolve blocos de catÃ¡logo para os Prompt IDs encontrados na folha DEBUG.
' - GitDebugExport_SelfTest_Basico() (Sub)
'   - Valida naming/sanitizaÃ§Ã£o de pasta de run e regista PASS/FAIL no DEBUG.
' =============================================================================

Public Sub PipelineGitDebug_ExportIfEnabled(Optional ByVal pipelineIndex As Long = 0)
    On Error GoTo EH

    Dim cfg As Object
    Set cfg = GH_Config_Load()

    If Not GH_Config_IsEnabled(cfg) Then Exit Sub

    Dim invalidReason As String
    If Not GH_Config_Validate(cfg, invalidReason) Then
        Call GH_LogWarn(0, "DEBUG", "GIT_DEBUG_EXPORT_DISABLED", invalidReason, _
                        "Preencha as chaves GIT_DEBUG_* na folha Config ou desative a feature.")
        Exit Sub
    End If

    Dim pipelineName As String
    pipelineName = M21_ReadPipelineName(pipelineIndex)

    Call M21_LogStage("start", "pipeline=" & pipelineName)

    If M21_ExportRunBundle(cfg, pipelineIndex, pipelineName) Then
        Call GH_LogInfo(0, "DEBUG", "GIT_DEBUG_EXPORT_OK", _
                        "Bundle de exportaÃ§Ã£o GitHub concluÃ­do.", _
                        "Verifique a pasta da run em logs/<YYYY-MM-SS - HHMM - [pipeline]>.")
        Call M21_LogStage("done", "pipeline=" & pipelineName)
    End If

    Exit Sub
EH:
    Call GH_LogError(0, "DEBUG", "GIT_DEBUG_EXPORT_EXCEPTION", _
                     "Erro inesperado no export GitHub: " & Err.Description, _
                     "Revise as configuraÃ§Ãµes GIT_DEBUG_* e o estado das folhas DEBUG/Seguimento.")
End Sub

Private Function M21_ExportRunBundle(ByVal cfg As Object, ByVal pipelineIndex As Long, ByVal pipelineName As String) As Boolean
    Dim catalogoTxt As String
    Dim debugTxt As String
    Dim seguimentoTxt As String

    Call M21_LogStage("collect_catalog", "")
    catalogoTxt = M21_SanitizeForExport(M21_CatalogosDosPromptsDoDebug())

    Call M21_LogStage("collect_debug", "")
    debugTxt = M21_SanitizeForExport(M21_SheetAsTsv("DEBUG"))

    Call M21_LogStage("collect_seguimento", "")
    seguimentoTxt = M21_SanitizeForExport(M21_SheetAsTsv("Seguimento"))

    If Len(catalogoTxt) = 0 Then catalogoTxt = "[Sem conteÃºdo de catÃ¡logo para exportar.]"
    If Len(debugTxt) = 0 Then debugTxt = "[Folha DEBUG sem conteÃºdo.]"
    If Len(seguimentoTxt) = 0 Then seguimentoTxt = "[Folha Seguimento sem conteÃºdo.]"

    Dim rootLogsPath As String
    rootLogsPath = M21_ResolveLogsRootPath(GH_Config_GetString(cfg, "path"))

    Dim runFolder As String
    runFolder = M21_BuildRunFolderName(pipelineName)

    Dim bundleBasePath As String
    bundleBasePath = rootLogsPath & "/" & runFolder

    Dim httpTimeoutMs As Long
    Dim httpMaxRetries As Long
    Dim httpBackoffMs As Long
    httpTimeoutMs = GH_Config_GetLong(cfg, "http_timeout_ms", 30000)
    httpMaxRetries = GH_Config_GetLong(cfg, "http_max_retries", 2)
    httpBackoffMs = GH_Config_GetLong(cfg, "http_backoff_ms", 800)

    Dim resultCatalogo As Object
    Dim resultDebug As Object
    Dim resultSeguimento As Object

    Set resultCatalogo = M21_UploadTextFile(cfg, bundleBasePath & "/catalogo_prompts.tsv", catalogoTxt, pipelineIndex, httpTimeoutMs, httpMaxRetries, httpBackoffMs)
    Set resultDebug = M21_UploadTextFile(cfg, bundleBasePath & "/debug.tsv", debugTxt, pipelineIndex, httpTimeoutMs, httpMaxRetries, httpBackoffMs)
    Set resultSeguimento = M21_UploadTextFile(cfg, bundleBasePath & "/seguimento.tsv", seguimentoTxt, pipelineIndex, httpTimeoutMs, httpMaxRetries, httpBackoffMs)

    Dim manifestJson As String
    manifestJson = M21_BuildManifestJson(pipelineIndex, pipelineName, bundleBasePath, resultCatalogo, resultDebug, resultSeguimento, httpTimeoutMs, httpMaxRetries, httpBackoffMs)

    Dim resultManifest As Object
    Set resultManifest = M21_UploadTextFile(cfg, bundleBasePath & "/manifest.json", manifestJson, pipelineIndex, httpTimeoutMs, httpMaxRetries, httpBackoffMs)

    M21_ExportRunBundle = (CBool(resultCatalogo("ok")) And CBool(resultDebug("ok")) And CBool(resultSeguimento("ok")) And CBool(resultManifest("ok")))
End Function

Private Function M21_UploadTextFile( _
    ByVal cfg As Object, _
    ByVal repoPath As String, _
    ByVal fileContent As String, _
    ByVal pipelineIndex As Long, _
    ByVal timeoutMs As Long, _
    ByVal maxRetries As Long, _
    ByVal backoffMs As Long) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("path") = repoPath
    result("ok") = False
    result("status") = 0
    result("attempts") = 0

    Dim url As String
    url = GH_TreeCommit_ContentsUrl( _
            GH_Config_GetString(cfg, "base_url"), _
            GH_Config_GetString(cfg, "owner"), _
            GH_Config_GetString(cfg, "repo"), _
            repoPath)

    Call M21_LogStage("get_sha", "path=" & repoPath)
    Dim existingSha As String
    existingSha = M21_GetRemoteSha(url, cfg, timeoutMs, maxRetries, backoffMs)

    Dim payload As String
    payload = GH_TreeCommit_BuildContentsPayload( _
                GH_TreeCommit_DefaultMessage(pipelineIndex), _
                GH_Config_GetString(cfg, "branch"), _
                GH_Blob_Base64FromText(fileContent), _
                existingSha)

    Dim statusCode As Long
    Dim responseText As String
    Dim errText As String
    Dim ok As Boolean
    Dim attemptsUsed As Long

    Call M21_LogStage("put_start", "path=" & repoPath & "|has_sha=" & IIf(existingSha <> "", "SIM", "NAO"))
    ok = GH_HTTP_RequestJson("PUT", url, GH_Config_GetString(cfg, "token"), payload, _
                             statusCode, responseText, errText, GH_Config_GetString(cfg, "user_agent"), _
                             timeoutMs, maxRetries, backoffMs, attemptsUsed)

    result("ok") = ok
    result("status") = statusCode
    result("attempts") = attemptsUsed

    If ok Then
        Call GH_LogInfo(0, "DEBUG", "GIT_DEBUG_EXPORT_FILE_OK", _
                        "Ficheiro exportado: " & repoPath & " (HTTP " & CStr(statusCode) & ", attempts=" & CStr(attemptsUsed) & ").", "")
    Else
        Call GH_LogError(0, "DEBUG", "GIT_DEBUG_EXPORT_FILE_FAIL", _
                         "Falha ao exportar " & repoPath & " (HTTP " & CStr(statusCode) & ", attempts=" & CStr(attemptsUsed) & "). " & _
                         M21_LeftSafe(responseText, 220), _
                         "Valide token/permissÃµes, branch e path do repositÃ³rio.")
        If Len(errText) > 0 Then
            Call GH_LogWarn(0, "DEBUG", "GIT_DEBUG_EXPORT_ENGINE", errText, _
                            "Confirme suporte WinHTTP/MSXML no host Office.")
        End If
    End If

    Set M21_UploadTextFile = result
End Function

Private Function M21_GetRemoteSha( _
    ByVal contentsUrl As String, _
    ByVal cfg As Object, _
    ByVal timeoutMs As Long, _
    ByVal maxRetries As Long, _
    ByVal backoffMs As Long) As String

    Dim statusCode As Long
    Dim responseText As String
    Dim errText As String
    Dim ok As Boolean
    Dim attemptsUsed As Long

    ok = GH_HTTP_RequestJson("GET", contentsUrl, GH_Config_GetString(cfg, "token"), "", _
                             statusCode, responseText, errText, GH_Config_GetString(cfg, "user_agent"), _
                             timeoutMs, maxRetries, backoffMs, attemptsUsed)

    If ok Then
        M21_GetRemoteSha = M21_ExtractJsonStringValue(responseText, "sha")
    Else
        M21_GetRemoteSha = ""
    End If
End Function

Private Function M21_BuildManifestJson( _
    ByVal pipelineIndex As Long, _
    ByVal pipelineName As String, _
    ByVal runPath As String, _
    ByVal resultCatalogo As Object, _
    ByVal resultDebug As Object, _
    ByVal resultSeguimento As Object, _
    ByVal timeoutMs As Long, _
    ByVal maxRetries As Long, _
    ByVal backoffMs As Long) As String

    Dim successAll As Boolean
    successAll = CBool(resultCatalogo("ok")) And CBool(resultDebug("ok")) And CBool(resultSeguimento("ok"))

    Dim json As String
    json = "{" & _
           """pipeline_index"":" & CStr(pipelineIndex) & "," & _
           """pipeline_name"":""" & GH_Blob_JsonEscape(pipelineName) & """," & _
           """run_path"":""" & GH_Blob_JsonEscape(runPath) & """," & _
           """generated_at"":""" & GH_Blob_JsonEscape(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & """," & _
           """http_timeout_ms"":" & CStr(timeoutMs) & "," & _
           """http_max_retries"":" & CStr(maxRetries) & "," & _
           """http_backoff_ms"":" & CStr(backoffMs) & "," & _
           """success_all"":" & LCase$(CStr(successAll)) & "," & _
           """artifacts"": [" & _
           M21_BuildManifestArtifactJson(resultCatalogo) & "," & _
           M21_BuildManifestArtifactJson(resultDebug) & "," & _
           M21_BuildManifestArtifactJson(resultSeguimento) & "]" & _
           "}"

    M21_BuildManifestJson = json
End Function

Private Function M21_BuildManifestArtifactJson(ByVal resultItem As Object) As String
    M21_BuildManifestArtifactJson = "{" & _
        """path"":""" & GH_Blob_JsonEscape(CStr(resultItem("path"))) & """," & _
        """ok"":" & LCase$(CStr(CBool(resultItem("ok")))) & "," & _
        """http_status"":" & CStr(CLng(resultItem("status"))) & "," & _
        """attempts"":" & CStr(CLng(resultItem("attempts"))) & _
        "}"
End Function

Private Sub M21_LogStage(ByVal stageName As String, ByVal contextText As String)
    Dim fullMsg As String
    fullMsg = "stage=" & stageName
    If Len(contextText) > 0 Then fullMsg = fullMsg & "|" & contextText
    Call GH_LogInfo(0, "DEBUG", "GH_EXPORT_STAGE", fullMsg, "")
End Sub

Private Function M21_SanitizeForExport(ByVal textIn As String) As String
    Dim s As String
    s = textIn

    s = M21_MaskAroundToken(s, "bearer ")
    s = M21_MaskAroundToken(s, "api_key")
    s = M21_MaskAroundToken(s, "token")
    s = M21_MaskAroundToken(s, "authorization")

    If Len(s) > 2000000 Then
        s = Left$(s, 2000000) & vbCrLf & "[TRUNCATED_FOR_EXPORT]"
    End If

    M21_SanitizeForExport = s
End Function

Private Function M21_MaskAroundToken(ByVal src As String, ByVal marker As String) As String
    Dim work As String
    work = src

    Dim pos As Long
    pos = InStr(1, LCase$(work), LCase$(marker), vbTextCompare)

    Do While pos > 0
        Dim startMask As Long
        startMask = pos + Len(marker)

        Dim endMask As Long
        endMask = InStr(startMask, work, vbCrLf)
        If endMask = 0 Then endMask = Len(work) + 1

        work = Left$(work, startMask - 1) & " ***" & Mid$(work, endMask)
        pos = InStr(startMask + 4, LCase$(work), LCase$(marker), vbTextCompare)
    Loop

    M21_MaskAroundToken = work
End Function

Public Sub GitDebugExport_SelfTest_Basico()
    On Error GoTo EH

    Dim f As String
    f = M21_BuildRunFolderName("Pipe: Test/Name")

    If InStr(1, f, "[") > 0 And InStr(1, f, "]") > 0 And InStr(1, f, "/") = 0 Then
        Call GH_LogInfo(0, "SELFTEST", "GH_EXPORT_SELFTEST", "PASS: naming/sanitize bÃ¡sico", "")
    Else
        Call GH_LogError(0, "SELFTEST", "GH_EXPORT_SELFTEST", "FAIL: naming/sanitize bÃ¡sico", "")
    End If
    Exit Sub
EH:
    Call GH_LogError(0, "SELFTEST", "GH_EXPORT_SELFTEST", "FAIL com exceÃ§Ã£o: " & Err.Description, "")
End Sub

Private Function M21_ExtractJsonStringValue(ByVal json As String, ByVal key As String) As String
    Dim marker As String
    marker = """" & key & """:"""

    Dim p As Long
    p = InStr(1, json, marker, vbTextCompare)
    If p <= 0 Then
        M21_ExtractJsonStringValue = ""
        Exit Function
    End If

    Dim startPos As Long
    startPos = p + Len(marker)

    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = startPos To Len(json)
        ch = Mid$(json, i, 1)
        If ch = """" Then
            If Mid$(json, i - 1, 1) <> "\" Then Exit For
        End If
        out = out & ch
    Next i

    M21_ExtractJsonStringValue = Replace$(out, "\" & Chr$(34), Chr$(34))
End Function

Private Function M21_ResolveLogsRootPath(ByVal configuredPath As String) As String
    Dim p As String
    p = Trim$(configuredPath)

    If p = "" Then
        M21_ResolveLogsRootPath = "logs"
        Exit Function
    End If

    If Right$(p, 1) = "/" Then p = Left$(p, Len(p) - 1)

    Dim slashPos As Long
    slashPos = InStrRev(p, "/")

    If InStrRev(p, ".") > slashPos And slashPos > 0 Then
        M21_ResolveLogsRootPath = Left$(p, slashPos - 1)
    ElseIf LCase$(Left$(p, 4)) <> "logs" Then
        M21_ResolveLogsRootPath = "logs"
    Else
        M21_ResolveLogsRootPath = p
    End If

    If Len(Trim$(M21_ResolveLogsRootPath)) = 0 Then M21_ResolveLogsRootPath = "logs"
End Function

Private Function M21_BuildRunFolderName(ByVal pipelineName As String) As String
    Dim prefix As String
    prefix = Format$(Now, "yyyy-mm-ss - hhnn")
    M21_BuildRunFolderName = prefix & " - [" & M21_SanitizePathPart(pipelineName) & "]"
End Function

Private Function M21_ReadPipelineName(ByVal pipelineIndex As Long) As String
    On Error GoTo Fallback

    If pipelineIndex <= 0 Then GoTo Fallback

    Dim wsPainel As Worksheet
    Set wsPainel = ThisWorkbook.Worksheets("PAINEL")

    Dim colIniciar As Long
    colIniciar = 2 + (pipelineIndex - 1) * 2

    Dim nome As String
    nome = Trim$(CStr(wsPainel.Cells(1, colIniciar).Value))
    If nome = "" Then GoTo Fallback

    M21_ReadPipelineName = nome
    Exit Function
Fallback:
    M21_ReadPipelineName = "Pipeline_" & Format$(pipelineIndex, "00")
End Function

Private Function M21_SanitizePathPart(ByVal s As String) As String
    Dim out As String
    out = Trim$(s)
    out = Replace$(out, "\", "_")
    out = Replace$(out, "/", "_")
    out = Replace$(out, ":", "_")
    out = Replace$(out, "*", "_")
    out = Replace$(out, "?", "_")
    out = Replace$(out, """", "_")
    out = Replace$(out, "<", "_")
    out = Replace$(out, ">", "_")
    out = Replace$(out, "|", "_")
    If out = "" Then out = "Pipeline"
    M21_SanitizePathPart = out
End Function

Private Function M21_CatalogosDosPromptsDoDebug() As String
    On Error GoTo EH

    Dim promptIds As Object
    Set promptIds = M21_PromptIdsFromDebug()

    If promptIds.Count = 0 Then
        M21_CatalogosDosPromptsDoDebug = "[Sem prompts encontrados no DEBUG.]"
        Exit Function
    End If

    Dim sb As String
    Dim promptId As Variant

    For Each promptId In promptIds.Keys
        sb = sb & M21_BlocosCatalogoPorPromptId(CStr(promptId)) & vbCrLf
    Next promptId

    M21_CatalogosDosPromptsDoDebug = Trim$(sb)
    Exit Function
EH:
    M21_CatalogosDosPromptsDoDebug = "[Erro ao montar catÃ¡logos: " & Err.Description & "]"
End Function

Private Function M21_PromptIdsFromDebug() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DEBUG")
    On Error GoTo 0

    If ws Is Nothing Then
        Set M21_PromptIdsFromDebug = dict
        Exit Function
    End If

    Dim colPrompt As Long
    colPrompt = M21_FindColumn(ws, "Prompt ID")
    If colPrompt = 0 Then
        Set M21_PromptIdsFromDebug = dict
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colPrompt).End(xlUp).Row

    Dim r As Long
    Dim promptId As String
    For r = 2 To lastRow
        promptId = Trim$(CStr(ws.Cells(r, colPrompt).Value))
        If promptId <> "" Then
            If UCase$(promptId) <> "DEBUG" And UCase$(promptId) <> "SELFTEST" Then
                dict(promptId) = True
            End If
        End If
    Next r

    Set M21_PromptIdsFromDebug = dict
End Function

Private Function M21_BlocosCatalogoPorPromptId(ByVal promptId As String) As String
    On Error GoTo EH

    Dim parts() As String
    parts = Split(promptId, "/")

    Dim sheetName As String
    sheetName = Trim$(parts(0))
    If sheetName = "" Then
        M21_BlocosCatalogoPorPromptId = "[Prompt ID sem prefixo de folha: " & promptId & "]"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim idCell As Range
    Set idCell = ws.Columns(1).Find(What:=promptId, LookIn:=xlValues, LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If idCell Is Nothing Then
        M21_BlocosCatalogoPorPromptId = "[Prompt ID nÃ£o encontrado no catÃ¡logo " & sheetName & ": " & promptId & "]"
        Exit Function
    End If

    Dim firstRow As Long
    firstRow = idCell.Row

    Dim blockRange As Range
    Set blockRange = ws.Range(ws.Cells(firstRow, 1), ws.Cells(firstRow + 3, 11))

    M21_BlocosCatalogoPorPromptId = "--- CatÃ¡logo " & sheetName & " | Prompt " & promptId & " ---" & vbCrLf & _
                                    M21_RangeAsTsv(blockRange)
    Exit Function
EH:
    M21_BlocosCatalogoPorPromptId = "[Erro ao ler catÃ¡logo de " & promptId & ": " & Err.Description & "]"
End Function

Private Function M21_SheetAsTsv(ByVal sheetName As String) As String
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If lastCell Is Nothing Then
        M21_SheetAsTsv = ""
        Exit Function
    End If

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = lastCell.Row

    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastCol = lastCell.Column

    M21_SheetAsTsv = M21_RangeAsTsv(ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)))
    Exit Function
EH:
    M21_SheetAsTsv = ""
End Function

Private Function M21_RangeAsTsv(ByVal rng As Range) As String
    Dim data As Variant
    data = rng.Value

    Dim r As Long
    Dim c As Long
    Dim out As String

    If IsArray(data) Then
        For r = LBound(data, 1) To UBound(data, 1)
            For c = LBound(data, 2) To UBound(data, 2)
                out = out & Replace$(Replace$(CStr(data(r, c)), vbCrLf, " "), vbTab, " ")
                If c < UBound(data, 2) Then out = out & vbTab
            Next c
            If r < UBound(data, 1) Then out = out & vbCrLf
        Next r
    Else
        out = Replace$(Replace$(CStr(data), vbCrLf, " "), vbTab, " ")
    End If

    M21_RangeAsTsv = out
End Function

Private Function M21_FindColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If Trim$(UCase$(CStr(ws.Cells(1, c).Value))) = Trim$(UCase$(headerText)) Then
            M21_FindColumn = c
            Exit Function
        End If
    Next c

    M21_FindColumn = 0
End Function

Private Function M21_LeftSafe(ByVal text As String, ByVal maxLen As Long) As String
    If maxLen <= 0 Then
        M21_LeftSafe = ""
    ElseIf Len(text) <= maxLen Then
        M21_LeftSafe = text
    Else
        M21_LeftSafe = Left$(text, maxLen) & "..."
    End If
End Function
