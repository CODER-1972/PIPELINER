Attribute VB_Name = "M27_GH_ContentsApi"
Option Explicit

' =============================================================================
' Modulo: M27_GH_ContentsApi
' Proposito:
' - Implementar upload GitHub via Contents API no runtime principal de debug export.
' - Processar ficheiros em serie (create/update) usando sha do ficheiro quando necessario.
' - Registar eventos canonicos de progresso/erro no DEBUG sem expor segredos.
'
' Atualizacoes:
' - 2026-03-08 | Codex | Reforca observabilidade por fase no contents_api
'   - Regista eventos de inicio GH_CONTENTS_CREATE_START/GH_CONTENTS_UPDATE_START e GH_FILE_FAILED.
'   - Valida path remoto vazio antes da chamada HTTP para evitar requests invalidos.
' - 2026-03-08 | Codex | Implementa modo contents_api com dispatch operacional
'   - Adiciona rotina GH_ContentsApi_UploadFiles para create/update serial por ficheiro.
'   - Implementa probe de existencia por path e captura de sha para updates seguros.
'   - Regista eventos GH_CONTENTS_* e sumario de lote com politica fail_fast/best_effort.
'
' Funcoes e procedimentos:
' - GH_ContentsApi_UploadFiles(cfg, files, pipelineNome, successCount, failCount, retryCount, errReason) As Boolean
'   - Executa lote serial via Contents API e devolve contadores/sucesso global.
' - GH_ContentsApi_ReadFileShaIfExists(cfg, repoPath, pipelineNome, existsOnRepo, fileSha, errReason) As Boolean
'   - Consulta metadados de um path no branch e devolve sha quando o ficheiro existe.
' =============================================================================

Public Function GH_ContentsApi_UploadFiles( _
    ByVal cfg As Object, _
    ByVal files As Collection, _
    ByVal pipelineNome As String, _
    ByRef successCount As Long, _
    ByRef failCount As Long, _
    ByRef retryCount As Long, _
    ByRef errReason As String) As Boolean

    On Error GoTo EH

    successCount = 0
    failCount = 0
    retryCount = 0
    errReason = ""

    Dim policy As String
    policy = LCase$(Trim$(GH_Config_GetString(cfg, "contents_batch_policy", "fail_fast")))
    If policy <> "best_effort" Then policy = "fail_fast"

    Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_BATCH_START, "Inicio do lote contents_api.", "files=" & CStr(files.Count) & " | policy=" & policy)

    Dim i As Long
    For i = 1 To files.Count
        Dim item As Object
        Set item = files(i)

        Dim repoPath As String
        repoPath = GH_ContentsApi_NormalizeRepoPath(CStr(item("path")))

        If repoPath = "" Then
            failCount = failCount + 1
            Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Path remoto invalido para contents_api.", "idx=" & CStr(i) & " | path_vazio")
            If policy = "fail_fast" Then
                errReason = "Path remoto vazio no item " & CStr(i)
                Exit Function
            End If
            GoTo ContinueLoop
        End If

        Call GH_LogInfo(0, pipelineNome, GH_EVT_FILE_BEGIN, "Processar ficheiro via contents_api.", "path=" & repoPath & " | idx=" & CStr(i))

        Dim existsOnRepo As Boolean
        Dim fileSha As String
        Dim probeReason As String

        If Not GH_ContentsApi_ReadFileShaIfExists(cfg, repoPath, pipelineNome, existsOnRepo, fileSha, probeReason) Then
            failCount = failCount + 1
            Call GH_LogError(0, pipelineNome, GH_EVT_FILE_PROBE_FAILED, "Falha no probe do ficheiro remoto.", "path=" & repoPath & " | " & probeReason)
            Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Ficheiro falhou no probe remoto.", "path=" & repoPath)
            If policy = "fail_fast" Then
                errReason = "Probe falhou: " & repoPath
                Exit Function
            End If
            GoTo ContinueLoop
        End If

        Dim reqBody As String
        reqBody = GH_ContentsApi_BuildUpsertBody(cfg, CStr(item("content")), fileSha)

        Dim upsertStatus As Long
        Dim upsertResp As String
        Dim upsertErr As String
        Dim upsertUrl As String

        upsertUrl = GH_ContentsApi_BuildContentUrl(cfg, repoPath)

        If existsOnRepo Then
            If fileSha = "" Then
                failCount = failCount + 1
                Call GH_LogError(0, pipelineNome, GH_EVT_FILE_SHA_MISSING_FOR_UPDATE, "Ficheiro existe sem sha para update.", "path=" & repoPath)
                Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Ficheiro falhou por sha ausente em update.", "path=" & repoPath)
                If policy = "fail_fast" Then
                    errReason = "sha em falta para update: " & repoPath
                    Exit Function
                End If
                GoTo ContinueLoop
            End If

            Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_UPDATE_START, "Inicio de update via Contents API.", "path=" & repoPath & " | sha=" & GH_ContentsApi_ShortSha(fileSha))
            If Not GH_HTTP_SendJson("PUT", upsertUrl, cfg, reqBody, upsertStatus, upsertResp, upsertErr, pipelineNome) Then
                failCount = failCount + 1
                Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_UPDATE_FAILED, "Falha no update via Contents API.", "path=" & repoPath & " | http_status=" & CStr(upsertStatus))
                Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Ficheiro falhou no update via contents_api.", "path=" & repoPath)
                If policy = "fail_fast" Then
                    errReason = "Update falhou: " & repoPath & " (status=" & CStr(upsertStatus) & ")"
                    Exit Function
                End If
                GoTo ContinueLoop
            End If

            successCount = successCount + 1
            Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_UPDATE_OK, "Update concluido via Contents API.", "path=" & repoPath & " | sha=" & GH_ContentsApi_ShortSha(GH_TreeCommit_JsonPick(upsertResp, "sha")))
        Else
            Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_CREATE_START, "Inicio de create via Contents API.", "path=" & repoPath)
            If Not GH_HTTP_SendJson("PUT", upsertUrl, cfg, reqBody, upsertStatus, upsertResp, upsertErr, pipelineNome) Then
                failCount = failCount + 1
                Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_CREATE_FAILED, "Falha no create via Contents API.", "path=" & repoPath & " | http_status=" & CStr(upsertStatus))
                Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Ficheiro falhou no create via contents_api.", "path=" & repoPath)
                If policy = "fail_fast" Then
                    errReason = "Create falhou: " & repoPath & " (status=" & CStr(upsertStatus) & ")"
                    Exit Function
                End If
                GoTo ContinueLoop
            End If

            successCount = successCount + 1
            Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_CREATE_OK, "Create concluido via Contents API.", "path=" & repoPath & " | sha=" & GH_ContentsApi_ShortSha(GH_TreeCommit_JsonPick(upsertResp, "sha")))
        End If

        Call GH_LogInfo(0, pipelineNome, GH_EVT_FILE_DONE, "Ficheiro concluido via contents_api.", "path=" & repoPath)

ContinueLoop:
    Next i

    Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_BATCH_DONE, "Lote contents_api terminado.", "success=" & CStr(successCount) & " | fail=" & CStr(failCount) & " | policy=" & policy)

    GH_ContentsApi_UploadFiles = (failCount = 0)
    If Not GH_ContentsApi_UploadFiles Then
        errReason = "Falhas no contents_api: " & CStr(failCount)
    End If
    Exit Function

EH:
    errReason = "Erro inesperado GH_ContentsApi_UploadFiles: " & Err.Description
End Function

Private Function GH_ContentsApi_ReadFileShaIfExists( _
    ByVal cfg As Object, _
    ByVal repoPath As String, _
    ByVal pipelineNome As String, _
    ByRef existsOnRepo As Boolean, _
    ByRef fileSha As String, _
    ByRef errReason As String) As Boolean

    existsOnRepo = False
    fileSha = ""
    errReason = ""

    Call GH_LogInfo(0, pipelineNome, GH_EVT_FILE_PROBE_START, "Probe do ficheiro remoto.", "path=" & repoPath)

    Dim statusCode As Long
    Dim responseText As String
    Dim httpErr As String
    Dim url As String

    url = GH_ContentsApi_BuildContentUrl(cfg, repoPath) & "?ref=" & GH_ContentsApi_UrlEncode(GH_Config_GetString(cfg, "branch"))

    If GH_HTTP_SendJson("GET", url, cfg, "", statusCode, responseText, httpErr, pipelineNome) Then
        existsOnRepo = True
        fileSha = GH_TreeCommit_JsonPick(responseText, "sha")
        Call GH_LogInfo(0, pipelineNome, GH_EVT_FILE_EXISTS_YES, "Ficheiro remoto existente.", "path=" & repoPath)
        If fileSha <> "" Then
            Call GH_LogInfo(0, pipelineNome, GH_EVT_FILE_SHA_OK, "SHA do ficheiro remoto obtido.", "path=" & repoPath & " | sha=" & GH_ContentsApi_ShortSha(fileSha))
        End If
        GH_ContentsApi_ReadFileShaIfExists = True
        Exit Function
    End If

    If statusCode = 404 Then
        existsOnRepo = False
        Call GH_LogWarn(0, pipelineNome, GH_EVT_FILE_EXISTS_NO, "Ficheiro remoto ainda nao existe.", "path=" & repoPath)
        GH_ContentsApi_ReadFileShaIfExists = True
        Exit Function
    End If

    errReason = "status=" & CStr(statusCode)
    If httpErr <> "" Then errReason = errReason & " | " & httpErr
End Function

Private Function GH_ContentsApi_BuildUpsertBody(ByVal cfg As Object, ByVal contentText As String, ByVal existingSha As String) As String
    Dim commitMsg As String
    commitMsg = GH_Config_GetString(cfg, "commit_message_template", "PIPELINER run {{RUN_ID}}")
    commitMsg = Replace(commitMsg, "{{RUN_ID}}", Format$(Now, "yyyymmdd_hhnnss"))

    Dim b64 As String
    b64 = GH_Blob_Base64FromText(contentText)

    GH_ContentsApi_BuildUpsertBody = "{""message"":""" & GH_Blob_JsonEscape(commitMsg) & """,""content"":""" & GH_Blob_JsonEscape(b64) & """,""branch"":""" & GH_Blob_JsonEscape(GH_Config_GetString(cfg, "branch")) & """"

    If Trim$(existingSha) <> "" Then
        GH_ContentsApi_BuildUpsertBody = GH_ContentsApi_BuildUpsertBody & ",""sha"":""" & GH_Blob_JsonEscape(existingSha) & """"
    End If

    GH_ContentsApi_BuildUpsertBody = GH_ContentsApi_BuildUpsertBody & "}"
End Function

Private Function GH_ContentsApi_BuildContentUrl(ByVal cfg As Object, ByVal repoPath As String) As String
    GH_ContentsApi_BuildContentUrl = GH_ContentsApi_ApiBase(cfg) & "/repos/" & _
                                     GH_ContentsApi_UrlEncode(GH_Config_GetString(cfg, "owner")) & "/" & _
                                     GH_ContentsApi_UrlEncode(GH_Config_GetString(cfg, "repo")) & "/contents/" & _
                                     GH_ContentsApi_EncodePath(repoPath)
End Function

Private Function GH_ContentsApi_ApiBase(ByVal cfg As Object) As String
    Dim apiBase As String
    apiBase = GH_Config_GetString(cfg, "api_base", "https://api.github.com")
    If Right$(apiBase, 1) = "/" Then apiBase = Left$(apiBase, Len(apiBase) - 1)
    GH_ContentsApi_ApiBase = apiBase
End Function

Private Function GH_ContentsApi_NormalizeRepoPath(ByVal repoPath As String) As String
    Dim outPath As String
    outPath = Replace$(Trim$(repoPath), "\", "/")

    Do While Left$(outPath, 1) = "/"
        outPath = Mid$(outPath, 2)
    Loop

    GH_ContentsApi_NormalizeRepoPath = outPath
End Function

Private Function GH_ContentsApi_EncodePath(ByVal repoPath As String) As String
    Dim parts() As String
    parts = Split(repoPath, "/")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = GH_ContentsApi_UrlEncode(parts(i))
    Next i

    GH_ContentsApi_EncodePath = Join(parts, "/")
End Function

Private Function GH_ContentsApi_UrlEncode(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        code = AscW(ch)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                out = out & ch
            Case Else
                out = out & "%" & Right$("0" & Hex$(code And &HFF), 2)
        End Select
    Next i

    GH_ContentsApi_UrlEncode = out
End Function

Private Function GH_ContentsApi_ShortSha(ByVal shaText As String) As String
    Dim trimmed As String
    trimmed = Trim$(shaText)
    If trimmed = "" Then Exit Function

    If Len(trimmed) > 10 Then
        GH_ContentsApi_ShortSha = Left$(trimmed, 10)
    Else
        GH_ContentsApi_ShortSha = trimmed
    End If
End Function
