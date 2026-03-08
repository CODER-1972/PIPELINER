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
' - 2026-03-08 | Codex | Reforca diagnostico por ficheiro e payload create/update
'   - Separa builders de payload para create e update, omitindo branch vazio e suportando author/committer opcionais.
'   - Acrescenta diagnostico file_kind/local_size_bytes no inicio de cada ficheiro para troubleshooting de 400 no contents_api.
' - 2026-03-08 | Codex | Validacao pre-PUT para create com diagnostico acionavel
'   - Adiciona GH_CONTENTS_CREATE_PAYLOAD_CHECK/REQUEST_READY e guardas para message/content/base64/sha/json.
'   - Enriquece GH_CONTENTS_CREATE_FAILED com endpoint/api_version/response_excerpt sem expor segredos.
' - 2026-03-08 | Codex | Melhora diagnostico de falhas HTTP 400/403 no contents_api
'   - Enriquece GH_CONTENTS_CREATE_FAILED/GH_CONTENTS_UPDATE_FAILED com resumo de erro HTTP e snippet da resposta.
'   - Adiciona sugestoes acionaveis para 400/401/403/404/422 no proprio detalhe de erro.
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
' - GH_ContentsApi_BuildHttpDiag(statusCode, httpErr, responseText, apiVersion, branch, hasMessage, hasContent, hasSha, endpointShort) As String
'   - Resume erro HTTP de forma curta/acionavel para DEBUG sem expor segredos.
' - GH_ContentsApi_ValidateCreatePayload(...) As Boolean
'   - Valida payload de create e regista resumo sanitizado pre-PUT no DEBUG.
' - GH_ContentsApi_BuildCreateBody/GH_ContentsApi_BuildUpdateBody(...)
'   - Montam JSON de create/update com regras explicitas por operacao e metadados opcionais de autoria.
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

        Call GH_LogInfo(0, pipelineNome, GH_EVT_FILE_BEGIN, "Processar ficheiro via contents_api.", "path=" & repoPath & " | idx=" & CStr(i) & " | " & GH_ContentsApi_FileDiag(CStr(item("content"))))

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

        Dim upsertStatus As Long
        Dim upsertResp As String
        Dim upsertErr As String
        Dim upsertUrl As String
        Dim upsertDiag As String

        upsertUrl = GH_ContentsApi_BuildContentUrl(cfg, repoPath)

        If existsOnRepo Then
            reqBody = GH_ContentsApi_BuildUpdateBody(cfg, CStr(item("content")), fileSha)
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
                upsertDiag = GH_ContentsApi_BuildHttpDiag(upsertStatus, upsertErr, upsertResp, GH_ContentsApi_NormalizeApiVersion(GH_Config_GetString(cfg, "api_version", "2022-11-28")), GH_Config_GetString(cfg, "branch", ""), True, True, True, "/repos/{owner}/{repo}/contents/{path}")
                Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_UPDATE_FAILED, "Falha no update via Contents API.", "path=" & repoPath & " | " & upsertDiag)
                Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Ficheiro falhou no update via contents_api.", "path=" & repoPath)
                If policy = "fail_fast" Then
                    errReason = "Update falhou: " & repoPath & " (" & upsertDiag & ")"
                    Exit Function
                End If
                GoTo ContinueLoop
            End If

            successCount = successCount + 1
            Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_UPDATE_OK, "Update concluido via Contents API.", "path=" & repoPath & " | sha=" & GH_ContentsApi_ShortSha(GH_TreeCommit_JsonPick(upsertResp, "sha")))
        Else
            reqBody = GH_ContentsApi_BuildCreateBody(cfg, CStr(item("content")))
            Dim apiVersion As String
            Dim endpointShort As String
            Dim createDiag As String
            apiVersion = GH_ContentsApi_NormalizeApiVersion(GH_Config_GetString(cfg, "api_version", "2022-11-28"))
            endpointShort = "/repos/{owner}/{repo}/contents/{path}"

            Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_CREATE_START, "Inicio de create via Contents API.", "path=" & repoPath)

            If Not GH_ContentsApi_ValidateCreatePayload(cfg, repoPath, reqBody, fileSha, apiVersion, endpointShort, pipelineNome, createDiag) Then
                failCount = failCount + 1
                Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Ficheiro falhou na validacao pre-PUT do create.", "path=" & repoPath)
                If policy = "fail_fast" Then
                    errReason = "Create invalido: " & repoPath & " (" & createDiag & ")"
                    Exit Function
                End If
                GoTo ContinueLoop
            End If

            Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_CREATE_REQUEST_READY, "Pedido create pronto para envio.", createDiag)

            If Not GH_HTTP_SendJson("PUT", upsertUrl, cfg, reqBody, upsertStatus, upsertResp, upsertErr, pipelineNome) Then
                failCount = failCount + 1
                upsertDiag = GH_ContentsApi_BuildHttpDiag(upsertStatus, upsertErr, upsertResp, apiVersion, GH_Config_GetString(cfg, "branch", ""), True, True, False, endpointShort)
                Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_CREATE_FAILED, "Falha no create via Contents API.", "path=" & repoPath & " | " & upsertDiag)
                Call GH_LogError(0, pipelineNome, GH_EVT_FILE_FAILED, "Ficheiro falhou no create via contents_api.", "path=" & repoPath)
                If policy = "fail_fast" Then
                    errReason = "Create falhou: " & repoPath & " (" & upsertDiag & ")"
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

Private Function GH_ContentsApi_BuildHttpDiag( _
    ByVal statusCode As Long, _
    ByVal httpErr As String, _
    ByVal responseText As String, _
    ByVal apiVersion As String, _
    ByVal branchName As String, _
    ByVal hasMessage As Boolean, _
    ByVal hasContent As Boolean, _
    ByVal hasSha As Boolean, _
    ByVal endpointShort As String) As String

    Dim diag As String
    diag = "http_status=" & CStr(statusCode)

    Dim msg As String
    msg = Trim$(GH_TreeCommit_JsonPick(responseText, "message"))
    If msg = "" Then msg = Trim$(httpErr)

    If msg <> "" Then
        If Len(msg) > 180 Then msg = Left$(msg, 180) & "..."
        diag = diag & " | err=" & Replace$(msg, "|", "/")
    End If

    diag = diag & " | method=PUT"
    diag = diag & " | endpoint_short=" & endpointShort
    diag = diag & " | api_version=" & apiVersion
    diag = diag & " | branch=" & IIf(Trim$(branchName) = "", "[default]", Replace$(Trim$(branchName), "|", "/"))
    diag = diag & " | has_message=" & IIf(hasMessage, "SIM", "NAO")
    diag = diag & " | has_content=" & IIf(hasContent, "SIM", "NAO")
    diag = diag & " | has_sha=" & IIf(hasSha, "SIM", "NAO")

    Dim responseExcerpt As String
    responseExcerpt = GH_ContentsApi_ShortTextPreview(responseText, 140)
    If responseExcerpt <> "" Then diag = diag & " | response_body_excerpt=" & responseExcerpt

    Select Case statusCode
        Case 400
            diag = diag & " | action=Validar GH_API_VERSION(YYYY-MM-DD), owner/repo/branch e payload."
        Case 401
            diag = diag & " | action=Token invalido/expirado; confirmar scope repo."
        Case 403
            diag = diag & " | action=Sem permissao ou limite; confirmar scope Content:write."
        Case 404
            diag = diag & " | action=Repo/branch/path inexistente ou sem acesso."
        Case 422
            diag = diag & " | action=Payload invalido (sha/branch/message/path); rever resposta."
    End Select

    GH_ContentsApi_BuildHttpDiag = diag
End Function

Private Function GH_ContentsApi_ValidateCreatePayload( _
    ByVal cfg As Object, _
    ByVal repoPath As String, _
    ByVal requestBody As String, _
    ByVal createSha As String, _
    ByVal apiVersion As String, _
    ByVal endpointShort As String, _
    ByVal pipelineNome As String, _
    ByRef diagOut As String) As Boolean

    Dim branchName As String
    Dim msg As String
    Dim b64 As String
    Dim hasMessage As Boolean
    Dim hasContent As Boolean
    Dim hasSha As Boolean

    branchName = Trim$(GH_Config_GetString(cfg, "branch", ""))
    msg = Trim$(GH_TreeCommit_JsonPick(requestBody, "message"))
    b64 = Trim$(GH_TreeCommit_JsonPick(requestBody, "content"))
    hasMessage = (msg <> "")
    hasContent = (b64 <> "")
    hasSha = (Trim$(createSha) <> "") Or (InStr(1, requestBody, """"sha"""", vbTextCompare) > 0)

    Dim b64Len As Long
    b64Len = Len(b64)

    diagOut = "path=" & repoPath & " | branch=" & IIf(branchName = "", "[default]", Replace$(branchName, "|", "/")) & _
              " | method=PUT | endpoint_short=" & endpointShort & _
              " | has_message=" & IIf(hasMessage, "SIM", "NAO") & _
              " | message_len=" & CStr(Len(msg)) & _
              " | has_content=" & IIf(hasContent, "SIM", "NAO") & _
              " | content_b64_len=" & CStr(b64Len) & _
              " | content_b64_mod4=" & CStr(b64Len Mod 4) & _
              " | content_b64_prefix=" & GH_ContentsApi_Base64PrefixMasked(b64) & _
              " | has_sha=" & IIf(hasSha, "SIM", "NAO") & _
              " | has_committer=" & IIf(InStr(1, requestBody, """"committer"""", vbTextCompare) > 0, "SIM", "NAO") & _
              " | has_author=" & IIf(InStr(1, requestBody, """"author"""", vbTextCompare) > 0, "SIM", "NAO") & _
              " | api_version=" & apiVersion & _
              " | json_len=" & CStr(Len(requestBody))

    Call GH_LogInfo(0, pipelineNome, GH_EVT_CONTENTS_CREATE_PAYLOAD_CHECK, "Validacao pre-PUT do payload de create.", diagOut)

    If Not hasMessage Then
        Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_CREATE_INVALID_MESSAGE, "Payload create sem message.", "path=" & repoPath & " | action=Preencher GH_COMMIT_MESSAGE_TEMPLATE.")
        Exit Function
    End If

    If Not hasContent Then
        Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_CREATE_INVALID_CONTENT, "Payload create sem content base64.", "path=" & repoPath & " | action=Validar conteudo local antes do upload.")
        Exit Function
    End If

    If Not GH_ContentsApi_IsValidBase64Shape(b64) Then
        Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_CREATE_INVALID_BASE64, "Payload create com base64 invalido.", "path=" & repoPath & " | action=Rever codificacao/base64 do conteudo.")
        Exit Function
    End If

    If hasSha Then
        Call GH_LogWarn(0, pipelineNome, GH_EVT_CONTENTS_CREATE_UNEXPECTED_SHA, "Create com sha inesperado; sha sera ignorado no create.", "path=" & repoPath)
    End If

    If Left$(Trim$(requestBody), 1) <> "{" Or Right$(Trim$(requestBody), 1) <> "}" Then
        Call GH_LogError(0, pipelineNome, GH_EVT_CONTENTS_CREATE_JSON_INVALID, "Payload create com JSON invalido.", "path=" & repoPath)
        Exit Function
    End If

    GH_ContentsApi_ValidateCreatePayload = True
End Function

Private Function GH_ContentsApi_IsValidBase64Shape(ByVal b64 As String) As Boolean
    Dim s As String
    s = Trim$(b64)
    If s = "" Then Exit Function

    If (Len(s) Mod 4) <> 0 Then Exit Function

    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=", ch, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i

    GH_ContentsApi_IsValidBase64Shape = True
End Function

Private Function GH_ContentsApi_Base64PrefixMasked(ByVal b64 As String) As String
    Dim s As String
    s = Trim$(b64)
    If s = "" Then
        GH_ContentsApi_Base64PrefixMasked = "[vazio]"
        Exit Function
    End If

    Dim pref As String
    pref = Left$(s, 10)
    GH_ContentsApi_Base64PrefixMasked = pref & "..."
End Function

Private Function GH_ContentsApi_ShortTextPreview(ByVal valueText As String, ByVal maxLen As Long) As String
    Dim s As String
    s = Trim$(Replace$(Replace$(valueText, vbCr, " "), vbLf, " "))
    s = Replace$(s, "|", "/")
    If s = "" Then Exit Function

    If Len(s) > maxLen Then s = Left$(s, maxLen) & "..."
    GH_ContentsApi_ShortTextPreview = s
End Function

Private Function GH_ContentsApi_NormalizeApiVersion(ByVal rawValue As String) As String
    Dim valueText As String
    valueText = Trim$(rawValue)

    If valueText = "" Then
        GH_ContentsApi_NormalizeApiVersion = "2022-11-28"
        Exit Function
    End If

    If valueText Like "####-##-##" Then
        GH_ContentsApi_NormalizeApiVersion = valueText
        Exit Function
    End If

    If valueText Like "##/##/####" Then
        GH_ContentsApi_NormalizeApiVersion = Right$(valueText, 4) & "-" & Mid$(valueText, 4, 2) & "-" & Left$(valueText, 2)
        Exit Function
    End If

    GH_ContentsApi_NormalizeApiVersion = "2022-11-28"
End Function

Private Function GH_ContentsApi_BuildCreateBody(ByVal cfg As Object, ByVal contentText As String) As String
    GH_ContentsApi_BuildCreateBody = GH_ContentsApi_BuildUpsertBody(cfg, contentText, "")
End Function

Private Function GH_ContentsApi_BuildUpdateBody(ByVal cfg As Object, ByVal contentText As String, ByVal existingSha As String) As String
    GH_ContentsApi_BuildUpdateBody = GH_ContentsApi_BuildUpsertBody(cfg, contentText, existingSha)
End Function

Private Function GH_ContentsApi_BuildUpsertBody(ByVal cfg As Object, ByVal contentText As String, ByVal existingSha As String) As String
    Dim commitMsg As String
    commitMsg = GH_Config_GetString(cfg, "commit_message_template", "PIPELINER run {{RUN_ID}}")
    commitMsg = Replace(commitMsg, "{{RUN_ID}}", Format$(Now, "yyyymmdd_hhnnss"))

    Dim b64 As String
    b64 = GH_Blob_Base64FromText(contentText)

    GH_ContentsApi_BuildUpsertBody = "{""message"":""" & GH_Blob_JsonEscape(commitMsg) & """,""content"":""" & GH_Blob_JsonEscape(b64) & """"

    Dim branchName As String
    branchName = Trim$(GH_Config_GetString(cfg, "branch"))
    If branchName <> "" Then
        GH_ContentsApi_BuildUpsertBody = GH_ContentsApi_BuildUpsertBody & ",""branch"":""" & GH_Blob_JsonEscape(branchName) & """"
    End If

    GH_ContentsApi_BuildUpsertBody = GH_ContentsApi_BuildUpsertBody & GH_ContentsApi_BuildAuthorCommitterJson(cfg)

    If Trim$(existingSha) <> "" Then
        GH_ContentsApi_BuildUpsertBody = GH_ContentsApi_BuildUpsertBody & ",""sha"":""" & GH_Blob_JsonEscape(existingSha) & """"
    End If

    GH_ContentsApi_BuildUpsertBody = GH_ContentsApi_BuildUpsertBody & "}"
End Function

Private Function GH_ContentsApi_BuildAuthorCommitterJson(ByVal cfg As Object) As String
    Dim authorName As String
    Dim authorEmail As String
    authorName = Trim$(GH_Config_GetString(cfg, "commit_author_name", ""))
    authorEmail = Trim$(GH_Config_GetString(cfg, "commit_author_email", ""))

    If authorName = "" Or authorEmail = "" Then Exit Function

    Dim personJson As String
    personJson = "{""name"":""" & GH_Blob_JsonEscape(authorName) & """,""email"":""" & GH_Blob_JsonEscape(authorEmail) & """}"
    GH_ContentsApi_BuildAuthorCommitterJson = ",""author"":" & personJson & ",""committer"":" & personJson
End Function

Private Function GH_ContentsApi_FileDiag(ByVal contentText As String) As String
    Dim fileKind As String
    If GH_ContentsApi_LooksBinary(contentText) Then
        fileKind = "binary"
    Else
        fileKind = "text"
    End If

    GH_ContentsApi_FileDiag = "file_kind=" & fileKind & " | local_size_bytes=" & CStr(LenB(contentText))
End Function

Private Function GH_ContentsApi_LooksBinary(ByVal contentText As String) As Boolean
    Dim i As Long
    Dim ch As Integer
    Dim suspicious As Long
    Dim total As Long
    total = Len(contentText)
    If total = 0 Then Exit Function

    For i = 1 To total
        ch = AscW(Mid$(contentText, i, 1))
        If ch = 0 Then
            GH_ContentsApi_LooksBinary = True
            Exit Function
        End If
        If ch < 32 And ch <> 9 And ch <> 10 And ch <> 13 Then suspicious = suspicious + 1
    Next i

    GH_ContentsApi_LooksBinary = (suspicious > (total \ 20))
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
