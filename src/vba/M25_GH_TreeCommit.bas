Attribute VB_Name = "M25_GH_TreeCommit"
Option Explicit

' =============================================================================
' Modulo: M25_GH_TreeCommit
' Proposito:
' - Orquestrar fluxo Git Data API: ref -> commit base -> blobs -> tree -> commit -> update ref.
' - Encapsular construcao de endpoints GitHub (git/ref, git/commits, git/trees, git/blobs).
' - Reutilizar HTTP/Blob/Logger mantendo M21 como facade de compatibilidade.
'
' Atualizacoes:
' - 2026-03-08 | Codex | Retry em conflito 409 no update de ref
'   - Implementa loop de retentativa com GH_RETRY_ON_CONFLICT e GH_MAX_RETRIES no update da branch.
'   - Regista eventos GH_REF_FETCH_START, GH_REF_CONFLICT_409 e GH_RETRY_SCHEDULED para troubleshooting.
' - 2026-03-07 | Codex | Corrige sintaxe JSON em payloads Git Data API
'   - Corrige escaping de aspas nas chaves `tree` e `parents` em requests de create tree/commit.
'   - Elimina `Compile error: Syntax error` causado por literais JSON malformadas.
' - 2026-03-04 | Codex | Refactor do fluxo tree/commit para modulo dedicado
'   - Move pipeline completo de commit para GH_TreeCommit_CommitFiles.
'   - Adiciona helpers publicos para endpoint de blobs e URL web da pasta do run.
'   - Centraliza parsing JSON simples (sha/tree.sha) para responses GitHub.
'
' Funcoes e procedimentos:
' - GH_TreeCommit_CommitFiles(cfg, files, pipelineNome, commitSha, errReason) As Boolean
'   - Executa fluxo completo de commit e atualiza a branch alvo.
' - GH_TreeCommit_GitBlobUrl(cfg As Object) As String
'   - Devolve endpoint /repos/{owner}/{repo}/git/blobs.
' - GH_TreeCommit_BuildWebFolderUrl(cfg, folderPath) As String
'   - Devolve URL web da pasta publicada no branch configurado.
' - GH_TreeCommit_JsonPick(body, keyName) As String
'   - Extrai valor string de chave JSON simples via regex.
' =============================================================================

Public Function GH_TreeCommit_CommitFiles( _
    ByVal cfg As Object, _
    ByVal files As Collection, _
    ByVal pipelineNome As String, _
    ByRef commitSha As String, _
    ByRef errReason As String, _
    Optional ByRef retryCount As Long = 0) As Boolean

    On Error GoTo EH

    commitSha = ""
    errReason = ""

    If files Is Nothing Or files.Count = 0 Then
        errReason = "Sem ficheiros para exportar"
        Exit Function
    End If

    Dim maxFiles As Long
    maxFiles = GH_Config_GetLong(cfg, "max_files", 200)
    If files.Count > maxFiles Then
        errReason = "Numero de ficheiros excede GH_MAX_FILES"
        Call GH_LogWarn(0, pipelineNome, GH_EVT_MAX_FILES, "files_count=" & CStr(files.Count), "GH_MAX_FILES=" & CStr(maxFiles))
        Exit Function
    End If

    retryCount = 0

    Dim allowRetry As Boolean
    allowRetry = GH_Config_GetBoolean(cfg, "retry_on_conflict", True)

    Dim maxRetries As Long
    maxRetries = GH_Config_GetLong(cfg, "max_retries", 3)
    If maxRetries < 0 Then maxRetries = 0

RetryFromHead:
    Dim headSha As String
    Dim baseTreeSha As String
    If Not GH_TreeCommit_LoadHeadAndBaseTree(cfg, pipelineNome, headSha, baseTreeSha, errReason) Then Exit Function

    Dim treeItems As String
    Dim i As Long
    For i = 1 To files.Count
        Dim f As Object
        Set f = files(i)

        Dim blobSha As String
        If Not GH_Blob_Create(cfg, CStr(f("path")), CStr(f("content")), pipelineNome, blobSha, errReason) Then Exit Function

        If treeItems <> "" Then treeItems = treeItems & ","
        treeItems = treeItems & "{""path"":""" & GH_Blob_JsonEscape(CStr(f("path")) ) & """,""mode"":""100644"",""type"":""blob"",""sha"":""" & blobSha & """}"
    Next i

    Dim treeSha As String
    If Not GH_TreeCommit_CreateTree(cfg, pipelineNome, baseTreeSha, treeItems, treeSha, errReason) Then Exit Function

    If Not GH_TreeCommit_CreateCommit(cfg, pipelineNome, headSha, treeSha, commitSha, errReason) Then Exit Function

    Dim updateStatus As Long
    If Not GH_TreeCommit_UpdateRef(cfg, pipelineNome, commitSha, errReason, updateStatus) Then
        If updateStatus = 409 And allowRetry And retryCount < maxRetries Then
            retryCount = retryCount + 1
            Call GH_LogWarn(0, pipelineNome, "GH_REF_CONFLICT_409", "Conflito 409 ao atualizar ref.", "retry=" & CStr(retryCount) & "/" & CStr(maxRetries))
            Call GH_LogInfo(0, pipelineNome, "GH_RETRY_SCHEDULED", "Nova tentativa apos conflito de ref.", "retry=" & CStr(retryCount))
            GoTo RetryFromHead
        End If
        Exit Function
    End If

    GH_TreeCommit_CommitFiles = True
    Exit Function

EH:
    errReason = "Erro inesperado GH_TreeCommit_CommitFiles: " & Err.Description
End Function

Public Function GH_TreeCommit_GitBlobUrl(ByVal cfg As Object) As String
    GH_TreeCommit_GitBlobUrl = GH_TreeCommit_ApiBase(cfg) & "/repos/" & _
                               GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "owner")) & "/" & _
                               GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "repo")) & "/git/blobs"
End Function

Public Function GH_TreeCommit_BuildWebFolderUrl(ByVal cfg As Object, ByVal folderPath As String) As String
    GH_TreeCommit_BuildWebFolderUrl = "https://github.com/" & GH_Config_GetString(cfg, "owner") & "/" & _
                                      GH_Config_GetString(cfg, "repo") & "/tree/" & _
                                      GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "branch")) & "/" & _
                                      GH_TreeCommit_EncodePath(folderPath)
End Function

Public Function GH_TreeCommit_JsonPick(ByVal body As String, ByVal keyName As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """" & keyName & """" & "\s*:\s*""([^""]+)"""
    If re.Test(body) Then GH_TreeCommit_JsonPick = re.Execute(body)(0).SubMatches(0)
End Function

Private Function GH_TreeCommit_JsonPickTreeSha(ByVal body As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """tree""\s*:\s*\{\s*""sha""\s*:\s*""([^""]+)"""
    If re.Test(body) Then GH_TreeCommit_JsonPickTreeSha = re.Execute(body)(0).SubMatches(0)
End Function

Private Function GH_TreeCommit_LoadHeadAndBaseTree( _
    ByVal cfg As Object, _
    ByVal pipelineNome As String, _
    ByRef headSha As String, _
    ByRef baseTreeSha As String, _
    ByRef errReason As String) As Boolean

    Dim headRefUrl As String
    headRefUrl = GH_TreeCommit_ApiBase(cfg) & "/repos/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "owner")) & _
                 "/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "repo")) & "/git/ref/heads/" & _
                 GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "branch"))

    Dim statusCode As Long
    Dim responseText As String
    Dim httpErr As String

    Call GH_LogInfo(0, pipelineNome, "GH_REF_FETCH_START", "A obter HEAD da branch.", "branch=" & GH_Config_GetString(cfg, "branch"))

    If Not GH_HTTP_SendJson("GET", headRefUrl, cfg, "", statusCode, responseText, httpErr, pipelineNome) Then
        errReason = "Falha a obter HEAD ref"
        Exit Function
    End If

    headSha = GH_TreeCommit_JsonPick(responseText, "sha")
    If headSha = "" Then
        errReason = "HEAD sem sha"
        Exit Function
    End If
    Call GH_LogInfo(0, pipelineNome, GH_EVT_REF_OK, "HEAD obtido", "sha=" & Left$(headSha, 10))

    Dim commitUrl As String
    commitUrl = GH_TreeCommit_ApiBase(cfg) & "/repos/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "owner")) & _
                "/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "repo")) & "/git/commits/" & headSha

    If Not GH_HTTP_SendJson("GET", commitUrl, cfg, "", statusCode, responseText, httpErr, pipelineNome) Then
        errReason = "Falha a obter commit base"
        Exit Function
    End If

    baseTreeSha = GH_TreeCommit_JsonPickTreeSha(responseText)
    If baseTreeSha = "" Then baseTreeSha = GH_TreeCommit_JsonPick(responseText, "sha")
    If baseTreeSha = "" Then
        errReason = "Commit base sem tree sha"
        Exit Function
    End If

    Call GH_LogInfo(0, pipelineNome, GH_EVT_BASE_TREE_OK, "Base tree resolvida", "tree_sha=" & Left$(baseTreeSha, 10))
    GH_TreeCommit_LoadHeadAndBaseTree = True
End Function

Private Function GH_TreeCommit_CreateTree( _
    ByVal cfg As Object, _
    ByVal pipelineNome As String, _
    ByVal baseTreeSha As String, _
    ByVal treeItems As String, _
    ByRef treeSha As String, _
    ByRef errReason As String) As Boolean

    Dim req As String
    req = "{""base_tree"":""" & baseTreeSha & """,""tree"": [" & treeItems & "]}"

    Dim statusCode As Long
    Dim responseText As String
    Dim httpErr As String
    Dim url As String

    url = GH_TreeCommit_ApiBase(cfg) & "/repos/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "owner")) & _
          "/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "repo")) & "/git/trees"

    If Not GH_HTTP_SendJson("POST", url, cfg, req, statusCode, responseText, httpErr, pipelineNome) Then
        errReason = "Falha a criar tree"
        Exit Function
    End If

    treeSha = GH_TreeCommit_JsonPick(responseText, "sha")
    If treeSha = "" Then
        errReason = "Resposta de tree sem sha"
        Exit Function
    End If

    Call GH_LogInfo(0, pipelineNome, GH_EVT_TREE_CREATED, "Tree criada", "sha=" & Left$(treeSha, 10))
    GH_TreeCommit_CreateTree = True
End Function

Private Function GH_TreeCommit_CreateCommit( _
    ByVal cfg As Object, _
    ByVal pipelineNome As String, _
    ByVal headSha As String, _
    ByVal treeSha As String, _
    ByRef commitSha As String, _
    ByRef errReason As String) As Boolean

    Dim commitMsg As String
    commitMsg = GH_Config_GetString(cfg, "commit_message_template", "PIPELINER run {{RUN_ID}}")
    commitMsg = Replace(commitMsg, "{{RUN_ID}}", Format$(Now, "yyyymmdd_hhnnss"))

    Dim req As String
    req = "{""message"":""" & GH_Blob_JsonEscape(commitMsg) & """,""tree"":""" & treeSha & """,""parents"": [""" & headSha & """]}"

    Dim statusCode As Long
    Dim responseText As String
    Dim httpErr As String
    Dim url As String

    url = GH_TreeCommit_ApiBase(cfg) & "/repos/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "owner")) & _
          "/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "repo")) & "/git/commits"

    If Not GH_HTTP_SendJson("POST", url, cfg, req, statusCode, responseText, httpErr, pipelineNome) Then
        errReason = "Falha a criar commit"
        Exit Function
    End If

    commitSha = GH_TreeCommit_JsonPick(responseText, "sha")
    If commitSha = "" Then
        errReason = "Resposta de commit sem sha"
        Exit Function
    End If

    Call GH_LogInfo(0, pipelineNome, GH_EVT_COMMIT_CREATED, "Commit criado", "sha=" & Left$(commitSha, 10))
    GH_TreeCommit_CreateCommit = True
End Function

Private Function GH_TreeCommit_UpdateRef( _
    ByVal cfg As Object, _
    ByVal pipelineNome As String, _
    ByVal commitSha As String, _
    ByRef errReason As String, _
    Optional ByRef statusCodeOut As Long = 0) As Boolean

    Dim req As String
    statusCodeOut = 0

    req = "{""sha"":""" & commitSha & """,""force"":" & LCase$(CStr(GH_Config_GetBoolean(cfg, "force_update", False))) & "}"

    Dim statusCode As Long
    Dim responseText As String
    Dim httpErr As String
    Dim url As String

    url = GH_TreeCommit_ApiBase(cfg) & "/repos/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "owner")) & _
          "/" & GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "repo")) & "/git/ref/heads/" & _
          GH_TreeCommit_UrlEncode(GH_Config_GetString(cfg, "branch"))

    If Not GH_HTTP_SendJson("PATCH", url, cfg, req, statusCode, responseText, httpErr, pipelineNome) Then
        statusCodeOut = statusCode
        errReason = "Falha a atualizar ref (status=" & CStr(statusCode) & ")"
        Exit Function
    End If

    Call GH_LogInfo(0, pipelineNome, GH_EVT_REF_UPDATED, "Ref atualizada", "sha=" & Left$(commitSha, 10))
    GH_TreeCommit_UpdateRef = True
End Function

Private Function GH_TreeCommit_ApiBase(ByVal cfg As Object) As String
    Dim apiBase As String
    apiBase = GH_Config_GetString(cfg, "api_base", "https://api.github.com")
    If Right$(apiBase, 1) = "/" Then apiBase = Left$(apiBase, Len(apiBase) - 1)
    GH_TreeCommit_ApiBase = apiBase
End Function

Private Function GH_TreeCommit_EncodePath(ByVal repoPath As String) As String
    Dim parts() As String
    parts = Split(repoPath, "/")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = GH_TreeCommit_UrlEncode(parts(i))
    Next i

    GH_TreeCommit_EncodePath = Join(parts, "/")
End Function

Private Function GH_TreeCommit_UrlEncode(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                out = out & ch
            Case Else
                out = out & "%" & Right$("0" & Hex$(code And &HFF), 2)
        End Select
    Next i

    GH_TreeCommit_UrlEncode = out
End Function
