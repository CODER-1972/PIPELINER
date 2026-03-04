Attribute VB_Name = "M25_GH_TreeCommit"
Option Explicit

' =============================================================================
' MÃ³dulo: M25_GH_TreeCommit
' PropÃ³sito:
' - Encapsular o fluxo GitHub Git Database (ref -> commit base -> blobs -> tree -> commit -> update ref).
' - Expor resultado estruturado (GH_Result) para o orquestrador decidir retry/bloqueio do passo.
' - Aplicar retry controlado em conflito de update de referÃªncia (PATCH /git/refs/heads/{branch}).
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | Implementa fluxo completo Tree Commit com retry em conflito
'   - Adiciona funÃ§Ã£o pÃºblica GH_TreeCommit_Execute com ciclo de retry desde leitura de HEAD.
'   - Regista GH_REF_CONFLICT, GH_RETRY_ATTEMPT e resultado final GH_DONE_OK/GH_DONE_FAIL no DEBUG.
'   - Introduz tipo GH_Result para retorno estruturado ao chamador.
'
' FunÃ§Ãµes e procedimentos:
' - GH_TreeCommit_Execute(...) As GH_Result
'   - Executa fluxo completo de commit GitHub e atualiza ref com retry em conflito.
' - GH_Result (Type)
'   - Contrato de retorno estruturado para decisÃ£o de repetiÃ§Ã£o/bloqueio.
' =============================================================================

Public Type GH_Result
    Success As Boolean
    HttpStatus As Long
    ErrorCode As String
    ErrorMessage As String
    BranchName As String
    HeadShaBefore As String
    BaseTreeSha As String
    NewTreeSha As String
    NewCommitSha As String
    Attempts As Long
    ConflictDetected As Boolean
    RetryEnabled As Boolean
    MaxRetries As Long
End Type

Private Const GH_API_BASE As String = "https://api.github.com"
Private Const GH_DEFAULT_MAX_RETRIES As Long = 3

Public Function GH_TreeCommit_Execute( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal branchName As String, _
    ByVal commitMessage As String, _
    ByVal entries As Collection, _
    ByVal githubToken As String, _
    Optional ByVal authorName As String = "", _
    Optional ByVal authorEmail As String = "") As GH_Result

    On Error GoTo EH

    Dim result As GH_Result
    result.BranchName = branchName
    result.RetryEnabled = GH_Config_GetBoolean("GH_RETRY_ON_CONFLICT", True)
    result.MaxRetries = GH_Config_GetLong("GH_MAX_RETRIES", GH_DEFAULT_MAX_RETRIES)
    If result.MaxRetries < 1 Then result.MaxRetries = 1

    Dim attempt As Long
    For attempt = 1 To result.MaxRetries
        result.Attempts = attempt
        If attempt > 1 Then
            DebugLog sevINFO, "GH_RETRY_ATTEMPT", "attempt=" & CStr(attempt) & "; max=" & CStr(result.MaxRetries) & "; branch=" & branchName
        End If

        Dim headSha As String
        Dim baseTreeSha As String
        Dim newTreeSha As String
        Dim newCommitSha As String

        headSha = GH_GetHeadCommitSha(owner, repo, branchName, githubToken, result)
        If headSha = "" Then GoTo Done
        result.HeadShaBefore = headSha

        baseTreeSha = GH_GetCommitTreeSha(owner, repo, headSha, githubToken, result)
        If baseTreeSha = "" Then GoTo Done
        result.BaseTreeSha = baseTreeSha

        newTreeSha = GH_CreateTree(owner, repo, baseTreeSha, entries, githubToken, result)
        If newTreeSha = "" Then GoTo Done
        result.NewTreeSha = newTreeSha

        newCommitSha = GH_CreateCommit(owner, repo, commitMessage, newTreeSha, headSha, githubToken, authorName, authorEmail, result)
        If newCommitSha = "" Then GoTo Done
        result.NewCommitSha = newCommitSha

        If GH_UpdateRef(owner, repo, branchName, newCommitSha, githubToken, result) Then
            result.Success = True
            result.ErrorCode = ""
            result.ErrorMessage = ""
            DebugLog sevINFO, "GH_DONE_OK", "branch=" & branchName & "; commit=" & newCommitSha & "; attempts=" & CStr(attempt)
            GH_TreeCommit_Execute = result
            Exit Function
        End If

        If Not result.ConflictDetected Then GoTo Done
        If Not result.RetryEnabled Then GoTo Done
        If attempt >= result.MaxRetries Then GoTo Done
    Next attempt

Done:
    result.Success = False
    If result.ErrorCode = "" Then result.ErrorCode = "GH_DONE_FAIL"
    If result.ErrorMessage = "" Then result.ErrorMessage = "Falha ao concluir tree commit no GitHub."
    DebugLog sevERRO, "GH_DONE_FAIL", "branch=" & branchName & "; code=" & result.ErrorCode & "; attempts=" & CStr(result.Attempts) & "; msg=" & result.ErrorMessage
    GH_TreeCommit_Execute = result
    Exit Function

EH:
    result.Success = False
    result.ErrorCode = "GH_UNEXPECTED"
    result.ErrorMessage = "Err " & CStr(Err.Number) & ": " & Err.Description
    DebugLog sevERRO, "GH_DONE_FAIL", "branch=" & branchName & "; code=" & result.ErrorCode & "; msg=" & result.ErrorMessage
    GH_TreeCommit_Execute = result
End Function

Private Function GH_GetHeadCommitSha( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal branchName As String, _
    ByVal githubToken As String, _
    ByRef result As GH_Result) As String

    Dim statusCode As Long, responseBody As String
    Dim endpoint As String
    endpoint = "/repos/" & owner & "/" & repo & "/git/refs/heads/" & branchName

    responseBody = GH_SendJsonRequest("GET", endpoint, "", githubToken, statusCode)
    result.HttpStatus = statusCode

    If statusCode < 200 Or statusCode >= 300 Then
        result.ErrorCode = "GH_REF_READ_FAIL"
        result.ErrorMessage = "Falha ao ler HEAD da branch. HTTP=" & CStr(statusCode)
        Exit Function
    End If

    GH_GetHeadCommitSha = GH_ExtractJsonField(responseBody, """sha"":")
    If GH_GetHeadCommitSha = "" Then
        result.ErrorCode = "GH_REF_HEAD_SHA_MISSING"
        result.ErrorMessage = "Resposta sem sha ao ler HEAD da branch."
    End If
End Function

Private Function GH_GetCommitTreeSha( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal commitSha As String, _
    ByVal githubToken As String, _
    ByRef result As GH_Result) As String

    Dim statusCode As Long, responseBody As String
    Dim endpoint As String
    endpoint = "/repos/" & owner & "/" & repo & "/git/commits/" & commitSha

    responseBody = GH_SendJsonRequest("GET", endpoint, "", githubToken, statusCode)
    result.HttpStatus = statusCode

    If statusCode < 200 Or statusCode >= 300 Then
        result.ErrorCode = "GH_COMMIT_BASE_READ_FAIL"
        result.ErrorMessage = "Falha ao ler commit base. HTTP=" & CStr(statusCode)
        Exit Function
    End If

    GH_GetCommitTreeSha = GH_ExtractNestedSha(responseBody, """tree""")
    If GH_GetCommitTreeSha = "" Then
        result.ErrorCode = "GH_BASE_TREE_MISSING"
        result.ErrorMessage = "Resposta sem tree.sha no commit base."
    End If
End Function

Private Function GH_CreateTree( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal baseTreeSha As String, _
    ByVal entries As Collection, _
    ByVal githubToken As String, _
    ByRef result As GH_Result) As String

    Dim blobsJson As String
    blobsJson = GH_CreateBlobsAndBuildTreeItems(owner, repo, entries, githubToken, result)
    If blobsJson = "" And entries.Count > 0 Then Exit Function

    Dim payload As String
    payload = "{""base_tree"":""" & GH_JsonEscape(baseTreeSha) & """,""tree"":" & IIf(entries.Count > 0, "[" & blobsJson & "]", "[]") & "}"

    Dim statusCode As Long, responseBody As String
    responseBody = GH_SendJsonRequest("POST", "/repos/" & owner & "/" & repo & "/git/trees", payload, githubToken, statusCode)
    result.HttpStatus = statusCode

    If statusCode < 200 Or statusCode >= 300 Then
        result.ErrorCode = "GH_TREE_CREATE_FAIL"
        result.ErrorMessage = "Falha ao criar tree. HTTP=" & CStr(statusCode)
        Exit Function
    End If

    GH_CreateTree = GH_ExtractJsonField(responseBody, """sha"":")
    If GH_CreateTree = "" Then
        result.ErrorCode = "GH_TREE_SHA_MISSING"
        result.ErrorMessage = "Resposta sem sha da tree criada."
    End If
End Function

Private Function GH_CreateBlobsAndBuildTreeItems( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal entries As Collection, _
    ByVal githubToken As String, _
    ByRef result As GH_Result) As String

    Dim i As Long
    Dim treeItems As String
    treeItems = ""

    For i = 1 To entries.Count
        Dim item As Object
        Set item = entries(i)

        Dim path As String
        Dim content As String
        Dim encoding As String
        Dim modeValue As String

        path = GH_Entry_Get(item, "path")
        content = GH_Entry_Get(item, "content")
        encoding = LCase$(Trim$(GH_Entry_Get(item, "encoding")))
        modeValue = GH_Entry_Get(item, "mode")

        If modeValue = "" Then modeValue = "100644"
        If encoding = "" Then encoding = "utf-8"

        Dim blobSha As String
        blobSha = GH_CreateBlob(owner, repo, content, encoding, githubToken, result)
        If blobSha = "" Then Exit Function

        If treeItems <> "" Then treeItems = treeItems & ","
        treeItems = treeItems & "{""path"":""" & GH_JsonEscape(path) & """,""mode"":""" & GH_JsonEscape(modeValue) & """,""type"":""blob"",""sha"":""" & GH_JsonEscape(blobSha) & """}"
    Next i

    GH_CreateBlobsAndBuildTreeItems = treeItems
End Function

Private Function GH_CreateBlob( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal content As String, _
    ByVal encoding As String, _
    ByVal githubToken As String, _
    ByRef result As GH_Result) As String

    Dim payload As String
    payload = "{""content"":""" & GH_JsonEscape(content) & """,""encoding"":""" & GH_JsonEscape(encoding) & """}"

    Dim statusCode As Long, responseBody As String
    responseBody = GH_SendJsonRequest("POST", "/repos/" & owner & "/" & repo & "/git/blobs", payload, githubToken, statusCode)
    result.HttpStatus = statusCode

    If statusCode < 200 Or statusCode >= 300 Then
        result.ErrorCode = "GH_BLOB_CREATE_FAIL"
        result.ErrorMessage = "Falha ao criar blob. HTTP=" & CStr(statusCode)
        Exit Function
    End If

    GH_CreateBlob = GH_ExtractJsonField(responseBody, """sha"":")
    If GH_CreateBlob = "" Then
        result.ErrorCode = "GH_BLOB_SHA_MISSING"
        result.ErrorMessage = "Resposta sem sha do blob criado."
    End If
End Function

Private Function GH_CreateCommit( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal commitMessage As String, _
    ByVal treeSha As String, _
    ByVal parentSha As String, _
    ByVal githubToken As String, _
    ByVal authorName As String, _
    ByVal authorEmail As String, _
    ByRef result As GH_Result) As String

    Dim payload As String
    payload = "{""message"":""" & GH_JsonEscape(commitMessage) & """,""tree"":""" & GH_JsonEscape(treeSha) & """,""parents"":[""" & GH_JsonEscape(parentSha) & """]"

    If Trim$(authorName) <> "" And Trim$(authorEmail) <> "" Then
        payload = payload & ",""author"":{""name"":""" & GH_JsonEscape(authorName) & """,""email"":""" & GH_JsonEscape(authorEmail) & """}"
    End If

    payload = payload & "}"

    Dim statusCode As Long, responseBody As String
    responseBody = GH_SendJsonRequest("POST", "/repos/" & owner & "/" & repo & "/git/commits", payload, githubToken, statusCode)
    result.HttpStatus = statusCode

    If statusCode < 200 Or statusCode >= 300 Then
        result.ErrorCode = "GH_COMMIT_CREATE_FAIL"
        result.ErrorMessage = "Falha ao criar commit. HTTP=" & CStr(statusCode)
        Exit Function
    End If

    GH_CreateCommit = GH_ExtractJsonField(responseBody, """sha"":")
    If GH_CreateCommit = "" Then
        result.ErrorCode = "GH_COMMIT_SHA_MISSING"
        result.ErrorMessage = "Resposta sem sha do commit criado."
    End If
End Function

Private Function GH_UpdateRef( _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal branchName As String, _
    ByVal commitSha As String, _
    ByVal githubToken As String, _
    ByRef result As GH_Result) As Boolean

    Dim payload As String
    payload = "{""sha"":""" & GH_JsonEscape(commitSha) & """,""force"":false}"

    Dim statusCode As Long, responseBody As String
    responseBody = GH_SendJsonRequest("PATCH", "/repos/" & owner & "/" & repo & "/git/refs/heads/" & branchName, payload, githubToken, statusCode)
    result.HttpStatus = statusCode

    If statusCode >= 200 And statusCode < 300 Then
        GH_UpdateRef = True
        result.ConflictDetected = False
        Exit Function
    End If

    result.ConflictDetected = GH_IsConflictStatus(statusCode, responseBody)
    If result.ConflictDetected Then
        DebugLog sevALERTA, "GH_REF_CONFLICT", "branch=" & branchName & "; status=" & CStr(statusCode) & "; retry_on_conflict=" & IIf(result.RetryEnabled, "TRUE", "FALSE")
        result.ErrorCode = "GH_REF_CONFLICT"
        result.ErrorMessage = "Conflito ao atualizar referÃªncia da branch. HTTP=" & CStr(statusCode)
    Else
        result.ErrorCode = "GH_REF_UPDATE_FAIL"
        result.ErrorMessage = "Falha ao atualizar referÃªncia da branch. HTTP=" & CStr(statusCode)
    End If
End Function

Private Function GH_SendJsonRequest( _
    ByVal method As String, _
    ByVal endpoint As String, _
    ByVal body As String, _
    ByVal githubToken As String, _
    ByRef statusCode As Long) As String

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open method, GH_API_BASE & endpoint, False
    http.SetRequestHeader "Accept", "application/vnd.github+json"
    http.SetRequestHeader "Authorization", "Bearer " & githubToken
    http.SetRequestHeader "X-GitHub-Api-Version", "2022-11-28"
    http.SetRequestHeader "User-Agent", "PIPELINER-VBA"

    If method = "POST" Or method = "PATCH" Or method = "PUT" Then
        http.SetRequestHeader "Content-Type", "application/json"
        http.Send body
    Else
        http.Send
    End If

    statusCode = CLng(http.Status)
    GH_SendJsonRequest = CStr(http.ResponseText)
    Exit Function

EH:
    statusCode = 0
    GH_SendJsonRequest = ""
End Function

Private Function GH_ExtractNestedSha(ByVal json As String, ByVal objectNameLiteral As String) As String
    Dim pObj As Long
    pObj = InStr(1, json, objectNameLiteral, vbTextCompare)
    If pObj = 0 Then Exit Function

    Dim pSha As Long
    pSha = InStr(pObj, json, """sha"":""", vbTextCompare)
    If pSha = 0 Then Exit Function

    GH_ExtractNestedSha = GH_ExtractJsonStringAt(json, pSha + Len("""sha"":"""))
End Function

Private Function GH_ExtractJsonField(ByVal json As String, ByVal token As String) As String
    Dim p As Long
    p = InStr(1, json, token, vbTextCompare)
    If p = 0 Then Exit Function

    GH_ExtractJsonField = GH_ExtractJsonStringAt(json, p + Len(token))
End Function

Private Function GH_ExtractJsonStringAt(ByVal json As String, ByVal startPos As Long) As String
    Dim firstQuote As Long
    firstQuote = InStr(startPos, json, """")
    If firstQuote = 0 Then Exit Function

    Dim i As Long
    For i = firstQuote + 1 To Len(json)
        If Mid$(json, i, 1) = """" And Mid$(json, i - 1, 1) <> "\" Then
            GH_ExtractJsonStringAt = Mid$(json, firstQuote + 1, i - firstQuote - 1)
            Exit Function
        End If
    Next i
End Function

Private Function GH_IsConflictStatus(ByVal statusCode As Long, ByVal responseBody As String) As Boolean
    If statusCode = 409 Then
        GH_IsConflictStatus = True
        Exit Function
    End If

    If statusCode = 422 Then
        If InStr(1, responseBody, "Reference update failed", vbTextCompare) > 0 _
            Or InStr(1, responseBody, "is at", vbTextCompare) > 0 Then
            GH_IsConflictStatus = True
        End If
    End If
End Function

Private Function GH_Config_GetBoolean(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    Dim raw As String
    raw = UCase$(Trim$(GH_Config_GetByKey(keyName, IIf(defaultValue, "TRUE", "FALSE"))))

    Select Case raw
        Case "1", "TRUE", "SIM", "YES", "Y", "ON"
            GH_Config_GetBoolean = True
        Case "0", "FALSE", "NAO", "NÃƒO", "NO", "N", "OFF"
            GH_Config_GetBoolean = False
        Case Else
            GH_Config_GetBoolean = defaultValue
    End Select
End Function

Private Function GH_Config_GetLong(ByVal keyName As String, ByVal defaultValue As Long) As Long
    Dim raw As String
    raw = Trim$(GH_Config_GetByKey(keyName, CStr(defaultValue)))
    If IsNumeric(raw) Then
        GH_Config_GetLong = CLng(raw)
    Else
        GH_Config_GetLong = defaultValue
    End If
End Function

Private Function GH_Config_GetByKey(ByVal keyName As String, ByVal defaultValue As String) As String
    On Error GoTo EH

    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets("Config")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row

    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), keyName, vbTextCompare) = 0 Then
            GH_Config_GetByKey = Trim$(CStr(ws.Cells(i, 2).Value))
            If GH_Config_GetByKey = "" Then GH_Config_GetByKey = defaultValue
            Exit Function
        End If
    Next i

EH:
    GH_Config_GetByKey = defaultValue
End Function

Private Function GH_Entry_Get(ByVal entry As Object, ByVal keyName As String) As String
    On Error GoTo EH

    If entry Is Nothing Then Exit Function

    GH_Entry_Get = CStr(entry(keyName))
    Exit Function

EH:
    GH_Entry_Get = ""
End Function

Private Function GH_JsonEscape(ByVal textIn As String) As String
    Dim s As String
    s = CStr(textIn)
    s = Replace(s, "\", "\\")
    s = Replace(s, Chr$(34), "\" & Chr$(34))
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    GH_JsonEscape = s
End Function
