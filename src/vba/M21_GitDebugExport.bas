Attribute VB_Name = "M21_GitDebugExport"
Option Explicit

' =============================================================================
' Modulo: M21_GitDebugExport
' Proposito:
' - Exportar artefactos de DEBUG/Seguimento no fim da execucao da pipeline.
' - Publicar ficheiros no GitHub via Git Data API (blobs -> tree -> commit -> update ref).
' - Registar o link da pasta remota em Seguimento/HISTORICO na coluna GIT_DEBUG.
'
' Atualizacoes:
' - 2026-03-04 | Codex | Macro de instalacao guiada dos parametros GH_* no Config
'   - Adiciona rotina para criar/atualizar chaves GH_* com default, explicacao pedagogica e valores possiveis.
'   - Mantem retrocompatibilidade (nao sobrescreve valores existentes por defeito).
' - 2026-03-04 | Codex | Implementa auto-upload Git debug por pipeline
'   - Le parametros GH_* da folha Config com defaults internos (retrocompativel).
'   - Ativa apenas quando auto-guardar contem "sim, todos" ou "debug".
'   - Gera 4 artefactos: DEBUG.csv, catalogo_prompts_executadas.csv, Seguimento.csv, painel_pipeline.txt.
'
' Funcoes e procedimentos:
' - PipelineGitDebug_ExportIfEnabled(...): entry point chamado no fim da pipeline.
' - GitDebug_Config_InstalarParametros(...): instala parametros GH_* no Config com ajuda para leigos.
' =============================================================================

Private Const SHEET_CONFIG As String = "Config"
Private Const SHEET_DEBUG As String = "DEBUG"
Private Const SHEET_SEGUIMENTO As String = "Seguimento"
Private Const SHEET_HIST As String = "HISTÓRICO"

Public Sub PipelineGitDebug_ExportIfEnabled(ByVal pipelineIndex As Long, ByVal pipelineNome As String, ByVal painelAutoSave As String)
    On Error GoTo EH

    If Not GitDebug_IsEnabled(painelAutoSave) Then Exit Sub

    Dim token As String
    token = GitCfg_ResolveToken()
    If Trim$(token) = "" Then
        Call Debug_Registar(0, pipelineNome, "ALERTA", "", "GH_TOKEN", "Auto-upload debug ativo mas token GitHub ausente.", "Defina GH_TOKEN_ENV ou GH_TOKEN_CONFIG na folha Config.")
        Exit Sub
    End If

    Dim owner As String: owner = GitCfg_Get("GH_OWNER", "")
    Dim repo As String: repo = GitCfg_Get("GH_REPO", "")
    Dim branch As String: branch = GitCfg_Get("GH_BRANCH", "main")
    If owner = "" Or repo = "" Then
        Call Debug_Registar(0, pipelineNome, "ALERTA", "", "GH_CONFIG", "GH_OWNER/GH_REPO em falta.", "Preencha GH_OWNER e GH_REPO no Config.")
        Exit Sub
    End If

    Dim ghFolder As String
    ghFolder = GitDebug_BuildRunFolder(pipelineNome)

    Dim files As Collection
    Set files = GitDebug_BuildFilesForUpload(pipelineIndex, pipelineNome, ghFolder)
    If files Is Nothing Or files.Count = 0 Then Exit Sub

    Dim commitSha As String
    commitSha = GitData_CommitFiles(owner, repo, branch, token, files, pipelineNome)
    If Trim$(commitSha) = "" Then Exit Sub

    Dim webUrl As String
    webUrl = "https://github.com/" & owner & "/" & repo & "/tree/" & Git_UrlEncodeSegment(branch) & "/" & Git_UrlEncodePath(ghFolder)

    Call GitDebug_WriteLinkToSeguimento(pipelineNome, webUrl)
    Call GitDebug_WriteLinkToHistorico(pipelineNome, webUrl)

    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_REF_UPDATED", "Debug export publicado no GitHub.", webUrl)
    Exit Sub
EH:
    Call Debug_Registar(0, pipelineNome, "ERRO", "", "GH_UPLOAD", "Falha no auto-upload de debug: " & Err.Description, "Validar parâmetros GH_* e conectividade com api.github.com.")
End Sub

Private Function GitDebug_IsEnabled(ByVal painelAutoSave As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(painelAutoSave))
    GitDebug_IsEnabled = (InStr(1, s, "sim, todos", vbTextCompare) > 0) Or (InStr(1, s, "debug", vbTextCompare) > 0)
End Function

Private Function GitDebug_BuildRunFolder(ByVal pipelineNome As String) As String
    GitDebug_BuildRunFolder = Format$(Now, "yyyy-mmm-dd") & "-" & Format$(Now, "hhnn") & " - " & Git_SanitizePathPart(pipelineNome)
End Function

Private Function GitDebug_BuildFilesForUpload(ByVal pipelineIndex As Long, ByVal pipelineNome As String, ByVal ghFolder As String) As Collection
    On Error GoTo EH

    Dim cfgBase As String
    cfgBase = Trim$(GitCfg_Get("GH_BASE_PATH", "pipeliner_runs"))
    If cfgBase <> "" Then ghFolder = cfgBase & "/" & ghFolder

    Dim wsDebug As Worksheet: Set wsDebug = ThisWorkbook.Worksheets(SHEET_DEBUG)
    Dim wsSeg As Worksheet: Set wsSeg = ThisWorkbook.Worksheets(SHEET_SEGUIMENTO)

    Dim csvDebug As String
    csvDebug = SheetToCsv(wsDebug)

    Dim csvSeg As String
    csvSeg = SheetToCsv(wsSeg)

    Dim csvCatalogo As String
    csvCatalogo = BuildExecutedCatalogCsv(wsSeg, pipelineNome)

    Dim txtPainel As String
    txtPainel = BuildPainelPipelineInfo(pipelineIndex)

    Dim files As New Collection
    files.Add GitFileItem(ghFolder & "/DEBUG.csv", csvDebug)
    files.Add GitFileItem(ghFolder & "/catalogo_prompts_executadas.csv", csvCatalogo)
    files.Add GitFileItem(ghFolder & "/Seguimento.csv", csvSeg)
    files.Add GitFileItem(ghFolder & "/painel_pipeline.txt", txtPainel)

    Set GitDebug_BuildFilesForUpload = files
    Exit Function
EH:
    Set GitDebug_BuildFilesForUpload = Nothing
End Function

Private Function BuildPainelPipelineInfo(ByVal pipelineIndex As Long) As String
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("PAINEL")
    Dim colIniciar As Long: colIniciar = 2 + (pipelineIndex - 1) * 2
    Dim colReg As Long: colReg = colIniciar + 1

    Dim txt As String
    txt = "Pipeline Index: " & CStr(pipelineIndex) & vbCrLf
    txt = txt & "Nome: " & CStr(ws.Cells(1, colIniciar).Value) & vbCrLf
    txt = txt & "INPUT Folder: " & CStr(ws.Cells(2, colIniciar).Value) & vbCrLf
    txt = txt & "OUTPUT Folder: " & CStr(ws.Cells(3, colIniciar).Value) & vbCrLf
    txt = txt & "Auto-guardar ficheiros: " & CStr(ws.Cells(4, colIniciar).Value) & vbCrLf
    txt = txt & "Max Steps: " & CStr(ws.Cells(5, colIniciar).Value) & vbCrLf
    txt = txt & "Max Repetitions: " & CStr(ws.Cells(6, colIniciar).Value) & vbCrLf
    txt = txt & "Primeiros IDs (INICIAR):" & vbCrLf

    Dim r As Long
    For r = 10 To 20
        txt = txt & "- " & CStr(ws.Cells(r, colIniciar).Value) & vbCrLf
    Next r

    txt = txt & "Primeiros IDs (REGISTAR):" & vbCrLf
    For r = 10 To 20
        txt = txt & "- " & CStr(ws.Cells(r, colReg).Value) & vbCrLf
    Next r

    BuildPainelPipelineInfo = txt
    Exit Function
EH:
    BuildPainelPipelineInfo = ""
End Function

Private Function BuildExecutedCatalogCsv(ByVal wsSeg As Worksheet, ByVal pipelineNome As String) As String
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim hMap As Object: Set hMap = HeaderMap(wsSeg)
    Dim cPipe As Long: cPipe = MapGet(hMap, "pipeline_name")
    Dim cPid As Long: cPid = MapGet(hMap, "Prompt ID")

    Dim lastRow As Long: lastRow = wsSeg.Cells(wsSeg.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(wsSeg.Cells(r, cPipe).Value)) = pipelineNome Then
            Dim pid As String
            pid = Trim$(CStr(wsSeg.Cells(r, cPid).Value))
            If pid <> "" And UCase$(pid) <> "STOP" Then d(pid) = 1
        End If
    Next r

    Dim out As String
    out = "prompt_id,catalogo,nome_curto,nome_descritivo,modelo,modos,storage" & vbCrLf

    Dim k As Variant
    For Each k In d.Keys
        Dim p As PromptDefinicao
        p = Catalogo_ObterPromptPorID(CStr(k))
        out = out & CsvRow(Array(p.Id, PrefixFromId(p.Id), p.NomeCurto, p.NomeDescritivo, p.modelo, p.modos, p.storage)) & vbCrLf
    Next k

    BuildExecutedCatalogCsv = out
End Function

Private Function PrefixFromId(ByVal promptId As String) As String
    Dim p As Long
    p = InStr(1, promptId, "/")
    If p > 1 Then
        PrefixFromId = Left$(promptId, p - 1)
    Else
        PrefixFromId = ""
    End If
End Function

Private Function SheetToCsv(ByVal ws As Worksheet) As String
    Dim lr As Long, lc As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim r As Long, c As Long, line As String, out As String
    For r = 1 To lr
        line = ""
        For c = 1 To lc
            If c > 1 Then line = line & ","
            line = line & CsvEscape(CStr(ws.Cells(r, c).Value))
        Next c
        out = out & line & vbCrLf
    Next r
    SheetToCsv = out
End Function

Private Function CsvRow(ByVal vals As Variant) As String
    Dim i As Long, s As String
    For i = LBound(vals) To UBound(vals)
        If i > LBound(vals) Then s = s & ","
        s = s & CsvEscape(CStr(vals(i)))
    Next i
    CsvRow = s
End Function

Private Function CsvEscape(ByVal s As String) As String
    s = Replace(s, """", """""")
    CsvEscape = """" & s & """"
End Function

Private Function GitFileItem(ByVal path As String, ByVal content As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("path") = path
    d("content") = content
    Set GitFileItem = d
End Function

Private Function GitData_CommitFiles(ByVal owner As String, ByVal repo As String, ByVal branch As String, ByVal token As String, ByVal files As Collection, ByVal pipelineNome As String) As String
    On Error GoTo EH

    Dim apiBase As String: apiBase = GitCfg_Get("GH_API_BASE", "https://api.github.com")
    Dim headRefUrl As String
    headRefUrl = apiBase & "/repos/" & owner & "/" & repo & "/git/ref/heads/" & branch

    Dim refBody As String
    refBody = Git_Http("GET", headRefUrl, token, "")
    Dim headSha As String: headSha = JsonPick(refBody, "sha")
    If headSha = "" Then Exit Function
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_REF_OK", "HEAD obtido.", "sha=" & Left$(headSha, 10))

    Dim commitBody As String
    commitBody = Git_Http("GET", apiBase & "/repos/" & owner & "/" & repo & "/git/commits/" & headSha, token, "")
    Dim baseTreeSha As String: baseTreeSha = JsonPickTreeSha(commitBody)
    If baseTreeSha = "" Then baseTreeSha = JsonPick(commitBody, "sha")

    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_BASE_TREE_OK", "Base tree resolvida.", "tree_sha=" & Left$(baseTreeSha, 10))

    Dim treeItems As String
    Dim i As Long
    For i = 1 To files.Count
        Dim f As Object: Set f = files(i)
        Dim blobSha As String
        blobSha = Git_CreateBlob(apiBase, owner, repo, token, CStr(f("content")))
        If blobSha = "" Then Exit Function
        If treeItems <> "" Then treeItems = treeItems & ","
        treeItems = treeItems & "{""path"":""" & Json_EscapeString(CStr(f("path")) ) & """,""mode"":""100644"",""type"":""blob"",""sha"":""" & blobSha & """}"
    Next i
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_BLOBS_CREATED", "Blobs criados.", "n=" & CStr(files.Count))

    Dim treeReq As String
    treeReq = "{""base_tree"":""" & baseTreeSha & """,""tree"": [" & treeItems & "]}"
    Dim treeResp As String
    treeResp = Git_Http("POST", apiBase & "/repos/" & owner & "/" & repo & "/git/trees", token, treeReq)
    Dim treeSha As String: treeSha = JsonPick(treeResp, "sha")
    If treeSha = "" Then Exit Function
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_TREE_CREATED", "Tree criada.", "sha=" & Left$(treeSha, 10))

    Dim commitMsg As String
    commitMsg = Replace(GitCfg_Get("GH_COMMIT_MESSAGE_TEMPLATE", "PIPELINER run {{RUN_ID}}"), "{{RUN_ID}}", Format$(Now, "yyyymmdd_hhnnss"))

    Dim newCommitReq As String
    newCommitReq = "{""message"":""" & Json_EscapeString(commitMsg) & """,""tree"":""" & treeSha & """,""parents"": [""" & headSha & """]}"
    Dim newCommitResp As String
    newCommitResp = Git_Http("POST", apiBase & "/repos/" & owner & "/" & repo & "/git/commits", token, newCommitReq)
    Dim newCommitSha As String: newCommitSha = JsonPick(newCommitResp, "sha")
    If newCommitSha = "" Then Exit Function
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_COMMIT_CREATED", "Commit criado.", "sha=" & Left$(newCommitSha, 10))

    Dim updReq As String
    updReq = "{""sha"":""" & newCommitSha & """,""force"":" & LCase$(GitCfg_Get("GH_FORCE_UPDATE", "false")) & "}"
    Call Git_Http("PATCH", headRefUrl, token, updReq)

    GitData_CommitFiles = newCommitSha
    Exit Function
EH:
    GitData_CommitFiles = ""
End Function

Private Function Git_CreateBlob(ByVal apiBase As String, ByVal owner As String, ByVal repo As String, ByVal token As String, ByVal content As String) As String
    Dim req As String
    req = "{""content"":""" & Json_EscapeString(content) & """,""encoding"":""utf-8""}"
    Dim resp As String
    resp = Git_Http("POST", apiBase & "/repos/" & owner & "/" & repo & "/git/blobs", token, req)
    Git_CreateBlob = JsonPick(resp, "sha")
End Function

Private Function Git_Http(ByVal method As String, ByVal url As String, ByVal token As String, ByVal body As String) As String
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Accept", "application/vnd.github+json"
    http.SetRequestHeader "X-GitHub-Api-Version", GitCfg_Get("GH_API_VERSION", "2022-11-28")
    http.SetRequestHeader "User-Agent", GitCfg_Get("GH_USER_AGENT", "PIPELINER-VBA")
    If body <> "" Then
        http.SetRequestHeader "Content-Type", "application/json"
        http.Send body
    Else
        http.Send
    End If

    Git_Http = CStr(http.ResponseText)
End Function

Private Function JsonPickTreeSha(ByVal body As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """tree""\s*:\s*\{\s*""sha""\s*:\s*""([^""]+)"""
    If re.Test(body) Then JsonPickTreeSha = re.Execute(body)(0).SubMatches(0)
End Function

Private Function JsonPick(ByVal body As String, ByVal keyName As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """" & keyName & """" & "\s*:\s*""([^""]+)"""
    If re.Test(body) Then JsonPick = re.Execute(body)(0).SubMatches(0)
End Function

Private Function GitCfg_Get(ByVal keyName As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), keyName, vbTextCompare) = 0 Then
            GitCfg_Get = Trim$(CStr(ws.Cells(i, 2).Value))
            If GitCfg_Get = "" Then GitCfg_Get = defaultValue
            Exit Function
        End If
    Next i
Fallback:
    GitCfg_Get = defaultValue
End Function

Private Function GitCfg_ResolveToken() As String
    Dim envKey As String: envKey = GitCfg_Get("GH_TOKEN_ENV", "GITHUB_TOKEN")
    Dim t As String
    t = Trim$(CStr(Environ$(envKey)))
    If t = "" Then
        t = GitCfg_Get("GH_TOKEN_CONFIG", "")
    End If
    GitCfg_ResolveToken = t
End Function

Private Sub GitDebug_WriteLinkToSeguimento(ByVal pipelineNome As String, ByVal link As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_SEGUIMENTO)
    Dim map As Object: Set map = HeaderMap(ws)
    Dim cPipe As Long: cPipe = MapGet(map, "pipeline_name")
    Dim cGit As Long: cGit = MapGet(map, "GIT_DEBUG")
    If cGit = 0 Then
        cGit = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, cGit).Value = "GIT_DEBUG"
    End If

    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, cPipe).Value)) = pipelineNome Then ws.Cells(r, cGit).Value = link
    Next r
End Sub

Private Sub GitDebug_WriteLinkToHistorico(ByVal pipelineNome As String, ByVal link As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_HIST)
    Dim map As Object: Set map = HeaderMap(ws)
    Dim cPipe As Long: cPipe = MapGet(map, "Nome do Pipeline")
    Dim cGit As Long: cGit = MapGet(map, "GIT_DEBUG")
    If cGit = 0 Then
        cGit = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, cGit).Value = "GIT_DEBUG"
    End If

    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, cPipe).Value)) = pipelineNome And Trim$(CStr(ws.Cells(r, cGit).Value)) = "" Then
            ws.Cells(r, cGit).Value = link
        End If
    Next r
End Sub

Private Function HeaderMap(ByVal ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim lc As Long: lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lc
        d(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c
    Set HeaderMap = d
End Function

Private Function MapGet(ByVal d As Object, ByVal key As String) As Long
    If d.exists(key) Then MapGet = CLng(d(key)) Else MapGet = 0
End Function

Private Function Git_SanitizePathPart(ByVal s As String) As String
    Dim out As String
    out = Trim$(s)
    out = Replace(out, "\", "-")
    out = Replace(out, "/", "-")
    out = Replace(out, ":", "-")
    out = Replace(out, "*", "-")
    out = Replace(out, "?", "-")
    out = Replace(out, """", "-")
    out = Replace(out, "<", "-")
    out = Replace(out, ">", "-")
    out = Replace(out, "|", "-")
    If out = "" Then out = "pipeline"
    Git_SanitizePathPart = out
End Function

Private Function Git_UrlEncodePath(ByVal p As String) As String
    Dim parts() As String
    parts = Split(p, "/")
    Dim i As Long, out As String
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then out = out & "/"
        out = out & Git_UrlEncodeSegment(CStr(parts(i)))
    Next i
    Git_UrlEncodePath = out
End Function

Private Function Git_UrlEncodeSegment(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "-" Or ch = "_" Or ch = "." Then
            out = out & ch
        Else
            out = out & "%" & Right$("0" & Hex$(code), 2)
        End If
    Next i
    Git_UrlEncodeSegment = out
End Function

Public Sub GitDebug_Config_InstalarParametros(Optional ByVal sobrescreverValores As Boolean = False)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    Call GitDebug_Config_EnsureGuideHeaders(ws)

    Dim defs As Collection
    Set defs = GitDebug_Config_Definitions()

    Dim i As Long
    Dim createdCount As Long
    Dim updatedCount As Long

    For i = 1 To defs.Count
        Dim d As Object
        Set d = defs(i)

        Dim rowKey As Long
        rowKey = GitDebug_Config_FindKeyRow(ws, CStr(d("key")))

        If rowKey = 0 Then
            rowKey = GitDebug_Config_NextRow(ws)
            ws.Cells(rowKey, 1).Value = CStr(d("key"))
            createdCount = createdCount + 1
        Else
            updatedCount = updatedCount + 1
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 2).Value)) = "" Then
            ws.Cells(rowKey, 2).Value = CStr(d("default"))
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 3).Value)) = "" Then
            ws.Cells(rowKey, 3).Value = CStr(d("help"))
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 4).Value)) = "" Then
            ws.Cells(rowKey, 4).Value = CStr(d("default"))
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 5).Value)) = "" Then
            ws.Cells(rowKey, 5).Value = CStr(d("allowed"))
        End If
    Next i

    MsgBox "Parametros GH_* preparados no Config." & vbCrLf & _
           "Criados: " & CStr(createdCount) & " | Atualizados/validados: " & CStr(updatedCount), vbInformation
    Exit Sub

EH:
    MsgBox "Erro em GitDebug_Config_InstalarParametros: " & Err.Description, vbExclamation
End Sub

Private Sub GitDebug_Config_EnsureGuideHeaders(ByVal ws As Worksheet)
    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then ws.Cells(1, 1).Value = "Key"
    If Trim$(CStr(ws.Cells(1, 2).Value)) = "" Then ws.Cells(1, 2).Value = "Value"
    If Trim$(CStr(ws.Cells(1, 3).Value)) = "" Then ws.Cells(1, 3).Value = "Explicacao (leigos)"
    If Trim$(CStr(ws.Cells(1, 4).Value)) = "" Then ws.Cells(1, 4).Value = "Default"
    If Trim$(CStr(ws.Cells(1, 5).Value)) = "" Then ws.Cells(1, 5).Value = "Valores possiveis / intervalo"
End Sub

Private Function GitDebug_Config_Definitions() As Collection
    Dim defs As New Collection

    Call GitDebug_Config_Add(defs, "GH_UPLOAD_MODE", "tree_commit", "Modo global do upload para GitHub no PIPELINER.", "contents_api | tree_commit")
    Call GitDebug_Config_Add(defs, "GH_OWNER", "cpsa-org", "Dono do repositorio (organizacao ou utilizador).", "texto nao vazio")
    Call GitDebug_Config_Add(defs, "GH_REPO", "pipeliner-data", "Nome do repositorio onde guardar os debug runs.", "texto nao vazio")
    Call GitDebug_Config_Add(defs, "GH_BRANCH", "main", "Branch alvo para criar commits de debug.", "branch existente")
    Call GitDebug_Config_Add(defs, "GH_API_BASE", "https://api.github.com", "URL base da API GitHub (ou GitHub Enterprise).", "URL valida")

    Call GitDebug_Config_Add(defs, "GH_AUTH_MODE", "PAT", "Modo de autenticacao. Hoje o fluxo usa token (PAT).", "PAT | GITHUB_APP")
    Call GitDebug_Config_Add(defs, "GH_TOKEN_ENV", "GITHUB_TOKEN", "Nome da variavel de ambiente que guarda o token.", "nome de variavel de ambiente")
    Call GitDebug_Config_Add(defs, "GH_TOKEN_CONFIG", "", "Fallback local para token quando ENV estiver vazio (evitar em producao).", "string vazia ou token")

    Call GitDebug_Config_Add(defs, "GH_COMMIT_PREFIX", "PIPELINER", "Prefixo visual para identificar commits automaticos.", "texto curto")
    Call GitDebug_Config_Add(defs, "GH_COMMIT_AUTHOR_NAME", "PIPELINER Bot", "Nome de autor para auditoria nos commits.", "texto")
    Call GitDebug_Config_Add(defs, "GH_COMMIT_AUTHOR_EMAIL", "bot@cpsa.pt", "Email de autor para auditoria nos commits.", "email")
    Call GitDebug_Config_Add(defs, "GH_COMMIT_MESSAGE_TEMPLATE", "PIPELINER run {{RUN_ID}}", "Template da mensagem de commit. {{RUN_ID}} e substituido no runtime.", "template com placeholders")

    Call GitDebug_Config_Add(defs, "GH_BATCH_MODE", "tree_commit", "Modo de upload em batch para este modulo.", "tree_commit")
    Call GitDebug_Config_Add(defs, "GH_MAX_FILES", "200", "Numero maximo de ficheiros por commit (protecao).", "1..1000")
    Call GitDebug_Config_Add(defs, "GH_MAX_FILE_MB", "50", "Tamanho maximo por ficheiro (MB).", "1..200")
    Call GitDebug_Config_Add(defs, "GH_ENCODING_TEXT", "utf-8", "Encoding dos ficheiros de texto enviados para blobs.", "utf-8")
    Call GitDebug_Config_Add(defs, "GH_BINARY_MODE", "base64", "Encoding recomendado para ficheiros binarios.", "base64")

    Call GitDebug_Config_Add(defs, "GH_BASE_PATH", "pipeliner_runs", "Pasta base no repo para agrupar execucoes.", "path relativo sem / inicial")
    Call GitDebug_Config_Add(defs, "GH_RUN_FOLDER_TEMPLATE", "{{DATE}}/{{RUN_ID}}", "Template opcional da subpasta do run.", "ex.: {{DATE}}/{{RUN_ID}}")
    Call GitDebug_Config_Add(defs, "GH_LOG_FOLDER", "logs", "Subpasta para logs complementares (quando aplicavel).", "path relativo")

    Call GitDebug_Config_Add(defs, "GH_RETRY_ON_CONFLICT", "true", "Se true, tenta novamente quando o HEAD muda durante commit.", "true | false")
    Call GitDebug_Config_Add(defs, "GH_MAX_RETRIES", "3", "Numero maximo de retries em conflito de ref.", "0..10")
    Call GitDebug_Config_Add(defs, "GH_FORCE_UPDATE", "false", "Se true, faz update forcado da ref (nao recomendado).", "true | false")

    Call GitDebug_Config_Add(defs, "GH_DEBUG_MODE", "true", "Liga registos de troubleshooting GH_* no DEBUG.", "true | false")
    Call GitDebug_Config_Add(defs, "GH_LOG_HTTP", "false", "Se true, regista requests/responses HTTP resumidos no DEBUG.", "true | false")
    Call GitDebug_Config_Add(defs, "GH_LOG_BLOB_SHA", "true", "Se true, mostra SHA curto dos blobs criados no DEBUG.", "true | false")

    Call GitDebug_Config_Add(defs, "GH_API_VERSION", "2022-11-28", "Versao da API GitHub enviada em header.", "YYYY-MM-DD")
    Call GitDebug_Config_Add(defs, "GH_USER_AGENT", "PIPELINER-VBA", "User-Agent usado nas chamadas a API.", "texto sem vazio")

    Set GitDebug_Config_Definitions = defs
End Function

Private Sub GitDebug_Config_Add(ByRef defs As Collection, ByVal keyName As String, ByVal defaultValue As String, ByVal helpText As String, ByVal allowed As String)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("key") = keyName
    d("default") = defaultValue
    d("help") = helpText
    d("allowed") = allowed
    defs.Add d
End Sub

Private Function GitDebug_Config_FindKeyRow(ByVal ws As Worksheet, ByVal keyName As String) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 1 To lr
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), keyName, vbTextCompare) = 0 Then
            GitDebug_Config_FindKeyRow = r
            Exit Function
        End If
    Next r

    GitDebug_Config_FindKeyRow = 0
End Function

Private Function GitDebug_Config_NextRow(ByVal ws As Worksheet) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 9 Then lr = 8
    GitDebug_Config_NextRow = lr + 1
End Function
