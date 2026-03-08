Attribute VB_Name = "M22_GH_Config"
Option Explicit

' =============================================================================
' Modulo: M22_GH_Config
' Proposito:
' - Centralizar leitura, normalizacao e validacao de parametros GH_* no Config.
' - Aplicar defaults internos para retrocompatibilidade com templates antigos.
' - Resolver enable/token de forma deterministica para o facade M21.
'
' Atualizacoes:
' - 2026-03-08 | Codex | Resolve modo de upload GH_UPLOAD_MODE com validacao canonica
'   - Adiciona helper GH_Config_ResolveUploadMode com default tree_commit quando vazio.
'   - Marca explicitamente valores invalidos para abortar com erro acionavel no runtime.
' - 2026-03-07 | Codex | Preferencia de token por ambiente GH_TOKEN
'   - Se GH_TOKEN_ENV estiver vazio, "Ambiente"/"Environment" ou erro, usa variavel de ambiente GH_TOKEN.
'   - Expoe token_source no cfg para auditoria no DEBUG sem revelar segredo.
' - 2026-03-07 | Codex | Validacao minima GH_* com mensagem acionavel
'   - Agrega campos obrigatorios em falta (owner/repo/branch/token/base_path) numa unica mensagem curta.
'   - Inclui acao recomendada para preencher Config (GH_OWNER/GH_REPO/GH_BRANCH/token/GH_BASE_PATH).
' - 2026-03-05 | Codex | Expor pasta de logs e template de run folder no cfg GH_*
'   - Passa a ler GH_LOG_FOLDER e GH_RUN_FOLDER_TEMPLATE para compor path remoto por execucao.
'   - Mantem defaults internos e compatibilidade quando as chaves nao existem.
' - 2026-03-04 | Codex | Refactor da configuracao GitHub para modulo dedicado
'   - Move leitura de GH_* (Config) para um dicionario normalizado.
'   - Implementa validacao canonica de campos obrigatorios e limites numericos.
'   - Mantem regra de enable por painelAutoSave ("sim, todos" ou "debug").
'
' Funcoes e procedimentos:
' - GH_Config_Load(painelAutoSave As String) As Object
'   - Le GH_* no Config, normaliza tipos e devolve dicionario pronto para runtime.
' - GH_Config_Validate(cfg As Object, reason As String) As Boolean
'   - Valida owner/repo/branch/token e limites de upload; devolve motivo curto.
' - GH_Config_GetString(cfg As Object, keyName As String, Optional fallback As String = "") As String
'   - Le string do dicionario com fallback seguro.
' - GH_Config_GetBoolean(cfg As Object, keyName As String, Optional fallback As Boolean = False) As Boolean
'   - Le booleano normalizado com fallback seguro.
' - GH_Config_AppendMissing(currentList As String, itemName As String) As String
'   - Helper para concatenar campos em falta nas mensagens de validacao GH_*.
' - GH_Config_ResolveToken(Optional ByRef sourceUsed As String = "") As String
'   - Resolve token priorizando GH_TOKEN quando indicado e devolve a fonte para auditoria.
' - GH_Config_ResolveUploadMode(cfg As Object, reason As String, Optional wasDefaulted As Boolean = False) As String
'   - Normaliza upload_mode e valida apenas tree_commit/contents_api com fallback seguro.
' =============================================================================

Private Const SHEET_CONFIG As String = "Config"

Public Function GH_Config_Load(ByVal painelAutoSave As String) As Object
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = 1

    cfg("enabled") = GH_Config_IsEnabledByPanel(painelAutoSave)
    cfg("upload_mode") = LCase$(GH_Config_Get("GH_UPLOAD_MODE", "tree_commit"))
    cfg("batch_mode") = LCase$(GH_Config_Get("GH_BATCH_MODE", "tree_commit"))

    cfg("owner") = GH_Config_Get("GH_OWNER", "")
    cfg("repo") = GH_Config_Get("GH_REPO", "")
    cfg("branch") = GH_Config_Get("GH_BRANCH", "main")
    cfg("api_base") = GH_Config_Get("GH_API_BASE", "https://api.github.com")
    cfg("api_version") = GH_Config_Get("GH_API_VERSION", "2022-11-28")
    cfg("user_agent") = GH_Config_Get("GH_USER_AGENT", "PIPELINER-VBA")

    Dim tokenSource As String
    cfg("token") = GH_Config_ResolveToken(tokenSource)
    cfg("token_source") = tokenSource

    cfg("base_path") = GH_Config_Get("GH_BASE_PATH", "pipeliner_runs")
    cfg("log_folder") = GH_Config_Get("GH_LOG_FOLDER", "logs")
    cfg("run_folder_template") = GH_Config_Get("GH_RUN_FOLDER_TEMPLATE", "{{YYYY}}-{{MM}}-{{DD}} - {{HHMM}} - [{{PIPELINE_NAME}}]")
    cfg("commit_message_template") = GH_Config_Get("GH_COMMIT_MESSAGE_TEMPLATE", "PIPELINER run {{RUN_ID}}")
    cfg("force_update") = GH_Config_ToBoolean(GH_Config_Get("GH_FORCE_UPDATE", "false"), False)

    cfg("max_files") = GH_Config_ToLong(GH_Config_Get("GH_MAX_FILES", "200"), 200)
    cfg("max_file_mb") = GH_Config_ToLong(GH_Config_Get("GH_MAX_FILE_MB", "50"), 50)
    cfg("encoding_text") = LCase$(GH_Config_Get("GH_ENCODING_TEXT", "utf-8"))
    cfg("binary_mode") = LCase$(GH_Config_Get("GH_BINARY_MODE", "base64"))

    cfg("debug_mode") = GH_Config_ToBoolean(GH_Config_Get("GH_DEBUG_MODE", "true"), True)
    cfg("log_http") = GH_Config_ToBoolean(GH_Config_Get("GH_LOG_HTTP", "false"), False)
    cfg("log_blob_sha") = GH_Config_ToBoolean(GH_Config_Get("GH_LOG_BLOB_SHA", "true"), True)
    cfg("retry_on_conflict") = GH_Config_ToBoolean(GH_Config_Get("GH_RETRY_ON_CONFLICT", "true"), True)
    cfg("max_retries") = GH_Config_ToLong(GH_Config_Get("GH_MAX_RETRIES", "3"), 3)
    cfg("contents_batch_policy") = LCase$(Trim$(GH_Config_Get("GH_CONTENTS_BATCH_POLICY", "fail_fast")))

    Set GH_Config_Load = cfg
End Function

Public Function GH_Config_ResolveUploadMode( _
    ByVal cfg As Object, _
    ByRef reason As String, _
    Optional ByRef wasDefaulted As Boolean = False) As String

    reason = ""
    wasDefaulted = False

    Dim modeValue As String
    modeValue = LCase$(Trim$(GH_Config_GetString(cfg, "upload_mode", "")))

    If modeValue = "" Then
        modeValue = "tree_commit"
        wasDefaulted = True
    End If

    Select Case modeValue
        Case "tree_commit", "contents_api"
            GH_Config_ResolveUploadMode = modeValue
        Case Else
            reason = "GH_UPLOAD_MODE invalido: " & modeValue & " | [ACTION] Usar tree_commit ou contents_api."
            GH_Config_ResolveUploadMode = ""
    End Select
End Function

Public Function GH_Config_Validate(ByVal cfg As Object, ByRef reason As String) As Boolean
    reason = ""

    If Not GH_Config_GetBoolean(cfg, "enabled", False) Then
        GH_Config_Validate = True
        Exit Function
    End If

    Dim missing As String
    missing = ""

    If GH_Config_GetString(cfg, "owner") = "" Then missing = GH_Config_AppendMissing(missing, "GH_OWNER")
    If GH_Config_GetString(cfg, "repo") = "" Then missing = GH_Config_AppendMissing(missing, "GH_REPO")
    If GH_Config_GetString(cfg, "branch") = "" Then missing = GH_Config_AppendMissing(missing, "GH_BRANCH")
    If GH_Config_GetString(cfg, "token") = "" Then missing = GH_Config_AppendMissing(missing, "GH_TOKEN_ENV/GH_TOKEN_CONFIG")
    If GH_Config_GetString(cfg, "base_path") = "" Then missing = GH_Config_AppendMissing(missing, "GH_BASE_PATH")

    If missing <> "" Then
        reason = "Campos GH_* obrigatorios em falta: " & missing & " | [ACTION] Na folha Config, preencher GH_OWNER, GH_REPO, GH_BRANCH, token (GH_TOKEN_ENV ou GH_TOKEN_CONFIG) e GH_BASE_PATH."
        Exit Function
    End If

    If GH_Config_GetLong(cfg, "max_files", 200) < 1 Then
        reason = "GH_MAX_FILES invalido"
        Exit Function
    End If

    If GH_Config_GetLong(cfg, "max_file_mb", 50) < 1 Then
        reason = "GH_MAX_FILE_MB invalido"
        Exit Function
    End If

    GH_Config_Validate = True
End Function

Public Function GH_Config_GetString(ByVal cfg As Object, ByVal keyName As String, Optional ByVal fallback As String = "") As String
    On Error GoTo Fallback
    If cfg.exists(keyName) Then
        GH_Config_GetString = Trim$(CStr(cfg(keyName)))
        If GH_Config_GetString = "" Then GH_Config_GetString = fallback
    Else
        GH_Config_GetString = fallback
    End If
    Exit Function
Fallback:
    GH_Config_GetString = fallback
End Function

Public Function GH_Config_GetBoolean(ByVal cfg As Object, ByVal keyName As String, Optional ByVal fallback As Boolean = False) As Boolean
    On Error GoTo Fallback
    If cfg.exists(keyName) Then
        GH_Config_GetBoolean = GH_Config_ToBoolean(cfg(keyName), fallback)
    Else
        GH_Config_GetBoolean = fallback
    End If
    Exit Function
Fallback:
    GH_Config_GetBoolean = fallback
End Function

Public Function GH_Config_GetLong(ByVal cfg As Object, ByVal keyName As String, Optional ByVal fallback As Long = 0) As Long
    On Error GoTo Fallback
    If cfg.exists(keyName) Then
        GH_Config_GetLong = GH_Config_ToLong(cfg(keyName), fallback)
    Else
        GH_Config_GetLong = fallback
    End If
    Exit Function
Fallback:
    GH_Config_GetLong = fallback
End Function

Public Function GH_Config_Get(ByVal keyName As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), keyName, vbTextCompare) = 0 Then
            GH_Config_Get = Trim$(CStr(ws.Cells(i, 2).Value))
            If GH_Config_Get = "" Then GH_Config_Get = defaultValue
            Exit Function
        End If
    Next i

Fallback:
    GH_Config_Get = defaultValue
End Function

Public Function GH_Config_ResolveToken(Optional ByRef sourceUsed As String = "") As String
    On Error GoTo Fallback

    Dim envKeyRaw As String
    envKeyRaw = GH_Config_Get("GH_TOKEN_ENV", "")

    Dim envKey As String
    envKey = Trim$(envKeyRaw)

    If envKey = "" Or UCase$(envKey) = "AMBIENTE" Or UCase$(envKey) = "ENVIRONMENT" Then
        envKey = "GH_TOKEN"
        sourceUsed = "ENV:GH_TOKEN (fallback)"
    Else
        sourceUsed = "ENV:" & envKey
    End If

    Dim tkn As String
    tkn = Trim$(CStr(Environ$(envKey)))

    If tkn = "" Then
        tkn = Trim$(GH_Config_Get("GH_TOKEN_CONFIG", ""))
        If tkn <> "" Then
            sourceUsed = "CONFIG:GH_TOKEN_CONFIG"
        Else
            sourceUsed = sourceUsed & " -> vazio"
        End If
    End If

    GH_Config_ResolveToken = tkn
    Exit Function

Fallback:
    GH_Config_ResolveToken = Trim$(CStr(Environ$("GH_TOKEN")))
    If GH_Config_ResolveToken <> "" Then
        sourceUsed = "ENV:GH_TOKEN (error_fallback)"
    Else
        sourceUsed = "nao_resolvido"
    End If
End Function

Private Function GH_Config_IsEnabledByPanel(ByVal painelAutoSave As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(painelAutoSave))
    GH_Config_IsEnabledByPanel = (InStr(1, s, "sim, todos", vbTextCompare) > 0) Or (InStr(1, s, "debug", vbTextCompare) > 0)
End Function

Private Function GH_Config_AppendMissing(ByVal currentList As String, ByVal itemName As String) As String
    If Trim$(currentList) = "" Then
        GH_Config_AppendMissing = itemName
    Else
        GH_Config_AppendMissing = currentList & ", " & itemName
    End If
End Function

Private Function GH_Config_ToBoolean(ByVal value As Variant, ByVal fallback As Boolean) As Boolean
    Dim raw As String
    raw = UCase$(Trim$(CStr(value)))

    Select Case raw
        Case "TRUE", "1", "SIM", "YES"
            GH_Config_ToBoolean = True
        Case "FALSE", "0", "NAO", "NÃO", "NO"
            GH_Config_ToBoolean = False
        Case Else
            GH_Config_ToBoolean = fallback
    End Select
End Function

Private Function GH_Config_ToLong(ByVal value As Variant, ByVal fallback As Long) As Long
    On Error GoTo Fallback
    GH_Config_ToLong = CLng(value)
    Exit Function
Fallback:
    GH_Config_ToLong = fallback
End Function
