Attribute VB_Name = "M23_GH_HTTP"
Option Explicit

' =============================================================================
' Modulo: M23_GH_HTTP
' Proposito:
' - Encapsular requests GitHub (GET/POST/PATCH) com headers canonicos.
' - Isolar fallback de engine HTTP (WinHTTP -> MSXML) para portabilidade.
' - Expor diagnostico minimo (status/erro/response) para fluxo de commit.
'
' Atualizacoes:
' - 2026-03-08 | Codex | Normaliza GH_API_VERSION para reduzir falhas 400 no GitHub
'   - Aceita formato dd/mm/yyyy e converte para yyyy-mm-dd antes de enviar header.
'   - Faz fallback seguro para 2022-11-28 quando o valor nao for valido.
' - 2026-03-04 | Codex | Refactor HTTP para modulo dedicado
'   - Move construcao de headers GitHub para helper canonico unico.
'   - Implementa send JSON com fallback WinHTTP/MSXML e output padronizado.
'   - Adiciona optional logging controlado por GH_LOG_HTTP.
'
' Funcoes e procedimentos:
' - GH_HTTP_SendJson(method, url, cfg, body, statusCode, responseText, errText, pipelineNome) As Boolean
'   - Executa request JSON autenticado e devolve sucesso por HTTP 2xx.
' - GH_HTTP_NormalizeApiVersion(rawValue As String) As String
'   - Normaliza GH_API_VERSION (yyyy-mm-dd) com fallback seguro para evitar 400 por formato invalido.
' =============================================================================

Public Function GH_HTTP_SendJson( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal cfg As Object, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String, _
    Optional ByVal pipelineNome As String = "") As Boolean

    statusCode = 0
    responseText = ""
    errText = ""

    If GH_HTTP_SendWithWinHttp(method, url, cfg, body, statusCode, responseText, errText) Then
        GH_HTTP_SendJson = (statusCode >= 200 And statusCode < 300)
        Call GH_HTTP_Log(method, url, statusCode, cfg, pipelineNome)
        Exit Function
    End If

    If GH_HTTP_SendWithMsxml(method, url, cfg, body, statusCode, responseText, errText) Then
        GH_HTTP_SendJson = (statusCode >= 200 And statusCode < 300)
        Call GH_HTTP_Log(method, url, statusCode, cfg, pipelineNome)
        Exit Function
    End If

    Call GH_HTTP_LogFailure(method, url, errText, cfg, pipelineNome)
    GH_HTTP_SendJson = False
End Function

Private Function GH_HTTP_SendWithWinHttp( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal cfg As Object, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open method, url, False
    Call GH_HTTP_ApplyGitHubHeaders(http, cfg)

    If body <> "" Then
        http.Send body
    Else
        http.Send
    End If

    statusCode = CLng(http.Status)
    responseText = CStr(http.ResponseText)
    GH_HTTP_SendWithWinHttp = True
    Exit Function

EH:
    errText = "WINHTTP: " & Err.Description
    GH_HTTP_SendWithWinHttp = False
End Function

Private Function GH_HTTP_SendWithMsxml( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal cfg As Object, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open method, url, False
    Call GH_HTTP_ApplyGitHubHeaders(http, cfg)

    If body <> "" Then
        http.Send body
    Else
        http.Send
    End If

    statusCode = CLng(http.Status)
    responseText = CStr(http.responseText)
    GH_HTTP_SendWithMsxml = True
    Exit Function

EH:
    If errText <> "" Then errText = errText & " | "
    errText = errText & "MSXML: " & Err.Description
    GH_HTTP_SendWithMsxml = False
End Function

Private Sub GH_HTTP_ApplyGitHubHeaders(ByVal http As Object, ByVal cfg As Object)
    http.setRequestHeader "Authorization", "Bearer " & GH_Config_GetString(cfg, "token")
    http.setRequestHeader "Accept", "application/vnd.github+json"
    http.setRequestHeader "X-GitHub-Api-Version", GH_HTTP_NormalizeApiVersion(GH_Config_GetString(cfg, "api_version", "2022-11-28"))
    http.setRequestHeader "User-Agent", GH_Config_GetString(cfg, "user_agent", "PIPELINER-VBA")
    http.setRequestHeader "Content-Type", "application/json"
End Sub

Private Function GH_HTTP_NormalizeApiVersion(ByVal rawValue As String) As String
    Dim valueText As String
    valueText = Trim$(rawValue)

    If valueText = "" Then
        GH_HTTP_NormalizeApiVersion = "2022-11-28"
        Exit Function
    End If

    If valueText Like "####-##-##" Then
        GH_HTTP_NormalizeApiVersion = valueText
        Exit Function
    End If

    If valueText Like "##/##/####" Then
        GH_HTTP_NormalizeApiVersion = Right$(valueText, 4) & "-" & Mid$(valueText, 4, 2) & "-" & Left$(valueText, 2)
        Exit Function
    End If

    GH_HTTP_NormalizeApiVersion = "2022-11-28"
End Function

Private Sub GH_HTTP_Log(ByVal method As String, ByVal url As String, ByVal statusCode As Long, ByVal cfg As Object, ByVal pipelineNome As String)
    If Not GH_Config_GetBoolean(cfg, "log_http", False) Then Exit Sub
    Call GH_LogInfo(0, pipelineNome, GH_EVT_HTTP, method & " " & url, "http_status=" & CStr(statusCode))
End Sub

Private Sub GH_HTTP_LogFailure(ByVal method As String, ByVal url As String, ByVal errText As String, ByVal cfg As Object, ByVal pipelineNome As String)
    If Not GH_Config_GetBoolean(cfg, "log_http", False) Then Exit Sub
    Call GH_LogWarn(0, pipelineNome, GH_EVT_HTTP_FAIL, method & " " & url, errText)
End Sub
