Attribute VB_Name = "M23_GH_HTTP"
Option Explicit

' =============================================================================
' Mï¿½dulo: M23_GH_HTTP
' Propï¿½sito:
' - Centralizar chamadas HTTP JSON para integraï¿½ï¿½es GitHub (GET/POST/PATCH).
' - Garantir headers comuns, timeout configurï¿½vel e retorno estruturado uniforme.
' - Encaminhar falhas HTTP/COM para logging dedicado em M26_GH_Logger.
'
' Atualizaï¿½ï¿½es:
' - 2026-03-04 | Codex | Criaï¿½ï¿½o do cliente HTTP GitHub com resposta estruturada
'   - Implementa GetJson, PostJson e PatchJson com retorno {status, body, ok, erro_curto}.
'   - Centraliza headers Authorization/Accept/X-GitHub-Api-Version/User-Agent.
'   - Adiciona timeout configurï¿½vel e captura robusta de exceï¿½ï¿½es WinHTTP.
'
' Funï¿½ï¿½es e procedimentos:
' - GetJson(endpointUrl, token, ...): executa GET e devolve dicionï¿½rio estruturado.
' - PostJson(endpointUrl, token, payloadJson, ...): executa POST JSON com retorno estruturado.
' - PatchJson(endpointUrl, token, payloadJson, ...): executa PATCH JSON com retorno estruturado.
' =============================================================================

Private Const GH_DEFAULT_ACCEPT As String = "application/vnd.github+json"
Private Const GH_DEFAULT_API_VERSION As String = "2022-11-28"
Private Const GH_DEFAULT_USER_AGENT As String = "PIPELINER-GH-Client/1.0"
Private Const GH_DEFAULT_TIMEOUT_MS As Long = 30000
Private Const GH_SNIPPET_MAX_LEN As Long = 500

Public Function GetJson( _
    ByVal endpointUrl As String, _
    ByVal token As String, _
    Optional ByVal stepName As String = "GH_GET", _
    Optional ByVal timeoutMs As Long = GH_DEFAULT_TIMEOUT_MS _
) As Object
    Set GetJson = GH_SendJson("GET", endpointUrl, token, vbNullString, stepName, timeoutMs)
End Function

Public Function PostJson( _
    ByVal endpointUrl As String, _
    ByVal token As String, _
    ByVal payloadJson As String, _
    Optional ByVal stepName As String = "GH_POST", _
    Optional ByVal timeoutMs As Long = GH_DEFAULT_TIMEOUT_MS _
) As Object
    Set PostJson = GH_SendJson("POST", endpointUrl, token, payloadJson, stepName, timeoutMs)
End Function

Public Function PatchJson( _
    ByVal endpointUrl As String, _
    ByVal token As String, _
    ByVal payloadJson As String, _
    Optional ByVal stepName As String = "GH_PATCH", _
    Optional ByVal timeoutMs As Long = GH_DEFAULT_TIMEOUT_MS _
) As Object
    Set PatchJson = GH_SendJson("PATCH", endpointUrl, token, payloadJson, stepName, timeoutMs)
End Function

Private Function GH_SendJson( _
    ByVal httpMethod As String, _
    ByVal endpointUrl As String, _
    ByVal token As String, _
    ByVal payloadJson As String, _
    ByVal stepName As String, _
    ByVal timeoutMs As Long _
) As Object
    On Error GoTo EH

    Dim result As Object
    Set result = GH_NewResult()

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call GH_ApplyTimeout(http, timeoutMs)
    Call http.Open(UCase$(Trim$(httpMethod)), endpointUrl, False)
    Call GH_ApplyHeaders(http, token)

    If UCase$(Trim$(httpMethod)) = "POST" Or UCase$(Trim$(httpMethod)) = "PATCH" Then
        http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        http.Send payloadJson
    Else
        http.Send
    End If

    result("status") = CLng(http.status)
    result("body") = CStr(http.ResponseText & "")
    result("ok") = (CLng(result("status")) >= 200 And CLng(result("status")) < 300)

    If Not CBool(result("ok")) Then
        result("erro_curto") = GH_BuildShortError(CLng(result("status")), CStr(result("body")))
        Call M26_GH_Logger.GH_LogHttpFailure(stepName, endpointUrl, CLng(result("status")), CStr(result("body")))
    End If

    Set GH_SendJson = result
    Exit Function

EH:
    result("ok") = False
    result("status") = 0
    result("body") = ""
    result("erro_curto") = GH_TrimOneLine("WINHTTP_EXCEPTION: " & Err.Number & " - " & Err.Description, GH_SNIPPET_MAX_LEN)
    Call M26_GH_Logger.GH_LogHttpFailure(stepName, endpointUrl, 0, CStr(result("erro_curto")))
    Set GH_SendJson = result
End Function

Private Sub GH_ApplyHeaders(ByVal http As Object, ByVal token As String)
    http.SetRequestHeader "Authorization", "Bearer " & Trim$(token)
    http.SetRequestHeader "Accept", GH_DEFAULT_ACCEPT
    http.SetRequestHeader "X-GitHub-Api-Version", GH_DEFAULT_API_VERSION
    http.SetRequestHeader "User-Agent", GH_DEFAULT_USER_AGENT
End Sub

Private Sub GH_ApplyTimeout(ByVal http As Object, ByVal timeoutMs As Long)
    Dim timeoutEffective As Long
    timeoutEffective = timeoutMs
    If timeoutEffective <= 0 Then timeoutEffective = GH_DEFAULT_TIMEOUT_MS

    On Error Resume Next
    http.SetTimeouts timeoutEffective, timeoutEffective, timeoutEffective, timeoutEffective
    On Error GoTo 0
End Sub

Private Function GH_NewResult() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d("status") = 0
    d("body") = ""
    d("ok") = False
    d("erro_curto") = ""

    Set GH_NewResult = d
End Function

Private Function GH_BuildShortError(ByVal statusCode As Long, ByVal responseBody As String) As String
    Dim prefix As String
    prefix = "HTTP " & CStr(statusCode)

    If Len(Trim$(responseBody)) = 0 Then
        GH_BuildShortError = prefix
    Else
        GH_BuildShortError = GH_TrimOneLine(prefix & " | " & responseBody, GH_SNIPPET_MAX_LEN)
    End If
End Function

Private Function GH_TrimOneLine(ByVal rawText As String, ByVal maxLen As Long) As String
    Dim txt As String
    txt = Replace(rawText, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Trim$(txt)

    If maxLen <= 3 Then maxLen = 3
    If Len(txt) > maxLen Then
        GH_TrimOneLine = Left$(txt, maxLen - 3) & "..."
    Else
        GH_TrimOneLine = txt
    End If
End Function
