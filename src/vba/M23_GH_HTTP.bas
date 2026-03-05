Attribute VB_Name = "M23_GH_HTTP"
Option Explicit

' =============================================================================
' MÃ³dulo: M23_GH_HTTP
' PropÃ³sito:
' - Encapsular chamadas HTTP para integraÃ§Ã£o GitHub da exportaÃ§Ã£o DEBUG.
' - Aplicar fallback entre WinHTTP e MSXML com interface Ãºnica.
' - Adicionar retry/backoff e timeout configurÃ¡veis para robustez operacional.
'
' AtualizaÃ§Ãµes:
' - 2026-03-05 | Codex | Hardening HTTP (PR2)
'   - Adiciona timeout/retries/backoff como parÃ¢metros opcionais.
'   - Implementa retry para 429/5xx com pausa incremental entre tentativas.
' - 2026-03-04 | Codex | CriaÃ§Ã£o do mÃ³dulo HTTP GitHub
'   - Adiciona request JSON com fallback de engine WinHTTP -> MSXML.
'   - ExpÃµe status/response para logging e decisÃ£o no mÃ³dulo facade.
'
' FunÃ§Ãµes e procedimentos:
' - GH_HTTP_RequestJson(method, url, token, body, statusCode, responseText, errText, userAgent, timeoutMs, maxRetries, backoffMs, attemptsUsed) As Boolean
'   - Executa chamada HTTP autenticada para API GitHub com retry e devolve sucesso/falha.
' =============================================================================

Public Function GH_HTTP_RequestJson( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String, _
    Optional ByVal userAgent As String = "PIPELINER-GitDebugExport", _
    Optional ByVal timeoutMs As Long = 30000, _
    Optional ByVal maxRetries As Long = 2, _
    Optional ByVal backoffMs As Long = 800, _
    Optional ByRef attemptsUsed As Long = 0) As Boolean

    statusCode = 0
    responseText = ""
    errText = ""
    attemptsUsed = 0

    If timeoutMs <= 0 Then timeoutMs = 30000
    If maxRetries < 0 Then maxRetries = 0
    If backoffMs < 0 Then backoffMs = 0

    Dim attempt As Long
    For attempt = 0 To maxRetries
        attemptsUsed = attempt + 1

        If GH_HTTP_RequestWithWinHttp(method, url, token, body, statusCode, responseText, errText, userAgent, timeoutMs) Then
            If statusCode >= 200 And statusCode < 300 Then
                GH_HTTP_RequestJson = True
                Exit Function
            End If
        ElseIf GH_HTTP_RequestWithMsxml(method, url, token, body, statusCode, responseText, errText, userAgent, timeoutMs) Then
            If statusCode >= 200 And statusCode < 300 Then
                GH_HTTP_RequestJson = True
                Exit Function
            End If
        End If

        If Not GH_HTTP_IsRetriableStatus(statusCode) Then Exit For
        Call GH_HTTP_PauseMs(backoffMs * (attempt + 1))
    Next attempt

    GH_HTTP_RequestJson = False
End Function

Private Function GH_HTTP_RequestWithWinHttp( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String, _
    ByVal userAgent As String, _
    ByVal timeoutMs As Long) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open method, url, False
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    Call GH_HTTP_ApplyHeaders(http, token, userAgent)
    http.Send body

    statusCode = CLng(http.Status)
    responseText = CStr(http.ResponseText)
    GH_HTTP_RequestWithWinHttp = True
    Exit Function
EH:
    errText = "WINHTTP: " & Err.Description
    GH_HTTP_RequestWithWinHttp = False
End Function

Private Function GH_HTTP_RequestWithMsxml( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String, _
    ByVal userAgent As String, _
    ByVal timeoutMs As Long) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open method, url, False
    http.setTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    Call GH_HTTP_ApplyHeaders(http, token, userAgent)
    http.Send body

    statusCode = CLng(http.Status)
    responseText = CStr(http.responseText)
    GH_HTTP_RequestWithMsxml = True
    Exit Function
EH:
    If Len(errText) > 0 Then errText = errText & " | "
    errText = errText & "MSXML: " & Err.Description
    GH_HTTP_RequestWithMsxml = False
End Function

Private Sub GH_HTTP_ApplyHeaders(ByVal http As Object, ByVal token As String, ByVal userAgent As String)
    http.setRequestHeader "Accept", "application/vnd.github+json"
    http.setRequestHeader "X-GitHub-Api-Version", "2022-11-28"
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "User-Agent", userAgent
    http.setRequestHeader "Content-Type", "application/json"
End Sub

Private Function GH_HTTP_IsRetriableStatus(ByVal statusCode As Long) As Boolean
    GH_HTTP_IsRetriableStatus = (statusCode = 429 Or (statusCode >= 500 And statusCode <= 599))
End Function

Private Sub GH_HTTP_PauseMs(ByVal delayMs As Long)
    If delayMs <= 0 Then Exit Sub

    Dim tEnd As Double
    tEnd = Timer + (delayMs / 1000#)

    Do While Timer < tEnd
        DoEvents
    Loop
End Sub
