Attribute VB_Name = "M23_GH_HTTP"
Option Explicit

' =============================================================================
' MÃ³dulo: M23_GH_HTTP
' PropÃ³sito:
' - Encapsular chamadas HTTP para integraÃ§Ã£o GitHub da exportaÃ§Ã£o DEBUG.
' - Aplicar fallback entre WinHTTP e MSXML com interface Ãºnica.
' - Isolar detalhes de headers/autenticaÃ§Ã£o do orquestrador.
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | CriaÃ§Ã£o do mÃ³dulo HTTP GitHub
'   - Adiciona request JSON com fallback de engine WinHTTP -> MSXML.
'   - ExpÃµe status/response para logging e decisÃ£o no mÃ³dulo facade.
'
' FunÃ§Ãµes e procedimentos:
' - GH_HTTP_RequestJson(method, url, token, body, statusCode, responseText, errText, userAgent) As Boolean
'   - Executa chamada HTTP autenticada para API GitHub e devolve sucesso/falha.
' =============================================================================

Public Function GH_HTTP_RequestJson( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String, _
    Optional ByVal userAgent As String = "PIPELINER-GitDebugExport") As Boolean

    statusCode = 0
    responseText = ""
    errText = ""

    If GH_HTTP_RequestWithWinHttp(method, url, token, body, statusCode, responseText, errText, userAgent) Then
        GH_HTTP_RequestJson = (statusCode >= 200 And statusCode < 300)
        Exit Function
    End If

    If GH_HTTP_RequestWithMsxml(method, url, token, body, statusCode, responseText, errText, userAgent) Then
        GH_HTTP_RequestJson = (statusCode >= 200 And statusCode < 300)
        Exit Function
    End If

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
    ByVal userAgent As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open method, url, False
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
    ByVal userAgent As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open method, url, False
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
