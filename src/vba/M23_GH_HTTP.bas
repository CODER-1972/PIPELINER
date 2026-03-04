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
' - 2026-03-04 | Codex | Resultado HTTP estruturado para integraÃ§Ã£o GitHub
'   - Adiciona GH_HTTP_RequestJsonResult com contrato GH_HttpCallResult.
'   - MantÃ©m compatibilidade via wrapper GH_HTTP_RequestJson.
' - 2026-03-04 | Codex | CriaÃ§Ã£o do mÃ³dulo HTTP GitHub
'   - Adiciona request JSON com fallback de engine WinHTTP -> MSXML.
'   - ExpÃµe status/response para logging e decisÃ£o no mÃ³dulo facade.
'
' FunÃ§Ãµes e procedimentos:
' - GH_HTTP_RequestJsonResult(method, url, token, body, stepName, userAgent) As GH_HttpCallResult
'   - Executa chamada HTTP autenticada e devolve resultado estruturado.
' - GH_HTTP_RequestJson(method, url, token, body, statusCode, responseText, errText, userAgent) As Boolean
'   - Wrapper retrocompatÃ­vel baseado em status/response/erro por referÃªncia.
' =============================================================================

Public Function GH_HTTP_RequestJsonResult( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByVal stepName As String, _
    Optional ByVal userAgent As String = "PIPELINER-GitDebugExport") As GH_HttpCallResult

    Dim result As GH_HttpCallResult
    result.stepName = stepName

    If GH_HTTP_RequestWithWinHttp(method, url, token, body, result, userAgent) Then
        GH_HTTP_RequestJsonResult = result
        Exit Function
    End If

    Call GH_HTTP_RequestWithMsxml(method, url, token, body, result, userAgent)
    GH_HTTP_RequestJsonResult = result
End Function

Public Function GH_HTTP_RequestJson( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByRef statusCode As Long, _
    ByRef responseText As String, _
    ByRef errText As String, _
    Optional ByVal userAgent As String = "PIPELINER-GitDebugExport") As Boolean

    Dim result As GH_HttpCallResult
    result = GH_HTTP_RequestJsonResult(method, url, token, body, method & " " & url, userAgent)

    statusCode = result.status
    responseText = result.body
    errText = result.errorDetail
    GH_HTTP_RequestJson = result.ok
End Function

Private Function GH_HTTP_RequestWithWinHttp( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByRef result As GH_HttpCallResult, _
    ByVal userAgent As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open method, url, False
    Call GH_HTTP_ApplyHeaders(http, token, userAgent)
    http.Send body

    result.status = CLng(http.Status)
    result.body = CStr(http.ResponseText)
    result.ok = (result.status >= 200 And result.status < 300)
    GH_HTTP_RequestWithWinHttp = True
    Exit Function
EH:
    result.ok = False
    result.status = 0
    result.body = ""
    result.errorDetail = "WINHTTP: " & Err.Description
    GH_HTTP_RequestWithWinHttp = False
End Function

Private Function GH_HTTP_RequestWithMsxml( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal token As String, _
    ByVal body As String, _
    ByRef result As GH_HttpCallResult, _
    ByVal userAgent As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open method, url, False
    Call GH_HTTP_ApplyHeaders(http, token, userAgent)
    http.Send body

    result.status = CLng(http.Status)
    result.body = CStr(http.responseText)
    result.ok = (result.status >= 200 And result.status < 300)
    GH_HTTP_RequestWithMsxml = True
    Exit Function
EH:
    result.ok = False
    result.status = 0
    result.body = ""
    If Len(result.errorDetail) > 0 Then result.errorDetail = result.errorDetail & " | "
    result.errorDetail = result.errorDetail & "MSXML: " & Err.Description
    GH_HTTP_RequestWithMsxml = False
End Function

Private Sub GH_HTTP_ApplyHeaders(ByVal http As Object, ByVal token As String, ByVal userAgent As String)
    http.setRequestHeader "Accept", "application/vnd.github+json"
    http.setRequestHeader "X-GitHub-Api-Version", "2022-11-28"
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "User-Agent", userAgent
    http.setRequestHeader "Content-Type", "application/json"
End Sub
