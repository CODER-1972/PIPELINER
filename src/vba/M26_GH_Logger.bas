Attribute VB_Name = "M26_GH_Logger"
Option Explicit

' =============================================================================
' Mï¿½dulo: M26_GH_Logger
' Propï¿½sito:
' - Registar falhas de integraï¿½ï¿½o GitHub em formato curto e acionï¿½vel.
' - Padronizar campos mï¿½nimos de troubleshooting HTTP (step, endpoint, status, snippet).
' - Evitar que erros de logging interrompam o fluxo principal.
'
' Atualizaï¿½ï¿½es:
' - 2026-03-04 | Codex | Criaï¿½ï¿½o do logger de falhas HTTP para integraï¿½ï¿½o GitHub
'   - Implementa GH_LogHttpFailure(step_name, endpoint, http_status, response_snippet).
'   - Encaminha eventos para a folha DEBUG via Debug_Registar com severidade ERRO.
'
' Funï¿½ï¿½es e procedimentos:
' - GH_LogHttpFailure(step_name, endpoint, http_status, response_snippet): grava falha HTTP no DEBUG.
' =============================================================================

Private Const GH_LOG_PARAM As String = "GH_HTTP"
Private Const GH_LOG_SNIPPET_MAX As Long = 500

Public Sub GH_LogHttpFailure( _
    ByVal step_name As String, _
    ByVal endpoint As String, _
    ByVal http_status As Long, _
    ByVal response_snippet As String _
)
    On Error GoTo EH

    Dim msg As String
    msg = "step_name=" & GH_Safe(step_name) & _
          " | endpoint=" & GH_Safe(endpoint) & _
          " | http_status=" & CStr(http_status) & _
          " | response_snippet=" & GH_Short(response_snippet, GH_LOG_SNIPPET_MAX)

    Debug_Registar 0, "", "ERRO", "", GH_LOG_PARAM, msg, "Rever permissï¿½es/token, endpoint e payload enviado ao GitHub."
    Exit Sub
EH:
    On Error GoTo 0
End Sub

Private Function GH_Short(ByVal rawText As String, ByVal maxLen As Long) As String
    Dim txt As String
    txt = Replace(rawText, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Trim$(txt)

    If maxLen <= 3 Then maxLen = 3
    If Len(txt) > maxLen Then
        GH_Short = Left$(txt, maxLen - 3) & "..."
    Else
        GH_Short = txt
    End If
End Function

Private Function GH_Safe(ByVal value As String) As String
    GH_Safe = Trim$(value & "")
End Function
