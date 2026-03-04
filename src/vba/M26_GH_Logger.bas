Attribute VB_Name = "M26_GH_Logger"
Option Explicit

' =============================================================================
' MÃ³dulo: M26_GH_Logger
' PropÃ³sito:
' - Centralizar logs do fluxo GitHub com esquema canÃ³nico estÃ¡vel.
' - Integrar registos GH na folha DEBUG reutilizando o logger existente (M02).
' - Sanitizar conteÃºdo sensÃ­vel (tokens/segredos) antes de persistir em DEBUG.
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | CriaÃ§Ã£o do logger canÃ³nico para o fluxo GitHub
'   - Adiciona API pÃºblica LogInfo/LogWarn/LogError/LogHttpFail.
'   - Introduz cÃ³digos canÃ³nicos GH_* e normalizaÃ§Ã£o de aliases legados.
'   - Garante sanitizaÃ§Ã£o de segredos antes de chamar Debug_Registar.
'
' FunÃ§Ãµes e procedimentos:
' - LogInfo(...): regista evento GH com severidade INFO.
' - LogWarn(...): regista evento GH com severidade ALERTA.
' - LogError(...): regista evento GH com severidade ERRO.
' - LogHttpFail(...): regista falha HTTP GH com severidade ERRO e contexto resumido.
' =============================================================================

Private Const GH_COMPONENT_DEFAULT As String = "GH_FLOW"
Private Const GH_DEBUG_FEATURE As String = "IntegraÃ§Ã£o GitHub: configuraÃ§Ã£o, referÃªncia, Ã¡rvore e commit."

' CÃ³digos canÃ³nicos (core)
Public Const GH_CONFIG_OK As String = "GH_CONFIG_OK"
Public Const GH_REF_OK As String = "GH_REF_OK"
Public Const GH_TREE_OK As String = "GH_TREE_OK"
Public Const GH_COMMIT_OK As String = "GH_COMMIT_OK"
Public Const GH_DONE_OK As String = "GH_DONE_OK"
Public Const GH_DONE_FAIL As String = "GH_DONE_FAIL"
Public Const GH_HTTP_FAIL As String = "GH_HTTP_FAIL"
Public Const GH_UNKNOWN As String = "GH_UNKNOWN"

Public Sub LogInfo( _
    ByVal pipelineName As String, _
    ByVal runId As String, _
    ByVal component As String, _
    ByVal eventCode As String, _
    ByVal details As String _
)
    GH_LogEvent pipelineName, runId, component, eventCode, "INFO", details, 0
End Sub

Public Sub LogWarn( _
    ByVal pipelineName As String, _
    ByVal runId As String, _
    ByVal component As String, _
    ByVal eventCode As String, _
    ByVal details As String _
)
    GH_LogEvent pipelineName, runId, component, eventCode, "ALERTA", details, 0
End Sub

Public Sub LogError( _
    ByVal pipelineName As String, _
    ByVal runId As String, _
    ByVal component As String, _
    ByVal eventCode As String, _
    ByVal details As String _
)
    GH_LogEvent pipelineName, runId, component, eventCode, "ERRO", details, 0
End Sub

Public Sub LogHttpFail( _
    ByVal pipelineName As String, _
    ByVal runId As String, _
    ByVal component As String, _
    ByVal eventCode As String, _
    ByVal httpStatus As Long, _
    ByVal endpoint As String, _
    ByVal details As String _
)
    Dim msg As String
    msg = "http_status=" & CStr(httpStatus) & " | endpoint=" & GH_SafeText(endpoint)
    If Trim$(details) <> "" Then msg = msg & " | " & GH_SafeText(details)

    GH_LogEvent pipelineName, runId, component, eventCode, "ERRO", msg, httpStatus
End Sub

Private Sub GH_LogEvent( _
    ByVal pipelineName As String, _
    ByVal runId As String, _
    ByVal component As String, _
    ByVal eventCode As String, _
    ByVal severity As String, _
    ByVal details As String, _
    ByVal httpStatus As Long _
)
    On Error GoTo Fim

    Dim ts As String
    ts = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Dim sev As String
    sev = GH_NormalizeSeverity(severity)

    Dim code As String
    code = GH_NormalizeEventCode(eventCode)

    Dim compSafe As String
    compSafe = Trim$(component)
    If compSafe = "" Then compSafe = GH_COMPONENT_DEFAULT
    compSafe = GH_SafeText(compSafe)

    Dim detailSafe As String
    detailSafe = GH_SafeText(details)

    Dim problema As String
    problema = "details=" & detailSafe

    Dim sugestao As String
    sugestao = "timestamp=" & ts & _
               " | pipeline_name=" & GH_SafeText(pipelineName) & _
               " | run_id=" & GH_SafeText(runId) & _
               " | component=" & compSafe & _
               " | event_code=" & code & _
               " | severity=" & sev

    If httpStatus > 0 Then sugestao = sugestao & " | http_status=" & CStr(httpStatus)

    Debug_Registar 0, GH_SafePromptId(runId), sev, "", code, problema, sugestao, GH_DEBUG_FEATURE

Fim:
    On Error GoTo 0
End Sub

Private Function GH_NormalizeEventCode(ByVal eventCode As String) As String
    Dim code As String
    code = UCase$(Trim$(eventCode))

    If code = "" Then
        GH_NormalizeEventCode = GH_UNKNOWN
        Exit Function
    End If

    Select Case code
        Case GH_CONFIG_OK, GH_REF_OK, GH_TREE_OK, GH_COMMIT_OK, GH_DONE_OK, GH_DONE_FAIL, GH_HTTP_FAIL
            GH_NormalizeEventCode = code
        Case "CONFIG_OK", "GH_CFG_OK"
            GH_NormalizeEventCode = GH_CONFIG_OK
        Case "REF_OK", "GH_REFERENCE_OK"
            GH_NormalizeEventCode = GH_REF_OK
        Case "TREE_OK", "GH_BLOB_TREE_OK"
            GH_NormalizeEventCode = GH_TREE_OK
        Case "COMMIT_OK", "GH_COMMIT_CREATED"
            GH_NormalizeEventCode = GH_COMMIT_OK
        Case "DONE_FAIL", "GH_FAILED"
            GH_NormalizeEventCode = GH_DONE_FAIL
        Case "HTTP_FAIL", "GH_API_FAIL"
            GH_NormalizeEventCode = GH_HTTP_FAIL
        Case Else
            GH_NormalizeEventCode = code
    End Select
End Function

Private Function GH_NormalizeSeverity(ByVal severity As String) As String
    Dim s As String
    s = UCase$(Trim$(severity))

    Select Case s
        Case "INFO", "INFORMACAO", "INFORMAÃ‡ÃƒO"
            GH_NormalizeSeverity = "INFO"
        Case "WARN", "WARNING", "ALERTA"
            GH_NormalizeSeverity = "ALERTA"
        Case "ERR", "ERROR", "ERRO"
            GH_NormalizeSeverity = "ERRO"
        Case Else
            GH_NormalizeSeverity = "INFO"
    End Select
End Function

Private Function GH_SafePromptId(ByVal runId As String) As String
    Dim s As String
    s = Trim$(runId)
    If s = "" Then s = "GH"
    GH_SafePromptId = Left$(GH_SafeText(s), 255)
End Function

Private Function GH_SafeText(ByVal valueText As String) As String
    Dim s As String
    s = CStr(valueText)

    ' RedaÃ§Ã£o de padrÃµes comuns de segredos/tokens.
    s = GH_RedactAfterKey(s, "token")
    s = GH_RedactAfterKey(s, "authorization")
    s = GH_RedactAfterKey(s, "api_key")
    s = GH_RedactAfterKey(s, "apikey")
    s = GH_RedactAfterKey(s, "password")

    s = GH_RedactPattern(s, "ghp_")
    s = GH_RedactPattern(s, "github_pat_")
    s = GH_RedactPattern(s, "gho_")
    s = GH_RedactPattern(s, "sk-")

    GH_SafeText = Left$(s, 1800)
End Function

Private Function GH_RedactAfterKey(ByVal txt As String, ByVal keyName As String) As String
    Dim s As String
    s = txt

    Dim look As String
    look = LCase$(s)

    Dim p As Long
    p = InStr(1, look, LCase$(keyName) & "=", vbTextCompare)
    If p = 0 Then p = InStr(1, look, LCase$(keyName) & ":", vbTextCompare)

    If p > 0 Then
        Dim pStart As Long
        pStart = p + Len(keyName) + 1

        Dim pEnd As Long
        pEnd = GH_FindSeparatorPos(s, pStart)
        If pEnd = 0 Then pEnd = Len(s) + 1

        s = Left$(s, pStart - 1) & "[REDACTED]" & Mid$(s, pEnd)
    End If

    GH_RedactAfterKey = s
End Function

Private Function GH_FindSeparatorPos(ByVal txt As String, ByVal startPos As Long) As Long
    Dim i As Long
    For i = startPos To Len(txt)
        Dim ch As String
        ch = Mid$(txt, i, 1)
        If ch = " " Or ch = "|" Or ch = ";" Or ch = "," Or ch = vbCr Or ch = vbLf Then
            GH_FindSeparatorPos = i
            Exit Function
        End If
    Next i
End Function

Private Function GH_RedactPattern(ByVal txt As String, ByVal tokenPrefix As String) As String
    Dim s As String
    s = txt

    Dim p As Long
    p = InStr(1, s, tokenPrefix, vbTextCompare)

    Do While p > 0
        Dim i As Long
        Dim j As Long
        j = p + Len(tokenPrefix)

        For i = j To Len(s)
            Dim ch As String
            ch = Mid$(s, i, 1)
            If Not GH_IsTokenChar(ch) Then
                Exit For
            End If
        Next i

        s = Left$(s, p - 1) & "[REDACTED_TOKEN]" & Mid$(s, i)
        p = InStr(p + Len("[REDACTED_TOKEN]"), s, tokenPrefix, vbTextCompare)
    Loop

    GH_RedactPattern = s
End Function

Private Function GH_IsTokenChar(ByVal ch As String) As Boolean
    If ch = "" Then Exit Function

    Dim c As Integer
    c = AscW(ch)

    GH_IsTokenChar = ((c >= 48 And c <= 57) Or _
                      (c >= 65 And c <= 90) Or _
                      (c >= 97 And c <= 122) Or _
                      ch = "_" Or ch = "-")
End Function
