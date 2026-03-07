Attribute VB_Name = "M26_GH_Logger"
Option Explicit

' =============================================================================
' Modulo: M26_GH_Logger
' Proposito:
' - Centralizar eventos/codigos canonicos GH_* para troubleshooting previsivel.
' - Encapsular chamadas a Debug_Registar para reduzir duplicacao de boilerplate.
' - Garantir mensagens curtas, sem segredos e com sugestoes acionaveis.
'
' Atualizacoes:
' - 2026-03-07 | Codex | Corrige constantes de evento GH para evitar erro de compilacao
'   - Substitui aliases legados nao declarados (GH_UNKNOWN/GH_CONFIG_OK/...) por codigos canonicos GH_EVT_*.
'   - Normaliza mapeamento em GH_NormalizeEventCode para nomes realmente definidos no modulo.
' - 2026-03-04 | Codex | Refactor de logging GitHub para modulo dedicado
'   - Move codigos de eventos GH_* para constantes publicas reutilizaveis.
'   - Mantem wrappers GH_LogInfo/GH_LogWarn/GH_LogError para padronizacao.
'
' Funcoes e procedimentos:
' - GH_LogInfo(stepNo, pipelineNome, eventCode, message, suggestion) (Sub)
'   - Regista evento INFO no DEBUG com codigo canonico.
' - GH_LogWarn(stepNo, pipelineNome, eventCode, message, suggestion) (Sub)
'   - Regista evento ALERTA no DEBUG com codigo canonico.
' - GH_LogError(stepNo, pipelineNome, eventCode, message, suggestion) (Sub)
'   - Regista evento ERRO no DEBUG com codigo canonico.
' =============================================================================

Public Const GH_EVT_CONFIG As String = "GH_CONFIG"
Public Const GH_EVT_UPLOAD As String = "GH_UPLOAD"
Public Const GH_EVT_HTTP As String = "GH_HTTP"
Public Const GH_EVT_HTTP_FAIL As String = "GH_HTTP_FAIL"
Public Const GH_EVT_REF_OK As String = "GH_REF_OK"
Public Const GH_EVT_BASE_TREE_OK As String = "GH_BASE_TREE_OK"
Public Const GH_EVT_BLOB_OK As String = "GH_BLOB_OK"
Public Const GH_EVT_BLOB_TOO_LARGE As String = "GH_BLOB_TOO_LARGE"
Public Const GH_EVT_TREE_CREATED As String = "GH_TREE_CREATED"
Public Const GH_EVT_COMMIT_CREATED As String = "GH_COMMIT_CREATED"
Public Const GH_EVT_REF_UPDATED As String = "GH_REF_UPDATED"
Public Const GH_EVT_MAX_FILES As String = "GH_MAX_FILES"

Public Sub GH_LogInfo(ByVal stepNo As Long, ByVal pipelineNome As String, ByVal eventCode As String, ByVal message As String, Optional ByVal suggestion As String = "")
    Call Debug_Registar(stepNo, pipelineNome, "INFO", "", eventCode, message, suggestion)
End Sub

Public Sub GH_LogWarn(ByVal stepNo As Long, ByVal pipelineNome As String, ByVal eventCode As String, ByVal message As String, Optional ByVal suggestion As String = "")
    Call Debug_Registar(stepNo, pipelineNome, "ALERTA", "", eventCode, message, suggestion)
End Sub

Public Sub GH_LogError(ByVal stepNo As Long, ByVal pipelineNome As String, ByVal eventCode As String, ByVal message As String, Optional ByVal suggestion As String = "")
    Call Debug_Registar(stepNo, pipelineNome, "ERRO", "", eventCode, message, suggestion)
End Sub

Private Function GH_NormalizeEventCode(ByVal eventCode As String) As String
    Dim code As String
    code = UCase$(Trim$(eventCode))

    If code = "" Then
        GH_NormalizeEventCode = GH_EVT_CONFIG
        Exit Function
    End If

    Select Case code
        Case GH_EVT_CONFIG, GH_EVT_UPLOAD, GH_EVT_HTTP, GH_EVT_HTTP_FAIL, _
             GH_EVT_REF_OK, GH_EVT_BASE_TREE_OK, GH_EVT_BLOB_OK, GH_EVT_BLOB_TOO_LARGE, _
             GH_EVT_TREE_CREATED, GH_EVT_COMMIT_CREATED, GH_EVT_REF_UPDATED, GH_EVT_MAX_FILES
            GH_NormalizeEventCode = code
        Case "CONFIG_OK", "GH_CFG_OK", "GH_CONFIG_OK"
            GH_NormalizeEventCode = GH_EVT_CONFIG
        Case "REF_OK", "GH_REFERENCE_OK", "GH_REF_OK"
            GH_NormalizeEventCode = GH_EVT_REF_OK
        Case "TREE_OK", "GH_TREE_OK", "GH_BLOB_TREE_OK"
            GH_NormalizeEventCode = GH_EVT_TREE_CREATED
        Case "COMMIT_OK", "GH_COMMIT_OK", "GH_COMMIT_CREATED"
            GH_NormalizeEventCode = GH_EVT_COMMIT_CREATED
        Case "HTTP_FAIL", "GH_API_FAIL", "GH_HTTP_FAIL"
            GH_NormalizeEventCode = GH_EVT_HTTP_FAIL
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
