Attribute VB_Name = "M22_GH_Config"
Option Explicit

' =============================================================================
' MÃ³dulo: M22_GH_Config
' PropÃ³sito:
' - Centralizar leitura de configuraÃ§Ã£o para exportaÃ§Ã£o de DEBUG para GitHub.
' - Fornecer defaults internos com tolerÃ¢ncia a Config incompleta.
' - Evitar dependÃªncias diretas de cÃ©lulas fixas fora deste mÃ³dulo.
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | CriaÃ§Ã£o do mÃ³dulo de configuraÃ§Ã£o GitHub
'   - Adiciona loader de configuraÃ§Ã£o com fallback por labels na folha Config.
'   - Implementa validaÃ§Ã£o mÃ­nima de obrigatoriedade e flag de enable.
'
' FunÃ§Ãµes e procedimentos:
' - GH_Config_Load() As Object
'   - Devolve dictionary com parÃ¢metros normalizados da exportaÃ§Ã£o GitHub.
' - GH_Config_IsEnabled(cfg As Object) As Boolean
'   - Informa se a exportaÃ§Ã£o estÃ¡ ativa para a execuÃ§Ã£o atual.
' - GH_Config_Validate(cfg As Object, reason As String) As Boolean
'   - Valida campos obrigatÃ³rios e devolve motivo curto em caso de bloqueio.
' =============================================================================

Public Function GH_Config_Load() As Object
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0

    cfg("enabled") = GH_Config_ToBoolean(GH_Config_ReadValue(ws, "GIT_DEBUG_EXPORT_ENABLED", "FALSE"), False)
    cfg("owner") = GH_Config_ReadValue(ws, "GIT_DEBUG_OWNER", "")
    cfg("repo") = GH_Config_ReadValue(ws, "GIT_DEBUG_REPO", "")
    cfg("branch") = GH_Config_ReadValue(ws, "GIT_DEBUG_BRANCH", "main")
    cfg("path") = GH_Config_ReadValue(ws, "GIT_DEBUG_PATH", "logs/debug_export.md")
    cfg("token") = GH_Config_ReadValue(ws, "GIT_DEBUG_TOKEN", "")
    cfg("base_url") = GH_Config_ReadValue(ws, "GIT_DEBUG_API_BASE_URL", "https://api.github.com")
    cfg("user_agent") = GH_Config_ReadValue(ws, "GIT_DEBUG_USER_AGENT", "PIPELINER-GitDebugExport")

    Set GH_Config_Load = cfg
End Function

Public Function GH_Config_IsEnabled(ByVal cfg As Object) As Boolean
    On Error GoTo Fallback
    GH_Config_IsEnabled = GH_Config_ToBoolean(cfg("enabled"), False)
    Exit Function
Fallback:
    GH_Config_IsEnabled = False
End Function

Public Function GH_Config_Validate(ByVal cfg As Object, ByRef reason As String) As Boolean
    reason = ""

    If Len(Trim$(GH_Config_GetString(cfg, "owner"))) = 0 Then
        reason = "Config em falta: GIT_DEBUG_OWNER"
        Exit Function
    End If

    If Len(Trim$(GH_Config_GetString(cfg, "repo"))) = 0 Then
        reason = "Config em falta: GIT_DEBUG_REPO"
        Exit Function
    End If

    If Len(Trim$(GH_Config_GetString(cfg, "token"))) = 0 Then
        reason = "Config em falta: GIT_DEBUG_TOKEN"
        Exit Function
    End If

    GH_Config_Validate = True
End Function

Public Function GH_Config_GetString(ByVal cfg As Object, ByVal key As String) As String
    On Error GoTo Fallback
    GH_Config_GetString = Trim$(CStr(cfg(key)))
    Exit Function
Fallback:
    GH_Config_GetString = ""
End Function

Private Function GH_Config_ReadValue(ByVal ws As Worksheet, ByVal keyName As String, ByVal defaultValue As String) As String
    If ws Is Nothing Then
        GH_Config_ReadValue = defaultValue
        Exit Function
    End If

    Dim found As Range
    On Error Resume Next
    Set found = ws.Columns(1).Find(What:=keyName, LookIn:=xlValues, LookAt:=xlWhole, _
                                   SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    On Error GoTo 0

    If Not found Is Nothing Then
        GH_Config_ReadValue = Trim$(CStr(ws.Cells(found.Row, 2).Value))
        If Len(GH_Config_ReadValue) = 0 Then GH_Config_ReadValue = defaultValue
        Exit Function
    End If

    GH_Config_ReadValue = defaultValue
End Function

Private Function GH_Config_ToBoolean(ByVal value As Variant, ByVal defaultValue As Boolean) As Boolean
    Dim raw As String
    raw = UCase$(Trim$(CStr(value)))

    If raw = "TRUE" Or raw = "1" Or raw = "SIM" Or raw = "YES" Then
        GH_Config_ToBoolean = True
    ElseIf raw = "FALSE" Or raw = "0" Or raw = "NAO" Or raw = "NÃƒO" Or raw = "NO" Then
        GH_Config_ToBoolean = False
    Else
        GH_Config_ToBoolean = defaultValue
    End If
End Function
