Attribute VB_Name = "M26_GH_Logger"
Option Explicit

' =============================================================================
' MÃ³dulo: M26_GH_Logger
' PropÃ³sito:
' - Uniformizar logs funcionais da exportaÃ§Ã£o GitHub no DEBUG.
' - Encapsular integraÃ§Ã£o com Debug_Registar para reduzir repetiÃ§Ã£o.
' - Garantir mensagens curtas e acionÃ¡veis para troubleshooting.
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | CriaÃ§Ã£o do logger dedicado GitHub
'   - Adiciona helpers GH_LogInfo/GH_LogWarn/GH_LogError.
'   - Normaliza parÃ¢metro e sugestÃ£o para eventos de exportaÃ§Ã£o.
'
' FunÃ§Ãµes e procedimentos:
' - GH_LogInfo(stepNo, promptId, paramName, message, suggestion) (Sub)
'   - Regista evento INFO da integraÃ§Ã£o GitHub.
' - GH_LogWarn(stepNo, promptId, paramName, message, suggestion) (Sub)
'   - Regista evento ALERTA da integraÃ§Ã£o GitHub.
' - GH_LogError(stepNo, promptId, paramName, message, suggestion) (Sub)
'   - Regista evento ERRO da integraÃ§Ã£o GitHub.
' =============================================================================

Public Sub GH_LogInfo(ByVal stepNo As Long, ByVal promptId As String, ByVal paramName As String, ByVal message As String, Optional ByVal suggestion As String = "")
    Call Debug_Registar(stepNo, promptId, "INFO", "", paramName, message, suggestion)
End Sub

Public Sub GH_LogWarn(ByVal stepNo As Long, ByVal promptId As String, ByVal paramName As String, ByVal message As String, Optional ByVal suggestion As String = "")
    Call Debug_Registar(stepNo, promptId, "ALERTA", "", paramName, message, suggestion)
End Sub

Public Sub GH_LogError(ByVal stepNo As Long, ByVal promptId As String, ByVal paramName As String, ByVal message As String, Optional ByVal suggestion As String = "")
    Call Debug_Registar(stepNo, promptId, "ERRO", "", paramName, message, suggestion)
End Sub
