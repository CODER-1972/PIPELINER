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
