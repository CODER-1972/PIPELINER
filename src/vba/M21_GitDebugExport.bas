Attribute VB_Name = "M21_GitDebugExport"
Option Explicit

' =============================================================================
' MÃ³dulo: M21_GitDebugExport
' PropÃ³sito:
' - Orquestrar exportaÃ§Ã£o opcional dos logs DEBUG/Seguimento para GitHub.
' - Manter a assinatura pÃºblica estÃ¡vel para chamadas de outros mÃ³dulos.
' - Delegar responsabilidades por camadas (config/http/blob/tree/logger).
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | RefatoraÃ§Ã£o para facade de alto nÃ­vel
'   - Preserva entry point pÃºblico PipelineGitDebug_ExportIfEnabled.
'   - Move leitura de config, HTTP, blob/base64 e logging para mÃ³dulos M22-M26.
'   - MantÃ©m fluxo resiliente: validaÃ§Ãµes + logs acionÃ¡veis sem quebrar pipeline.
'
' FunÃ§Ãµes e procedimentos:
' - PipelineGitDebug_ExportIfEnabled(Optional pipelineIndex As Long = 0) (Sub)
'   - Entry point pÃºblico; avalia enable/config e executa exportaÃ§Ã£o para GitHub.
' =============================================================================

Public Sub PipelineGitDebug_ExportIfEnabled(Optional ByVal pipelineIndex As Long = 0)
    On Error GoTo EH

    Dim cfg As Object
    Set cfg = GH_Config_Load()

    If Not GH_Config_IsEnabled(cfg) Then
        Exit Sub
    End If

    Dim invalidReason As String
    If Not GH_Config_Validate(cfg, invalidReason) Then
        Call GH_LogWarn(0, "DEBUG", "GIT_DEBUG_EXPORT_DISABLED", invalidReason, _
                        "Preencha as chaves GIT_DEBUG_* na folha Config ou desative a feature.")
        Exit Sub
    End If

    Dim markdown As String
    markdown = M21_BuildDebugMarkdown(pipelineIndex)
    If Len(markdown) = 0 Then
        Call GH_LogWarn(0, "DEBUG", "GIT_DEBUG_EXPORT_EMPTY", _
                        "Sem conteÃºdo para exportar (DEBUG/Seguimento vazios).", _
                        "Confirme se a execuÃ§Ã£o gerou logs antes de exportar.")
        Exit Sub
    End If

    Dim url As String
    url = GH_TreeCommit_ContentsUrl( _
            GH_Config_GetString(cfg, "base_url"), _
            GH_Config_GetString(cfg, "owner"), _
            GH_Config_GetString(cfg, "repo"), _
            GH_Config_GetString(cfg, "path"))

    Dim payload As String
    payload = GH_TreeCommit_BuildContentsPayload( _
                GH_TreeCommit_DefaultMessage(pipelineIndex), _
                GH_Config_GetString(cfg, "branch"), _
                GH_Blob_Base64FromText(markdown))

    Dim statusCode As Long
    Dim responseText As String
    Dim errText As String
    Dim ok As Boolean

    ok = GH_HTTP_RequestJson("PUT", url, GH_Config_GetString(cfg, "token"), payload, _
                             statusCode, responseText, errText, GH_Config_GetString(cfg, "user_agent"))

    If ok Then
        Call GH_LogInfo(0, "DEBUG", "GIT_DEBUG_EXPORT_OK", _
                        "ExportaÃ§Ã£o DEBUG enviada para GitHub (HTTP " & CStr(statusCode) & ").", _
                        "Verifique o ficheiro no repositÃ³rio remoto.")
    Else
        Call GH_LogError(0, "DEBUG", "GIT_DEBUG_EXPORT_FAIL", _
                         "Falha ao exportar DEBUG para GitHub (HTTP " & CStr(statusCode) & "). " & _
                         M21_LeftSafe(responseText, 220), _
                         "Valide token/permissÃµes, owner/repo/path e branch configurados.")
        If Len(errText) > 0 Then
            Call GH_LogWarn(0, "DEBUG", "GIT_DEBUG_EXPORT_ENGINE", errText, _
                            "Confirme suporte WinHTTP/MSXML no host Office.")
        End If
    End If

    Exit Sub
EH:
    Call GH_LogError(0, "DEBUG", "GIT_DEBUG_EXPORT_EXCEPTION", _
                     "Erro inesperado no export GitHub: " & Err.Description, _
                     "Revise as configuraÃ§Ãµes GIT_DEBUG_* e o estado das folhas DEBUG/Seguimento.")
End Sub

Private Function M21_BuildDebugMarkdown(ByVal pipelineIndex As Long) As String
    Dim nowText As String
    nowText = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Dim debugSection As String
    Dim followSection As String
    debugSection = M21_SheetAsTsv("DEBUG")
    followSection = M21_SheetAsTsv("Seguimento")

    If Len(debugSection) = 0 And Len(followSection) = 0 Then
        M21_BuildDebugMarkdown = ""
        Exit Function
    End If

    Dim out As String
    out = "# PIPELINER Debug Export" & vbCrLf & vbCrLf & _
          "- Generated at: " & nowText & vbCrLf & _
          "- Pipeline index: " & CStr(pipelineIndex) & vbCrLf & vbCrLf & _
          "## DEBUG" & vbCrLf & "```tsv" & vbCrLf & debugSection & vbCrLf & "```" & vbCrLf & vbCrLf & _
          "## Seguimento" & vbCrLf & "```tsv" & vbCrLf & followSection & vbCrLf & "```"

    M21_BuildDebugMarkdown = out
End Function

Private Function M21_SheetAsTsv(ByVal sheetName As String) As String
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim usedRg As Range
    Set usedRg = ws.UsedRange

    If usedRg Is Nothing Then
        M21_SheetAsTsv = ""
        Exit Function
    End If

    Dim rowsCount As Long
    Dim colsCount As Long
    rowsCount = usedRg.Rows.Count
    colsCount = usedRg.Columns.Count

    Dim r As Long
    Dim c As Long
    Dim sb As String
    Dim lineText As String

    For r = 1 To rowsCount
        lineText = ""
        For c = 1 To colsCount
            If c > 1 Then lineText = lineText & vbTab
            lineText = lineText & Replace$(Replace$(CStr(usedRg.Cells(r, c).Value), vbCrLf, " "), vbTab, " ")
        Next c
        sb = sb & lineText & vbCrLf
    Next r

    M21_SheetAsTsv = Trim$(sb)
    Exit Function
EH:
    M21_SheetAsTsv = ""
End Function

Private Function M21_LeftSafe(ByVal text As String, ByVal maxLen As Long) As String
    If maxLen <= 0 Then
        M21_LeftSafe = ""
    ElseIf Len(text) <= maxLen Then
        M21_LeftSafe = text
    Else
        M21_LeftSafe = Left$(text, maxLen) & "..."
    End If
End Function
