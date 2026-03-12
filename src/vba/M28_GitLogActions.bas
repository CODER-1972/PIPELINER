Attribute VB_Name = "M28_GitLogActions"
Option Explicit

' =============================================================================
' Modulo: M28_GitLogActions
' Proposito:
' - Gerir a coluna "Eliminar" na folha GIT LOG com acao por linha.
' - Resolver paths remotos associados a cada registo e executar delete no GitHub.
' - Aplicar politica de remocao local apos delete remoto com logging GH_* no DEBUG.
'
' Atualizacoes:
' - 2026-03-11 | Codex | Implementa delete por linha no GIT LOG
'   - Adiciona criacao/refresh da coluna Eliminar com botao/shape por linha para GitLog_DeleteEntry.
'   - Persiste paths tecnicos em coluna oculta GH_REMOTE_PATHS (e SHA opcional) por linha.
'   - Executa delete remoto com retries para 409/422 e tratamento acionavel de 404/409/422.
'
' Funcoes e procedimentos:
' - GitLog_EnsureDeleteColumn(Optional sheetName As String = "GIT LOG")
'   - Garante colunas tecnicas/cabecalhos e cria links "Eliminar" por linha.
' - GitLog_BindRemotePaths(targetRow As Long, remotePaths As String, Optional remoteShas As String = "", Optional sheetName As String = "GIT LOG")
'   - Guarda metadados tecnicos de delete remoto na linha de log.
' - GitLog_DeleteEntry()
'   - Entry point para acao da linha: apaga remoto e remove/manter linha local segundo politica.
' - GitLog_PurgeDeleteButtons/GitLog_UpsertDeleteButton (Private Sub)
'   - Mantem a UI idempotente, removendo botoes antigos e recriando um por linha valida.
' =============================================================================

Private Const GITLOG_SHEET As String = "GIT LOG"
Private Const HDR_ELIMINAR As String = "Eliminar"
Private Const HDR_REMOTE_PATHS As String = "GH_REMOTE_PATHS"
Private Const HDR_REMOTE_SHAS As String = "GH_REMOTE_SHAS"
Private Const HDR_DELETE_STATUS As String = "DELETE_STATUS"
Private Const BTN_DELETE_PREFIX As String = "BTN_GITLOG_DELETE_"

Public Sub GitLog_EnsureDeleteColumn(Optional ByVal sheetName As String = GITLOG_SHEET)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim cEliminar As Long
    Dim cPaths As Long
    Dim cShas As Long
    Dim cStatus As Long

    cEliminar = GitLog_EnsureHeader(ws, HDR_ELIMINAR)
    cPaths = GitLog_EnsureHeader(ws, HDR_REMOTE_PATHS)
    cShas = GitLog_EnsureHeader(ws, HDR_REMOTE_SHAS)
    cStatus = GitLog_EnsureHeader(ws, HDR_DELETE_STATUS)

    ws.Columns(cPaths).Hidden = True
    ws.Columns(cShas).Hidden = True

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Call GitLog_PurgeDeleteButtons(ws)

    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            ws.Cells(r, cEliminar).Value = "Eliminar"
            Call GitLog_UpsertDeleteButton(ws, r, cEliminar)

            If Trim$(CStr(ws.Cells(r, cStatus).Value)) = "" Then ws.Cells(r, cStatus).Value = "PENDENTE"
        End If
    Next r

    Exit Sub
EH:
    Call GH_LogError(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Falha ao preparar coluna Eliminar no GIT LOG.", "err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 180))
End Sub

Public Sub GitLog_BindRemotePaths( _
    ByVal targetRow As Long, _
    ByVal remotePaths As String, _
    Optional ByVal remoteShas As String = "", _
    Optional ByVal sheetName As String = GITLOG_SHEET)

    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim cPaths As Long
    Dim cShas As Long
    cPaths = GitLog_EnsureHeader(ws, HDR_REMOTE_PATHS)
    cShas = GitLog_EnsureHeader(ws, HDR_REMOTE_SHAS)

    ws.Cells(targetRow, cPaths).Value = Trim$(remotePaths)
    ws.Cells(targetRow, cShas).Value = Trim$(remoteShas)

    Dim cEliminar As Long
    cEliminar = GitLog_EnsureHeader(ws, HDR_ELIMINAR)
    ws.Cells(targetRow, cEliminar).Value = "Eliminar"
    Call GitLog_UpsertDeleteButton(ws, targetRow, cEliminar)

    Exit Sub
EH:
    Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Falha ao guardar metadados tecnicos de delete no GIT LOG.", "row=" & CStr(targetRow) & " | err=" & Left$(Err.Description, 180))
End Sub

Public Sub GitLog_DeleteEntry()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Or StrComp(ws.Name, GITLOG_SHEET, vbTextCompare) <> 0 Then
        Set ws = ThisWorkbook.Worksheets(GITLOG_SHEET)
    End If

    Dim rowTarget As Long
    rowTarget = GitLog_ResolveTargetRow(ws)
    If rowTarget < 2 Then
        Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Linha alvo invalida para delete no GIT LOG.", "action=Selecionar a celula da linha/usar hyperlink Eliminar.")
        Exit Sub
    End If

    Dim cPaths As Long
    Dim cShas As Long
    Dim cStatus As Long
    cPaths = GitLog_EnsureHeader(ws, HDR_REMOTE_PATHS)
    cShas = GitLog_EnsureHeader(ws, HDR_REMOTE_SHAS)
    cStatus = GitLog_EnsureHeader(ws, HDR_DELETE_STATUS)

    Dim rawPaths As String
    rawPaths = Trim$(CStr(ws.Cells(rowTarget, cPaths).Value))
    If rawPaths = "" Then
        rawPaths = GitLog_ReadPathsFromComment(ws.Cells(rowTarget, cPaths))
    End If
    If rawPaths = "" Then rawPaths = GitLog_DerivePathsFromGitDebugUrl(ws, rowTarget)

    If rawPaths = "" Then
        ws.Cells(rowTarget, cStatus).Value = "ERRO: sem paths"
        Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Linha sem GH_REMOTE_PATHS para delete.", "row=" & CStr(rowTarget) & " | action=Guardar path remoto na coluna tecnica oculta.")
        Exit Sub
    End If

    Dim cfg As Object
    Set cfg = GH_Config_Load("debug")
    cfg("enabled") = True

    Dim reason As String
    If Not GH_Config_Validate(cfg, reason) Then
        ws.Cells(rowTarget, cStatus).Value = "ERRO: config"
        Call GH_LogError(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Configuracao GH invalida para delete no GIT LOG.", reason)
        Exit Sub
    End If

    Dim maxRetries As Long
    maxRetries = GH_Config_GetLong(cfg, "max_retries", 3)
    If maxRetries < 0 Then maxRetries = 0

    Dim allOk As Boolean
    allOk = True

    Dim paths As Variant
    paths = Split(rawPaths, ";")

    Dim shas As Variant
    shas = Split(CStr(ws.Cells(rowTarget, cShas).Value), ";")

    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        Dim repoPath As String
        repoPath = GitLog_NormalizePath(CStr(paths(i)))
        If repoPath = "" Then GoTo ContinuePath

        Dim fileSha As String
        fileSha = ""
        If i <= UBound(shas) Then fileSha = Trim$(CStr(shas(i)))

        Dim attempt As Long
        Dim okDelete As Boolean
        Dim statusCode As Long
        Dim diag As String

        For attempt = 0 To maxRetries
            If fileSha = "" Then
                Call GH_ContentsApi_GetFileSha(cfg, repoPath, "", fileSha, statusCode, diag)
                If statusCode = 404 Then
                    okDelete = True
                    Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_NOT_FOUND, "Ficheiro remoto ja nao existe (404).", "path=" & repoPath & " | action=Linha pode ser removida localmente.")
                    Exit For
                End If
            End If

            okDelete = GH_ContentsApi_DeleteFile(cfg, repoPath, fileSha, "Delete from GIT LOG row " & CStr(rowTarget), "", statusCode, diag)
            If okDelete Then Exit For

            If statusCode = 404 Then
                okDelete = True
                Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_NOT_FOUND, "Delete devolveu 404; ficheiro considerado ausente.", "path=" & repoPath & " | action=Validar se ja foi removido noutra execucao.")
                Exit For
            End If

            If (statusCode = 409 Or statusCode = 422) And attempt < maxRetries Then
                Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_RETRY, "Retry agendado para delete de ficheiro remoto.", "path=" & repoPath & " | status=" & CStr(statusCode) & " | tentativa=" & CStr(attempt + 1) & "/" & CStr(maxRetries))
                fileSha = ""
            Else
                Exit For
            End If
        Next attempt

        If Not okDelete Then
            allOk = False
            ws.Cells(rowTarget, cStatus).Value = "ERRO remoto: " & CStr(statusCode)
            Call GH_LogError(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Falha no delete remoto a partir do GIT LOG.", "path=" & repoPath & " | " & diag)
        Else
            Call GH_LogInfo(0, "", GH_EVT_GITLOG_DELETE_OK, "Delete remoto concluido para linha do GIT LOG.", "path=" & repoPath)
        End If

ContinuePath:
    Next i

    Dim policy As String
    policy = LCase$(Trim$(GH_Config_Get("GH_GITLOG_DELETE_POLICY", "after_remote_success")))
    If policy = "" Then policy = "after_remote_success"

    Select Case policy
        Case "after_remote_success", "always", "keep_local"
            ' ok
        Case Else
            Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_ROW_KEPT, "GH_GITLOG_DELETE_POLICY invalida; aplicado fallback seguro.", "raw=" & policy & " | normalized=after_remote_success")
            policy = "after_remote_success"
    End Select

    If (allOk And policy = "after_remote_success") Or policy = "always" Then
        ws.Rows(rowTarget).Delete
        Call GH_LogInfo(0, "", GH_EVT_GITLOG_DELETE_ROW_REMOVED, "Linha removida do GIT LOG apos delete remoto.", "row=" & CStr(rowTarget) & " | policy=" & policy)
    Else
        If allOk Then
            ws.Cells(rowTarget, cStatus).Value = "REMOTO_OK"
        Else
            ws.Cells(rowTarget, cStatus).Value = "ERRO remoto"
        End If
        Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_ROW_KEPT, "Linha mantida no GIT LOG apos delete.", "row=" & CStr(rowTarget) & " | policy=" & policy & " | all_ok=" & IIf(allOk, "SIM", "NAO"))
    End If

    Exit Sub

EH:
    Call GH_LogError(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Erro inesperado em GitLog_DeleteEntry.", "err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 180))
End Sub

Private Function GitLog_ResolveTargetRow(ByVal ws As Worksheet) As Long
    On Error GoTo Fallback

    Dim callerName As String
    callerName = CStr(Application.Caller)

    If callerName <> "" Then
        Dim shp As Shape
        Set shp = ws.Shapes(callerName)
        GitLog_ResolveTargetRow = shp.TopLeftCell.Row
        Exit Function
    End If

Fallback:
    On Error Resume Next
    GitLog_ResolveTargetRow = ActiveCell.Row
End Function

Private Sub GitLog_UpsertDeleteButton(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal targetCol As Long)
    On Error GoTo EH

    Dim shpName As String
    shpName = BTN_DELETE_PREFIX & Format$(targetRow, "00000")

    On Error Resume Next
    ws.Shapes(shpName).Delete
    On Error GoTo EH

    Dim targetCell As Range
    Set targetCell = ws.Cells(targetRow, targetCol)

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(Type:=xlButtonControl, _
                                       Left:=targetCell.Left + 1, _
                                       Top:=targetCell.Top + 1, _
                                       Width:=Application.Max(targetCell.Width - 2, 8), _
                                       Height:=Application.Max(targetCell.Height - 2, 8))

    shp.Name = shpName
    shp.TextFrame.Characters.Text = "Eliminar"
    shp.OnAction = "GitLog_DeleteEntry"
    Exit Sub
EH:
    Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Falha a criar botao Eliminar no GIT LOG.", "row=" & CStr(targetRow) & " | err=" & Left$(Err.Description, 180))
End Sub

Private Sub GitLog_PurgeDeleteButtons(ByVal ws As Worksheet)
    On Error GoTo EH

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(BTN_DELETE_PREFIX)) = BTN_DELETE_PREFIX Then shp.Delete
    Next shp
    Exit Sub
EH:
    Call GH_LogWarn(0, "", GH_EVT_GITLOG_DELETE_FAILED, "Falha a limpar botoes antigos de delete no GIT LOG.", "err=" & Left$(Err.Description, 180))
End Sub

Private Function GitLog_EnsureHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lc As Long
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lc
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), headerName, vbTextCompare) = 0 Then
            GitLog_EnsureHeader = c
            Exit Function
        End If
    Next c

    GitLog_EnsureHeader = lc + 1
    ws.Cells(1, GitLog_EnsureHeader).Value = headerName
End Function

Private Function GitLog_ReadPathsFromComment(ByVal targetCell As Range) As String
    On Error Resume Next
    GitLog_ReadPathsFromComment = ""
    If Not targetCell.Comment Is Nothing Then GitLog_ReadPathsFromComment = Trim$(targetCell.Comment.Text)
End Function


Private Function GitLog_DerivePathsFromGitDebugUrl(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    On Error GoTo Fallback

    Dim cGit As Long
    cGit = GitLog_EnsureHeader(ws, "GIT_DEBUG")

    Dim linkText As String
    linkText = Trim$(CStr(ws.Cells(targetRow, cGit).Value))
    If linkText = "" Then Exit Function

    Dim marker As String
    marker = "/tree/"

    Dim pos As Long
    pos = InStr(1, linkText, marker, vbTextCompare)
    If pos = 0 Then Exit Function

    Dim rel As String
    rel = Mid$(linkText, pos + Len(marker))

    Dim slashPos As Long
    slashPos = InStr(1, rel, "/")
    If slashPos = 0 Then Exit Function

    rel = Mid$(rel, slashPos + 1)
    If Trim$(rel) = "" Then Exit Function

    GitLog_DerivePathsFromGitDebugUrl = rel & "/DEBUG.csv;" & rel & "/SEGUIMENTO.csv;" & rel & "/PAINEL.txt;" & rel & "/catalogo_prompts_executadas.csv"
    Exit Function

Fallback:
    GitLog_DerivePathsFromGitDebugUrl = ""
End Function

Private Function GitLog_NormalizePath(ByVal repoPath As String) As String
    Dim p As String
    p = Trim$(repoPath)
    p = Replace$(p, vbCr, "")
    p = Replace$(p, vbLf, "")
    Do While Len(p) > 0 And Left$(p, 1) = "/"
        p = Mid$(p, 2)
    Loop
    GitLog_NormalizePath = p
End Function
