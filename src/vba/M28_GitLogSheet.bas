Attribute VB_Name = "M28_GitLogSheet"
Option Explicit

' =============================================================================
' Modulo: M28_GitLogSheet
' Proposito:
' - Garantir a existencia e o schema operacional da folha "GIT LOG".
' - Registar eventos de run/upload num formato tabular canonico para auditoria.
' - Manter formatacao idempotente e separar visualmente runs diferentes.
'
' Atualizacoes:
' - 2026-03-12 | Codex | Corrige run_id e preserva separadores entre runs
'   - Garante run_id mesmo quando chamadas chegam sem runId (ex.: eventos de upload) usando cache/fallback da ultima run.
'   - Evita limpar formataÃ§Ã£o de rows separadoras (6pt preta) durante normalizaÃ§Ã£o idempotente da folha.
' - 2026-03-12 | Codex | Normaliza schema/estilo para formato final do GIT LOG
'   - Atualiza headers para: Timestamp|Pipeline|PromptID|Version|Success|New version|Analysis Link|New Prompt Link|Eliminar|Summary.
'   - Corrige timestamp para yyyy-mm-dd hh:mm e limpa formatos de dados (sem fundo azul).
'   - Adiciona separador visual (row 6pt preta) entre runs diferentes.
' - 2026-03-12 | Codex | Adiciona registo de eventos no GIT LOG
'   - Inclui GitLog_AppendEvent para gravar linhas de auditoria por run/pipeline.
'   - Mantem bootstrap idempotente e adiciona helper para proxima linha de escrita.
' - 2026-03-11 | Codex | Criacao do modulo de bootstrap da folha GIT LOG
'   - Adiciona GitLog_EnsureSheet para criar/normalizar a folha "GIT LOG".
'   - Define headers explicitos, estilo do header, WrapText em Summary,
'     larguras iniciais e congelamento de painel na linha 2.
'
' Funcoes e procedimentos:
' - GitLog_EnsureSheet() As Worksheet
'   - Garante folha GIT LOG com schema final e layout idempotente.
' - GitLog_AppendEvent(...)
'   - Regista evento da run/upload numa nova linha, com run_id normalizado no Summary.
' - GitLog_ResolveRunId(runId As String, ws As Worksheet) As String (Private Function)
'   - Resolve run_id efetivo (normalizado/cached/fallback da ultima run observada).
' - GitLog_BuildPromptLabel(promptId As String) As String (Private Function)
'   - Converte Prompt ID completo para formato curto <ordem>_<nomeCurto>.
' =============================================================================

Private Const GITLOG_SHEET_NAME As String = "GIT LOG"
Private Const GITLOG_HEADER_ROW As Long = 1
Private Const GITLOG_DATA_START_ROW As Long = 2
Private Const GITLOG_SUMMARY_HEADER As String = "Summary"
Private Const GITLOG_SUMMARY_COL As Long = 10
Private Const GITLOG_SEPARATOR_HEIGHT As Double = 6#

Private mGitLogLastRunId As String

Public Function GitLog_EnsureSheet() As Worksheet
    On Error GoTo EH

    Dim ws As Worksheet
    Dim created As Boolean

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GITLOG_SHEET_NAME)
    On Error GoTo EH

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = GITLOG_SHEET_NAME
        created = True
    End If

    If GitLog_HeaderNeedsReset(ws) Then
        ws.Cells.Clear
        created = True
    End If

    Call GitLog_EnsureHeaders(ws)
    Call GitLog_ApplyHeaderStyle(ws, UBound(GitLog_Headers()) + 1)
    Call GitLog_EnsureSummaryWrap(ws)
    Call GitLog_NormalizeDataArea(ws)

    If created Then
        Call GitLog_ApplyInitialColumnWidths(ws)
    End If

    Call GitLog_EnsureFreezeTopRow(ws)

    Set GitLog_EnsureSheet = ws
    Exit Function
EH:
    Set GitLog_EnsureSheet = Nothing
End Function

Public Sub GitLog_AppendEvent(ByVal runId As String, ByVal stepNumber As Long, ByVal pipelineName As String, ByVal promptId As String, _
    ByVal severity As String, ByVal eventCode As String, ByVal componentName As String, ByVal summary As String, ByVal details As String)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = GitLog_EnsureSheet()
    If ws Is Nothing Then Exit Sub

    Dim runNorm As String
    runNorm = GitLog_ResolveRunId(runId, ws)

    Call GitLog_InsertRunSeparatorIfNeeded(ws, runNorm)

    Dim targetRow As Long
    targetRow = GitLog_NextDataRow(ws)

    Dim promptLabel As String
    Dim promptVersion As String
    promptLabel = GitLog_BuildPromptLabel(promptId)
    promptVersion = GitLog_ExtractPromptVersion(promptId)

    ws.Cells(targetRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn")
    ws.Cells(targetRow, 2).Value = Trim$(pipelineName)
    ws.Cells(targetRow, 3).Value = promptLabel
    ws.Cells(targetRow, 4).Value = promptVersion
    ws.Cells(targetRow, 5).Value = GitLog_MapSuccess(severity, eventCode)
    ws.Cells(targetRow, 6).Value = ""
    ws.Cells(targetRow, 7).Value = GitLog_ExtractLink(details)
    ws.Cells(targetRow, 8).Value = ""
    ws.Cells(targetRow, 9).Value = ""
    ws.Cells(targetRow, 10).Value = GitLog_BuildSummary(runNorm, summary, eventCode, componentName, details, stepNumber)

    Call GitLog_ApplyDataRowStyle(ws, targetRow)
    Exit Sub
EH:
    ' Nao bloquear fluxo da pipeline por falha de log auxiliar.
End Sub

Private Sub GitLog_EnsureHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    headers = GitLog_Headers()

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(GITLOG_HEADER_ROW, i + 1).Value = CStr(headers(i))
    Next i
End Sub

Private Sub GitLog_ApplyHeaderStyle(ByVal ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(GITLOG_HEADER_ROW, 1), ws.Cells(GITLOG_HEADER_ROW, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub GitLog_EnsureSummaryWrap(ByVal ws As Worksheet)
    ws.Columns(GITLOG_SUMMARY_COL).WrapText = True
End Sub

Private Sub GitLog_NormalizeDataArea(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < GITLOG_DATA_START_ROW Then Exit Sub

    Dim r As Long
    For r = GITLOG_DATA_START_ROW To lastRow
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(GitLog_Headers()) + 1))) > 0 Then
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(GitLog_Headers()) + 1))
                .Interior.Pattern = xlNone
                .Font.Bold = False
                .VerticalAlignment = xlVAlignCenter
                .RowHeight = 15
            End With
        End If
    Next r

    ws.Columns(1).NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Columns(3).HorizontalAlignment = xlLeft
End Sub

Private Sub GitLog_ApplyInitialColumnWidths(ByVal ws As Worksheet)
    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 28
    ws.Columns(3).ColumnWidth = 28
    ws.Columns(4).ColumnWidth = 10
    ws.Columns(5).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 14
    ws.Columns(7).ColumnWidth = 40
    ws.Columns(8).ColumnWidth = 40
    ws.Columns(9).ColumnWidth = 10
    ws.Columns(10).ColumnWidth = 80
End Sub

Private Sub GitLog_EnsureFreezeTopRow(ByVal ws As Worksheet)
    On Error GoTo Fim

    If Application.ActiveWindow Is Nothing Then Exit Sub

    Dim activeSheetName As String
    activeSheetName = Application.ActiveSheet.Name

    ws.Activate
    ws.Cells(GITLOG_DATA_START_ROW, 1).Select

    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With

Fim:
    On Error Resume Next
    ThisWorkbook.Worksheets(activeSheetName).Activate
    On Error GoTo 0
End Sub

Private Sub GitLog_ApplyDataRowStyle(ByVal ws As Worksheet, ByVal rowNum As Long)
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, UBound(GitLog_Headers()) + 1))
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .VerticalAlignment = xlVAlignCenter
        .RowHeight = 15
    End With

    ws.Cells(rowNum, 1).NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Cells(rowNum, 5).HorizontalAlignment = xlCenter
End Sub

Private Sub GitLog_InsertRunSeparatorIfNeeded(ByVal ws As Worksheet, ByVal runNorm As String)
    Dim lastDataRow As Long
    lastDataRow = GitLog_LastContentRow(ws)
    If lastDataRow < GITLOG_DATA_START_ROW Then Exit Sub

    Dim prevRun As String
    prevRun = GitLog_ExtractRunIdFromSummary(CStr(ws.Cells(lastDataRow, GITLOG_SUMMARY_COL).Value))

    If runNorm = "" Or prevRun = "" Then Exit Sub
    If StrComp(prevRun, runNorm, vbTextCompare) = 0 Then Exit Sub

    Dim sepRow As Long
    sepRow = GitLog_NextDataRow(ws)

    With ws.Range(ws.Cells(sepRow, 1), ws.Cells(sepRow, UBound(GitLog_Headers()) + 1))
        .ClearContents
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(0, 0, 0)
        .RowHeight = GITLOG_SEPARATOR_HEIGHT
    End With
End Sub

Private Function GitLog_LastContentRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To GITLOG_DATA_START_ROW Step -1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(GitLog_Headers()) + 1))) > 0 Then
            GitLog_LastContentRow = r
            Exit Function
        End If
    Next r
    GitLog_LastContentRow = GITLOG_HEADER_ROW
End Function

Private Function GitLog_NextDataRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < GITLOG_DATA_START_ROW Then
        GitLog_NextDataRow = GITLOG_DATA_START_ROW
    Else
        GitLog_NextDataRow = lastRow + 1
    End If
End Function

Private Function GitLog_HeaderNeedsReset(ByVal ws As Worksheet) As Boolean
    Dim headers As Variant
    headers = GitLog_Headers()

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If CStr(ws.Cells(GITLOG_HEADER_ROW, i + 1).Value) <> CStr(headers(i)) Then
            GitLog_HeaderNeedsReset = True
            Exit Function
        End If
    Next i
End Function

Private Function GitLog_NormalizeRunId(ByVal runId As String) As String
    Dim s As String
    s = Trim$(runId)

    If Len(s) >= 4 Then
        If UCase$(Left$(s, 4)) = "RUN|" Then s = Mid$(s, 5)
    End If

    GitLog_NormalizeRunId = s
End Function

Private Function GitLog_ResolveRunId(ByVal runId As String, ByVal ws As Worksheet) As String
    Dim norm As String
    norm = GitLog_NormalizeRunId(runId)

    If norm <> "" Then
        mGitLogLastRunId = norm
        GitLog_ResolveRunId = norm
        Exit Function
    End If

    If Trim$(mGitLogLastRunId) <> "" Then
        GitLog_ResolveRunId = mGitLogLastRunId
        Exit Function
    End If

    Dim lastDataRow As Long
    lastDataRow = GitLog_LastContentRow(ws)
    If lastDataRow >= GITLOG_DATA_START_ROW Then
        mGitLogLastRunId = GitLog_ExtractRunIdFromSummary(CStr(ws.Cells(lastDataRow, GITLOG_SUMMARY_COL).Value))
    End If

    GitLog_ResolveRunId = mGitLogLastRunId
End Function

Private Function GitLog_BuildPromptLabel(ByVal promptId As String) As String
    Dim cleanId As String
    cleanId = Trim$(promptId)
    If cleanId = "" Then
        GitLog_BuildPromptLabel = ""
        Exit Function
    End If

    Dim parts() As String
    parts = Split(cleanId, "/")
    If UBound(parts) >= 2 Then
        GitLog_BuildPromptLabel = Trim$(parts(1)) & "_" & Trim$(parts(2))
    Else
        GitLog_BuildPromptLabel = cleanId
    End If
End Function

Private Function GitLog_ExtractPromptVersion(ByVal promptId As String) As String
    Dim cleanId As String
    cleanId = Trim$(promptId)
    If cleanId = "" Then Exit Function

    Dim parts() As String
    parts = Split(cleanId, "/")
    If UBound(parts) >= 3 Then GitLog_ExtractPromptVersion = Trim$(parts(3))
End Function

Private Function GitLog_MapSuccess(ByVal severity As String, ByVal eventCode As String) As String
    Dim sev As String
    sev = UCase$(Trim$(severity))

    If InStr(1, UCase$(eventCode), "FAILED", vbTextCompare) > 0 Or sev = "ERRO" Or sev = "ERROR" Then
        GitLog_MapSuccess = "NAO"
    ElseIf sev = "ALERTA" Or sev = "WARN" Or sev = "WARNING" Then
        GitLog_MapSuccess = "PARCIAL"
    Else
        GitLog_MapSuccess = "SIM"
    End If
End Function

Private Function GitLog_ExtractLink(ByVal details As String) As String
    Dim p As Long
    p = InStr(1, details, "http", vbTextCompare)
    If p > 0 Then GitLog_ExtractLink = Trim$(Mid$(details, p))
End Function

Private Function GitLog_BuildSummary(ByVal runNorm As String, ByVal summary As String, ByVal eventCode As String, ByVal componentName As String, ByVal details As String, ByVal stepNumber As Long) As String
    Dim out As String
    If runNorm <> "" Then out = "run_id=" & runNorm
    If stepNumber > 0 Then out = GitLog_JoinSummary(out, "step=" & CStr(stepNumber))
    out = GitLog_JoinSummary(out, "event=" & Trim$(eventCode))
    out = GitLog_JoinSummary(out, "component=" & Trim$(componentName))
    out = GitLog_JoinSummary(out, Trim$(summary))

    Dim detShort As String
    detShort = Trim$(details)
    If Len(detShort) > 240 Then detShort = Left$(detShort, 240) & "..."
    out = GitLog_JoinSummary(out, detShort)

    GitLog_BuildSummary = out
End Function

Private Function GitLog_JoinSummary(ByVal currentTxt As String, ByVal piece As String) As String
    Dim p As String
    p = Trim$(piece)
    If p = "" Then
        GitLog_JoinSummary = currentTxt
    ElseIf Trim$(currentTxt) = "" Then
        GitLog_JoinSummary = p
    Else
        GitLog_JoinSummary = currentTxt & " | " & p
    End If
End Function

Private Function GitLog_ExtractRunIdFromSummary(ByVal summaryText As String) As String
    Dim txt As String
    txt = Trim$(summaryText)
    If txt = "" Then Exit Function

    Dim parts() As String
    parts = Split(txt, "|")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim token As String
        token = Trim$(parts(i))
        If InStr(1, token, "run_id=", vbTextCompare) = 1 Then
            GitLog_ExtractRunIdFromSummary = Trim$(Mid$(token, 8))
            Exit Function
        End If
    Next i
End Function

Private Function GitLog_Headers() As Variant
    GitLog_Headers = Array( _
        "Timestamp", _
        "Pipeline", _
        "PromptID", _
        "Version", _
        "Success", _
        "New version", _
        "Analysis Link", _
        "New Prompt Link", _
        "Eliminar", _
        "Summary" _
    )
End Function
