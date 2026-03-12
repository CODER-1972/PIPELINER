Attribute VB_Name = "M28_GitLogSheet"
Option Explicit

' =============================================================================
' Modulo: M28_GitLogSheet
' Proposito:
' - Garantir a existencia e o schema base da folha "GIT LOG".
' - Aplicar layout inicial nao destrutivo (headers, formato e vista) para registo
'   de eventos Git por run/pipeline.
' - Fornecer rotina idempotente para chamadas repetidas sem duplicar headers.
'
' Atualizacoes:
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
'   - Garante folha "GIT LOG" pronta para uso; cria se nao existir e aplica
'     configuracoes idempotentes (headers/formato base/painel).
' - GitLog_AppendEvent(runId As String, stepNumber As Long, pipelineName As String, promptId As String, severity As String, eventCode As String, componentName As String, summary As String, details As String)
'   - Acrescenta um registo no GIT LOG com timestamp e colunas canonicas.
' - GitLog_FindHeaderColumn(ws As Worksheet, headerName As String) As Long
'   - Localiza a coluna de um header por comparacao case-insensitive.
' - GitLog_ApplyHeaderStyle(ws As Worksheet, lastCol As Long)
'   - Formata apenas a linha de header (fundo azul claro, negrito, alinhamento).
' =============================================================================

Private Const GITLOG_SHEET_NAME As String = "GIT LOG"
Private Const GITLOG_HEADER_ROW As Long = 1
Private Const GITLOG_DATA_START_ROW As Long = 2
Private Const GITLOG_SUMMARY_HEADER As String = "Summary"

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

    Call GitLog_EnsureHeaders(ws)
    Call GitLog_ApplyHeaderStyle(ws, UBound(GitLog_Headers()) + 1)
    Call GitLog_EnsureSummaryWrap(ws)

    If created Then
        Call GitLog_ApplyInitialColumnWidths(ws)
    End If

    Call GitLog_EnsureFreezeTopRow(ws)

    Set GitLog_EnsureSheet = ws
    Exit Function
EH:
    Set GitLog_EnsureSheet = Nothing
End Function

Private Sub GitLog_EnsureHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    headers = GitLog_Headers()

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If CStr(ws.Cells(GITLOG_HEADER_ROW, i + 1).Value) <> CStr(headers(i)) Then
            ws.Cells(GITLOG_HEADER_ROW, i + 1).Value = CStr(headers(i))
        End If
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
    Dim summaryCol As Long
    summaryCol = GitLog_FindHeaderColumn(ws, GITLOG_SUMMARY_HEADER)
    If summaryCol <= 0 Then Exit Sub

    ws.Columns(summaryCol).WrapText = True
End Sub

Private Sub GitLog_ApplyInitialColumnWidths(ByVal ws As Worksheet)
    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 12
    ws.Columns(4).ColumnWidth = 26
    ws.Columns(5).ColumnWidth = 18
    ws.Columns(6).ColumnWidth = 16
    ws.Columns(7).ColumnWidth = 20
    ws.Columns(8).ColumnWidth = 18
    ws.Columns(9).ColumnWidth = 80
    ws.Columns(10).ColumnWidth = 32
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


Public Sub GitLog_AppendEvent(ByVal runId As String, ByVal stepNumber As Long, ByVal pipelineName As String, ByVal promptId As String, _
    ByVal severity As String, ByVal eventCode As String, ByVal componentName As String, ByVal summary As String, ByVal details As String)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = GitLog_EnsureSheet()
    If ws Is Nothing Then Exit Sub

    Dim targetRow As Long
    targetRow = GitLog_NextDataRow(ws)

    ws.Cells(targetRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(targetRow, 2).Value = Trim$(runId)
    If stepNumber > 0 Then
        ws.Cells(targetRow, 3).Value = stepNumber
    Else
        ws.Cells(targetRow, 3).Value = ""
    End If
    ws.Cells(targetRow, 4).Value = Trim$(pipelineName)
    ws.Cells(targetRow, 5).Value = Trim$(promptId)
    ws.Cells(targetRow, 6).Value = UCase$(Trim$(severity))
    ws.Cells(targetRow, 7).Value = Trim$(eventCode)
    ws.Cells(targetRow, 8).Value = Trim$(componentName)
    ws.Cells(targetRow, 9).Value = summary
    ws.Cells(targetRow, 10).Value = details
    Exit Sub
EH:
    ' Nao bloquear fluxo da pipeline por falha de log auxiliar.
End Sub

Private Function GitLog_NextDataRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < GITLOG_DATA_START_ROW Then
        GitLog_NextDataRow = GITLOG_DATA_START_ROW
    Else
        GitLog_NextDataRow = lastRow + 1
    End If
End Function
Private Function GitLog_FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(GITLOG_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(GITLOG_HEADER_ROW, c).Value)), headerName, vbTextCompare) = 0 Then
            GitLog_FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function GitLog_Headers() As Variant
    GitLog_Headers = Array( _
        "Timestamp", _
        "Run ID", _
        "Step", _
        "Pipeline", _
        "Prompt ID", _
        "Severity", _
        "Event Code", _
        "Component", _
        "Summary", _
        "Details" _
    )
End Function
