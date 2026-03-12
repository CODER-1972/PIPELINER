Attribute VB_Name = "M28_GitLog"
Option Explicit

' =============================================================================
' Modulo: M28_GitLog
' Proposito:
' - Escrever entradas de execucao Git LOG no topo da folha de log (linha 2).
' - Separar visualmente runs distintas sem quebrar a continuidade entre prompts da mesma run.
' - Persistir metadado de run por linha atraves de coluna auxiliar oculta.
'
' Atualizacoes:
' - 2026-03-12 | Codex | Diagnostico e resolucao robusta da folha alvo do Git LOG
'   - Passa a resolver folha alvo por candidatos: "GIT LOG", "GIT_LOG", "HISTORICO" e "HISTÃ“RICO".
'   - Adiciona GitLog_DiagnoseTarget para registar diagnostico acionavel no DEBUG quando nao houver escrita.
'   - Exponibiliza estado de sucesso/falha em GitLog_InsertEntryTop e GitLog_InsertRunSeparatorIfNeeded.
' - 2026-03-11 | Codex | Cria modulo de insercao top-down para Git LOG por run
'   - Implementa separador de run (linha preta com 6 pt) apenas quando a run muda.
'   - Implementa insercao no topo (linha 2), empurrando historico para baixo.
'   - Persiste `run_id` em coluna auxiliar oculta (`__RUN_ID_META`) para distinguir grupos da mesma run.
'   - Garante timestamp textual no formato `yyyy-mm-dd hh:mm`.
'
' Funcoes e procedimentos:
' - GitLog_DiagnoseTarget(ByRef detail As String) As Boolean
'   - Valida e descreve a folha alvo/cabecalhos para troubleshooting.
' - GitLog_InsertRunSeparatorIfNeeded(runId As String, ...)
'   - Insere linha separadora no topo quando a run corrente difere da ultima run registada.
' - GitLog_InsertEntryTop(...)
'   - Escreve entrada na linha 2 e garante metadados/continuidade do bloco da run.
' =============================================================================

Private Const GITLOG_META_HEADER As String = "__RUN_ID_META"
Private Const GITLOG_TOP_ROW As Long = 2
Private Const GITLOG_SEPARATOR_HEIGHT As Double = 6#

Private Const GITLOG_SHEET_A As String = "GIT LOG"
Private Const GITLOG_SHEET_B As String = "GIT_LOG"
Private Const GITLOG_SHEET_C As String = "HISTORICO"

Public Function GitLog_DiagnoseTarget(ByRef detail As String) As Boolean
    On Error GoTo EH

    detail = ""

    Dim ws As Worksheet
    Dim resolvedName As String
    Set ws = GitLog_GetSheet(resolvedName)

    If ws Is Nothing Then
        detail = "target_sheet=NAO_ENCONTRADA | candidatos=GIT LOG;GIT_LOG;HISTORICO;HISTÃ“RICO"
        GitLog_DiagnoseTarget = False
        Exit Function
    End If

    Dim map As Object
    Set map = GitLog_HeaderMap(ws)

    Dim hasTs As Boolean
    Dim hasPrompt As Boolean
    hasTs = map.exists("Timestamp")
    hasPrompt = map.exists("Prompt ID")

    Dim metaCol As Long
    metaCol = GitLog_EnsureMetaColumn(ws)

    detail = "target_sheet=" & resolvedName & _
             " | headers_count=" & CStr(map.Count) & _
             " | has_timestamp=" & IIf(hasTs, "SIM", "NAO") & _
             " | has_prompt_id=" & IIf(hasPrompt, "SIM", "NAO") & _
             " | meta_col=" & CStr(metaCol)

    GitLog_DiagnoseTarget = hasTs
    Exit Function

EH:
    detail = "diag_err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 160)
    GitLog_DiagnoseTarget = False
End Function

Public Sub GitLog_InsertRunSeparatorIfNeeded( _
    ByVal runId As String, _
    Optional ByRef outOk As Boolean = False, _
    Optional ByRef outReason As String = "")

    On Error GoTo EH

    outOk = False
    outReason = ""

    Dim ws As Worksheet
    Dim resolvedName As String
    Set ws = GitLog_GetSheet(resolvedName)
    If ws Is Nothing Then
        outReason = "folha alvo nao encontrada (GIT LOG/GIT_LOG/HISTORICO/HISTÃ“RICO)"
        Exit Sub
    End If

    Dim metaCol As Long
    metaCol = GitLog_EnsureMetaColumn(ws)

    Dim topRunId As String
    topRunId = Trim$(CStr(ws.Cells(GITLOG_TOP_ROW, metaCol).Value))

    If topRunId = "" Then
        outOk = True
        outReason = "sem separador: topo vazio"
        Exit Sub
    End If

    If StrComp(topRunId, Trim$(runId), vbTextCompare) = 0 Then
        outOk = True
        outReason = "sem separador: mesma run"
        Exit Sub
    End If

    ws.Rows(GITLOG_TOP_ROW).Insert Shift:=xlDown
    With ws.Rows(GITLOG_TOP_ROW)
        .RowHeight = GITLOG_SEPARATOR_HEIGHT
        .Interior.Color = vbBlack
        .Font.Color = vbWhite
    End With

    ws.Cells(GITLOG_TOP_ROW, metaCol).Value = "__RUN_SEPARATOR__"
    outOk = True
    outReason = "separador inserido"
    Exit Sub

EH:
    outReason = "separator_err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 160)
End Sub

Public Sub GitLog_InsertEntryTop( _
    ByVal runId As String, _
    ByVal pipelineNome As String, _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal httpStatus As Long, _
    ByVal responseId As String, _
    Optional ByVal outputResumo As String = "", _
    Optional ByVal nextPromptDecidido As String = "", _
    Optional ByRef outOk As Boolean = False, _
    Optional ByRef outReason As String = "")

    On Error GoTo EH

    outOk = False
    outReason = ""

    Dim ws As Worksheet
    Dim resolvedName As String
    Set ws = GitLog_GetSheet(resolvedName)
    If ws Is Nothing Then
        outReason = "folha alvo nao encontrada (GIT LOG/GIT_LOG/HISTORICO/HISTÃ“RICO)"
        Exit Sub
    End If

    Dim sepOk As Boolean
    Dim sepReason As String
    Call GitLog_InsertRunSeparatorIfNeeded(runId, sepOk, sepReason)
    If Not sepOk Then
        outReason = "falha separador: " & sepReason
        Exit Sub
    End If

    Dim map As Object
    Set map = GitLog_HeaderMap(ws)

    Dim metaCol As Long
    metaCol = GitLog_EnsureMetaColumn(ws)

    ws.Rows(GITLOG_TOP_ROW).Insert Shift:=xlDown

    If map.exists("Timestamp") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("Timestamp"))).Value = Format$(Now, "yyyy-mm-dd hh:mm")
    If map.exists("Nome do Pipeline") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("Nome do Pipeline"))).Value = pipelineNome
    If map.exists("Passo") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("Passo"))).Value = passo
    If map.exists("Prompt ID") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("Prompt ID"))).Value = promptId
    If map.exists("HTTP Status") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("HTTP Status"))).Value = httpStatus
    If map.exists("Response ID") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("Response ID"))).Value = responseId
    If map.exists("Output (texto)") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("Output (texto)"))).Value = outputResumo
    If map.exists("Next prompt decidido") Then ws.Cells(GITLOG_TOP_ROW, CLng(map("Next prompt decidido"))).Value = nextPromptDecidido

    ws.Cells(GITLOG_TOP_ROW, metaCol).Value = Trim$(runId)

    If map.exists("Timestamp") Then
        ws.Cells(GITLOG_TOP_ROW, CLng(map("Timestamp"))).NumberFormat = "yyyy-mm-dd hh:mm"
    End If

    outOk = True
    outReason = "ok | target_sheet=" & resolvedName & " | sep=" & sepReason
    Exit Sub

EH:
    outReason = "entry_err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 160)
End Sub

Private Function GitLog_GetSheet(Optional ByRef resolvedName As String = "") As Worksheet
    On Error Resume Next

    resolvedName = ""

    Set GitLog_GetSheet = ThisWorkbook.Worksheets(GITLOG_SHEET_A)
    If Not GitLog_GetSheet Is Nothing Then
        resolvedName = GITLOG_SHEET_A
        GoTo CleanExit
    End If

    Set GitLog_GetSheet = ThisWorkbook.Worksheets(GITLOG_SHEET_B)
    If Not GitLog_GetSheet Is Nothing Then
        resolvedName = GITLOG_SHEET_B
        GoTo CleanExit
    End If

    Set GitLog_GetSheet = ThisWorkbook.Worksheets(GITLOG_SHEET_C)
    If Not GitLog_GetSheet Is Nothing Then
        resolvedName = GITLOG_SHEET_C
        GoTo CleanExit
    End If

    Set GitLog_GetSheet = ThisWorkbook.Worksheets("HIST" & ChrW$(&HD3) & "RICO")
    If Not GitLog_GetSheet Is Nothing Then resolvedName = "HISTÃ“RICO"

CleanExit:
    On Error GoTo 0
End Function

Private Function GitLog_EnsureMetaColumn(ByVal ws As Worksheet) As Long
    Dim map As Object
    Set map = GitLog_HeaderMap(ws)

    If map.exists(GITLOG_META_HEADER) Then
        GitLog_EnsureMetaColumn = CLng(map(GITLOG_META_HEADER))
    Else
        GitLog_EnsureMetaColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        If GitLog_EnsureMetaColumn < 1 Then GitLog_EnsureMetaColumn = 1
        ws.Cells(1, GitLog_EnsureMetaColumn).Value = GITLOG_META_HEADER
    End If

    ws.Columns(GitLog_EnsureMetaColumn).EntireColumn.Hidden = True
End Function

Private Function GitLog_HeaderMap(ByVal ws As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then
        Set GitLog_HeaderMap = d
        Exit Function
    End If

    Dim c As Long
    For c = 1 To lastCol
        Dim h As String
        h = Trim$(CStr(ws.Cells(1, c).Value))
        If h <> "" Then
            If Not d.exists(h) Then d.Add h, c
        End If
    Next c

    Set GitLog_HeaderMap = d
End Function
