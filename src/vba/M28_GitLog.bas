Attribute VB_Name = "M28_GitLog"
Option Explicit

' =============================================================================
' Modulo: M28_GitLog
' Proposito:
' - Escrever entradas de execucao Git LOG no topo da folha HISTORICO.
' - Separar visualmente runs distintas sem quebrar a continuidade entre prompts da mesma run.
' - Persistir metadado de run por linha atraves de coluna auxiliar oculta.
'
' Atualizacoes:
' - 2026-03-11 | Codex | Cria modulo de insercao top-down para Git LOG por run
'   - Implementa separador de run (linha preta com 6 pt) apenas quando a run muda.
'   - Implementa insercao no topo (linha 2), empurrando historico para baixo.
'   - Persiste `run_id` em coluna auxiliar oculta (`__RUN_ID_META`) para distinguir grupos da mesma run.
'   - Garante timestamp textual no formato `yyyy-mm-dd hh:mm`.
'
' Funcoes e procedimentos:
' - GitLog_InsertRunSeparatorIfNeeded(runId As String)
'   - Insere linha separadora no topo quando a run corrente difere da ultima run registada.
' - GitLog_InsertEntryTop(...)
'   - Escreve entrada na linha 2 e garante metadados/continuidade do bloco da run.
' =============================================================================

Private Const GITLOG_SHEET As String = "HISTORICO"
Private Const GITLOG_META_HEADER As String = "__RUN_ID_META"
Private Const GITLOG_TOP_ROW As Long = 2
Private Const GITLOG_SEPARATOR_HEIGHT As Double = 6#

Public Sub GitLog_InsertRunSeparatorIfNeeded(ByVal runId As String)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = GitLog_GetSheet()
    If ws Is Nothing Then Exit Sub

    Dim metaCol As Long
    metaCol = GitLog_EnsureMetaColumn(ws)

    Dim topRunId As String
    topRunId = Trim$(CStr(ws.Cells(GITLOG_TOP_ROW, metaCol).Value))

    If topRunId = "" Then Exit Sub
    If StrComp(topRunId, Trim$(runId), vbTextCompare) = 0 Then Exit Sub

    ws.Rows(GITLOG_TOP_ROW).Insert Shift:=xlDown
    With ws.Rows(GITLOG_TOP_ROW)
        .RowHeight = GITLOG_SEPARATOR_HEIGHT
        .Interior.Color = vbBlack
        .Font.Color = vbWhite
    End With

    ws.Cells(GITLOG_TOP_ROW, metaCol).Value = "__RUN_SEPARATOR__"
    Exit Sub

EH:
    Call Debug_Registar(0, "", "ALERTA", "", "GIT_LOG", _
        "Falha ao inserir separador de run no Git LOG: " & Err.Description, _
        "Validar existencia da folha HISTORICO e permissao de escrita.")
End Sub

Public Sub GitLog_InsertEntryTop( _
    ByVal runId As String, _
    ByVal pipelineNome As String, _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal httpStatus As Long, _
    ByVal responseId As String, _
    Optional ByVal outputResumo As String = "", _
    Optional ByVal nextPromptDecidido As String = "")

    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = GitLog_GetSheet()
    If ws Is Nothing Then Exit Sub

    Call GitLog_InsertRunSeparatorIfNeeded(runId)

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

    Exit Sub

EH:
    Call Debug_Registar(0, promptId, "ALERTA", "", "GIT_LOG", _
        "Falha ao inserir entrada no topo do Git LOG: " & Err.Description, _
        "Validar cabecalhos no HISTORICO e existencia da coluna auxiliar de metadado.")
End Sub

Private Function GitLog_GetSheet() As Worksheet
    On Error Resume Next
    Set GitLog_GetSheet = ThisWorkbook.Worksheets(GITLOG_SHEET)
    If GitLog_GetSheet Is Nothing Then Set GitLog_GetSheet = ThisWorkbook.Worksheets("HIST" & ChrW$(&HD3) & "RICO")
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
