Attribute VB_Name = "M20_DebugClipboard"
Option Explicit

' =============================================================================
' Modulo: M20_DebugClipboard
' Proposito:
' - Criar/atualizar botao na folha DEBUG para copiar um pacote de diagnostico.
' - Montar texto unico com catalogos dos prompts executados + tabelas DEBUG e Seguimento.
' - Copiar o pacote para o clipboard com fallback resiliente entre hosts Office.
'
' Atualizacoes:
' - 2026-03-03 | Codex | Criacao do modulo de clipboard para troubleshooting guiado
'   - Adiciona macro publica para instalar botao na folha DEBUG com OnAction dedicado.
'   - Implementa geracao do pacote textual com catalogos usados no DEBUG e logs completos.
'   - Implementa copia para clipboard com DataObject e fallback HTMLFile + logging em DEBUG.
'
' Funcoes e procedimentos:
' - DebugClipboard_InstalarBotao() (Sub)
'   - Garante botao idempotente na folha DEBUG e aponta para macro de copia.
' - DebugClipboard_CopiarPacoteDiagnostico() (Sub)
'   - Gera o texto completo do diagnostico e copia para clipboard.
' =============================================================================

Private Const SHEET_DEBUG As String = "DEBUG"
Private Const SHEET_SEGUIMENTO As String = "Seguimento"
Private Const BUTTON_NAME As String = "btnDebugClipboardBundle"
Private Const BUTTON_CAPTION As String = "Copiar pacote diagnóstico"
Private Const BUTTON_ONACTION As String = "DebugClipboard_CopiarPacoteDiagnostico"

Public Sub DebugClipboard_InstalarBotao()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEBUG)

    Dim alvo As Range
    Set alvo = ws.Range("L1")

    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(BUTTON_NAME)
    On Error GoTo EH

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(Type:=xlButtonControl, _
                                           Left:=alvo.Left, _
                                           Top:=alvo.Top, _
                                           Width:=220, _
                                           Height:=24)
        shp.Name = BUTTON_NAME
    End If

    shp.OnAction = BUTTON_ONACTION
    shp.TextFrame.Characters.Text = BUTTON_CAPTION

    Call Debug_Registar(0, "DEBUG", "INFO", "", "DEBUG_CLIPBOARD_BUTTON", _
        "Botão de pacote de diagnóstico instalado/atualizado.", _
        "Use o botão na folha DEBUG para copiar o conteúdo para o clipboard.")
    Exit Sub
EH:
    Call Debug_Registar(0, "DEBUG", "ERRO", "", "DEBUG_CLIPBOARD_BUTTON_FAIL", _
        "Falha ao instalar botão de pacote de diagnóstico: " & Err.Description, _
        "Verifique se a folha DEBUG está disponível e se shapes estão desbloqueados.")
End Sub

Public Sub DebugClipboard_CopiarPacoteDiagnostico()
    On Error GoTo EH

    Dim payload As String
    payload = DebugClipboard_MontarPacoteDiagnostico()

    If Len(payload) = 0 Then
        Call Debug_Registar(0, "DEBUG", "ALERTA", "", "DEBUG_CLIPBOARD_EMPTY", _
            "Pacote de diagnóstico vazio.", _
            "Confirme se as folhas DEBUG/Seguimento têm dados.")
        Exit Sub
    End If

    If DebugClipboard_SetClipboardText(payload) Then
        Call Debug_Registar(0, "DEBUG", "INFO", "", "DEBUG_CLIPBOARD_OK", _
            "Pacote de diagnóstico copiado para clipboard.", _
            "Cole o conteúdo no chat para análise de problemas.")
        MsgBox "Pacote de diagnóstico copiado para o clipboard.", vbInformation
    Else
        Call Debug_Registar(0, "DEBUG", "ERRO", "", "DEBUG_CLIPBOARD_FAIL", _
            "Não foi possível copiar para clipboard neste host.", _
            "Execute em ambiente Windows com suporte a Forms.DataObject.")
        MsgBox "Falha ao copiar para o clipboard. Veja DEBUG para diagnóstico.", vbExclamation
    End If
    Exit Sub
EH:
    Call Debug_Registar(0, "DEBUG", "ERRO", "", "DEBUG_CLIPBOARD_EXCEPTION", _
        "Erro inesperado ao montar/copy pacote de diagnóstico: " & Err.Description, _
        "Verifique a integridade das folhas DEBUG/Seguimento e catálogos.")
End Sub

Private Function DebugClipboard_MontarPacoteDiagnostico() As String
    Dim catalogoTxt As String
    Dim debugTxt As String
    Dim seguimentoTxt As String

    catalogoTxt = DebugClipboard_CatalogosDosPromptsDoDebug()
    debugTxt = DebugClipboard_SheetAsTsv(SHEET_DEBUG)
    seguimentoTxt = DebugClipboard_SheetAsTsv(SHEET_SEGUIMENTO)

    DebugClipboard_MontarPacoteDiagnostico = _
        "O catálogo da prompt abaixo:" & vbCrLf & _
        catalogoTxt & vbCrLf & vbCrLf & vbCrLf & _
        "Resultou neste DEBUG:" & vbCrLf & _
        debugTxt & vbCrLf & vbCrLf & vbCrLf & _
        "E neste Seguimento:" & vbCrLf & _
        seguimentoTxt & vbCrLf & vbCrLf & vbCrLf & _
        "Faça uma lista de problemas a diagnosticar, qual a razão mais provável e o que sugere fazer-se." & vbCrLf
End Function

Private Function DebugClipboard_CatalogosDosPromptsDoDebug() As String
    On Error GoTo EH

    Dim promptIds As Object
    Set promptIds = DebugClipboard_PromptIdsFromDebug()

    If promptIds.Count = 0 Then
        DebugClipboard_CatalogosDosPromptsDoDebug = "[Sem prompts encontrados no DEBUG.]"
        Exit Function
    End If

    Dim sb As String
    Dim promptId As Variant

    For Each promptId In promptIds.Keys
        sb = sb & DebugClipboard_BlocosCatalogoPorPromptId(CStr(promptId)) & vbCrLf
    Next promptId

    If Len(Trim$(sb)) = 0 Then sb = "[Sem blocos de catálogo resolvidos para os prompts do DEBUG.]"
    DebugClipboard_CatalogosDosPromptsDoDebug = Trim$(sb)
    Exit Function
EH:
    DebugClipboard_CatalogosDosPromptsDoDebug = "[Erro ao montar catálogos: " & Err.Description & "]"
End Function

Private Function DebugClipboard_PromptIdsFromDebug() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEBUG)

    Dim colPrompt As Long
    colPrompt = DebugClipboard_FindColumn(ws, "Prompt ID")
    If colPrompt = 0 Then
        Set DebugClipboard_PromptIdsFromDebug = dict
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colPrompt).End(xlUp).Row

    Dim r As Long
    Dim promptId As String
    For r = 2 To lastRow
        promptId = Trim$(CStr(ws.Cells(r, colPrompt).value))
        If promptId <> "" Then
            If UCase$(promptId) <> "DEBUG" And UCase$(promptId) <> "SELFTEST" Then
                dict(promptId) = True
            End If
        End If
    Next r

    Set DebugClipboard_PromptIdsFromDebug = dict
End Function

Private Function DebugClipboard_BlocosCatalogoPorPromptId(ByVal promptId As String) As String
    On Error GoTo EH

    Dim parts() As String
    parts = Split(promptId, "/")
    If UBound(parts) < 0 Then
        DebugClipboard_BlocosCatalogoPorPromptId = "[Prompt ID inválido no DEBUG: " & promptId & "]"
        Exit Function
    End If

    Dim sheetName As String
    sheetName = Trim$(parts(0))
    If sheetName = "" Then
        DebugClipboard_BlocosCatalogoPorPromptId = "[Prompt ID sem prefixo de folha: " & promptId & "]"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim idCell As Range
    Set idCell = ws.Columns(1).Find(What:=promptId, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If idCell Is Nothing Then
        DebugClipboard_BlocosCatalogoPorPromptId = "[Prompt ID não encontrado no catálogo " & sheetName & ": " & promptId & "]"
        Exit Function
    End If

    Dim firstRow As Long
    firstRow = idCell.Row

    Dim blockRange As Range
    Set blockRange = ws.Range(ws.Cells(firstRow, 1), ws.Cells(firstRow + 3, 11))

    DebugClipboard_BlocosCatalogoPorPromptId = "--- Catálogo " & sheetName & " | Prompt " & promptId & " ---" & vbCrLf & _
        DebugClipboard_RangeAsTsv(blockRange)
    Exit Function
EH:
    DebugClipboard_BlocosCatalogoPorPromptId = "[Erro ao ler catálogo de " & promptId & ": " & Err.Description & "]"
End Function

Private Function DebugClipboard_SheetAsTsv(ByVal sheetName As String) As String
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column

    If lastRow < 1 Or lastCol < 1 Then
        DebugClipboard_SheetAsTsv = "[Folha " & sheetName & " sem conteúdo.]"
    Else
        DebugClipboard_SheetAsTsv = DebugClipboard_RangeAsTsv(ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)))
    End If
    Exit Function
EH:
    DebugClipboard_SheetAsTsv = "[Erro ao ler folha " & sheetName & ": " & Err.Description & "]"
End Function

Private Function DebugClipboard_RangeAsTsv(ByVal rng As Range) As String
    Dim data As Variant
    data = rng.value

    Dim r As Long, c As Long
    Dim out As String

    If IsArray(data) Then
        For r = LBound(data, 1) To UBound(data, 1)
            For c = LBound(data, 2) To UBound(data, 2)
                out = out & DebugClipboard_SanitizeCell(CStr(data(r, c)))
                If c < UBound(data, 2) Then out = out & vbTab
            Next c
            If r < UBound(data, 1) Then out = out & vbCrLf
        Next r
    Else
        out = DebugClipboard_SanitizeCell(CStr(data))
    End If

    DebugClipboard_RangeAsTsv = out
End Function

Private Function DebugClipboard_SanitizeCell(ByVal v As String) As String
    Dim t As String
    t = Replace(v, vbCrLf, " [NL] ")
    t = Replace(t, vbCr, " [NL] ")
    t = Replace(t, vbLf, " [NL] ")
    t = Replace(t, vbTab, " ")
    DebugClipboard_SanitizeCell = t
End Function

Private Function DebugClipboard_FindColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            DebugClipboard_FindColumn = c
            Exit Function
        End If
    Next c

    DebugClipboard_FindColumn = 0
End Function

Private Function DebugClipboard_SetClipboardText(ByVal txt As String) As Boolean
    On Error GoTo TryHtml

    Dim dobj As Object
    Set dobj = CreateObject("Forms.DataObject")
    dobj.SetText txt
    dobj.PutInClipboard
    DebugClipboard_SetClipboardText = True
    Exit Function

TryHtml:
    On Error GoTo FailAll

    Dim html As Object
    Set html = CreateObject("htmlfile")
    html.ParentWindow.ClipboardData.SetData "Text", txt
    DebugClipboard_SetClipboardText = True
    Exit Function

FailAll:
    DebugClipboard_SetClipboardText = False
End Function
