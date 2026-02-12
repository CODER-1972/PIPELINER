Attribute VB_Name = "M08_VBA_Code_Manager"
Option Explicit

' =============================================================================
' Módulo: M08_VBA_Code_Manager
' Propósito:
' - Automatizar import/export e manutenção de módulos VBA no projeto.
' - Apoiar tarefas de engenharia de código e sincronização de ficheiros .bas.
'
' Atualizações:
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - VbaCode_Setup (Sub): rotina pública do módulo.
' - VbaCode_Download (Sub): rotina pública do módulo.
' - VbaCode_Upload (Sub): rotina pública do módulo.
' =============================================================================

' ============================================================
' M08_VBA_Code_Manager (versao limpa)
'
' Objetivo:
'   - Criar uma folha "VBA CODE" para gerir codigo VBA dentro do Excel.
'   - DOWNLOAD: exporta codigo do projecto para a folha, por modulo e por procedimento.
'   - UPLOAD  : carrega codigo da folha para o projecto, com backup e rollback (transacional).
'
' Nota importante:
'   - Para aceder ao codigo do projecto, o Excel tem de permitir:
'       Trust access to the VBA project object model
'     (Trust Center > Macro Settings).
'
' Nota sobre HASH:
'   - SHA-256 via CryptoAPI falha em muitos ambientes institucionais (keyset/provider).
'   - Esta versao usa um hash deterministico em VBA puro (FNV-1a 32-bit, modulo 2^32),
'     suficiente para detetar alteracoes e suportar auditoria.
'
' Regra de trabalho (Alt 1):
'   - Fonte ativa = header + procedimentos.
'   - MODULE_WHOLE = cache/backup visual.
'   - Se MODULE_WHOLE for alterado, o VBA resegmenta e atualiza as linhas antes do upload.
'
' Nota sobre caracteres:
'   - Para reduzir risco de caracteres esquisitos em .bas, este modulo evita acentos.
' ============================================================

Private Const SHEET_VBACODE As String = "VBA CODE"

' Tabela principal (codigo)
Private Const ROW_HEADERS As Long = 3
Private Const ROW_DATA As Long = 4

Private Const COL_MODULO As Long = 1      ' A
Private Const COL_META_MOD As Long = 2    ' B
Private Const COL_PROC As Long = 3        ' C
Private Const COL_META_PROC As Long = 4   ' D
Private Const COL_CODE As Long = 5        ' E
Private Const COL_NOTES As Long = 6       ' F

' Controlo (linha 1)
Private Const CELL_FOLDER As String = "D1"
Private Const CELL_ARMED As String = "F1"

' Limite de caracteres por celula (Excel)
Private Const CELL_LIMIT As Long = 32767
Private Const WHOLE_SAFE_CHARS As Long = 30000  ' margem de seguranca

' Log (na mesma folha, colunas H..L)
Private Const LOG_COL As Long = 8   ' H
Private Const LOG_ROW As Long = 4   ' primeira linha de log

' VBComponent.Type (valores numericos para late binding)
Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_CLASSMODULE As Long = 2
Private Const VBEXT_CT_MSFORM As Long = 3
Private Const VBEXT_CT_DOCUMENT As Long = 100

' vbext_ProcKind (valores numericos para late binding)
Private Const VBEXT_PK_PROC As Long = 0
Private Const VBEXT_PK_LET As Long = 1
Private Const VBEXT_PK_SET As Long = 2
Private Const VBEXT_PK_GET As Long = 3

' Prefixo de botoes (Shapes)
Private Const BTN_PREFIX As String = "BTN_VBACODE_"
Private Const BTN_DOWNLOAD As String = "BTN_VBACODE_DOWNLOAD"
Private Const BTN_UPLOAD As String = "BTN_VBACODE_UPLOAD"


' ============================================================
' Entrada principal (setup)
' ============================================================

Public Sub VbaCode_Setup()
    ' Cria a folha (se necessario), define defaults e cria botoes.
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Set ws = VbaCode_ObterFolha()

    ' Default folder: pasta do workbook
    If Trim$(CStr(ws.Range(CELL_FOLDER).value)) = "" Then
        ws.Range(CELL_FOLDER).value = ThisWorkbook.path
    End If

    If Trim$(CStr(ws.Range(CELL_ARMED).value)) = "" Then
        ws.Range(CELL_ARMED).value = "FALSE"
    End If

    Call VbaCode_CriarBotoes(ws)
    Exit Sub

TrataErro:
    MsgBox "Erro em VbaCode_Setup: " & Err.Description, vbExclamation
End Sub


Public Sub VbaCode_Download()
    ' Botao DOWNLOAD
    On Error GoTo TrataErro

    Call VbaCode_Setup

    If Not VbaCode_TemAcessoVBIDE(True) Then Exit Sub

    Dim ws As Worksheet
    Set ws = VbaCode_ObterFolha()

    Call VbaCode_LimparTabela(ws)

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim comp As Object
    Dim linhaOut As Long
    linhaOut = ROW_DATA

    Dim totalMods As Long
    totalMods = 0

    For Each comp In vbProj.VBComponents
        totalMods = totalMods + 1
        linhaOut = VbaCode_EscreverModulo(ws, comp, linhaOut)
    Next comp

    Call VbaCode_Log(ws, "DOWNLOAD", "Modulos=" & CStr(totalMods), "", VbaCode_ResolverFolder(ws))
    MsgBox "DOWNLOAD concluido. Modulos: " & CStr(totalMods), vbInformation
    Exit Sub

TrataErro:
    MsgBox "Erro em VbaCode_Download: " & Err.Description, vbExclamation
End Sub


Public Sub VbaCode_Upload()
    ' Botao UPLOAD (transacional)
    On Error GoTo TrataErro

    Call VbaCode_Setup

    Dim ws As Worksheet
    Set ws = VbaCode_ObterFolha()

    If Not VbaCode_Armed(ws) Then
        MsgBox "UPLOAD bloqueado: defina ARMED=TRUE na folha 'VBA CODE'.", vbExclamation
        Call VbaCode_Log(ws, "UPLOAD", "BLOQUEADO (ARMED=FALSE)", "", VbaCode_ResolverFolder(ws))
        Exit Sub
    End If

    If Not VbaCode_TemAcessoVBIDE(True) Then Exit Sub

    Dim pastaBase As String
    pastaBase = VbaCode_ResolverFolder(ws)

    If Trim$(pastaBase) = "" Then
        MsgBox "Folder vazio. Preencha 'Folder:' em VBA CODE!D1.", vbExclamation
        Exit Sub
    End If

    If Not VbaCode_PastaExiste(pastaBase) Then
        MsgBox "Folder nao existe: " & pastaBase, vbExclamation
        Exit Sub
    End If

    Dim stamp As String
    stamp = Format$(Now, "yyyy-mm-dd_hhmm")

    Dim pastaBackup As String
    pastaBackup = VbaCode_GarantirSubPasta(pastaBase, "_VBA_BACKUPS")

    ' 1) Snapshot em memoria (para rollback)
    Dim snapshot As Object
    Set snapshot = CreateObject("Scripting.Dictionary")

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim comp As Object
    For Each comp In vbProj.VBComponents
        snapshot.Add comp.name, VbaCode_LerCodigoComp(comp)
    Next comp

    ' 2) Exportar backup para ficheiros .bas/.cls + manifesto
    Dim manifestPath As String
    manifestPath = pastaBackup & "\Manifest_" & stamp & ".txt"
    Call VbaCode_ExportarBackups(vbProj, pastaBackup, stamp, manifestPath)

    ' 3) Ler a folha e preparar novo codigo por modulo
    Dim novos As Object
    Set novos = VbaCode_LerNovosModulosDaFolha(ws, vbProj)

    If novos.Count = 0 Then
        MsgBox "Nao encontrei codigo na folha para carregar (sem modulos na tabela).", vbExclamation
        Exit Sub
    End If

    ' 4) Aplicar alteracoes (transacional)
    Dim erros As String
    erros = ""

    Dim k As Variant
    For Each k In novos.keys
        Dim nomeModulo As String
        nomeModulo = CStr(k)

        If Not VbaCode_ExisteVBComponent(vbProj, nomeModulo) Then
            erros = erros & "Modulo nao existe no projecto: " & nomeModulo & vbCrLf
        Else
            Dim novoCodigo As String
            novoCodigo = CStr(novos(nomeModulo))

            If Trim$(novoCodigo) = "" Then
    ' Se nao ha codigo para este modulo, nao e erro.
    ' E normal em ThisWorkbook/Sheets quando nao ha macros nesse modulo.
    Call VbaCode_Log(ws, "UPLOAD", "SKIP sem codigo: " & nomeModulo, "", pastaBackup)
Else
    Call VbaCode_SubstituirCodigo(vbProj.VBComponents(nomeModulo), novoCodigo)
End If

        End If
    Next k

    If erros <> "" Then
        ' Rollback total
        Call VbaCode_Rollback(vbProj, snapshot)
        Call VbaCode_Log(ws, "UPLOAD", "FALHOU - rollback executado", erros, pastaBackup)
        MsgBox "UPLOAD falhou. Rollback executado. Ver LOG e erros." & vbCrLf & vbCrLf & erros, vbCritical
        Exit Sub
    End If

    ' 5) Atualizar metadados na folha (last_upload/hash/chars)
    Call VbaCode_AtualizarMetadadosFolhaAposUpload(ws, novos)

    ' 6) Desarmar por defeito
    ws.Range(CELL_ARMED).value = "FALSE"

    Call VbaCode_Log(ws, "UPLOAD", "SUCESSO - modulos=" & CStr(novos.Count), "", pastaBackup)
    MsgBox "UPLOAD concluido. Modulos atualizados: " & CStr(novos.Count) & vbCrLf & _
           "Sugestao: Debug > Compile VBAProject", vbInformation
    Exit Sub

TrataErro:
    On Error Resume Next
    If Not snapshot Is Nothing Then
        Call VbaCode_Rollback(ThisWorkbook.VBProject, snapshot)
    End If
    On Error GoTo 0

    MsgBox "Erro em VbaCode_Upload: " & Err.Description, vbExclamation
End Sub


' ============================================================
' Folha e botoes
' ============================================================

Private Function VbaCode_ObterFolha() As Worksheet
    On Error GoTo Criar
    Set VbaCode_ObterFolha = ThisWorkbook.Worksheets(SHEET_VBACODE)
    Exit Function
Criar:
    Set VbaCode_ObterFolha = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    VbaCode_ObterFolha.name = SHEET_VBACODE
    Call VbaCode_CriarLayoutBasico(VbaCode_ObterFolha)
End Function


Private Sub VbaCode_CriarLayoutBasico(ByVal ws As Worksheet)
    ws.Range("A1").value = "DOWNLOAD (botao via macro)"
    ws.Range("B1").value = "UPLOAD (botao via macro)"
    ws.Range("C1").value = "Folder:"
    ws.Range("D1").value = ""
    ws.Range("E1").value = "ARMED:"
    ws.Range("F1").value = "FALSE"

    ws.Range("A2").value = "Nota: Pode editar qualquer celula (MODULE_WHOLE, header ou procedimentos). No UPLOAD, o codigo e sempre reconstruido por header + procedimentos. Se MODULE_WHOLE for alterado, o VBA resegmenta e atualiza as linhas automaticamente."
    ws.Range("A2:F2").Merge

    ws.Cells(ROW_HEADERS, COL_MODULO).value = "Modulo"
    ws.Cells(ROW_HEADERS, COL_META_MOD).value = "Metadados modulo"
    ws.Cells(ROW_HEADERS, COL_PROC).value = "Procedimento"
    ws.Cells(ROW_HEADERS, COL_META_PROC).value = "Metadados procedimento"
    ws.Cells(ROW_HEADERS, COL_CODE).value = "Codigo"
    ws.Cells(ROW_HEADERS, COL_NOTES).value = "Notas"

    ws.Range("H2").value = "LOG: ultimas operacoes de DOWNLOAD/UPLOAD."
    ws.Range("H2:L2").Merge
    ws.Cells(ROW_HEADERS, LOG_COL + 0).value = "LOG Timestamp"
    ws.Cells(ROW_HEADERS, LOG_COL + 1).value = "Operacao"
    ws.Cells(ROW_HEADERS, LOG_COL + 2).value = "Resumo"
    ws.Cells(ROW_HEADERS, LOG_COL + 3).value = "Erros/Avisos"
    ws.Cells(ROW_HEADERS, LOG_COL + 4).value = "Pasta"
End Sub


Private Sub VbaCode_CriarBotoes(ByVal ws As Worksheet)
    On Error GoTo TrataErro

    Call VbaCode_ApagarBotoesExistentes(ws)

    Call VbaCode_CriarBotao(ws, ws.Range("A1"), BTN_DOWNLOAD, "DOWNLOAD", "VbaCode_Download")
    Call VbaCode_CriarBotao(ws, ws.Range("B1"), BTN_UPLOAD, "UPLOAD", "VbaCode_Upload")
    Exit Sub

TrataErro:
End Sub


Private Sub VbaCode_ApagarBotoesExistentes(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.name, Len(BTN_PREFIX)) = BTN_PREFIX Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0
End Sub


Private Sub VbaCode_CriarBotao(ByVal ws As Worksheet, ByVal alvo As Range, ByVal nome As String, ByVal texto As String, ByVal macroOnAction As String)
    Dim margem As Double
    margem = 1

    Dim leftPos As Double, topPos As Double, w As Double, h As Double
    leftPos = alvo.Left + margem
    topPos = alvo.Top + margem
    w = Application.Max(20, alvo.Width - 2 * margem)
    h = Application.Max(14, alvo.Height - 2 * margem)

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(Type:=xlButtonControl, Left:=leftPos, Top:=topPos, Width:=w, Height:=h)
    shp.name = nome
    shp.OnAction = macroOnAction

    On Error Resume Next
    shp.TextFrame.Characters.text = texto
    On Error GoTo 0
End Sub


' ============================================================
' Validacoes / helpers
' ============================================================

Private Function VbaCode_Armed(ByVal ws As Worksheet) As Boolean
    Dim v As String
    v = UCase$(Trim$(CStr(ws.Range(CELL_ARMED).value)))
    VbaCode_Armed = (v = "TRUE" Or v = "VERDADEIRO" Or v = "1" Or v = "SIM")
End Function


Private Function VbaCode_TemAcessoVBIDE(ByVal mostrarMsg As Boolean) As Boolean
    On Error GoTo Falha
    Dim n As Long
    n = ThisWorkbook.VBProject.VBComponents.Count
    VbaCode_TemAcessoVBIDE = True
    Exit Function
Falha:
    VbaCode_TemAcessoVBIDE = False
    If mostrarMsg Then
        MsgBox "Sem acesso ao modelo de objectos do projecto VBA." & vbCrLf & vbCrLf & _
               "Ative: Trust access to the VBA project object model" & vbCrLf & _
               "(File > Options > Trust Center > Trust Center Settings > Macro Settings).", vbExclamation
    End If
End Function


Private Function VbaCode_ResolverFolder(ByVal ws As Worksheet) As String
    Dim p As String
    p = Trim$(CStr(ws.Range(CELL_FOLDER).value))
    If p = "" Then p = ThisWorkbook.path
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    VbaCode_ResolverFolder = p
End Function


Private Function VbaCode_PastaExiste(ByVal p As String) As Boolean
    On Error Resume Next
    VbaCode_PastaExiste = (Len(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function


Private Function VbaCode_GarantirSubPasta(ByVal base As String, ByVal nome As String) As String
    Dim p As String
    p = base
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    p = p & "\" & nome

    If Len(Dir$(p, vbDirectory)) = 0 Then
        MkDir p
    End If
    VbaCode_GarantirSubPasta = p
End Function


' ============================================================
' DOWNLOAD: escrever na folha
' ============================================================

Private Sub VbaCode_LimparTabela(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rowS.Count, COL_MODULO).End(xlUp).Row
    If lastRow < ROW_DATA Then lastRow = ROW_DATA
    ws.Range(ws.Cells(ROW_DATA, COL_MODULO), ws.Cells(lastRow, COL_NOTES)).ClearContents
End Sub


Private Function VbaCode_EscreverModulo(ByVal ws As Worksheet, ByVal comp As Object, ByVal linhaOut As Long) As Long
    On Error GoTo TrataErro

    Dim nome As String
    nome = comp.name

    Dim tipo As String
    tipo = VbaCode_TipoComponente(comp.Type)

    Dim codigoTodo As String
    codigoTodo = VbaCode_LerCodigoComp(comp)

    Dim chars As Long
    chars = Len(codigoTodo)

    Dim whoRisk As Double
    whoRisk = 0
    If CELL_LIMIT > 0 Then whoRisk = chars / CELL_LIMIT

    Dim wholeOk As Boolean
    wholeOk = (chars <= WHOLE_SAFE_CHARS)

    Dim hashMod As String
    hashMod = VbaCode_HashTexto(codigoTodo)

    ' Linha MODULE_WHOLE (se couber)
    ws.Cells(linhaOut, COL_MODULO).value = nome
    ws.Cells(linhaOut, COL_META_MOD).value = VbaCode_FormatMetaModulo(tipo, Now, "", hashMod, chars, wholeOk, whoRisk)
    ws.Cells(linhaOut, COL_PROC).value = "MODULE_WHOLE"
    ws.Cells(linhaOut, COL_META_PROC).value = VbaCode_FormatMetaProc("WHOLE", Now, "", hashMod, chars)

    If wholeOk Then
        ws.Cells(linhaOut, COL_CODE).value = codigoTodo
    Else
        ws.Cells(linhaOut, COL_CODE).value = ""
        ws.Cells(linhaOut, COL_NOTES).value = "WHOLE omitido (risco de exceder limite de celula)."
    End If
    linhaOut = linhaOut + 1

    ' Header (declaracoes antes do 1o procedimento)
    Dim headerCode As String
    headerCode = VbaCode_ExtrairHeader(comp)

    ws.Cells(linhaOut, COL_MODULO).value = nome
    ws.Cells(linhaOut, COL_META_MOD).value = ""
    ws.Cells(linhaOut, COL_PROC).value = ""
    ws.Cells(linhaOut, COL_META_PROC).value = ""
    ws.Cells(linhaOut, COL_CODE).value = headerCode
    linhaOut = linhaOut + 1

    ' Procedimentos
    Dim procs As Collection
    Set procs = VbaCode_ListarProcedimentos(comp)

    Dim i As Long
    For i = 1 To procs.Count
        Dim pack As Variant
        pack = procs(i)

        Dim nomeProc As String
        nomeProc = CStr(pack(0))

        Dim kind As Long
        kind = CLng(pack(1))

        Dim codigoProc As String
        codigoProc = CStr(pack(2))

        Dim tituloProc As String
        tituloProc = VbaCode_TituloProcedimento(codigoProc, nomeProc, kind)

        ws.Cells(linhaOut, COL_MODULO).value = nome
        ws.Cells(linhaOut, COL_PROC).value = tituloProc
        ws.Cells(linhaOut, COL_META_PROC).value = VbaCode_FormatMetaProc(VbaCode_ProcKindToText(kind), Now, "", VbaCode_HashTexto(codigoProc), Len(codigoProc))
        ws.Cells(linhaOut, COL_CODE).value = codigoProc
        linhaOut = linhaOut + 1
    Next i

    VbaCode_EscreverModulo = linhaOut
    Exit Function

TrataErro:
    ws.Cells(linhaOut, COL_MODULO).value = comp.name
    ws.Cells(linhaOut, COL_NOTES).value = "Erro ao descarregar modulo: " & Err.Description
    VbaCode_EscreverModulo = linhaOut + 1
End Function


Private Function VbaCode_LerCodigoComp(ByVal comp As Object) As String
    On Error GoTo Falha
    Dim cm As Object
    Set cm = comp.CodeModule
    VbaCode_LerCodigoComp = cm.lines(1, cm.CountOfLines)
    Exit Function
Falha:
    VbaCode_LerCodigoComp = ""
End Function


Private Function VbaCode_ExtrairHeader(ByVal comp As Object) As String
    On Error GoTo Falha
    Dim cm As Object
    Set cm = comp.CodeModule

    Dim total As Long
    total = cm.CountOfLines
    If total <= 0 Then
        VbaCode_ExtrairHeader = ""
        Exit Function
    End If

    Dim firstProcLine As Long
    firstProcLine = VbaCode_PrimeiraLinhaDeProcedimento(cm, total)

    If firstProcLine <= 1 Then
        VbaCode_ExtrairHeader = cm.lines(1, 1)
    Else
        VbaCode_ExtrairHeader = cm.lines(1, firstProcLine - 1)
    End If
    Exit Function

Falha:
    VbaCode_ExtrairHeader = ""
End Function


Private Function VbaCode_PrimeiraLinhaDeProcedimento(ByVal cm As Object, ByVal total As Long) As Long
    On Error GoTo Falha

    Dim ln As Long
    For ln = 1 To total
        Dim k As Long
        Dim pName As String
        pName = cm.ProcOfLine(ln, k)
        If Trim$(pName) <> "" Then
            VbaCode_PrimeiraLinhaDeProcedimento = ln
            Exit Function
        End If
    Next ln

    VbaCode_PrimeiraLinhaDeProcedimento = total + 1
    Exit Function

Falha:
    VbaCode_PrimeiraLinhaDeProcedimento = total + 1
End Function


Private Function VbaCode_ListarProcedimentos(ByVal comp As Object) As Collection
    Dim lista As New Collection
    On Error GoTo Falha

    Dim cm As Object
    Set cm = comp.CodeModule

    Dim total As Long
    total = cm.CountOfLines
    If total <= 0 Then
        Set VbaCode_ListarProcedimentos = lista
        Exit Function
    End If

    Dim ln As Long
    ln = 1

    Do While ln <= total
        Dim kind As Long
        Dim pName As String
        pName = cm.ProcOfLine(ln, kind)

        If Trim$(pName) <> "" Then
            Dim startLine As Long
            startLine = cm.ProcStartLine(pName, kind)

            Dim nLines As Long
            nLines = cm.ProcCountLines(pName, kind)

            Dim codigoProc As String
            codigoProc = cm.lines(startLine, nLines)

            Dim pack(2) As Variant
            pack(0) = pName
            pack(1) = kind
            pack(2) = codigoProc

            lista.Add pack

            ln = startLine + nLines
        Else
            ln = ln + 1
        End If
    Loop

    Set VbaCode_ListarProcedimentos = lista
    Exit Function

Falha:
    Set VbaCode_ListarProcedimentos = lista
End Function


Private Function VbaCode_TituloProcedimento(ByVal codigoProc As String, ByVal nomeProc As String, ByVal kind As Long) As String
    Dim firstLine As String
    firstLine = VbaCode_PrimeiraLinhaSignificativa(codigoProc)

    Dim t As String
    t = LCase$(firstLine)

    If InStr(1, t, "sub ", vbTextCompare) > 0 Then
        VbaCode_TituloProcedimento = "Sub " & nomeProc
    ElseIf InStr(1, t, "function ", vbTextCompare) > 0 Then
        VbaCode_TituloProcedimento = "Function " & nomeProc
    ElseIf InStr(1, t, "property get", vbTextCompare) > 0 Then
        VbaCode_TituloProcedimento = "Property Get " & nomeProc
    ElseIf InStr(1, t, "property let", vbTextCompare) > 0 Then
        VbaCode_TituloProcedimento = "Property Let " & nomeProc
    ElseIf InStr(1, t, "property set", vbTextCompare) > 0 Then
        VbaCode_TituloProcedimento = "Property Set " & nomeProc
    Else
        VbaCode_TituloProcedimento = nomeProc & " (" & VbaCode_ProcKindToText(kind) & ")"
    End If
End Function


Private Function VbaCode_PrimeiraLinhaSignificativa(ByVal s As String) As String
    Dim linhas() As String
    linhas = Split(Replace(s, vbCrLf, vbLf), vbLf)

    Dim i As Long
    For i = LBound(linhas) To UBound(linhas)
        Dim ln As String
        ln = Trim$(CStr(linhas(i)))
        If ln <> "" Then
            If Left$(ln, 1) <> "'" Then
                VbaCode_PrimeiraLinhaSignificativa = ln
                Exit Function
            End If
        End If
    Next i

    VbaCode_PrimeiraLinhaSignificativa = ""
End Function


Private Function VbaCode_TipoComponente(ByVal t As Long) As String
    Select Case t
        Case VBEXT_CT_STDMODULE: VbaCode_TipoComponente = "standard"
        Case VBEXT_CT_CLASSMODULE: VbaCode_TipoComponente = "class"
        Case VBEXT_CT_DOCUMENT: VbaCode_TipoComponente = "document"
        Case VBEXT_CT_MSFORM: VbaCode_TipoComponente = "msform"
        Case Else: VbaCode_TipoComponente = "unknown(" & CStr(t) & ")"
    End Select
End Function


Private Function VbaCode_ProcKindToText(ByVal k As Long) As String
    Select Case k
        Case VBEXT_PK_PROC: VbaCode_ProcKindToText = "proc"
        Case VBEXT_PK_GET: VbaCode_ProcKindToText = "get"
        Case VBEXT_PK_LET: VbaCode_ProcKindToText = "let"
        Case VBEXT_PK_SET: VbaCode_ProcKindToText = "set"
        Case Else: VbaCode_ProcKindToText = "kind(" & CStr(k) & ")"
    End Select
End Function


Private Function VbaCode_FormatMetaModulo( _
    ByVal tipo As String, _
    ByVal lastDownload As Date, _
    ByVal lastUpload As String, _
    ByVal hash As String, _
    ByVal nChars As Long, _
    ByVal wholeOk As Boolean, _
    ByVal wholeRisk As Double _
) As String
    Dim s As String
    s = "TYPE=" & tipo & vbCrLf & _
        "LAST_DOWNLOAD=" & Format$(lastDownload, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "LAST_UPLOAD=" & lastUpload & vbCrLf & _
        "HASH=" & hash & vbCrLf & _
        "CHARS=" & CStr(nChars) & vbCrLf & _
        "WHOLE_OK=" & IIf(wholeOk, "TRUE", "FALSE") & vbCrLf & _
        "WHOLE_RISK=" & Format$(wholeRisk, "0.00%")
    VbaCode_FormatMetaModulo = s
End Function


Private Function VbaCode_FormatMetaProc( _
    ByVal tipo As String, _
    ByVal lastDownload As Date, _
    ByVal lastUpload As String, _
    ByVal hash As String, _
    ByVal nChars As Long _
) As String
    Dim s As String
    s = "TYPE=" & tipo & vbCrLf & _
        "LAST_DOWNLOAD=" & Format$(lastDownload, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "LAST_UPLOAD=" & lastUpload & vbCrLf & _
        "HASH=" & hash & vbCrLf & _
        "CHARS=" & CStr(nChars)
    VbaCode_FormatMetaProc = s
End Function


' ============================================================
' UPLOAD: ler da folha e aplicar
' ============================================================
Private Function VbaCode_LerNovosModulosDaFolha(ByVal ws As Worksheet, ByVal vbProj As Object) As Object
    ' Devolve dict: modulo -> codigo completo (string).
    '
    ' Regra (Alternativa 1):
    '   - O UPLOAD e sempre feito por HEADER + PROCEDIMENTOS.
    '   - MODULE_WHOLE e apenas cache/backup visual.
    '   - Se MODULE_WHOLE foi alterado (hash diferente do guardado em meta),
    '     o VBA resegmenta esse texto para HEADER + PROCEDIMENTOS e atualiza
    '     a folha antes de reconstruir.

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = ROW_DATA

    Do While True
        Dim modName As String
        modName = Trim$(CStr(ws.Cells(r, COL_MODULO).value))

        Dim procName As String
        procName = Trim$(CStr(ws.Cells(r, COL_PROC).value))

        Dim codeCell As String
        codeCell = CStr(ws.Cells(r, COL_CODE).value)

        If modName = "" And procName = "" And Trim$(codeCell) = "" Then Exit Do

        If modName <> "" Then
            ' Ler bloco deste modulo (linhas consecutivas com o mesmo modulo em A)
            Dim startR As Long
            startR = r

            Dim endR As Long
            endR = r
            Do While Trim$(CStr(ws.Cells(endR, COL_MODULO).value)) = modName
                endR = endR + 1
            Loop
            endR = endR - 1

            ' Se MODULE_WHOLE foi alterado, sincronizar header/procs na folha
            Call VbaCode_SincronizarSeWholeAlterado(ws, vbProj, modName, startR, endR)

            Dim codigoFinal As String
            codigoFinal = VbaCode_ReconstruirCodigoModulo(ws, startR, endR)

            If Not d.exists(modName) Then
                d.Add modName, codigoFinal
            Else
                ' Se duplicado, concatenar (tolerante; nao recomendado)
                d(modName) = CStr(d(modName)) & vbCrLf & vbCrLf & codigoFinal
            End If

            r = endR + 1
        Else
            r = r + 1
        End If

        If r > ws.rowS.Count Then Exit Do
    Loop

    Set VbaCode_LerNovosModulosDaFolha = d
End Function

Private Function VbaCode_ReconstruirCodigoModulo(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long) As String
    ' Reconstrucao SEMPRE por header + procedimentos.
    ' MODULE_WHOLE e ignorado como fonte (apenas cache).

    Dim headerCode As String
    headerCode = ""

    Dim procs As String
    procs = ""

    Dim r As Long
    For r = r1 To r2
        Dim p As String
        p = Trim$(CStr(ws.Cells(r, COL_PROC).value))

        Dim c As String
        c = CStr(ws.Cells(r, COL_CODE).value)

        If UCase$(p) = "MODULE_WHOLE" Then
            ' Ignorar como fonte
        ElseIf p = "" Then
            If headerCode = "" Then headerCode = c
        Else
            If Trim$(c) <> "" Then
                If procs <> "" Then procs = procs & vbCrLf & vbCrLf
                procs = procs & c
            End If
        End If
    Next r

    Dim s As String
    s = ""
    If Trim$(headerCode) <> "" Then s = headerCode
    If Trim$(procs) <> "" Then
        If Trim$(s) <> "" Then s = s & vbCrLf & vbCrLf
        s = s & procs
    End If

    VbaCode_ReconstruirCodigoModulo = s
End Function


' ============================================================
' UPLOAD - suporte a edicao em MODULE_WHOLE (cache) ou em PROCS
' ============================================================

Private Sub VbaCode_SincronizarSeWholeAlterado(ByVal ws As Worksheet, ByVal vbProj As Object, ByVal modName As String, ByVal r1 As Long, ByRef r2 As Long)
    ' Se a linha MODULE_WHOLE tiver sido alterada desde o ultimo DOWNLOAD/UPLOAD,
    ' resegmenta o seu texto em header + procedimentos e atualiza o bloco na folha.

    On Error GoTo Sai

    Dim rWhole As Long
    rWhole = VbaCode_EncontrarLinhaWhole(ws, r1, r2)
    If rWhole = 0 Then Exit Sub

    Dim wholeText As String
    wholeText = CStr(ws.Cells(rWhole, COL_CODE).value)
    If Trim$(wholeText) = "" Then Exit Sub

    ' Se exceder o limite seguro, nao usar como fonte (pode estar truncado)
    If Len(wholeText) > WHOLE_SAFE_CHARS Then
        ws.Cells(rWhole, COL_NOTES).value = "WHOLE demasiado grande; edite por procedimentos."
        Exit Sub
    End If

    Dim hashAtual As String
    hashAtual = VbaCode_HashTexto(wholeText)

    Dim metaProc As String
    metaProc = CStr(ws.Cells(rWhole, COL_META_PROC).value)

    Dim hashGuardado As String
    hashGuardado = VbaCode_MetaGet(metaProc, "HASH")

    If hashGuardado = "" Then
        Call VbaCode_SincronizarBlocoAPartirDoWhole(ws, vbProj, modName, r1, r2, wholeText)
    ElseIf UCase$(Trim$(hashAtual)) <> UCase$(Trim$(hashGuardado)) Then
        Call VbaCode_SincronizarBlocoAPartirDoWhole(ws, vbProj, modName, r1, r2, wholeText)
    End If

Sai:
End Sub


Private Function VbaCode_EncontrarLinhaWhole(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim r As Long
    For r = r1 To r2
        If UCase$(Trim$(CStr(ws.Cells(r, COL_PROC).value))) = "MODULE_WHOLE" Then
            VbaCode_EncontrarLinhaWhole = r
            Exit Function
        End If
    Next r
    VbaCode_EncontrarLinhaWhole = 0
End Function


Private Sub VbaCode_SincronizarBlocoAPartirDoWhole(ByVal ws As Worksheet, ByVal vbProj As Object, ByVal modName As String, ByVal r1 As Long, ByRef r2 As Long, ByVal wholeText As String)
    ' Converte MODULE_WHOLE -> header + procedimentos (usando VBIDE para parsing),
    ' e reescreve o bloco do modulo na folha (A:F) de forma consistente.

    Dim headerCode As String
    Dim procs As Collection

    Call VbaCode_SegmentarCodigoWhole(vbProj, wholeText, headerCode, procs)
    Call VbaCode_AplicarSegmentacaoNoBloco(ws, modName, r1, r2, headerCode, procs)
End Sub


Private Sub VbaCode_SegmentarCodigoWhole(ByVal vbProj As Object, ByVal wholeText As String, ByRef headerCode As String, ByRef procs As Collection)
    ' Usa um modulo temporario para aproveitar ProcOfLine/ProcStartLine/ProcCountLines.
    ' Se falhar, devolve header=wholeText e procs vazia (mantem a edicao do utilizador).

    headerCode = ""
    Set procs = New Collection

    Dim tempComp As Object
    Dim tempName As String
    tempName = "zzTempParse_" & Format$(Now, "hhmmss") & "_" & CStr(Int((Rnd() + 0.001) * 100000))

    On Error GoTo Falha

    Set tempComp = vbProj.VBComponents.Add(VBEXT_CT_STDMODULE)
    tempComp.name = tempName

    tempComp.CodeModule.AddFromString wholeText

    headerCode = VbaCode_ExtrairHeader(tempComp)
    Set procs = VbaCode_ListarProcedimentos(tempComp)

Limpar:
    On Error Resume Next
    If Not tempComp Is Nothing Then
        vbProj.VBComponents.Remove tempComp
    End If
    On Error GoTo 0
    Exit Sub

Falha:
    headerCode = wholeText
    Set procs = New Collection
    Resume Limpar
End Sub


Private Sub VbaCode_AplicarSegmentacaoNoBloco(ByVal ws As Worksheet, ByVal modName As String, ByVal r1 As Long, ByRef r2 As Long, ByVal headerCode As String, ByVal procs As Collection)
    ' Reescreve o bloco do modulo (A:F) para ficar coerente com a segmentacao:
    '   r1   : MODULE_WHOLE (mantem o texto editado)
    '   r1+1 : header (proc vazio)
    '   r1+2.. : procedimentos (um por linha)

    Dim headerRow As Long
    headerRow = r1 + 1

    ' Garantir que existe linha de header
    If r2 < headerRow Then
        Call VbaCode_InserirLinhasTabela(ws, headerRow, headerRow - r2)
        r2 = headerRow
    End If

    Dim newProcCount As Long
    newProcCount = procs.Count

    Dim desiredEnd As Long
    desiredEnd = (r1 + 1) + newProcCount

    If desiredEnd > r2 Then
        Call VbaCode_InserirLinhasTabela(ws, r2 + 1, desiredEnd - r2)
        r2 = desiredEnd
    ElseIf desiredEnd < r2 Then
        Call VbaCode_ApagarLinhasTabela(ws, desiredEnd + 1, r2 - desiredEnd)
        r2 = desiredEnd
    End If

    ' Fixar o cabecalho do bloco
    ws.Cells(r1, COL_MODULO).value = modName
    ws.Cells(r1, COL_PROC).value = "MODULE_WHOLE"

    ws.Cells(headerRow, COL_MODULO).value = modName
    ws.Cells(headerRow, COL_META_MOD).value = ""
    ws.Cells(headerRow, COL_PROC).value = ""
    ws.Cells(headerRow, COL_META_PROC).value = ""
    ws.Cells(headerRow, COL_CODE).value = headerCode
    ws.Cells(headerRow, COL_NOTES).value = ""

    ' Procedimentos
    Dim i As Long
    For i = 1 To newProcCount
        Dim pack As Variant
        pack = procs(i)

        Dim nomeProc As String
        nomeProc = CStr(pack(0))

        Dim kind As Long
        kind = CLng(pack(1))

        Dim codigoProc As String
        codigoProc = CStr(pack(2))

        Dim tituloProc As String
        tituloProc = VbaCode_TituloProcedimento(codigoProc, nomeProc, kind)

        Dim rowProc As Long
        rowProc = headerRow + i

        ws.Cells(rowProc, COL_MODULO).value = modName
        ws.Cells(rowProc, COL_META_MOD).value = ""
        ws.Cells(rowProc, COL_PROC).value = tituloProc
        ws.Cells(rowProc, COL_META_PROC).value = ""
        ws.Cells(rowProc, COL_CODE).value = codigoProc
        ws.Cells(rowProc, COL_NOTES).value = ""
    Next i
End Sub


Private Sub VbaCode_InserirLinhasTabela(ByVal ws As Worksheet, ByVal atRow As Long, ByVal n As Long)
    ' Insere "n" linhas de celulas na tabela A:F, deslocando para baixo (nao mexe no LOG em H..).
    If n <= 0 Then Exit Sub

    ws.Range(ws.Cells(atRow, COL_MODULO), ws.Cells(atRow + n - 1, COL_NOTES)).Insert Shift:=xlDown
End Sub


Private Sub VbaCode_ApagarLinhasTabela(ByVal ws As Worksheet, ByVal atRow As Long, ByVal n As Long)
    ' Apaga "n" linhas de celulas na tabela A:F, deslocando para cima (nao mexe no LOG em H..).
    If n <= 0 Then Exit Sub

    ws.Range(ws.Cells(atRow, COL_MODULO), ws.Cells(atRow + n - 1, COL_NOTES)).Delete Shift:=xlUp
End Sub



Private Sub VbaCode_SubstituirCodigo(ByVal comp As Object, ByVal novoCodigo As String)
    On Error GoTo TrataErro
    Dim cm As Object
    Set cm = comp.CodeModule

    If cm.CountOfLines > 0 Then
        cm.DeleteLines 1, cm.CountOfLines
    End If

    If Trim$(novoCodigo) <> "" Then
        cm.AddFromString novoCodigo
    End If
    Exit Sub

TrataErro:
    Err.Raise Err.Number, "VbaCode_SubstituirCodigo(" & comp.name & ")", Err.Description
End Sub


Private Function VbaCode_ExisteVBComponent(ByVal vbProj As Object, ByVal nome As String) As Boolean
    On Error GoTo Falha
    Dim x As Object
    Set x = vbProj.VBComponents(nome)
    VbaCode_ExisteVBComponent = True
    Exit Function
Falha:
    VbaCode_ExisteVBComponent = False
End Function


Private Sub VbaCode_Rollback(ByVal vbProj As Object, ByVal snapshot As Object)
    On Error Resume Next
    Dim k As Variant
    For Each k In snapshot.keys
        Dim comp As Object
        Set comp = vbProj.VBComponents(CStr(k))
        Dim cm As Object
        Set cm = comp.CodeModule
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        cm.AddFromString CStr(snapshot(k))
    Next k
    On Error GoTo 0
End Sub


' ============================================================
' Backup export (.bas/.cls) + manifesto
' ============================================================

Private Sub VbaCode_ExportarBackups(ByVal vbProj As Object, ByVal pastaBackup As String, ByVal stamp As String, ByVal manifestPath As String)
    On Error GoTo TrataErro

    Dim f As Integer
    f = FreeFile
    Open manifestPath For Output As #f
    Print #f, "Manifest generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "Backup folder: " & pastaBackup
    Print #f, ""

    Dim comp As Object
    For Each comp In vbProj.VBComponents
        Dim tipo As String
        tipo = VbaCode_TipoComponente(comp.Type)

        Dim ext As String
        ext = ".bas"
        If comp.Type = VBEXT_CT_CLASSMODULE Then ext = ".cls"
        If comp.Type = VBEXT_CT_MSFORM Then ext = ".frm"

        Dim pathOut As String
        pathOut = pastaBackup & "\" & comp.name & "_" & stamp & ext

        On Error Resume Next
        comp.Export pathOut
        If Err.Number <> 0 Then
            Err.Clear
            Call VbaCode_EscreverTextoEmFicheiro(pathOut, VbaCode_LerCodigoComp(comp))
        End If
        On Error GoTo TrataErro

        Dim codeText As String
        codeText = VbaCode_LerCodigoComp(comp)
        Dim h As String
        h = VbaCode_HashTexto(codeText)

        Print #f, comp.name & " | " & tipo & " | CHARS=" & CStr(Len(codeText)) & " | HASH=" & h & " | FILE=" & pathOut
    Next comp

    Close #f
    Exit Sub

TrataErro:
    On Error Resume Next
    Close #f
End Sub


Private Sub VbaCode_EscreverTextoEmFicheiro(ByVal path As String, ByVal texto As String)
    On Error GoTo TrataErro
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, texto
    Close #f
    Exit Sub
TrataErro:
    On Error Resume Next
    Close #f
End Sub


' ============================================================
' Atualizar metadados na folha apos UPLOAD
' ============================================================
Private Sub VbaCode_AtualizarMetadadosFolhaAposUpload(ByVal ws As Worksheet, ByVal novos As Object)
    ' Atualiza metadados (HASH/CHARS/LAST_UPLOAD) e refresca MODULE_WHOLE como cache
    ' (apenas se for seguro escrever o codigo inteiro numa celula).

    On Error GoTo Falha

    Dim stamp As String
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Dim r As Long
    r = ROW_DATA

    Do While True
        Dim modName As String
        modName = Trim$(CStr(ws.Cells(r, COL_MODULO).value))

        Dim procName As String
        procName = Trim$(CStr(ws.Cells(r, COL_PROC).value))

        Dim codeCell As String
        codeCell = CStr(ws.Cells(r, COL_CODE).value)

        If modName = "" And procName = "" And Trim$(codeCell) = "" Then Exit Do

        If modName <> "" Then
            If novos.exists(modName) Then
                ' Meta de modulo: assume-se que a 1a linha do bloco e a linha MODULE_WHOLE
                Dim metaMod As String
                metaMod = CStr(ws.Cells(r, COL_META_MOD).value)

                Dim novoCodigo As String
                novoCodigo = CStr(novos(modName))

                Dim chars As Long
                chars = Len(novoCodigo)

                Dim whoRisk As Double
                whoRisk = 0
                If CELL_LIMIT > 0 Then whoRisk = chars / CELL_LIMIT

                Dim wholeOk As Boolean
                wholeOk = (chars <= WHOLE_SAFE_CHARS)

                ' Refrescar MODULE_WHOLE (cache) apenas se couber com margem de seguranca
                If wholeOk Then
                    ws.Cells(r, COL_CODE).value = novoCodigo
                    ws.Cells(r, COL_NOTES).value = ""
                Else
                    ws.Cells(r, COL_CODE).value = ""
                    ws.Cells(r, COL_NOTES).value = "WHOLE omitido (risco de exceder limite de celula)."
                End If

                metaMod = VbaCode_MetaSet(metaMod, "LAST_UPLOAD", stamp)
                metaMod = VbaCode_MetaSet(metaMod, "HASH", VbaCode_HashTexto(novoCodigo))
                metaMod = VbaCode_MetaSet(metaMod, "CHARS", CStr(chars))
                metaMod = VbaCode_MetaSet(metaMod, "WHOLE_OK", IIf(wholeOk, "TRUE", "FALSE"))
                metaMod = VbaCode_MetaSet(metaMod, "WHOLE_RISK", Format$(whoRisk, "0.00%"))
                ws.Cells(r, COL_META_MOD).value = metaMod

                ' Meta por procedimento (inclui MODULE_WHOLE e procedimentos reais)
                Dim rr As Long
                rr = r
                Do While Trim$(CStr(ws.Cells(rr, COL_MODULO).value)) = modName
                    Dim p As String
                    p = Trim$(CStr(ws.Cells(rr, COL_PROC).value))
                    If p <> "" Then
                        Dim metaProc As String
                        metaProc = CStr(ws.Cells(rr, COL_META_PROC).value)

                        Dim codeProc As String
                        codeProc = CStr(ws.Cells(rr, COL_CODE).value)

                        metaProc = VbaCode_MetaSet(metaProc, "LAST_UPLOAD", stamp)
                        metaProc = VbaCode_MetaSet(metaProc, "HASH", VbaCode_HashTexto(codeProc))
                        metaProc = VbaCode_MetaSet(metaProc, "CHARS", CStr(Len(codeProc)))

                        ws.Cells(rr, COL_META_PROC).value = metaProc
                    End If
                    rr = rr + 1
                Loop
            End If

            ' Saltar bloco inteiro do modulo
            Dim endR As Long
            endR = r
            Do While Trim$(CStr(ws.Cells(endR, COL_MODULO).value)) = modName
                endR = endR + 1
            Loop
            r = endR
        Else
            r = r + 1
        End If

        If r > ws.rowS.Count Then Exit Do
    Loop

    Exit Sub

Falha:
    ' Metadados sao auditoria; se falhar nao bloqueia o upload
End Sub

Private Function VbaCode_MetaSet(ByVal meta As String, ByVal key As String, ByVal value As String) As String
    Dim texto As String
    texto = Replace(meta, vbCrLf, vbLf)

    Dim linhas() As String
    linhas = Split(texto, vbLf)

    Dim i As Long
    Dim found As Boolean
    found = False

    For i = LBound(linhas) To UBound(linhas)
        Dim ln As String
        ln = CStr(linhas(i))
        If UCase$(Left$(Trim$(ln), Len(key) + 1)) = UCase$(key & "=") Then
            linhas(i) = key & "=" & value
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        If Trim$(texto) = "" Then
            texto = key & "=" & value
        Else
            texto = texto & vbLf & key & "=" & value
        End If
        VbaCode_MetaSet = Replace(texto, vbLf, vbCrLf)
        Exit Function
    End If

    texto = Join(linhas, vbLf)
    VbaCode_MetaSet = Replace(texto, vbLf, vbCrLf)
End Function


Private Function VbaCode_MetaGet(ByVal meta As String, ByVal key As String) As String
    ' Meta em linhas "KEY=valor"
    Dim texto As String
    texto = Replace(meta, vbCrLf, vbLf)

    Dim linhas() As String
    linhas = Split(texto, vbLf)

    Dim i As Long
    For i = LBound(linhas) To UBound(linhas)
        Dim ln As String
        ln = Trim$(CStr(linhas(i)))
        If UCase$(Left$(ln, Len(key) + 1)) = UCase$(key & "=") Then
            VbaCode_MetaGet = Mid$(ln, Len(key) + 2)
            Exit Function
        End If
    Next i

    VbaCode_MetaGet = ""
End Function




' ============================================================
' Log (na mesma folha)
' ============================================================

Private Sub VbaCode_Log(ByVal ws As Worksheet, ByVal operacao As String, ByVal resumo As String, ByVal erros As String, ByVal pasta As String)
    On Error Resume Next

    Dim r As Long
    r = ws.Cells(ws.rowS.Count, LOG_COL).End(xlUp).Row + 1
    If r < LOG_ROW Then r = LOG_ROW

    ws.Cells(r, LOG_COL + 0).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, LOG_COL + 1).value = operacao
    ws.Cells(r, LOG_COL + 2).value = resumo
    ws.Cells(r, LOG_COL + 3).value = erros
    ws.Cells(r, LOG_COL + 4).value = pasta
End Sub


' ============================================================
' Hash (VBA puro): FNV-1a 32-bit (mod 2^32)
' ============================================================

Private Function VbaCode_HashTexto(ByVal texto As String) As String
    ' FNV-1a 32-bit sem overflow (multiplicacao mod 2^32 por decomposicao 16-bit)
    Const OFFSET_BASIS As Long = &H811C9DC5
    Const FNV_PRIME As Long = &H1000193

    Dim h As Long
    h = OFFSET_BASIS

    Dim b() As Byte
    b = StrConv(texto, vbFromUnicode)

    Dim i As Long
    If (Not Not b) <> 0 Then
        For i = LBound(b) To UBound(b)
            h = (h Xor CLng(b(i)))
            h = VbaCode_MulMod32(h, FNV_PRIME)
        Next i
    End If

    VbaCode_HashTexto = "FNV32-" & Right$("00000000" & Hex$(h), 8)
End Function


Private Function VbaCode_MulMod32(ByVal a As Long, ByVal b As Long) As Long
    ' (a*b) mod 2^32 usando decomposicao em 16-bit e Double (sem overflow e com precisao)
    Dim a0 As Long, a1 As Long, b0 As Long, b1 As Long
    a0 = a And &HFFFF&
    a1 = (a And &HFFFF0000) \ &H10000
    b0 = b And &HFFFF&
    b1 = (b And &HFFFF0000) \ &H10000

    Dim res As Double
    res = (CDbl(a0) * CDbl(b0)) + (CDbl(a0) * CDbl(b1) + CDbl(a1) * CDbl(b0)) * 65536#

    ' mod 2^32
    res = res - Int(res / 4294967296#) * 4294967296#

    ' converter para Long assinado
    If res >= 2147483648# Then res = res - 4294967296#
    VbaCode_MulMod32 = CLng(res)
End Function


