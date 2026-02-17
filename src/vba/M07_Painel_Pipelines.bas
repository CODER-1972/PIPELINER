Attribute VB_Name = "M07_Painel_Pipelines"
Option Explicit

' =============================================================================
' Módulo: M07_Painel_Pipelines
' Propósito:
' - Orquestrar execução de pipelines a partir da folha PAINEL e ações de botões.
' - Gerir limites, fluxo de passos, integração com catálogo/API/logs e geração de mapa/registo.
'
' Atualizações:
' - 2026-02-17 | Codex | Encadeamento robusto de previous_response_id
'   - Só encadeia previous_response_id quando a resposta anterior foi persistida (store=TRUE).
'   - Evita erro HTTP 400 previous_response_not_found em pipelines com prompts store=FALSE.
' - 2026-02-17 | Codex | INPUTS híbrido (append + extração de variáveis)
'   - Adiciona INPUTS_APPEND_MODE (OFF|SAFE|RAW; default RAW) para anexar INPUTS ao prompt final.
'   - Extrai pares chave/valor de INPUTS (":" e "=") para mapa normalizado com logging no DEBUG.
'   - Grava variáveis extraídas no Seguimento (captured_vars) via ContextKV, sem substituição automática no prompt.
' - 2026-02-16 | Codex | Resolução de API key via ambiente com fallback compatível
'   - Substitui leitura direta de Config!B1 por resolver central (M14_ConfigApiKey).
'   - Regista ALERTA/ERRO no DEBUG para origem/falhas da credencial sem expor segredo.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - Painel_CriarBotoes (Sub): rotina pública do módulo.
' - Painel_Click_Iniciar (Sub): rotina pública do módulo.
' - Painel_Click_Registar (Sub): rotina pública do módulo.
' - Painel_Click_SetDefault (Sub): rotina pública do módulo.
' - Painel_Click_CriarMapa (Sub): rotina pública do módulo.
' =============================================================================

' ============================================================
' M07_Painel_Pipelines  (VERSAO REVISTA)
'
' Objetivo:
'   - Painel com 10 pipelines (pares de colunas):
'       * INICIAR    : executa pipeline (API) a partir do 1o ID
'       * REGISTAR   : mapeia sequencia a partir do 1o ID (Next PROMPT)
'       * SET DEFAULT: para as prompts na 2a coluna, Next PROMPT = default
'       * Criar mapa : gera ficheiro Word (docx) no INPUT Folder com mapping
'
' Integra:
'   - M03: Catalogo_ObterPromptPorID
'   - M04: ConfigExtra_Converter
'   - M05: OpenAI_Executar
'   - M02: Debug_Registar / Seguimento_Registar
'   - M06: Seguimento_ArquivarLimpar
'   - M09: Files_PrepararContextoDaPrompt
'
' Requisitos adicionados (user):
'   1) Ao clicar INICIAR: foco em Seguimento!A1
'   2) Ao clicar INICIAR: limpar DEBUG (sessao anterior) sem ativar a folha
'   3) Status bar: "(hh:mm) Step: x of y  |  Retry: z"
'   4) Check/diagnostico de FILES (3 checks) + logging util
' ============================================================

Private Const SHEET_PAINEL As String = "PAINEL"
Private Const SHEET_CONFIG As String = "Config"
Private Const SHEET_SEGUIMENTO As String = "Seguimento"
Private Const SHEET_DEBUG As String = "DEBUG"

Private Const LIST_START_ROW As Long = 10
Private Const LIST_MAX_ROWS As Long = 40

Private Const PIPELINES As Long = 10

' Prefixos de nomes de botoes (Shapes) criados no PAINEL
Private Const BTN_PREFIX As String = "BTN_"
Private Const BTN_INICIAR As String = "BTN_INICIAR_"
Private Const BTN_REGISTAR As String = "BTN_REGISTAR_"
Private Const BTN_SETDEFAULT As String = "BTN_SETDEFAULT_"
Private Const BTN_MAPA As String = "BTN_MAPA_"

Private Const INPUTS_APPEND_HEADER As String = "### INPUTS_RESOLVIDOS"

' ============================================================
' 1) Instalacao / UI
' ============================================================

Public Sub Painel_CriarBotoes()
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PAINEL)

    Call Painel_ApagarBotoesExistentes(ws)

    Dim i As Long
    For i = 1 To PIPELINES
        Dim colIniciar As Long, colRegistar As Long
        Call Painel_ObterColunasPipeline(i, colIniciar, colRegistar)

        ' Botoes principais (linha 8)
        Call Painel_CriarBotaoEmCelula(ws, ws.Cells(8, colIniciar), BTN_INICIAR & Format$(i, "00"), "INICIAR", "Painel_Click_Iniciar")
        Call Painel_CriarBotaoEmCelula(ws, ws.Cells(8, colRegistar), BTN_REGISTAR & Format$(i, "00"), "REGISTAR", "Painel_Click_Registar")

        ' Botoes auxiliares (linha 5/6 na 2a coluna do par)
        Call Painel_CriarBotaoEmCelula(ws, ws.Cells(6, colRegistar), BTN_MAPA & Format$(i, "00"), "Criar mapa", "Painel_Click_CriarMapa")
        Call Painel_CriarBotaoEmCelula(ws, ws.Cells(6, colRegistar), BTN_SETDEFAULT & Format$(i, "00"), "Set Default", "Painel_Click_SetDefault")
    Next i

    MsgBox "Botoes criados no PAINEL.", vbInformation
    Exit Sub

TrataErro:
    MsgBox "Erro ao criar botoes no PAINEL: " & Err.Description, vbExclamation
End Sub

Private Sub Painel_ApagarBotoesExistentes(ByVal ws As Worksheet)
    On Error Resume Next

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.name, Len(BTN_PREFIX)) = BTN_PREFIX Then
            shp.Delete
        End If
    Next shp

    On Error GoTo 0
End Sub

Private Sub Painel_CriarBotaoEmCelula(ByVal ws As Worksheet, ByVal alvo As Range, ByVal nomeBotao As String, ByVal texto As String, ByVal macroOnAction As String)
    Dim margem As Double
    margem = 1

    Dim leftPos As Double, topPos As Double, w As Double, h As Double
    leftPos = alvo.Left + margem
    topPos = alvo.Top + margem
    w = Application.Max(20, alvo.Width - 2 * margem)
    h = Application.Max(14, alvo.Height - 2 * margem)

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(Type:=xlButtonControl, Left:=leftPos, Top:=topPos, Width:=w, Height:=h)

    shp.name = nomeBotao
    shp.OnAction = macroOnAction

    On Error Resume Next
    shp.TextFrame.Characters.text = texto
    On Error GoTo 0
End Sub

' ============================================================
' 2) Macros chamadas pelos botoes (sem argumentos)
' ============================================================

Public Sub Painel_Click_Iniciar()
    Dim idx As Long
    idx = Painel_ExtrairIndiceDoCaller(BTN_INICIAR)
    If idx = 0 Then Exit Sub

    ' (1) Limpar DEBUG de sessoes previas (sem ativar folha)
    Call Painel_LimparDebugSessaoAnterior

    ' Nota: o foco em Seguimento e feito dentro do Painel_IniciarPipeline,
    ' apos arquivar/limpar o Seguimento e repor a altura das linhas.
    Call Painel_IniciarPipeline(idx)
End Sub


Public Sub Painel_Click_Registar()
    Dim idx As Long
    idx = Painel_ExtrairIndiceDoCaller(BTN_REGISTAR)
    If idx = 0 Then Exit Sub
    Call Painel_RegistarPipeline(idx)
End Sub

Public Sub Painel_Click_SetDefault()
    Dim idx As Long
    idx = Painel_ExtrairIndiceDoCaller(BTN_SETDEFAULT)
    If idx = 0 Then Exit Sub
    Call Painel_SetDefaultPipeline(idx)
End Sub

Public Sub Painel_Click_CriarMapa()
    Dim idx As Long
    idx = Painel_ExtrairIndiceDoCaller(BTN_MAPA)
    If idx = 0 Then Exit Sub
    Call Painel_CriarMapaPipeline(idx)
End Sub

Private Function Painel_ExtrairIndiceDoCaller(ByVal prefixo As String) As Long
    On Error GoTo Falha

    Dim callerName As String
    callerName = CStr(Application.Caller)

    If InStr(1, callerName, prefixo, vbTextCompare) <> 1 Then
        Painel_ExtrairIndiceDoCaller = 0
        Exit Function
    End If

    Dim s As String
    s = Mid$(callerName, Len(prefixo) + 1)

    Painel_ExtrairIndiceDoCaller = CLng(val(s))
    Exit Function

Falha:
    Painel_ExtrairIndiceDoCaller = 0
End Function

' ============================================================
' 3) Pipeline: REGISTAR / SET DEFAULT / CRIAR MAPA
' ============================================================

Private Sub Painel_RegistarPipeline(ByVal pipelineIndex As Long)
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PAINEL)

    Dim colIniciar As Long, colRegistar As Long
    Call Painel_ObterColunasPipeline(pipelineIndex, colIniciar, colRegistar)

    Dim maxSteps As Long, maxRep As Long
    Call Painel_LerLimitesPipeline(ws, pipelineIndex, maxSteps, maxRep)

    Dim startId As String
    startId = Painel_LerPrimeiroID(ws, colRegistar)
    If startId = "" Then startId = Painel_LerPrimeiroID(ws, colIniciar)

    Call Painel_LimparLista(ws, colRegistar)

    If startId = "" Or Painel_EhSTOP(startId) Then
        ws.Cells(LIST_START_ROW, colRegistar).value = "STOP"
        Exit Sub
    End If

    Dim vistos As Object
    Set vistos = CreateObject("Scripting.Dictionary")

    Dim atual As String
    atual = startId

    Dim writeRow As Long
    writeRow = LIST_START_ROW

    Dim i As Long
    For i = 1 To maxSteps
        If atual = "" Or Painel_EhSTOP(atual) Then
            ws.Cells(writeRow, colRegistar).value = "STOP"
            Exit Sub
        End If

        If vistos.exists(UCase$(atual)) Then
            Call Debug_Registar(0, atual, "ALERTA", "", "Next PROMPT", _
                "Ciclo detetado ao registar pipeline.", _
                "Sugestao: reveja Next PROMPT / allowed ou use STOP.")
            ws.Cells(writeRow, colRegistar).value = "STOP"
            Exit Sub
        End If

        vistos.Add UCase$(atual), True
        ws.Cells(writeRow, colRegistar).value = atual
        writeRow = writeRow + 1
        If writeRow > LIST_START_ROW + LIST_MAX_ROWS - 1 Then Exit For

        Dim nextPrompt As String, nextDefault As String, nextAllowed As String
        Call Catalogo_LerNextConfig(atual, nextPrompt, nextDefault, nextAllowed)

        Dim candidato As String
        candidato = Painel_ResolverNextDeterministico(nextPrompt, nextDefault)
        candidato = Painel_ValidarAllowedEExistencia(candidato, nextDefault, nextAllowed, 0, atual)

        If candidato = "" Or Painel_EhSTOP(candidato) Then
            ws.Cells(writeRow, colRegistar).value = "STOP"
            Exit Sub
        End If

        atual = candidato
    Next i

    Call Debug_Registar(0, startId, "ALERTA", "", "MaxSteps", _
        "MaxSteps atingido ao registar pipeline.", _
        "Sugestao: aumente Max Steps no PAINEL ou introduza STOP.")
    ws.Cells(writeRow, colRegistar).value = "STOP"
    Exit Sub

TrataErro:
    Call Debug_Registar(0, "PIPELINE_" & CStr(pipelineIndex), "ERRO", "", "REGISTAR", _
        Err.Description, _
        "Sugestao: verifique IDs e folhas.")
End Sub

Private Sub Painel_SetDefaultPipeline(ByVal pipelineIndex As Long)
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PAINEL)

    Dim colIniciar As Long, colRegistar As Long
    Call Painel_ObterColunasPipeline(pipelineIndex, colIniciar, colRegistar)

    Dim r As Long
    For r = LIST_START_ROW To LIST_START_ROW + LIST_MAX_ROWS - 1
        Dim pid As String
        pid = Trim$(CStr(ws.Cells(r, colRegistar).value))

        If pid = "" Then Exit For
        If Painel_EhSTOP(pid) Then Exit For

        Dim celId As Range
        Set celId = Catalogo_EncontrarCelulaID(pid)

        If celId Is Nothing Then
            Call Debug_Registar(0, pid, "ALERTA", "", "SET DEFAULT", _
                "ID nao encontrado no catalogo.", _
                "Sugestao: confirme o ID e a folha.")
        Else
            Dim txtDefault As String
            txtDefault = CStr(celId.Offset(2, 1).value) ' Next PROMPT default
            Dim valDefault As String
            valDefault = Painel_ExtrairValorAposDoisPontos(txtDefault)
            If valDefault = "" Then valDefault = "STOP"

            celId.Offset(1, 1).value = "Next PROMPT: " & valDefault
        End If
    Next r

    Call Painel_RegistarPipeline(pipelineIndex)
    Exit Sub

TrataErro:
    Call Debug_Registar(0, "PIPELINE_" & CStr(pipelineIndex), "ERRO", "", "SET DEFAULT", _
        Err.Description, _
        "Sugestao: verifique o catalogo.")
End Sub

Private Sub Painel_CriarMapaPipeline(ByVal pipelineIndex As Long)
    On Error GoTo TrataErro

    Dim wsPainel As Worksheet
    Set wsPainel = ThisWorkbook.Worksheets(SHEET_PAINEL)

    Dim colIniciar As Long, colRegistar As Long
    Call Painel_ObterColunasPipeline(pipelineIndex, colIniciar, colRegistar)

    Dim nomePipeline As String
    nomePipeline = Painel_LerNomePipeline(wsPainel, pipelineIndex)

    Dim inputFolder As String
    inputFolder = Trim$(CStr(wsPainel.Cells(2, colIniciar).value))
    If inputFolder = "" Then
        Call Debug_Registar(0, nomePipeline, "ALERTA", "", "INPUT Folder", _
            "INPUT Folder vazio no PAINEL.", _
            "Sugestao: preencha o caminho na linha INPUT Folder.")
        Exit Sub
    End If
    If Right$(inputFolder, 1) = "\" Then inputFolder = Left$(inputFolder, Len(inputFolder) - 1)

    Dim listaIDs As Collection
    Set listaIDs = Painel_LerListaIDsUnica(wsPainel, colRegistar)
    If listaIDs.Count = 0 Then
        Call Debug_Registar(0, nomePipeline, "ALERTA", "", "CRIAR MAPA", _
            "Lista de prompts vazia na 2a coluna.", _
            "Sugestao: clique REGISTAR antes de Criar mapa.")
        Exit Sub
    End If

    Dim fileName As String
    fileName = Painel_SanitizarNomeFicheiro(nomePipeline) & "_PromptMapping_" & Format$(Now, "yyyy-mm-dd_hhmmss") & ".docx"

    Dim fullPath As String
    fullPath = inputFolder & "\" & fileName

    Dim wordApp As Object, wordDoc As Object
    Dim tbl As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add

    wordDoc.content.InsertAfter "Prompt Mapping - " & nomePipeline & vbCrLf
    wordDoc.content.InsertAfter "Gerado em: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf

    Dim totalRows As Long
    totalRows = 1 ' header

    Dim i As Long
    For i = 1 To listaIDs.Count
        totalRows = totalRows + 1 ' linha principal
        totalRows = totalRows + 3 ' next config
        totalRows = totalRows + 1 ' separador
    Next i

    Set tbl = wordDoc.Tables.Add(Range:=wordDoc.Paragraphs(wordDoc.Paragraphs.Count).Range, NumRows:=totalRows, NumColumns:=4)
    tbl.Borders.Enable = True

    Call Word_SetCellText(tbl, 1, 1, "ID")
    Call Word_SetCellText(tbl, 1, 2, "Nome curto")
    Call Word_SetCellText(tbl, 1, 3, "Nome descritivo")
    Call Word_SetCellText(tbl, 1, 4, "Texto prompt")

    Dim rowPtr As Long
    rowPtr = 2

    For i = 1 To listaIDs.Count
        Dim pid As String
        pid = CStr(listaIDs(i))

        Dim p As PromptDefinicao
        p = Catalogo_ObterPromptPorID(pid)

        Call Word_SetCellText(tbl, rowPtr, 1, p.Id)
        Call Word_SetCellText(tbl, rowPtr, 2, p.NomeCurto)
        Call Word_SetCellText(tbl, rowPtr, 3, p.NomeDescritivo)
        Call Word_SetCellText(tbl, rowPtr, 4, p.textoPrompt)
        rowPtr = rowPtr + 1

        Dim celId As Range
        Set celId = Catalogo_EncontrarCelulaID(pid)

        If Not celId Is Nothing Then
            Call Word_SetCellText(tbl, rowPtr, 1, "")
            Call Word_SetCellText(tbl, rowPtr, 2, CStr(celId.Offset(1, 1).value))
            Call Word_SetCellText(tbl, rowPtr, 3, CStr(celId.Offset(1, 2).value))
            Call Word_SetCellText(tbl, rowPtr, 4, CStr(celId.Offset(1, 3).value))
            rowPtr = rowPtr + 1

            Call Word_SetCellText(tbl, rowPtr, 1, "")
            Call Word_SetCellText(tbl, rowPtr, 2, CStr(celId.Offset(2, 1).value))
            Call Word_SetCellText(tbl, rowPtr, 3, CStr(celId.Offset(2, 2).value))
            Call Word_SetCellText(tbl, rowPtr, 4, CStr(celId.Offset(2, 3).value))
            rowPtr = rowPtr + 1

            Call Word_SetCellText(tbl, rowPtr, 1, "")
            Call Word_SetCellText(tbl, rowPtr, 2, CStr(celId.Offset(3, 1).value))
            Call Word_SetCellText(tbl, rowPtr, 3, CStr(celId.Offset(3, 2).value))
            Call Word_SetCellText(tbl, rowPtr, 4, CStr(celId.Offset(3, 3).value))
            rowPtr = rowPtr + 1
        Else
            Call Word_SetCellText(tbl, rowPtr, 1, "")
            Call Word_SetCellText(tbl, rowPtr, 2, "Next PROMPT: (nao encontrado)")
            Call Word_SetCellText(tbl, rowPtr, 3, "")
            Call Word_SetCellText(tbl, rowPtr, 4, "")
            rowPtr = rowPtr + 3
        End If

        ' separador
        Call Word_SetCellText(tbl, rowPtr, 1, "")
        Call Word_SetCellText(tbl, rowPtr, 2, "")
        Call Word_SetCellText(tbl, rowPtr, 3, "")
        Call Word_SetCellText(tbl, rowPtr, 4, "")
        rowPtr = rowPtr + 1
    Next i

    wordDoc.SaveAs2 fileName:=fullPath
    wordDoc.Close SaveChanges:=False
    wordApp.Quit

    Set wordDoc = Nothing
    Set wordApp = Nothing

    MsgBox "Mapa criado: " & fullPath, vbInformation
    Exit Sub

TrataErro:
    Call Debug_Registar(0, "PIPELINE_" & CStr(pipelineIndex), "ERRO", "", "CRIAR MAPA", _
        Err.Description, _
        "Sugestao: confirme pasta, permissao e Word instalado.")
    On Error Resume Next
    If Not wordApp Is Nothing Then wordApp.Quit
End Sub

Private Sub Word_SetCellText(ByVal tbl As Object, ByVal r As Long, ByVal c As Long, ByVal texto As String)
    On Error Resume Next
    tbl.Cell(r, c).Range.text = texto
    On Error GoTo 0
End Sub

' ============================================================
' 4) Pipeline: INICIAR (executa chamadas a API e segue Next PROMPT)
' ============================================================

Private Sub Painel_IniciarPipeline(ByVal pipelineIndex As Long)
    On Error GoTo TrataErro

    Dim oldDisplayStatusBar As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldStatusBar As Variant

    ' ---- FILE OUTPUT (declarações) ----
    Dim fo_outputKind As String, fo_processMode As String, fo_autoSave As String, fo_overwriteMode As String
    Dim fo_prefixTmpl As String, fo_subfolderTmpl As String, fo_structuredMode As String
    Dim fo_pptxMode As String, fo_xlsxMode As String, fo_pdfMode As String, fo_imageMode As String
    Dim modosEfetivo As String, extraFragmentFO As String
    Dim fo_filesUsedOut As String, fo_filesOpsOut As String, fo_logSeguimento As String
    Dim textoSeguimento As String, filesUsedResumo As String, filesOpsResumo As String


    oldDisplayStatusBar = Application.DisplayStatusBar
    oldEnableEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    Dim wsPainel As Worksheet
    Set wsPainel = ThisWorkbook.Worksheets(SHEET_PAINEL)

    Dim colIniciar As Long, colRegistar As Long
    Call Painel_ObterColunasPipeline(pipelineIndex, colIniciar, colRegistar)

    Dim pipelineNome As String
    pipelineNome = Painel_LerNomePipeline(wsPainel, pipelineIndex)

    Dim inputFolder As String
    inputFolder = Trim$(CStr(wsPainel.Cells(2, colIniciar).value))
    If Right$(inputFolder, 1) = "\" Then inputFolder = Left$(inputFolder, Len(inputFolder) - 1)

    ' Output Folder (base) + toggle pipeline para auto-guardar ficheiros
    Dim outputFolderBase As String
    outputFolderBase = Trim$(CStr(wsPainel.Cells(3, colIniciar).value))
    If Right$(outputFolderBase, 1) = "\" Then outputFolderBase = Left$(outputFolderBase, Len(outputFolderBase) - 1)

    Dim painelAutoSave As String
    painelAutoSave = Trim$(CStr(wsPainel.Cells(4, colIniciar).value))
    If painelAutoSave = "" Then painelAutoSave = "Sim"


    Dim maxSteps As Long, maxRep As Long
    Call Painel_LerLimitesPipeline(wsPainel, pipelineIndex, maxSteps, maxRep)

    ' Reset visual da coluna INICIAR
    Call Painel_ReporFormatoColunaIDs(wsPainel, colIniciar)

    ' Ler API key (ambiente primeiro) e defaults
    Dim apiKey As String
    Dim apiKeySource As String
    Dim apiKeyAlert As String
    Dim apiKeyError As String

    If Not Config_ResolveOpenAIApiKey(apiKey, apiKeySource, apiKeyAlert, apiKeyError) Then
        Call Debug_Registar(0, pipelineNome, "ERRO", "", "OPENAI_API_KEY", _
            apiKeyError, _
            "Defina a variável de ambiente OPENAI_API_KEY (recomendado) ou um fallback válido em Config!B1.")
        MsgBox "OPENAI_API_KEY ausente. Verifique a variável de ambiente OPENAI_API_KEY ou o fallback em Config!B1.", vbExclamation
        GoTo SaidaLimpa
    End If

    If Trim$(apiKeyAlert) <> "" Then
        Call Debug_Registar(0, pipelineNome, "ALERTA", "", "OPENAI_API_KEY", apiKeyAlert, "Sem ação imediata; recomendado migrar para variável de ambiente.")
    End If

    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Worksheets(SHEET_CONFIG)

    Call Painel_EnsureConfigInputsModes(wsCfg)

    Dim modeloDefault As String
    modeloDefault = Trim$(CStr(wsCfg.Range("B2").value))
    If modeloDefault = "" Then modeloDefault = "gpt-4.1"

    Dim temperaturaDefault As Double
    temperaturaDefault = 0#
    On Error Resume Next
    If Trim$(CStr(wsCfg.Range("B3").value)) <> "" Then temperaturaDefault = CDbl(wsCfg.Range("B3").value)
    On Error GoTo TrataErro

    Dim maxTokensDefault As Long
    maxTokensDefault = 0
    On Error Resume Next
    If Trim$(CStr(wsCfg.Range("B4").value)) <> "" Then maxTokensDefault = CLng(wsCfg.Range("B4").value)
    On Error GoTo TrataErro

    If maxTokensDefault <= 0 Then maxTokensDefault = 800

    ' Limpar/arquivar Seguimento (rotina existente)
    Seguimento_ArquivarLimpar

    ' Garantir foco em Seguimento (pedido do utilizador)
    Call Painel_FocarSeguimentoA1

    ' Determinar ID inicial
    Dim startId As String
    startId = Painel_LerPrimeiroID(wsPainel, colIniciar)
    If startId = "" Or Painel_EhSTOP(startId) Then
        startId = Painel_LerPrimeiroID(wsPainel, colRegistar)
    End If

    If startId = "" Or Painel_EhSTOP(startId) Then
        Call Debug_Registar(0, "PIPELINE_" & CStr(pipelineIndex), "ALERTA", "", "StartID", _
            "Nao foi encontrado um ID inicial valido nas colunas INICIAR/REGISTAR.", _
            "Sugestao: coloque um Prompt ID na primeira linha da lista (linha " & CStr(LIST_START_ROW) & ").")
        GoTo SaidaLimpa
    End If

    Call Debug_Registar(0, pipelineNome, "INFO", "", "PIPELINE", _
        "Inicio pipeline. startId=[" & startId & "] maxSteps=" & CStr(maxSteps) & " maxRep=" & CStr(maxRep) & " inputFolder=[" & inputFolder & "]", _
        "OK")

    Dim runToken As String
    runToken = "RUN|" & pipelineNome & "|" & Format$(Now, "yyyymmdd_hhnnss") & "|P" & Format$(pipelineIndex, "00")

    ' Definir token de run (PAINEL) para separador visual na folha FILES_MANAGEMENT (M09)
    Call Files_SetRunToken(runToken)

    ' Garantir colunas/headers de ContextKV (idempotente)
    On Error Resume Next
    Call ContextKV_EnsureLayout
    On Error GoTo TrataErro
    ' Estruturas de controlo
    Dim visitas As Object
    Set visitas = CreateObject("Scripting.Dictionary")

    Dim ultimos4(1 To 4) As String
    ultimos4(1) = "": ultimos4(2) = "": ultimos4(3) = "": ultimos4(4) = ""

    Dim prevResponseId As String
    Dim prevResponseReusable As Boolean
    prevResponseId = ""
    prevResponseReusable = False

    Dim inicioHHMM As String
    inicioHHMM = Format$(Now, "hh:nn")

    Dim execCount As Long
    execCount = 0

    ' Execucao
    Dim atual As String
    atual = startId

    Dim cursorRow As Long
    cursorRow = LIST_START_ROW

    Dim passo As Long
    For passo = 1 To maxSteps

        Call Painel_StatusBar_Set(inicioHHMM, passo, maxSteps, execCount)
        DoEvents

        wsPainel.Cells(cursorRow, colIniciar).value = atual

        ' Controlo de repeticoes por ID
        Dim key As String
        key = UCase$(Trim$(atual))
        If Not visitas.exists(key) Then visitas.Add key, 0
        visitas(key) = CLng(visitas(key)) + 1

        If CLng(visitas(key)) > maxRep Then
            Call Debug_Registar(passo, atual, "ALERTA", "", "MaxRepetitions", _
                "Max Repetitions excedido para o ID: " & atual, _
                "Sugestao: reduza ciclos, ajuste Next PROMPT, ou aumente Max Repetitions no PAINEL.")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        ' Ler definicao da prompt
        Dim prompt As PromptDefinicao
        prompt = Catalogo_ObterPromptPorID(atual)

        If Trim$(prompt.textoPrompt) = "" Then
            Call Debug_Registar(passo, atual, "ERRO", "", "Catalogo", _
                "Prompt ID nao encontrado no catalogo: " & atual, _
                "Sugestao: confirme se existe uma linha com este ID e se o nome da folha coincide com o prefixo do ID.")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        Dim modeloUsado As String
        modeloUsado = Trim$(prompt.modelo)
        If modeloUsado = "" Then modeloUsado = modeloDefault

        Dim promptTextFinal As String
        Dim injectedVarsJson As String
        Dim extractedInputVarsJson As String
        Dim injectErro As String
        Dim injectOk As Boolean
        promptTextFinal = prompt.textoPrompt
        injectedVarsJson = ""
        extractedInputVarsJson = ""
        injectErro = ""
        injectOk = True

        Dim rawInputsText As String
        rawInputsText = Painel_TryReadInputsTextByPromptId(prompt.Id)

        Dim appendMode As String
        appendMode = Painel_GetConfigByKey(wsCfg, "INPUTS_APPEND_MODE", "RAW")

        Dim autoExtractInputVars As Boolean
        autoExtractInputVars = Painel_ConfigBoolByKey(wsCfg, "AUTO_INJECT_INPUT_VARS", True)

        Call Painel_ProcessInputsHybrid(pipelineNome, passo, prompt.Id, rawInputsText, appendMode, autoExtractInputVars, promptTextFinal, extractedInputVarsJson)

        ' ================================
        ' CONTEXTKV - INJECAO (ANTES DA API)
        ' ================================
        On Error Resume Next
        injectOk = ContextKV_InjectForStep(pipelineNome, passo, prompt.Id, outputFolderBase, runToken, promptTextFinal, injectedVarsJson, injectErro)
        If Err.Number <> 0 Then
            Call Debug_Registar(passo, prompt.Id, "ALERTA", "", "CONTEXT_KV", _
                "Erro ao executar ContextKV_InjectForStep: " & Err.Description, "")
            Err.Clear
            injectOk = True
            injectErro = ""
            injectedVarsJson = ""
            promptTextFinal = prompt.textoPrompt
        End If
        On Error GoTo TrataErro

        If injectOk = False Then
            Call Debug_Registar(passo, prompt.Id, "ERRO", "", "CONTEXT_KV", injectErro, "")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        ' Converter Config extra (amigavel) -> JSON (audit) / input override / extra fragment
        Dim auditJson As String, inputJsonLiteral As String, extraFragment As String
        Call ConfigExtra_Converter(prompt.ConfigExtra, prompt.textoPrompt, passo, prompt.Id, auditJson, inputJsonLiteral, extraFragment)

        ' Encadear previous_response_id, apenas se o config extra nao tiver conversation/previous_response_id
        If prevResponseId <> "" And prevResponseReusable Then
            If InStr(1, auditJson, """conversation""", vbTextCompare) = 0 And _
               InStr(1, auditJson, """previous_response_id""", vbTextCompare) = 0 Then
                extraFragment = Painel_AdicionarCampoJson(extraFragment, "previous_response_id", prevResponseId)
            End If
        End If

        ' -------------------------------
        ' FILES MANAGEMENT (M09)
        ' -------------------------------
        Dim inputJsonFinal As String
        Dim filesUsed As String, filesOps As String, fileIds As String
        Dim falhaCriticaFiles As Boolean, erroFiles As String

        Dim promptTemFiles As Boolean, promptTemRequiredFiles As Boolean
        Dim linhaFilesLista As String
        Call Painel_DeterminarFlagsFiles(atual, promptTemFiles, promptTemRequiredFiles, linhaFilesLista)

        ' Check 1 (antes do M09): prompt declara FILES mas INPUT Folder invalido
        If promptTemFiles Then
            If (Trim$(inputFolder) = "") Or (Dir(inputFolder, vbDirectory) = "") Then
                Call Debug_Registar(passo, prompt.Id, IIf(promptTemRequiredFiles, "ERRO", "ALERTA"), "", "INPUT Folder", _
                    "Prompt tem FILES mas INPUT Folder nao existe/esta vazio. inputFolder=[" & inputFolder & "]", _
                    "Sugestao: preencha INPUT Folder no PAINEL (linha 2) para esta pipeline.")
                If promptTemRequiredFiles Then
                    Call Seguimento_Registar(passo, prompt, modeloUsado, auditJson, 0, "", _
                        "[ERRO FILES] INPUT Folder invalido: " & inputFolder)
                    wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
                    GoTo SaidaLimpa
                End If
            End If
        End If

        Dim okFiles As Boolean
        If promptTemFiles Then
            okFiles = Files_PrepararContextoDaPrompt( _
                apiKey, _
                pipelineNome, _
                inputFolder, _
                prompt.Id, _
                prompt.textoPrompt, _
                inputJsonLiteral, _
                inputJsonFinal, _
                filesUsed, _
                filesOps, _
                fileIds, _
                falhaCriticaFiles, _
                erroFiles, _
                False, _
                passo _
            )
        Else
            ' Sem FILES declarados: nao chamar M09 (mais rapido / menos ruido)
            okFiles = True
            inputJsonFinal = inputJsonLiteral
            filesUsed = ""
            filesOps = ""
            fileIds = ""
            falhaCriticaFiles = False
            erroFiles = ""
        End If

        If (Not okFiles) Or falhaCriticaFiles Then
            Call Debug_Registar(passo, prompt.Id, "ERRO", "", "FILES", _
                "Falha critica a preparar contexto de ficheiros. inputFolder=[" & inputFolder & "] erro=[" & erroFiles & "]", _
                "Sugestao: confirme INPUT Folder e declaracoes FILES: no catalogo. Verifique tambem o upload (M09).")
            Call Seguimento_Registar(passo, prompt, modeloUsado, auditJson, 0, "", _
                "[ERRO FILES] " & erroFiles)
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        ' Check 2 + 3 (apos M09): input_file/input_image/text_embed e consistencia com fileIds
        If promptTemFiles Then
            Dim okFilesChecks As Boolean
            okFilesChecks = Painel_Files_Checks_Debug(passo, prompt.Id, pipelineNome, inputFolder, _
                                                     promptTemFiles, promptTemRequiredFiles, linhaFilesLista, _
                                                     inputJsonFinal, filesUsed, filesOps, fileIds)
            If Not okFilesChecks Then
                Call Seguimento_Registar(passo, prompt, modeloUsado, auditJson, 0, "", _
                    "[ERRO FILES] Prompt tem (required) mas nenhum ficheiro/imagem foi anexado.")
                wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
                GoTo SaidaLimpa
            End If
        End If

        Call Debug_Registar(passo, prompt.Id, "INFO", "", "REQ_INPUT_JSON", _
            "len=" & Len(inputJsonFinal) & _
            " | has_input_file=" & IIf(InStr(1, inputJsonFinal, """type"":""input_file""", vbTextCompare) > 0, "SIM", "NAO") & _
            " | has_input_image=" & IIf(InStr(1, inputJsonFinal, """type"":""input_image""", vbTextCompare) > 0, "SIM", "NAO") & _
            " | preview=" & Left$(inputJsonFinal, 350), _
            "Se has_input_file=NAO, o modelo nao pode receber PDF.")

        ' -------------------------------
        ' Chamada a API (1 chamada / passo)
        ' -------------------------------
        Dim resultado As ApiResultado
                ' -------------------------------
        ' FILE OUTPUT (pre-request): resolver config + preparar request (tools / structured outputs)
        ' -------------------------------
        Call FileOutput_ResolveEffectiveConfig(pipelineIndex, pipelineNome, prompt.Id, painelAutoSave, _
            fo_outputKind, fo_processMode, fo_autoSave, fo_overwriteMode, fo_prefixTmpl, fo_subfolderTmpl, _
            fo_structuredMode, fo_pptxMode, fo_xlsxMode, fo_pdfMode, fo_imageMode, prompt.ConfigExtra)

        modosEfetivo = prompt.modos
        extraFragmentFO = extraFragment
        Call FileOutput_PrepareRequest(fo_outputKind, fo_processMode, fo_structuredMode, modosEfetivo, extraFragmentFO)


        resultado = OpenAI_Executar(apiKey, modeloUsado, promptTextFinal, temperaturaDefault, maxTokensDefault, _
                                    modosEfetivo, prompt.storage, inputJsonFinal, extraFragmentFO, prompt.Id)

        execCount = execCount + 1
        Call Painel_StatusBar_Set(inicioHHMM, passo, maxSteps, execCount)
        DoEvents

                ' -------------------------------
        ' FILE OUTPUT (pos-response): guardar raw + ficheiros (metadata / code_interpreter)
        ' -------------------------------
        fo_filesUsedOut = "": fo_filesOpsOut = "": fo_logSeguimento = ""
        fo_logSeguimento = FileOutput_ProcessAfterResponse(apiKey, outputFolderBase, pipelineNome, pipelineIndex, passo, prompt.Id, resultado, _
            fo_outputKind, fo_processMode, fo_autoSave, fo_overwriteMode, fo_prefixTmpl, fo_subfolderTmpl, _
            fo_pptxMode, fo_xlsxMode, fo_pdfMode, fo_imageMode, fo_filesUsedOut, fo_filesOpsOut)

        If Trim$(resultado.Erro) <> "" Then
            textoSeguimento = "[ERRO] " & resultado.Erro
        ElseIf Trim$(fo_logSeguimento) <> "" Then
            textoSeguimento = fo_logSeguimento
        Else
            textoSeguimento = resultado.outputText
        End If

        filesUsedResumo = filesUsed
        If Trim$(fo_filesUsedOut) <> "" Then
            If Trim$(filesUsedResumo) <> "" Then
                filesUsedResumo = filesUsedResumo & " | " & fo_filesUsedOut
            Else
                filesUsedResumo = fo_filesUsedOut
            End If
        End If

        filesOpsResumo = filesOps
        If Trim$(fo_filesOpsOut) <> "" Then
            If Trim$(filesOpsResumo) <> "" Then
                filesOpsResumo = filesOpsResumo & " | " & fo_filesOpsOut
            Else
                filesOpsResumo = fo_filesOpsOut
            End If
        End If

        Call Seguimento_Registar(passo, prompt, modeloUsado, auditJson, resultado.httpStatus, resultado.responseId, _
            textoSeguimento, pipelineNome, "", filesUsedResumo, filesOpsResumo, fileIds)

        ' ================================
        ' CONTEXTKV - REGISTAR + CAPTURAR
        ' ================================
        On Error Resume Next
        Call ContextKV_WriteInjectedVars(pipelineNome, passo, prompt.Id, injectedVarsJson, outputFolderBase, runToken)
        Call Painel_WriteCapturedInputVars(pipelineNome, passo, prompt.Id, extractedInputVarsJson)
        Call ContextKV_CaptureRow(pipelineNome, passo, prompt.Id, outputFolderBase, runToken)
        If Err.Number <> 0 Then
            Call Debug_Registar(passo, prompt.Id, "ALERTA", "", "CONTEXT_KV", _
                "Erro em WriteInjectedVars/WriteCapturedInputVars/CaptureRow: " & Err.Description, "")
            Err.Clear
        End If
        On Error GoTo TrataErro


        If Trim$(resultado.Erro) <> "" Then
            Call Debug_Registar(passo, atual, "ERRO", "", "API", _
                resultado.Erro, _
                "Sugestao: verifique modelo, quota, payload e configuracao.")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        prevResponseId = resultado.responseId
        prevResponseReusable = (prompt.storage And Trim$(resultado.responseId) <> "")

        ' Ler Next config
        Dim nextPrompt As String, nextDefault As String, nextAllowed As String
        Call Catalogo_LerNextConfig(atual, nextPrompt, nextDefault, nextAllowed)

        ' Resolver proximo esperado com output (AUTO tenta extrair; senao default)
        Dim proximoEsperado As String
        proximoEsperado = Painel_ResolverNextComOutput(nextPrompt, nextDefault, resultado.outputText)
        If proximoEsperado = "" Then proximoEsperado = nextDefault

        If proximoEsperado = "" Or Painel_EhSTOP(proximoEsperado) Then
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        ' Integracao com a lista do PAINEL
        Dim nextRow As Long
        nextRow = cursorRow + 1

        Dim nextCursorRow As Long
        nextCursorRow = nextRow

        Dim idNaLinhaSeguinte As String
        idNaLinhaSeguinte = Trim$(CStr(wsPainel.Cells(nextRow, colIniciar).value))

        Dim proximoFinal As String
        proximoFinal = ""

        If idNaLinhaSeguinte <> "" Then
            If Painel_AllowedContem(nextAllowed, idNaLinhaSeguinte) Then
                If UCase$(Trim$(idNaLinhaSeguinte)) = UCase$(Trim$(proximoEsperado)) Then
                    Call Painel_FormatarCelulaID(wsPainel.Cells(nextRow, colIniciar), "AZUL", False)
                Else
                    Call Painel_FormatarCelulaID(wsPainel.Cells(nextRow, colIniciar), "LARANJA", False)
                End If
                proximoFinal = idNaLinhaSeguinte
            Else
                Call Painel_FormatarCelulaID(wsPainel.Cells(nextRow, colIniciar), "VERMELHO", True)

                If nextDefault = "" Then
                    Call Debug_Registar(passo, atual, "ALERTA", "", "NextDefault", _
                        "ID na linha seguinte nao permitido e Next PROMPT default esta vazio.", _
                        "Sugestao: defina Next PROMPT default na folha do catalogo.")
                    wsPainel.Cells(nextRow + 1, colIniciar).value = "STOP"
                    GoTo SaidaLimpa
                End If

                Call Painel_InserirCelulaNaColuna(wsPainel, colIniciar, nextRow + 1)

                wsPainel.Cells(nextRow + 1, colIniciar).value = nextDefault
                Call Painel_FormatarCelulaID(wsPainel.Cells(nextRow + 1, colIniciar), "VERDE", False)

                proximoFinal = nextDefault
                nextCursorRow = nextRow + 1
            End If
        Else
            If nextDefault = "" Then
                Call Debug_Registar(passo, atual, "ALERTA", "", "NextDefault", _
                    "Linha seguinte vazia e Next PROMPT default esta vazio.", _
                    "Sugestao: defina Next PROMPT default na folha do catalogo.")
                wsPainel.Cells(nextRow, colIniciar).value = "STOP"
                GoTo SaidaLimpa
            End If

            wsPainel.Cells(nextRow, colIniciar).value = nextDefault
            Call Painel_FormatarCelulaID(wsPainel.Cells(nextRow, colIniciar), "VERDE", False)
            proximoFinal = nextDefault
        End If

        If proximoFinal = "" Then proximoFinal = proximoEsperado

        proximoFinal = Painel_ValidarAllowedEExistencia(proximoFinal, nextDefault, nextAllowed, passo, atual)
        If proximoFinal = "" Or Painel_EhSTOP(proximoFinal) Then
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        ' Detecao ABAB
        Call Painel_AtualizarUltimos4(ultimos4, atual)
        If Painel_DetetouAlternanciaABAB(ultimos4) Then
            Call Debug_Registar(passo, atual, "ALERTA", "", "Ciclos", _
                "Detetada alternancia A-B-A-B. Pipeline interrompida.", _
                "Sugestao: adicione condicao de saida, restrinja allowed, ou introduza STOP.")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        atual = proximoFinal
        cursorRow = nextCursorRow

        If Painel_EhSTOP(atual) Then
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        If cursorRow > LIST_START_ROW + LIST_MAX_ROWS - 2 Then
            Call Debug_Registar(passo, atual, "ALERTA", "", "LimiteLista", _
                "A lista INICIAR excedeu o limite de linhas do PAINEL.", _
                "Sugestao: aumente LIST_MAX_ROWS no codigo ou reduza a pipeline.")
            wsPainel.Cells(LIST_START_ROW + LIST_MAX_ROWS - 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

    Next passo

    Call Debug_Registar(maxSteps, startId, "ALERTA", "", "MaxSteps", _
        "Max Steps atingido. Pipeline terminou por limite.", _
        "Sugestao: aumente Max Steps no PAINEL ou defina STOP.")
    wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"

SaidaLimpa:
    Application.StatusBar = False
    Application.DisplayStatusBar = oldDisplayStatusBar
    Application.EnableEvents = oldEnableEvents
    Exit Sub

TrataErro:
    Call Debug_Registar(0, "PIPELINE_" & CStr(pipelineIndex), "ERRO", "", "VBA", _
        "Erro inesperado: " & Err.Description, _
        "Sugestao: verifique IDs, folhas e referencias. Compile o VBAProject.")
    Resume SaidaLimpa
End Sub

' ============================================================
' 4.x) Ajudas pedidas: foco/DEBUG/status bar/checks de FILES
' ============================================================

Private Sub Painel_FocarSeguimentoA1()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEGUIMENTO)

    ws.Activate

    ' Garantir que a vista fica no canto superior esquerdo
    Application.GoTo ws.Range("A1"), True
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If

    ws.Range("A1").Select
    On Error GoTo 0
End Sub


Private Sub Painel_LimparDebugSessaoAnterior()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEBUG)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ws.rowS("2:" & CStr(lastRow)).ClearContents
    On Error GoTo 0
End Sub

Private Sub Painel_StatusBar_Set(ByVal inicioHHMM As String, ByVal passo As Long, ByVal total As Long, ByVal execCount As Long)
    On Error Resume Next

    Dim passoTxt As String
    If total > 10 Then
        passoTxt = Format$(passo, "00")
    Else
        passoTxt = CStr(passo)
    End If

    Application.StatusBar = "(" & inicioHHMM & ") Step: " & passoTxt & " of " & CStr(total) & "  |  Retry: " & CStr(execCount)
    On Error GoTo 0
End Sub

Private Sub Painel_ProcessInputsHybrid( _
    ByVal pipelineNome As String, _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal rawInputs As String, _
    ByVal appendModeRaw As String, _
    ByVal autoExtractVars As Boolean, _
    ByRef ioPromptText As String, _
    ByRef outExtractedVarsJson As String _
)
    Dim modeNorm As String
    modeNorm = UCase$(Trim$(appendModeRaw))
    If modeNorm = "" Then modeNorm = "RAW"
    If modeNorm <> "OFF" And modeNorm <> "SAFE" And modeNorm <> "RAW" Then
        Call Debug_Registar(passo, promptId, "ALERTA", "", "INPUTS_APPEND_MODE", _
            "Valor invalido para INPUTS_APPEND_MODE=['" & appendModeRaw & "']. A aplicar RAW.", _
            "Use OFF | SAFE | RAW na folha Config (coluna A/B).")
        modeNorm = "RAW"
    End If

    outExtractedVarsJson = ""
    If Trim$(rawInputs) = "" Then Exit Sub

    Dim lines() As String
    lines = Split(Painel_NormalizarQuebras(rawInputs), vbLf)

    Dim dictExtracted As Object
    Set dictExtracted = CreateObject("Scripting.Dictionary")

    Dim sbSafe As String
    Dim sbRaw As String
    sbSafe = ""
    sbRaw = ""

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lnRaw As String
        lnRaw = CStr(lines(i))

        Dim ln As String
        ln = Trim$(lnRaw)

        If ln <> "" Then
            If sbRaw <> "" Then sbRaw = sbRaw & vbCrLf
            sbRaw = sbRaw & ln
        End If

        If ln = "" Then GoTo NextLine

        Dim isTechnical As Boolean
        isTechnical = Painel_IsLinhaTecnicaInputs(ln)

        If Not isTechnical Then
            If sbSafe <> "" Then sbSafe = sbSafe & vbCrLf
            sbSafe = sbSafe & ln
        End If

        If autoExtractVars Then
            If Not isTechnical Then
                Call Painel_TryExtractInputVarLine(dictExtracted, pipelineNome, passo, promptId, ln)
            End If
        End If

NextLine:
    Next i

    If autoExtractVars Then
        outExtractedVarsJson = Painel_DictToJsonString(dictExtracted)
        If dictExtracted.Count > 0 Then
            Call Debug_Registar(passo, promptId, "INFO", "", "INPUTS_VARS", _
                "Variaveis extraidas de INPUTS: " & CStr(dictExtracted.Count) & ".", _
                "As keys foram normalizadas e guardadas em captured_vars no Seguimento.")
        Else
            Call Debug_Registar(passo, promptId, "ALERTA", "", "INPUTS_VARS", _
                "AUTO_INJECT_INPUT_VARS=TRUE mas nao foram extraidas variaveis de INPUTS.", _
                "Use linhas no formato CHAVE: valor ou CHAVE=valor.")
        End If
    End If

    Dim toAppend As String
    toAppend = ""
    If modeNorm = "RAW" Then toAppend = sbRaw
    If modeNorm = "SAFE" Then toAppend = sbSafe

    If modeNorm <> "OFF" And Trim$(toAppend) <> "" Then
        ioPromptText = ioPromptText & vbCrLf & vbCrLf & INPUTS_APPEND_HEADER & vbCrLf & toAppend
        Call Debug_Registar(passo, promptId, "INFO", "", "INPUTS_APPEND", _
            "INPUTS anexado ao prompt final (modo=" & modeNorm & "; chars=" & CStr(Len(toAppend)) & ").", _
            "Use INPUTS_APPEND_MODE=OFF para desativar, SAFE para remover linhas tecnicas.")
    End If
End Sub

Private Function Painel_TryReadInputsTextByPromptId(ByVal promptId As String) As String
    On Error GoTo EH

    Dim celId As Range
    Set celId = Catalogo_EncontrarCelulaID(promptId)
    If celId Is Nothing Then Exit Function

    Painel_TryReadInputsTextByPromptId = CStr(celId.Offset(2, 3).value)
    Exit Function
EH:
    Painel_TryReadInputsTextByPromptId = ""
End Function

Private Function Painel_NormalizarQuebras(ByVal s As String) As String
    Dim t As String
    t = Replace$(CStr(s), vbCrLf, vbLf)
    t = Replace$(t, vbCr, vbLf)
    Painel_NormalizarQuebras = t
End Function

Private Function Painel_IsLinhaTecnicaInputs(ByVal ln As String) As Boolean
    Dim u As String
    u = UCase$(Painel_RemoverAcentos(Trim$(ln)))

    If Left$(u, 6) = "FILES:" Or Left$(u, 7) = "FILES :" Then
        Painel_IsLinhaTecnicaInputs = True
        Exit Function
    End If
    If Left$(u, 10) = "FICHEIROS:" Or Left$(u, 11) = "FICHEIROS :" Then
        Painel_IsLinhaTecnicaInputs = True
        Exit Function
    End If
    If Left$(u, 22) = "OPERACOES COM FICHEIROS" Then
        Painel_IsLinhaTecnicaInputs = True
        Exit Function
    End If

    If InStr(1, u, "[PDF_UPLOAD]", vbTextCompare) > 0 Or _
       InStr(1, u, "[IMAGE_UPLOAD]", vbTextCompare) > 0 Or _
       InStr(1, u, "[TEXT_EMBED]", vbTextCompare) > 0 Or _
       InStr(1, u, "[INPUT_FILE]", vbTextCompare) > 0 Then
        Painel_IsLinhaTecnicaInputs = True
    End If
End Function

Private Sub Painel_TryExtractInputVarLine( _
    ByRef dictExtracted As Object, _
    ByVal pipelineNome As String, _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal ln As String _
)
    Dim pColon As Long, pEq As Long, pSep As Long
    pColon = InStr(1, ln, ":", vbTextCompare)
    pEq = InStr(1, ln, "=", vbTextCompare)

    If pColon > 0 And pEq > 0 Then
        pSep = IIf(pColon < pEq, pColon, pEq)
    ElseIf pColon > 0 Then
        pSep = pColon
    ElseIf pEq > 0 Then
        pSep = pEq
    Else
        Call Debug_Registar(passo, promptId, "ALERTA", "", "INPUTS_VARS", _
            "Linha de INPUTS ignorada (sem ':' ou '='): " & Left$(ln, 180), _
            "Use CHAVE: valor ou CHAVE=valor para extracao automatica.")
        Exit Sub
    End If

    Dim kRaw As String, vRaw As String
    kRaw = Trim$(Left$(ln, pSep - 1))
    vRaw = Trim$(Mid$(ln, pSep + 1))

    If kRaw = "" Then
        Call Debug_Registar(passo, promptId, "ALERTA", "", "INPUTS_VARS", _
            "Linha de INPUTS ignorada (chave vazia): " & Left$(ln, 180), _
            "Defina a chave antes de ':' ou '='.")
        Exit Sub
    End If

    Dim keyNorm As String
    keyNorm = Painel_NormalizarInputKey(kRaw)
    If keyNorm = "" Then
        Call Debug_Registar(passo, promptId, "ALERTA", "", "INPUTS_VARS", _
            "Linha de INPUTS ignorada (chave invalida): " & Left$(ln, 180), _
            "Use caracteres alfanumericos/underscore na chave.")
        Exit Sub
    End If

    If dictExtracted.exists(keyNorm) Then
        If CStr(dictExtracted(keyNorm)) <> vRaw Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "INPUTS_VARS", _
                "Conflito para chave '" & keyNorm & "' em INPUTS. Mantido primeiro valor.", _
                "Valor antigo=['" & Left$(CStr(dictExtracted(keyNorm)), 100) & "']; novo=['" & Left$(vRaw, 100) & "']")
        End If
        Exit Sub
    End If

    dictExtracted.Add keyNorm, vRaw
    Call Debug_Registar(passo, promptId, "INFO", "", "INPUTS_VARS", _
        "INPUTS extraido: " & keyNorm & "=['" & Left$(vRaw, 120) & "']", _
        "Disponivel em captured_vars deste passo.")
End Sub

Private Function Painel_NormalizarInputKey(ByVal k As String) As String
    Dim t As String
    t = UCase$(Trim$(Painel_RemoverAcentos(k)))
    t = Replace$(t, " ", "_")
    t = Replace$(t, "-", "_")
    t = Replace$(t, "/", "_")

    Dim i As Long, ch As String, sb As String
    sb = ""
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            sb = sb & ch
        End If
    Next i

    Do While InStr(1, sb, "__", vbTextCompare) > 0
        sb = Replace$(sb, "__", "_")
    Loop

    sb = Trim$(sb)
    If Len(sb) > 0 Then
        If Left$(sb, 1) = "_" Then sb = Mid$(sb, 2)
        If Len(sb) > 0 And Right$(sb, 1) = "_" Then sb = Left$(sb, Len(sb) - 1)
    End If

    Painel_NormalizarInputKey = sb
End Function

Private Function Painel_DictToJsonString(ByVal d As Object) As String
    On Error GoTo EH

    If d Is Nothing Then
        Painel_DictToJsonString = "{}"
        Exit Function
    End If

    Dim k As Variant
    Dim sb As String
    sb = "{"

    For Each k In d.keys
        If Len(sb) > 1 Then sb = sb & ","
        sb = sb & """" & JsonEscapar(CStr(k)) & """:""" & JsonEscapar(CStr(d(k))) & """"
    Next k

    sb = sb & "}"
    Painel_DictToJsonString = sb
    Exit Function
EH:
    Painel_DictToJsonString = "{}"
End Function

Private Function Painel_RemoverAcentos(ByVal s As String) As String
    Dim t As String
    t = s

    t = Replace$(t, ChrW(225), "a"): t = Replace$(t, ChrW(224), "a"): t = Replace$(t, ChrW(226), "a")
    t = Replace$(t, ChrW(227), "a"): t = Replace$(t, ChrW(228), "a")
    t = Replace$(t, ChrW(193), "A"): t = Replace$(t, ChrW(192), "A"): t = Replace$(t, ChrW(194), "A")
    t = Replace$(t, ChrW(195), "A"): t = Replace$(t, ChrW(196), "A")

    t = Replace$(t, ChrW(233), "e"): t = Replace$(t, ChrW(232), "e"): t = Replace$(t, ChrW(234), "e")
    t = Replace$(t, ChrW(235), "e"): t = Replace$(t, ChrW(201), "E"): t = Replace$(t, ChrW(200), "E")
    t = Replace$(t, ChrW(202), "E"): t = Replace$(t, ChrW(203), "E")

    t = Replace$(t, ChrW(237), "i"): t = Replace$(t, ChrW(236), "i"): t = Replace$(t, ChrW(238), "i")
    t = Replace$(t, ChrW(239), "i"): t = Replace$(t, ChrW(205), "I"): t = Replace$(t, ChrW(204), "I")
    t = Replace$(t, ChrW(206), "I"): t = Replace$(t, ChrW(207), "I")

    t = Replace$(t, ChrW(243), "o"): t = Replace$(t, ChrW(242), "o"): t = Replace$(t, ChrW(244), "o")
    t = Replace$(t, ChrW(245), "o"): t = Replace$(t, ChrW(246), "o")
    t = Replace$(t, ChrW(211), "O"): t = Replace$(t, ChrW(210), "O"): t = Replace$(t, ChrW(212), "O")
    t = Replace$(t, ChrW(213), "O"): t = Replace$(t, ChrW(214), "O")

    t = Replace$(t, ChrW(250), "u"): t = Replace$(t, ChrW(249), "u"): t = Replace$(t, ChrW(251), "u")
    t = Replace$(t, ChrW(252), "u"): t = Replace$(t, ChrW(218), "U"): t = Replace$(t, ChrW(217), "U")
    t = Replace$(t, ChrW(219), "U"): t = Replace$(t, ChrW(220), "U")

    t = Replace$(t, ChrW(231), "c"): t = Replace$(t, ChrW(199), "C")

    Painel_RemoverAcentos = t
End Function

Private Function Painel_GetConfigByKey(ByVal ws As Worksheet, ByVal keyName As String, ByVal defaultValue As String) As String
    On Error GoTo EH
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 1 To lastR
        If StrComp(Trim$(CStr(ws.Cells(i, 1).value)), keyName, vbTextCompare) = 0 Then
            Dim v As String
            v = Trim$(CStr(ws.Cells(i, 2).value))
            If v = "" Then v = defaultValue
            Painel_GetConfigByKey = v
            Exit Function
        End If
    Next i

    Painel_GetConfigByKey = defaultValue
    Exit Function
EH:
    Painel_GetConfigByKey = defaultValue
End Function

Private Function Painel_ConfigBoolByKey(ByVal ws As Worksheet, ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    Dim t As String
    t = UCase$(Painel_GetConfigByKey(ws, keyName, IIf(defaultValue, "TRUE", "FALSE")))
    If t = "TRUE" Or t = "VERDADEIRO" Or t = "SIM" Or t = "1" Then
        Painel_ConfigBoolByKey = True
    ElseIf t = "FALSE" Or t = "FALSO" Or t = "NAO" Or t = "NÃO" Or t = "0" Then
        Painel_ConfigBoolByKey = False
    Else
        Painel_ConfigBoolByKey = defaultValue
    End If
End Function

Private Sub Painel_EnsureConfigInputsModes(ByVal ws As Worksheet)
    On Error Resume Next
    Call Painel_EnsureConfigKey(ws, "INPUTS_APPEND_MODE", "RAW", "Controla como o texto de INPUTS e anexado ao prompt final: OFF=nao anexa; SAFE=remove linhas tecnicas (FILES/FICHEIROS/metadados internos); RAW=anexa integralmente. Default: RAW.")
    Call Painel_EnsureConfigKey(ws, "AUTO_INJECT_INPUT_VARS", "TRUE", "Quando TRUE, extrai pares CHAVE:valor ou CHAVE=valor da secao INPUTS para captured_vars (Seguimento) e logs no DEBUG. Nao substitui placeholders automaticamente no prompt.")
    On Error GoTo 0
End Sub

Private Sub Painel_EnsureConfigKey(ByVal ws As Worksheet, ByVal keyName As String, ByVal defaultValue As String, ByVal descriptionText As String)
    Dim lastR As Long, i As Long, r As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 8 Then lastR = 8
    r = 0

    For i = 1 To lastR
        If StrComp(Trim$(CStr(ws.Cells(i, 1).value)), keyName, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i

    If r = 0 Then r = lastR + 1

    ws.Cells(r, 1).value = keyName
    If Trim$(CStr(ws.Cells(r, 2).value)) = "" Then ws.Cells(r, 2).value = defaultValue
    If Trim$(CStr(ws.Cells(r, 3).value)) = "" Then ws.Cells(r, 3).value = descriptionText
End Sub

Private Sub Painel_WriteCapturedInputVars(ByVal pipelineNome As String, ByVal passo As Long, ByVal promptId As String, ByVal capturedVarsJson As String)
    On Error GoTo EH

    If Trim$(capturedVarsJson) = "" Then Exit Sub

    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SHEET_SEGUIMENTO)
    If wsS Is Nothing Then Exit Sub

    Dim rowS As Long
    rowS = Painel_FindSeguimentoRow(wsS, pipelineNome, passo, promptId)
    If rowS = 0 Then Exit Sub

    Dim colCap As Long, colMeta As Long
    colCap = Painel_FindColumnByHeader(wsS, "captured_vars")
    colMeta = Painel_FindColumnByHeader(wsS, "captured_vars_meta")
    If colCap = 0 Then Exit Sub

    wsS.Cells(rowS, colCap).value = capturedVarsJson
    If colMeta > 0 Then wsS.Cells(rowS, colMeta).value = "{""source"":""inputs_extract"",""mode"":""normalized_kv""}"

    Exit Sub
EH:
    Call Debug_Registar(passo, promptId, "ALERTA", "", "INPUTS_VARS", "Falha a gravar captured_vars no Seguimento: " & Err.Description, "")
End Sub

Private Function Painel_FindSeguimentoRow(ByVal wsS As Worksheet, ByVal pipelineNome As String, ByVal passo As Long, ByVal promptId As String) As Long
    On Error GoTo EH

    Dim colPasso As Long, colPrompt As Long, colPipe As Long
    colPasso = Painel_FindColumnByHeader(wsS, "Passo")
    colPrompt = Painel_FindColumnByHeader(wsS, "Prompt ID")
    colPipe = Painel_FindColumnByHeader(wsS, "Pipeline")

    If colPasso = 0 Or colPrompt = 0 Then Exit Function

    Dim lastR As Long, r As Long
    lastR = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If CLng(Val(CStr(wsS.Cells(r, colPasso).value))) = passo Then
            If StrComp(Trim$(CStr(wsS.Cells(r, colPrompt).value)), Trim$(promptId), vbTextCompare) = 0 Then
                If colPipe > 0 Then
                    If StrComp(Trim$(CStr(wsS.Cells(r, colPipe).value)), Trim$(pipelineNome), vbTextCompare) <> 0 Then GoTo NextR
                End If
                Painel_FindSeguimentoRow = r
                Exit Function
            End If
        End If
NextR:
    Next r
    Exit Function
EH:
    Painel_FindSeguimentoRow = 0
End Function

Private Function Painel_FindColumnByHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long
    On Error GoTo EH

    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            Painel_FindColumnByHeader = c
            Exit Function
        End If
    Next c

    Painel_FindColumnByHeader = 0
    Exit Function
EH:
    Painel_FindColumnByHeader = 0
End Function

Private Sub Painel_DeterminarFlagsFiles(ByVal promptId As String, ByRef outTemFiles As Boolean, ByRef outTemRequired As Boolean, ByRef outListaFiles As String)
    outTemFiles = False
    outTemRequired = False
    outListaFiles = ""

    Dim celId As Range
    Set celId = Catalogo_EncontrarCelulaID(promptId)
    If celId Is Nothing Then Exit Sub

    ' Layout do catalogo:
    '   - INPUTS: 2 linhas abaixo do ID, coluna D (Offset(2,3))
    Dim textoInputs As String
    textoInputs = CStr(celId.Offset(2, 3).value)

    outListaFiles = Painel_ExtrairListaAposTagFiles(textoInputs)

    If Trim$(outListaFiles) <> "" Then
        outTemFiles = True

        Dim low As String
        low = LCase$(outListaFiles)

        outTemRequired = (InStr(1, low, "(required)", vbTextCompare) > 0) Or _
                         (InStr(1, low, "(obrigatorio)", vbTextCompare) > 0) Or _
                         (InStr(1, low, "(obrigatoria)", vbTextCompare) > 0)
    End If
End Sub

Private Function Painel_ExtrairListaAposTagFiles(ByVal textoInputs As String) As String
    On Error GoTo Falha

    Dim t As String
    t = CStr(textoInputs)

    Dim p As Long
    p = InStr(1, t, "FILES:", vbTextCompare)
    If p = 0 Then p = InStr(1, t, "FILES :", vbTextCompare)
    If p = 0 Then p = InStr(1, t, "FICHEIROS:", vbTextCompare)
    If p = 0 Then p = InStr(1, t, "FICHEIROS :", vbTextCompare)

    If p = 0 Then
        Painel_ExtrairListaAposTagFiles = ""
        Exit Function
    End If

    Dim depois As String
    depois = Mid$(t, p)

    Dim p2 As Long
    p2 = InStr(1, depois, ":", vbTextCompare)
    If p2 = 0 Then
        Painel_ExtrairListaAposTagFiles = ""
        Exit Function
    End If

    Dim lista As String
    lista = Mid$(depois, p2 + 1)

    Dim eol As Long
    eol = InStr(1, lista, vbCrLf)
    If eol = 0 Then eol = InStr(1, lista, vbLf)
    If eol > 0 Then lista = Left$(lista, eol - 1)

    Painel_ExtrairListaAposTagFiles = Trim$(lista)
    Exit Function

Falha:
    Painel_ExtrairListaAposTagFiles = ""
End Function

Private Function Painel_Files_Checks_Debug( _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal pipelineNome As String, _
    ByVal inputFolder As String, _
    ByVal promptTemFiles As Boolean, _
    ByVal promptTemRequiredFiles As Boolean, _
    ByVal listaFiles As String, _
    ByVal inputJsonFinal As String, _
    ByVal filesUsed As String, _
    ByVal filesOps As String, _
    ByVal fileIds As String _
) As Boolean

    ' Roteiro (3 checks):
    '   Check 1: INPUT Folder valido quando FILES (feito antes de chamar M09)
    '   Check 2: inputJsonFinal contem input_file / input_image OU text_embed
    '   Check 3: logging de coerencia:
    '            - FILE_ID (upload): pode existir "file_id" e fileIds tipicamente nao vazio (apenas diagnostico)
    '            - INLINE_BASE64: existe "file_data"/"image_url" com data:... e e normal fileIds ficar vazio

    Painel_Files_Checks_Debug = True
    If Not promptTemFiles Then Exit Function

    Dim temInputFileOuImagem As Boolean
    temInputFileOuImagem = (InStr(1, inputJsonFinal, """type"":""input_file""", vbTextCompare) > 0) Or _
                           (InStr(1, inputJsonFinal, """type"":""input_image""", vbTextCompare) > 0)

    Dim temTextEmbed As Boolean
    temTextEmbed = (InStr(1, inputJsonFinal, "----- BEGIN FILE:", vbTextCompare) > 0)

    If (Not temInputFileOuImagem) And (Not temTextEmbed) Then

        Dim diag As String
        diag = Painel_DiagnosticoFilesNoFolder(inputFolder, listaFiles)

        Call Debug_Registar(passo, promptId, IIf(promptTemRequiredFiles, "ERRO", "ALERTA"), "", "FILES", _
            "Nenhum input_file/input_image/text_embed anexado. pipeline=[" & pipelineNome & _
            "] inputFolder=[" & inputFolder & _
            "] filesUsed=[" & filesUsed & _
            "] fileIds=[" & fileIds & _
            "] diag=[" & diag & "]", _
            "Sugestao: confirme a linha FILES: no INPUTS do catalogo. Se houver upload, verifique M09 (boundary/Send).")

        If promptTemRequiredFiles Then
            Painel_Files_Checks_Debug = False
        End If
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' Check 3 (CORRIGIDO): nao falhar em INLINE_BASE64 so por fileIds=""
    ' ------------------------------------------------------------
    If temInputFileOuImagem Then

        Dim temInlineBase64 As Boolean
        temInlineBase64 = (InStr(1, inputJsonFinal, """file_data"":""data:", vbTextCompare) > 0) Or _
                          (InStr(1, inputJsonFinal, """image_url"":""data:", vbTextCompare) > 0)

        Dim temFileIdNoJson As Boolean
        temFileIdNoJson = (InStr(1, inputJsonFinal, """file_id"":""", vbTextCompare) > 0)

        If temInlineBase64 Then
            Call Debug_Registar(passo, promptId, "INFO", "", "FILES", _
                "Anexacao OK via INLINE_BASE64 (file_data/image_url data:...). fileIds=[" & fileIds & "]", _
                "Nota: em INLINE_BASE64 e normal fileIds ficar vazio; o ficheiro segue inline no JSON.")
        ElseIf temFileIdNoJson Then
            If Trim$(fileIds) <> "" Then
                Call Debug_Registar(passo, promptId, "INFO", "", "FILES", _
                    "Anexacao OK (input_file/input_image com file_id). fileIds=[" & fileIds & "]", _
                    "Nota: file_id presente implica que houve upload anterior ou reutilizacao valida.")
            Else
                ' Nao falhar: o anexo existe (file_id no JSON). Isto e apenas incoerencia de logging.
                Call Debug_Registar(passo, promptId, "ALERTA", "", "FILES", _
                    "input_file/input_image com file_id no JSON, mas fileIds vazio (inconsistencia de logging).", _
                    "Sugestao: confirme se M09 devolve outFileIdsUsed e se a string fileIds esta a ser propagada no M07.")
            End If
        Else
            ' Formato inesperado (mas ainda existe input_file/input_image no JSON)
            Call Debug_Registar(passo, promptId, "ALERTA", "", "FILES", _
                "input_file/input_image presente, mas nao detetei file_id nem file_data/image_url data:. Formato inesperado.", _
                "Sugestao: reveja o JSON final e M09 (construcao do input_file/input_image).")
        End If

    ElseIf temTextEmbed Then
        Call Debug_Registar(passo, promptId, "INFO", "", "FILES", _
            "Anexacao OK via text_embed (conteudo extraido e inserido no input_text).", _
            "Nota: para PDFs/imagens, prefira input_file/input_image quando necessario.")
    End If
End Function

Private Function Painel_DiagnosticoFilesNoFolder(ByVal inputFolder As String, ByVal listaFiles As String) As String
    On Error GoTo Falha

    Dim folderOk As Boolean
    folderOk = (Trim$(inputFolder) <> "") And (Dir(inputFolder, vbDirectory) <> "")

    If Not folderOk Then
        Painel_DiagnosticoFilesNoFolder = "inputFolder_invalido"
        Exit Function
    End If

    If Trim$(listaFiles) = "" Then
        Painel_DiagnosticoFilesNoFolder = "sem_lista_FILES"
        Exit Function
    End If

    Dim itens() As String
    itens = Split(listaFiles, ";")

    Dim sb As String
    sb = ""

    Dim i As Long
    For i = LBound(itens) To UBound(itens)
        Dim raw As String
        raw = Trim$(itens(i))
        If raw <> "" Then
            Dim required As Boolean
            required = (InStr(1, raw, "(required)", vbTextCompare) > 0) Or _
                       (InStr(1, raw, "(obrigatorio)", vbTextCompare) > 0) Or _
                       (InStr(1, raw, "(obrigatoria)", vbTextCompare) > 0)

            Dim nome As String
            nome = raw
            Dim p As Long
            p = InStr(1, nome, "(", vbTextCompare)
            If p > 0 Then nome = Trim$(Left$(nome, p - 1))

            nome = Replace(nome, """", "")
            nome = Replace(nome, "'", "")
            nome = Trim$(nome)

            Dim found As String
            found = "NAO"

            Dim nExact As Long, nWild As Long, nSub As Long
            nExact = 0: nWild = 0: nSub = 0

            If nome <> "" Then
                If InStr(1, nome, "*", vbTextCompare) > 0 Then
                    nWild = Painel_DirCountMatches(inputFolder, nome)
                    If nWild > 0 Then found = "SIM"
                Else
                    If Len(Dir(inputFolder & "\" & nome)) > 0 Then
                        nExact = 1
                        found = "SIM"
                    Else
                        nSub = Painel_DirCountSubstringMatches(inputFolder, nome, 10)
                        If nSub > 0 Then found = "SIM"
                    End If
                End If
            End If

            Dim token As String
            token = nome & "|found=" & found & "|exact=" & CStr(nExact) & "|wild=" & CStr(nWild) & "|sub=" & CStr(nSub)
            If required Then token = token & "|required=SIM"

            If sb = "" Then
                sb = token
            Else
                sb = sb & " || " & token
            End If

            If Len(sb) > 600 Then
                sb = Left$(sb, 600) & "..."
                Exit For
            End If
        End If
    Next i

    Painel_DiagnosticoFilesNoFolder = sb
    Exit Function

Falha:
    Painel_DiagnosticoFilesNoFolder = "diagnostico_falhou"
End Function

Private Function Painel_DirCountMatches(ByVal folder As String, ByVal pattern As String) As Long
    On Error GoTo Falha
    Painel_DirCountMatches = 0

    Dim f As String
    f = Dir(folder & "\" & pattern)
    Do While f <> ""
        Painel_DirCountMatches = Painel_DirCountMatches + 1
        f = Dir()
        If Painel_DirCountMatches > 50 Then Exit Do
    Loop
    Exit Function

Falha:
    Painel_DirCountMatches = 0
End Function

Private Function Painel_DirCountSubstringMatches(ByVal folder As String, ByVal needle As String, ByVal maxCount As Long) As Long
    On Error GoTo Falha
    Painel_DirCountSubstringMatches = 0

    Dim f As String
    f = Dir(folder & "\*.*")

    Dim lowNeedle As String
    lowNeedle = LCase$(needle)

    Do While f <> ""
        If InStr(1, LCase$(f), lowNeedle, vbTextCompare) > 0 Then
            Painel_DirCountSubstringMatches = Painel_DirCountSubstringMatches + 1
            If Painel_DirCountSubstringMatches >= maxCount Then Exit Do
        End If
        f = Dir()
    Loop

    Exit Function

Falha:
    Painel_DirCountSubstringMatches = 0
End Function

' ============================================================
' ----- apoio visual / lista -----
' ============================================================

Private Sub Painel_ReporFormatoColunaIDs(ByVal ws As Worksheet, ByVal col As Long)
    Dim r As Long
    For r = LIST_START_ROW To LIST_START_ROW + LIST_MAX_ROWS - 1
        Dim v As String
        v = Trim$(CStr(ws.Cells(r, col).value))
        If v <> "" Then
            ws.Cells(r, col).Font.Color = vbBlack
            ws.Cells(r, col).Font.Bold = False
        End If
    Next r
End Sub

Private Sub Painel_FormatarCelulaID(ByVal cel As Range, ByVal corNome As String, ByVal negrito As Boolean)
    Dim cor As Long
    Select Case UCase$(Trim$(corNome))
        Case "AZUL": cor = RGB(0, 112, 192)
        Case "LARANJA": cor = RGB(255, 192, 0)
        Case "VERDE": cor = RGB(0, 176, 80)
        Case "VERMELHO": cor = vbRed
        Case Else: cor = vbBlack
    End Select

    cel.Font.Color = cor
    cel.Font.Bold = negrito
End Sub

Private Sub Painel_InserirCelulaNaColuna(ByVal ws As Worksheet, ByVal col As Long, ByVal insertRow As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rowS.Count, col).End(xlUp).Row
    If lastRow < insertRow Then Exit Sub

    If lastRow < LIST_START_ROW Then lastRow = LIST_START_ROW
    If lastRow > LIST_START_ROW + LIST_MAX_ROWS - 1 Then lastRow = LIST_START_ROW + LIST_MAX_ROWS - 1

    ws.Range(ws.Cells(insertRow, col), ws.Cells(lastRow, col)).Insert Shift:=xlDown
End Sub

' ============================================================
' 5) Catalogo: ler / validar Next PROMPT
' ============================================================

Private Sub Catalogo_LerNextConfig(ByVal promptId As String, ByRef nextPrompt As String, ByRef nextDefault As String, ByRef nextAllowed As String)
    nextPrompt = ""
    nextDefault = ""
    nextAllowed = "ALL; STOP"

    Dim celId As Range
    Set celId = Catalogo_EncontrarCelulaID(promptId)
    If celId Is Nothing Then Exit Sub

    Dim txtNext As String, txtDef As String, txtAllowed As String
    txtNext = CStr(celId.Offset(1, 1).value)
    txtDef = CStr(celId.Offset(2, 1).value)
    txtAllowed = CStr(celId.Offset(3, 1).value)

    nextPrompt = Painel_ExtrairValorAposDoisPontos(txtNext)
    nextDefault = Painel_ExtrairValorAposDoisPontos(txtDef)
    nextAllowed = Painel_ExtrairValorAposDoisPontos(txtAllowed)

    If nextAllowed = "" Then nextAllowed = "ALL; STOP"
End Sub

Private Function Catalogo_EncontrarCelulaID(ByVal promptId As String) As Range
    On Error GoTo Falha

    Dim nomeFolha As String
    nomeFolha = Painel_ExtrairFolhaDoID(promptId)
    If nomeFolha = "" Then GoTo Falha

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomeFolha)

    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=Trim$(promptId), LookIn:=xlValues, LookAt:=xlWhole)

    Set Catalogo_EncontrarCelulaID = rng
    Exit Function

Falha:
    Set Catalogo_EncontrarCelulaID = Nothing
End Function

' ============================================================
' 6) Resolucao do proximo passo
' ============================================================

Private Function Painel_ResolverNextDeterministico(ByVal nextPrompt As String, ByVal nextDefault As String) As String
    Dim n As String
    n = UCase$(Trim$(nextPrompt))

    If n = "" Then
        Painel_ResolverNextDeterministico = "STOP"
    ElseIf n = "AUTO" Then
        If Trim$(nextDefault) = "" Then
            Painel_ResolverNextDeterministico = "STOP"
        Else
            Painel_ResolverNextDeterministico = Trim$(nextDefault)
        End If
    Else
        Painel_ResolverNextDeterministico = Trim$(nextPrompt)
    End If
End Function

Private Function Painel_ResolverNextComOutput(ByVal nextPrompt As String, ByVal nextDefault As String, ByVal outputText As String) As String
    Dim n As String
    n = UCase$(Trim$(nextPrompt))

    If n = "" Then
        Painel_ResolverNextComOutput = "STOP"
        Exit Function
    End If

    If n = "STOP" Then
        Painel_ResolverNextComOutput = "STOP"
        Exit Function
    End If

    If n <> "AUTO" Then
        Painel_ResolverNextComOutput = Trim$(nextPrompt)
        Exit Function
    End If

    Dim extraido As String
    extraido = Painel_ExtrairNextPromptIdDoOutput(outputText)

    If extraido <> "" Then
        Painel_ResolverNextComOutput = extraido
    ElseIf Trim$(nextDefault) <> "" Then
        Painel_ResolverNextComOutput = Trim$(nextDefault)
    Else
        Painel_ResolverNextComOutput = "STOP"
    End If
End Function

Private Function Painel_ExtrairNextPromptIdDoOutput(ByVal outputText As String) As String
    Dim lines() As String
    lines = Split(outputText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(lines(i))

        If UCase$(Left$(ln, 14)) = "NEXT_PROMPT_ID" Then
            Dim p As Long
            p = InStr(1, ln, ":", vbTextCompare)
            If p > 0 Then
                Painel_ExtrairNextPromptIdDoOutput = Trim$(Mid$(ln, p + 1))
                Exit Function
            End If
        End If
    Next i

    Painel_ExtrairNextPromptIdDoOutput = ""
End Function

Private Function Painel_ValidarAllowedEExistencia(ByVal candidato As String, ByVal fallbackDefault As String, ByVal allowedList As String, ByVal passo As Long, ByVal promptAtual As String) As String
    Dim c As String
    c = Trim$(candidato)

    If c = "" Then
        Painel_ValidarAllowedEExistencia = "STOP"
        Exit Function
    End If

    If Painel_EhSTOP(c) Then
        Painel_ValidarAllowedEExistencia = "STOP"
        Exit Function
    End If

    If Not Painel_AllowedContem(allowedList, c) Then
        If Trim$(fallbackDefault) <> "" And Painel_AllowedContem(allowedList, Trim$(fallbackDefault)) Then
            Call Debug_Registar(passo, promptAtual, "ALERTA", "", "Next allowed", _
                "NEXT_PROMPT nao permitido; usei default.", _
                "Sugestao: ajuste allowed ou output em AUTO.")
            c = Trim$(fallbackDefault)
        Else
            Call Debug_Registar(passo, promptAtual, "ALERTA", "", "Next allowed", _
                "NEXT_PROMPT nao permitido; terminei em STOP.", _
                "Sugestao: ajuste allowed/default.")
            Painel_ValidarAllowedEExistencia = "STOP"
            Exit Function
        End If
    End If

    Dim p As PromptDefinicao
    p = Catalogo_ObterPromptPorID(c)

    If Trim$(p.textoPrompt) = "" Then
        Call Debug_Registar(passo, promptAtual, "ALERTA", "", "Next prompt", _
            "NEXT_PROMPT nao existe no catalogo; terminei em STOP.", _
            "Sugestao: confirme o ID e a folha.")
        Painel_ValidarAllowedEExistencia = "STOP"
        Exit Function
    End If

    Painel_ValidarAllowedEExistencia = c
End Function

Private Function Painel_AllowedContem(ByVal allowedList As String, ByVal candidato As String) As Boolean
    Dim a As String
    a = UCase$(Trim$(allowedList))

    If a = "" Then
        Painel_AllowedContem = True
        Exit Function
    End If

    If InStr(1, a, "ALL", vbTextCompare) > 0 Then
        Painel_AllowedContem = True
        Exit Function
    End If

    Dim itens() As String
    itens = Split(a, ";")

    Dim c As String
    c = UCase$(Trim$(candidato))

    Dim i As Long
    For i = LBound(itens) To UBound(itens)
        If UCase$(Trim$(itens(i))) = c Then
            Painel_AllowedContem = True
            Exit Function
        End If
    Next i

    If c = "STOP" Then
        Painel_AllowedContem = True
        Exit Function
    End If

    Painel_AllowedContem = False
End Function

' ============================================================
' 7) Utilitarios PAINEL
' ============================================================

Private Sub Painel_ObterColunasPipeline(ByVal pipelineIndex As Long, ByRef colIniciar As Long, ByRef colRegistar As Long)
    colIniciar = 2 + (pipelineIndex - 1) * 2
    colRegistar = colIniciar + 1
End Sub

Private Sub Painel_LerLimitesPipeline(ByVal wsPainel As Worksheet, ByVal pipelineIndex As Long, ByRef maxSteps As Long, ByRef maxRep As Long)
    Dim colIniciar As Long, colRegistar As Long
    Call Painel_ObterColunasPipeline(pipelineIndex, colIniciar, colRegistar)

    maxSteps = CLng(val(wsPainel.Cells(6, colIniciar).value))
    maxRep = CLng(val(wsPainel.Cells(6, colIniciar).value))

    If maxSteps <= 0 Then maxSteps = 20
    If maxRep <= 0 Then maxRep = 3
End Sub

Private Function Painel_LerNomePipeline(ByVal wsPainel As Worksheet, ByVal pipelineIndex As Long) As String
    Dim colIniciar As Long, colRegistar As Long
    Call Painel_ObterColunasPipeline(pipelineIndex, colIniciar, colRegistar)
    Painel_LerNomePipeline = Trim$(CStr(wsPainel.Cells(1, colIniciar).value))
    If Painel_LerNomePipeline = "" Then Painel_LerNomePipeline = "Pipeline_" & Format$(pipelineIndex, "00")
End Function

Private Function Painel_LerPrimeiroID(ByVal wsPainel As Worksheet, ByVal col As Long) As String
    Dim v As String
    v = Trim$(CStr(wsPainel.Cells(LIST_START_ROW, col).value))
    If v = "" Then
        Painel_LerPrimeiroID = ""
    Else
        Painel_LerPrimeiroID = v
    End If
End Function

Private Sub Painel_LimparLista(ByVal wsPainel As Worksheet, ByVal col As Long)
    Dim r As Long
    For r = LIST_START_ROW To LIST_START_ROW + LIST_MAX_ROWS - 1
        wsPainel.Cells(r, col).value = ""
    Next r
End Sub

Private Function Painel_LerListaIDsUnica(ByVal wsPainel As Worksheet, ByVal col As Long) As Collection
    Dim lista As New Collection
    Dim vistos As Object
    Set vistos = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = LIST_START_ROW To LIST_START_ROW + LIST_MAX_ROWS - 1
        Dim pid As String
        pid = Trim$(CStr(wsPainel.Cells(r, col).value))

        If pid = "" Then Exit For
        If Painel_EhSTOP(pid) Then Exit For

        Dim k As String
        k = UCase$(pid)
        If Not vistos.exists(k) Then
            vistos.Add k, True
            lista.Add pid
        End If
    Next r

    Set Painel_LerListaIDsUnica = lista
End Function

Private Function Painel_ExtrairFolhaDoID(ByVal promptId As String) As String
    Dim p As Long
    p = InStr(1, promptId, "/")
    If p = 0 Then
        Painel_ExtrairFolhaDoID = ""
    Else
        Painel_ExtrairFolhaDoID = Left$(promptId, p - 1)
    End If
End Function

Private Function Painel_ExtrairValorAposDoisPontos(ByVal texto As String) As String
    Dim t As String
    t = Trim$(texto)
    If t = "" Then
        Painel_ExtrairValorAposDoisPontos = ""
        Exit Function
    End If

    Dim p As Long
    p = InStr(1, t, ":", vbTextCompare)
    If p = 0 Then
        Painel_ExtrairValorAposDoisPontos = Trim$(t)
    Else
        Painel_ExtrairValorAposDoisPontos = Trim$(Mid$(t, p + 1))
    End If
End Function

Private Function Painel_EhSTOP(ByVal s As String) As Boolean
    Painel_EhSTOP = (UCase$(Trim$(s)) = "STOP")
End Function

Private Function Painel_SanitizarNomeFicheiro(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)

    t = Replace(t, "\", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, ":", "_")
    t = Replace(t, "*", "_")
    t = Replace(t, "?", "_")
    t = Replace(t, """", "_")
    t = Replace(t, "<", "_")
    t = Replace(t, ">", "_")
    t = Replace(t, "|", "_")

    If t = "" Then t = "Pipeline"
    Painel_SanitizarNomeFicheiro = t
End Function

' ---- ABAB (array de 4 ultimos IDs) ----

Private Sub Painel_AtualizarUltimos4(ByRef ultimos4() As String, ByVal atual As String)
    ultimos4(1) = ultimos4(2)
    ultimos4(2) = ultimos4(3)
    ultimos4(3) = ultimos4(4)
    ultimos4(4) = atual
End Sub

Private Function Painel_DetetouAlternanciaABAB(ByRef ultimos4() As String) As Boolean
    If Trim$(ultimos4(1)) = "" Then
        Painel_DetetouAlternanciaABAB = False
        Exit Function
    End If

    If (UCase$(ultimos4(1)) = UCase$(ultimos4(3))) And (UCase$(ultimos4(2)) = UCase$(ultimos4(4))) Then
        If UCase$(ultimos4(1)) <> UCase$(ultimos4(2)) Then
            Painel_DetetouAlternanciaABAB = True
            Exit Function
        End If
    End If

    Painel_DetetouAlternanciaABAB = False
End Function

' ---- JSON helper para acrescentar campo ao fragmento ----

Private Function Painel_AdicionarCampoJson(ByVal extraFragmentSemInput As String, ByVal chave As String, ByVal valor As String) As String
    Dim frag As String
    frag = Trim$(extraFragmentSemInput)

    Dim par As String
    par = """" & chave & """:""" & JsonEscaparSimples(valor) & """"

    If frag = "" Then
        Painel_AdicionarCampoJson = par
    Else
        Painel_AdicionarCampoJson = frag & "," & par
    End If
End Function

Private Function JsonEscaparSimples(ByVal s As String) As String
    JsonEscaparSimples = Json_EscapeString(CStr(s))
End Function
