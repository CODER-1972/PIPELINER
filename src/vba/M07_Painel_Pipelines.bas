Attribute VB_Name = "M07_Painel_Pipelines"
Option Explicit

' =============================================================================
' Módulo: M07_Painel_Pipelines
' Propósito:
' - Orquestrar execução de pipelines a partir da folha PAINEL e ações de botões.
' - Gerir limites, fluxo de passos, integração com catálogo/API/logs e geração de mapa/registo.
'
' Atualizações:
' - 2026-02-28 | Codex | Ajusta total exibido em "Step x of y" para total planeado da lista
'   - Passa a calcular total visível por passo com base em "Row n de z" (prompts planeados).
'   - Mantém Max Steps como limite duro de execução, evitando mostrar total inferior ao passo atual.
' - 2026-02-28 | Codex | Status bar passa a exibir Prompt ID completo por fase
'   - Inclui o prompt ID entre "Row n de z" e o detalhe da fase (preparacao/upload/execucao/resposta).
'   - Mantem formato existente de Step/Retry/Row para preservar compatibilidade visual no PAINEL.
' - 2026-02-27 | Codex | Corrige leitura dos limites Max Steps/Max Repetitions no PAINEL
'   - Ajusta Painel_LerLimitesPipeline para ler Max Steps da linha 5 e Max Repetitions da linha 6.
'   - Elimina bug em que Max Steps herdava indevidamente o valor de Max Repetitions.
' - 2026-02-27 | Codex | Fingerprint operacional e mensagens FILES mais explicativas
'   - Injeta fingerprint textual (pipeline/step/prompt/mode) na chamada M05 para correlação com M10.
'   - Refina mensagens REQ_INPUT_JSON e text_embed para diferenciar anexação textual de upload com file_id.
' - 2026-02-26 | Codex | STOP passa a encerrar contagem de Row n de z no PAINEL
'   - Ajusta Painel_ContarPromptsPlaneados para terminar no primeiro STOP encontrado.
'   - Evita que IDs residuais abaixo de STOP inflem o total mostrado na status bar.
' - 2026-02-26 | Codex | Lookup robusto de IDs no catálogo para evitar STOP falso
'   - Adiciona fallback por varrimento normalizado de IDs quando o Find exato não encontra a célula.
'   - Normaliza quebras de linha/NBSP/TAB em IDs para tolerar colagens de DOCX/CSV com lixo invisível.
' - 2026-02-26 | Codex | Contagem robusta de "Row n de z" com lista INICIAR esparsa
'   - Ignora STOP/lacunas intermedias quando ainda existem IDs validos abaixo na coluna INICIAR.
'   - Calcula rowPos pelo indice logico de prompts validos, evitando "Row 1 de 1" falso com listas maiores.
' - 2026-02-26 | Codex | Status bar com posicao da linha no PAINEL
'   - Mostra "Row n de z" com base na lista INICIAR (exclui STOP) para contexto visual do utilizador.
'   - Mantem "Step x of y" como limite de execucao (Max Steps), sem quebrar compatibilidade.
' - 2026-02-23 | Codex | Execução de Output Orders (EXECUTE: LOAD_CSV)
'   - Executa parser/whitelist de ordens após File Output em respostas com sucesso.
'   - Acrescenta logs de importação CSV em files_ops_log sem quebrar fluxo sem EXECUTE.
' - 2026-02-18 | Codex | Enriquecimento da status bar por fase operacional
'   - Permite detalhar a fase atual (ex.: preparacao, upload de ficheiros, chamada API).
'   - Atualiza a barra de estado antes de operacoes criticas para feedback em tempo real.
' - 2026-02-17 | Codex | Corrige fonte do texto enviado ao M09
'   - Passa promptTextFinal (com INPUTS_DECLARADOS_NO_CATALOGO) para Files_PrepararContextoDaPrompt.
'   - Evita perda de URLS_ENTRADA/FILES no input_text quando há anexos.
' - 2026-02-17 | Codex | Injecao explicita de INPUTS (incluindo FILES/FICHEIROS) no texto enviado ao modelo
'   - Anexa ao prompt final as linhas operacionais do INPUTS (URLS_ENTRADA, MODO_DE_VERIFICACAO e FILES/FICHEIROS).
'   - Mantem o anexo tecnico dos ficheiros no fluxo M09; bloco textual passa a ser informativo para o modelo.
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
'   3) Status bar: "(hh:mm) Step: x of y  |  Retry: z  |  Row n de z"
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
    prevResponseId = ""

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


        wsPainel.Cells(cursorRow, colIniciar).value = atual

        Dim rowPos As Long
        Dim rowTotal As Long
        Dim stepTotalVisivel As Long
        rowPos = Painel_PosicaoPromptPlaneado(wsPainel, colIniciar, cursorRow)
        rowTotal = Painel_ContarPromptsPlaneados(wsPainel, colIniciar)
        If rowTotal < rowPos Then rowTotal = rowPos
        stepTotalVisivel = Painel_TotalVisivelStep(maxSteps, rowTotal, passo)

        Call Painel_StatusBar_Set(inicioHHMM, passo, stepTotalVisivel, execCount, "A preparar passo", rowPos, rowTotal)
        DoEvents

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
        Dim injectErro As String
        Dim injectOk As Boolean
        promptTextFinal = prompt.textoPrompt
        injectedVarsJson = ""
        injectErro = ""
        injectOk = True

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

        ' INPUTS declarados no catalogo (incluindo FILES/FICHEIROS) seguem para o modelo.
        ' O bloco textual e informativo; o anexo tecnico de ficheiros continua em M09.
        Call Painel_AnexarInputsTextuaisAoPrompt(prompt.Id, promptTextFinal)

        ' Converter Config extra (amigavel) -> JSON (audit) / input override / extra fragment
        Dim auditJson As String, inputJsonLiteral As String, extraFragment As String
        Call ConfigExtra_Converter(prompt.ConfigExtra, promptTextFinal, passo, prompt.Id, auditJson, inputJsonLiteral, extraFragment)

        ' Encadear previous_response_id, apenas se o config extra nao tiver conversation/previous_response_id
        If prevResponseId <> "" Then
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
            Call Painel_StatusBar_Set(inicioHHMM, passo, stepTotalVisivel, execCount, "Uploading file", rowPos, rowTotal)
            DoEvents

            okFiles = Files_PrepararContextoDaPrompt( _
                apiKey, _
                pipelineNome, _
                inputFolder, _
                prompt.Id, _
                promptTextFinal, _
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
            "Input final construído; este retrato confirma se anexos seguiram como file/image ou apenas texto. Se esperava PDF e has_input_file=NAO, corrigir FILES/mode.")

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


        Call Painel_StatusBar_Set(inicioHHMM, passo, stepTotalVisivel, execCount, "A executar prompt", rowPos, rowTotal)
        DoEvents

        Dim debugFingerprintSeed As String
        debugFingerprintSeed = "pipeline=" & pipelineNome & "|step=" & CStr(passo) & "|prompt=" & prompt.Id & "|mode=" & LCase$(Trim$(fo_outputKind)) & "/" & LCase$(Trim$(fo_processMode))

        resultado = OpenAI_Executar(apiKey, modeloUsado, promptTextFinal, temperaturaDefault, maxTokensDefault, _
                                    modosEfetivo, prompt.storage, inputJsonFinal, extraFragmentFO, prompt.Id, debugFingerprintSeed)

        execCount = execCount + 1
        Call Painel_StatusBar_Set(inicioHHMM, passo, stepTotalVisivel, execCount, "Resposta recebida", rowPos, rowTotal)
        DoEvents

                ' -------------------------------
        ' FILE OUTPUT (pos-response): guardar raw + ficheiros (metadata / code_interpreter)
        ' -------------------------------
        fo_filesUsedOut = "": fo_filesOpsOut = "": fo_logSeguimento = ""
        fo_logSeguimento = FileOutput_ProcessAfterResponse(apiKey, outputFolderBase, pipelineNome, pipelineIndex, passo, prompt.Id, resultado, _
            fo_outputKind, fo_processMode, fo_autoSave, fo_overwriteMode, fo_prefixTmpl, fo_subfolderTmpl, _
            fo_pptxMode, fo_xlsxMode, fo_pdfMode, fo_imageMode, fo_filesUsedOut, fo_filesOpsOut)
        Dim fo_executeOpsLog As String
        fo_executeOpsLog = ""
        If Trim$(resultado.Erro) = "" And resultado.httpStatus >= 200 And resultado.httpStatus < 300 Then
            fo_executeOpsLog = OutputOrders_TryExecute(passo, prompt.Id, resultado.responseId, resultado.outputText, outputFolderBase, fo_filesOpsOut)
            If Trim$(fo_executeOpsLog) <> "" Then
                If Trim$(fo_filesOpsOut) <> "" Then
                    fo_filesOpsOut = fo_filesOpsOut & " | " & fo_executeOpsLog
                Else
                    fo_filesOpsOut = fo_executeOpsLog
                End If
            End If
        End If

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
        Call ContextKV_CaptureRow(pipelineNome, passo, prompt.Id, outputFolderBase, runToken)
        If Err.Number <> 0 Then
            Call Debug_Registar(passo, prompt.Id, "ALERTA", "", "CONTEXT_KV", _
                "Erro em WriteInjectedVars/CaptureRow: " & Err.Description, "")
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

Private Function Painel_TotalVisivelStep(ByVal maxSteps As Long, ByVal rowTotal As Long, ByVal passoAtual As Long) As Long
    Dim total As Long

    total = rowTotal
    If total <= 0 Then total = maxSteps
    If total <= 0 Then total = passoAtual
    If total < passoAtual Then total = passoAtual

    Painel_TotalVisivelStep = total
End Function

Private Sub Painel_StatusBar_Set(ByVal inicioHHMM As String, ByVal passo As Long, ByVal total As Long, ByVal execCount As Long, Optional ByVal detalhe As String = "", Optional ByVal rowPos As Long = 0, Optional ByVal rowTotal As Long = 0)
    On Error Resume Next

    Dim passoTxt As String
    If total > 10 Then
        passoTxt = Format$(passo, "00")
    Else
        passoTxt = CStr(passo)
    End If

    Dim detalheLimpo As String
    detalheLimpo = Trim$(CStr(detalhe))

    Dim promptIdLimpo As String
    promptIdLimpo = Trim$(CStr(promptId))

    Dim rowLabel As String
    rowLabel = ""
    If rowTotal > 0 Then
        If rowPos <= 0 Then rowPos = 1
        If rowPos > rowTotal Then rowPos = rowTotal
        rowLabel = "  |  Row " & CStr(rowPos) & " de " & CStr(rowTotal)
    End If

    Application.StatusBar = "(" & inicioHHMM & ") Step: " & passoTxt & " of " & CStr(total) & "  |  Retry: " & CStr(execCount) & _
                            rowLabel & IIf(promptIdLimpo = "", "", "  |  " & promptIdLimpo) & _
                            IIf(detalheLimpo = "", "", "  |  " & detalheLimpo)
    On Error GoTo 0
End Sub

Private Function Painel_ContarPromptsPlaneados(ByVal wsPainel As Worksheet, ByVal colIniciar As Long) As Long
    On Error GoTo Falha

    Dim total As Long
    total = 0

    Dim r As Long
    For r = LIST_START_ROW To LIST_START_ROW + LIST_MAX_ROWS - 1
        Dim v As String
        v = Trim$(CStr(wsPainel.Cells(r, colIniciar).value))

        If v = "" Then
            If Painel_ExistePromptValidoAbaixo(wsPainel, colIniciar, r + 1) Then
                GoTo ProximaLinha
            Else
                Exit For
            End If
        End If

        If Painel_EhSTOP(v) Then Exit For

        total = total + 1
ProximaLinha:
    Next r

    Painel_ContarPromptsPlaneados = total
    Exit Function

Falha:
    Painel_ContarPromptsPlaneados = 0
End Function

Private Function Painel_PosicaoPromptPlaneado(ByVal wsPainel As Worksheet, ByVal colIniciar As Long, ByVal targetRow As Long) As Long
    On Error GoTo Falha

    If targetRow < LIST_START_ROW Then
        Painel_PosicaoPromptPlaneado = 1
        Exit Function
    End If

    Dim total As Long
    total = 0

    Dim r As Long
    For r = LIST_START_ROW To targetRow
        Dim v As String
        v = Trim$(CStr(wsPainel.Cells(r, colIniciar).value))

        If v <> "" And Not Painel_EhSTOP(v) Then total = total + 1
    Next r

    If total <= 0 Then total = 1
    Painel_PosicaoPromptPlaneado = total
    Exit Function

Falha:
    Painel_PosicaoPromptPlaneado = 1
End Function

Private Function Painel_ExistePromptValidoAbaixo(ByVal wsPainel As Worksheet, ByVal colIniciar As Long, ByVal fromRow As Long) As Boolean
    On Error GoTo Falha

    If fromRow > LIST_START_ROW + LIST_MAX_ROWS - 1 Then Exit Function

    Dim r As Long
    For r = fromRow To LIST_START_ROW + LIST_MAX_ROWS - 1
        Dim v As String
        v = Trim$(CStr(wsPainel.Cells(r, colIniciar).value))
        If v <> "" And Not Painel_EhSTOP(v) Then
            Painel_ExistePromptValidoAbaixo = True
            Exit Function
        End If
    Next r

    Exit Function

Falha:
    Painel_ExistePromptValidoAbaixo = False
End Function


Private Sub Painel_AnexarInputsTextuaisAoPrompt(ByVal promptId As String, ByRef ioPromptText As String)
    On Error GoTo Falha

    Dim blocoInputs As String
    blocoInputs = Painel_ExtrairInputsTextuais(promptId)

    If Trim$(blocoInputs) = "" Then Exit Sub

    ioPromptText = RTrim$(CStr(ioPromptText)) & vbCrLf & vbCrLf & _
                   "INPUTS_DECLARADOS_NO_CATALOGO:" & vbCrLf & blocoInputs
    Exit Sub

Falha:
    ' Silencioso por compatibilidade retroativa.
End Sub

Private Function Painel_ExtrairInputsTextuais(ByVal promptId As String) As String
    On Error GoTo Falha

    Dim celId As Range
    Set celId = Catalogo_EncontrarCelulaID(promptId)
    If celId Is Nothing Then Exit Function

    Dim textoInputs As String
    textoInputs = CStr(celId.Offset(2, 3).value)
    If Trim$(textoInputs) = "" Then Exit Function

    Dim linhas() As String
    linhas = Split(Replace(textoInputs, vbCrLf, vbLf), vbLf)

    Dim i As Long
    Dim acc As String
    acc = ""

    For i = LBound(linhas) To UBound(linhas)
        Dim linha As String
        linha = Trim$(CStr(linhas(i)))

        If linha <> "" Then
            If acc <> "" Then acc = acc & vbCrLf
            acc = acc & linha
        End If
    Next i

    Painel_ExtrairInputsTextuais = acc
    Exit Function

Falha:
    Painel_ExtrairInputsTextuais = ""
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
            "Anexo processado por extração/embebido de texto; neste modo não existe file_id.", _
            "Se precisares de layout visual (ex.: PDF), mudar para modo upload.")
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

    If rng Is Nothing Then
        Dim alvoNormalizado As String
        alvoNormalizado = Painel_NormalizarPromptIdKey(promptId)

        If alvoNormalizado <> "" Then
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            Dim r As Long
            For r = 1 To lastRow
                If Painel_NormalizarPromptIdKey(CStr(ws.Cells(r, 1).value)) = alvoNormalizado Then
                    Set rng = ws.Cells(r, 1)
                    Exit For
                End If
            Next r
        End If
    End If

    Set Catalogo_EncontrarCelulaID = rng
    Exit Function

Falha:
    Set Catalogo_EncontrarCelulaID = Nothing
End Function

Private Function Painel_NormalizarPromptIdKey(ByVal rawId As String) As String
    Dim k As String
    k = CStr(rawId)

    k = Replace(k, vbCr, "")
    k = Replace(k, vbLf, "")
    k = Replace(k, vbTab, "")
    k = Replace(k, Chr$(160), " ")
    k = Application.WorksheetFunction.Trim(k)

    Painel_NormalizarPromptIdKey = UCase$(Trim$(k))
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

    maxSteps = CLng(val(wsPainel.Cells(5, colIniciar).value))
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
