Attribute VB_Name = "M07_Painel_Pipelines"
Option Explicit

' =============================================================================
' Modulo: M07_Painel_Pipelines
' Proposito:
' - Orquestrar execucao de pipelines a partir da folha PAINEL e acoes de botoes.
' - Gerir limites, fluxo de passos, integracao com catalogo/API/logs e geracao de mapa/registo.
'
' Atualizações:
' - 2026-03-12 | Codex | `Step x of y` passa a usar estimativa da execucao real
'   - Substitui `y` baseado na lista planeada do PAINEL por estimativa dinamica da cadeia Next PROMPT a partir da prompt atual.
'   - Em prompts AUTO, usa fallback deterministico (Next default) e interrompe estimativa em STOP/loop/ID invalido ou Max Steps.
' - 2026-03-12 | Codex | Retry da status bar passa a refletir novas tentativas reais de API
'   - Substitui contador de chamadas por acumulador de retries reais devolvidos pelo M05 (`ApiResultado.retryCount`).
'   - Mantem formato visual da barra e atualiza o valor apos cada resposta recebida.
' - 2026-03-11 | Codex | Ativa bootstrap da folha GIT LOG no arranque da pipeline
'   - Quando o toggle Git LOG estiver ON, chama `GitLog_EnsureSheet` antes da execucao para garantir schema base da folha.
'   - Em falha de inicializacao, regista ALERTA no DEBUG e segue o run sem bloquear a pipeline.
' - 2026-03-11 | Codex | Corrige compilacao em exportacao TSV do DEBUG
'   - Alinha salto condicional com label local `NextRow` em `Painel_DebugSheetToTsv`.
'   - Elimina referencia a `ProximaLinha` inexistente que gerava `Compile error: Label not defined`.
' - 2026-03-09 | Codex | Snapshot DEBUG atualizado no fecho do passo antes de sair/avancar
'   - Regrava o espelho no catalogo imediatamente antes de cada saida antecipada do passo.
'   - Regrava no caminho nominal apos validacoes de NEXT para capturar logs emitidos apos a chamada API.
'   - Reduz perdas de linhas da prompt causadas por snapshot tirado demasiado cedo no fluxo do passo.
' - 2026-03-09 | Codex | Corrige desvio de fluxo no snapshot DEBUG (label de salto)
'   - Substitui salto para label inexistente por label local `NextRow`, removendo `Compile error: Label not defined`.
'   - Mantem filtro por Prompt ID/Passo, ignorando apenas linhas fora de contexto no loop de exportacao TSV.
' - 2026-03-09 | Codex | Snapshot DEBUG focado no contexto da prompt executada
'   - Filtra linhas por Prompt ID e por Passo (fallback) para manter eventos da prompt mesmo quando o Prompt ID vem vazio.
'   - Mantem cabecalho da linha 1 e ordem original das linhas do DEBUG no TSV.
'   - Evita truncagem prematura por ruido de outras prompts ao espelhar apenas o contexto relevante no catalogo.
' - 2026-03-09 | Codex | Espelho do DEBUG passa a copiar desde a linha 1 e quebra por caracteres
'   - Remove filtro por Prompt ID no snapshot para refletir integralmente a folha DEBUG no catalogo.
'   - Inclui a linha 1 (cabecalhos) e preserva a numeracao original das linhas no TSV.
'   - Divide texto entre linhas 3/4 por comprimento real (32767 chars por celula), continuando no ponto exato da interrupcao.
' - 2026-03-08 | Codex | Espelho DEBUG filtrado por prompt com quebra automatica para 4a linha
'   - Copia apenas linhas do DEBUG da prompt executada, prefixando cada linha com o numero da linha original.
'   - Mantem TSV sem wrap na coluna Notas para desenvolvimento e divide entre linhas +2/+3 quando excede o limite da celula.
'   - Emite alerta quando excede tambem a 4a linha e ocorre truncagem controlada por limite de 32767 caracteres.
' - 2026-03-08 | Codex | Espelho do DEBUG no catalogo apos cada prompt executada
'   - Copia a tabela DEBUG para uma unica celula TSV sem wrap em Notas para desenvolvimento (linha +2 do bloco).
'   - Escreve cabecalho "DEBUG [dd-mm-yyyy hh:mm]" a negrito em Notas para desenvolvimento (linha +1) e aplica fundo salmao claro.
'   - Regista no DEBUG sucesso/alerta/erro da persistencia para diagnostico leigo.
' - 2026-03-07 | Codex | Feedback visual do toggle Git LOG no PAINEL
'   - Aplica fundo azul claro acinzentado + texto a negrito quando Git LOG estiver ON.
'   - Repoe estilo padrao do botao quando Git LOG estiver OFF.
' - 2026-03-07 | Codex | Toggle Git LOG por pipeline no PAINEL para gatilho de auto-upload
'   - Cria botao "Git LOG ON/OFF" abaixo de INICIAR (linha 9) para cada pipeline com estado independente.
'   - Quando ON, ativa export Git no fim do run como equivalente ao auto-save=debug, sem exigir mudanca estrutural no PAINEL.
' - 2026-03-04 | Codex | Auto-upload Git de artefactos de debug por pipeline
'   - Ativa exportacao no fim da execucao quando auto-guardar contem "sim, todos" ou "debug".
'   - Publica CSV de DEBUG/Seguimento/catalogo e TXT do PAINEL via Git Data API usando GH_* no Config.
' - 2026-03-04 | Codex | Hardening de gates por etapa + validacoes preventivas de Next
'   - Adiciona validacao preventiva de coerencia default em allowed antes da resolucao do proximo passo.
'   - Emite diagnostico especifico quando NEXT_PROMPT_ID nao e encontrado em modo AUTO (fallback para default/STOP).
'   - Inclui lint operacional para passos CSV/EXECUTE sem diagnostic_contract ativo no Config extra.
'   - Lint reconhece aliases de chave de contrato, reporta process_mode e reduz falso alerta por token generico CSV.
' - 2026-03-02 | Codex | Foco inicial no DEBUG e destaque visual de fim de prompt
'   - Ao clicar INICIAR, ativa DEBUG!A1 para acompanhar novas linhas em tempo real.
'   - Regista stage=step_completed no fim de cada passo para suportar destaque verde no DEBUG.
' - 2026-03-01 | Codex | DEBUG_SCHEMA_VERSION=2 com relatorio paralelo DEBUG_DIAG
'   - Mede tempos de subfases (files/api/directives) e regista diagnostico aditivo em DEBUG_DIAG quando DEBUG_LEVEL>=DIAG.
'   - Mantem DEBUG legado intacto e adiciona bundle opcional via DEBUG_BUNDLE sem impacto em BASE.
' - 2026-03-01 | Codex | Passa intencao CI resolvida para o M05 (debug e anti-falso-supressao)
'   - Calcula ciIntentResolved via modo efetivo file/code_interpreter e envia para OpenAI_Executar.
'   - Mantem compatibilidade para chamadas antigas (parametro opcional no M05).
' - 2026-03-01 | Codex | Breadcrumbs STEP_STAGE para diagnostico pre-API
'   - Regista fases enter_step/catalog_loaded/config_parsed/files_prepare/before_api no DEBUG.
'   - Facilita triagem quando a pipeline para sem erro e sem novas linhas em Seguimento.
' - 2026-03-01 | Codex | Aumenta granularidade dos breadcrumbs para isolamento de bloqueios
'   - Adiciona marcos antes/depois de ContextKV, anexacao de INPUTS e inicio do parse de Config extra.
'   - Permite distinguir bloqueio em pre-processamento interno vs parse/FILES vs chamada API.
' - 2026-03-01 | Codex | Evita falha silenciosa sem linha no Seguimento em erro inesperado
'   - Captura ultimo stage do passo (mStepLastStage) e inclui no erro de excecao do pipeline.
'   - Em `TrataErro`, tenta escrever linha tecnica no Seguimento com passo/prompt/stage para auditoria minima.
' - 2026-03-01 | Codex | Refina alerta de downgrade para reduzir falso positivo
'   - Limita M07_FILEOUTPUT_MODE_MISMATCH a cenarios com intencao explicita de File Output no Config extra/outputKind.
'   - Padroniza mensagem com PROBLEMA/IMPACTO/ACAO/DETALHE para triagem rapida no DEBUG.
' - 2026-03-01 | Codex | Alerta preventivo para downgrade silencioso de modo de File Output
'   - Regista M07_FILEOUTPUT_MODE_MISMATCH quando Code Interpreter esta ativo nos modos mas o modo efetivo nao e file/code_interpreter.
'   - Regista M07_FILEOUTPUT_PARSE_GUARD quando Config extra menciona output_kind/process_mode mas cai em text/metadata por parseavel invalido.
' - 2026-03-01 | Codex | Corrige dependencia invalida de helper privado entre modulos
'   - Substitui chamadas `Nz(...)` por helper local `Painel_Nz(...)` nas rotinas de validacao FILES do modulo.
'   - Elimina `Compile error: Sub or Function not defined` sem alterar comportamento funcional.
' - 2026-03-01 | Codex | D1 bloqueante aware de wildcard/latest
'   - Resolve padroes FILES com wildcard/latest para nomes reais antes da comparacao com filesUsed.
'   - Evita falso negativo de INPUTFILES_MISSING quando M09 seleciona nome final por pattern.
' - 2026-02-28 | Codex | Corrige compile error na status bar por variavel nao declarada
'   - Adiciona parametro opcional promptId em Painel_StatusBar_Set para suportar exibicao do ID atual.
'   - Atualiza chamadas durante o passo (preparacao/upload/execucao/resposta) para passar prompt ID consistente.
' - 2026-02-28 | Codex | Ajusta total exibido em "Step x of y" para total planeado da lista
'   - Passa a calcular total visivel por passo com base em "Row n de z" (prompts planeados).
'   - Mantem Max Steps como limite duro de execucao, evitando mostrar total inferior ao passo atual.
' - 2026-02-28 | Codex | Status bar passa a exibir Prompt ID completo por fase
'   - Inclui o prompt ID entre "Row n de z" e o detalhe da fase (preparacao/upload/execucao/resposta).
'   - Mantem formato existente de Step/Retry/Row para preservar compatibilidade visual no PAINEL.
' - 2026-02-27 | Codex | Corrige leitura dos limites Max Steps/Max Repetitions no PAINEL
'   - Ajusta Painel_LerLimitesPipeline para ler Max Steps da linha 5 e Max Repetitions da linha 6.
'   - Elimina bug em que Max Steps herdava indevidamente o valor de Max Repetitions.
' - 2026-02-27 | Codex | Fingerprint operacional e mensagens FILES mais explicativas
'   - Injeta fingerprint textual (pipeline/step/prompt/mode) na chamada M05 para correlacao com M10.
'   - Refina mensagens REQ_INPUT_JSON e text_embed para diferenciar anexacao textual de upload com file_id.
' - 2026-02-26 | Codex | STOP passa a encerrar contagem de Row n de z no PAINEL
'   - Ajusta Painel_ContarPromptsPlaneados para terminar no primeiro STOP encontrado.
'   - Evita que IDs residuais abaixo de STOP inflem o total mostrado na status bar.
' - 2026-02-26 | Codex | Lookup robusto de IDs no catalogo para evitar STOP falso
'   - Adiciona fallback por varrimento normalizado de IDs quando o Find exato nao encontra a celula.
'   - Normaliza quebras de linha/NBSP/TAB em IDs para tolerar colagens de DOCX/CSV com lixo invisivel.
' - 2026-02-26 | Codex | Contagem robusta de "Row n de z" com lista INICIAR esparsa
'   - Ignora STOP/lacunas intermedias quando ainda existem IDs validos abaixo na coluna INICIAR.
'   - Calcula rowPos pelo indice logico de prompts validos, evitando "Row 1 de 1" falso com listas maiores.
' - 2026-02-26 | Codex | Status bar com posicao da linha no PAINEL
'   - Mostra "Row n de z" com base na lista INICIAR (exclui STOP) para contexto visual do utilizador.
'   - Mantem "Step x of y" como limite de execucao (Max Steps), sem quebrar compatibilidade.
' - 2026-02-23 | Codex | Execucao de Output Orders (EXECUTE: LOAD_CSV)
'   - Executa parser/whitelist de ordens apos File Output em respostas com sucesso.
'   - Acrescenta logs de importacao CSV em files_ops_log sem quebrar fluxo sem EXECUTE.
' - 2026-02-18 | Codex | Enriquecimento da status bar por fase operacional
'   - Permite detalhar a fase atual (ex.: preparacao, upload de ficheiros, chamada API).
'   - Atualiza a barra de estado antes de operacoes criticas para feedback em tempo real.
' - 2026-02-17 | Codex | Corrige fonte do texto enviado ao M09
'   - Passa promptTextFinal (com INPUTS_DECLARADOS_NO_CATALOGO) para Files_PrepararContextoDaPrompt.
'   - Evita perda de URLS_ENTRADA/FILES no input_text quando ha anexos.
' - 2026-02-17 | Codex | Injecao explicita de INPUTS (incluindo FILES/FICHEIROS) no texto enviado ao modelo
'   - Anexa ao prompt final as linhas operacionais do INPUTS (URLS_ENTRADA, MODO_DE_VERIFICACAO e FILES/FICHEIROS).
'   - Mantem o anexo tecnico dos ficheiros no fluxo M09; bloco textual passa a ser informativo para o modelo.
' - 2026-02-16 | Codex | Resolucao de API key via ambiente com fallback compativel
'   - Substitui leitura direta de Config!B1 por resolver central (M14_ConfigApiKey).
'   - Regista ALERTA/ERRO no DEBUG para origem/falhas da credencial sem expor segredo.
' - 2026-02-12 | Codex | Implementacao do padrao de header obrigatorio
'   - Adiciona proposito, historico de alteracoes e inventario de rotinas publicas.
'   - Mantem documentacao tecnica do modulo alinhada com AGENTS.md.
'
' Funcoes e procedimentos (inventario publico):
' - Painel_CriarBotoes (Sub): rotina publica do modulo.
' - Painel_Click_Iniciar (Sub): rotina publica do modulo.
' - Painel_Click_GitLog (Sub): alterna estado ON/OFF do gatilho Git LOG por pipeline.
' - Painel_GitLog_ApplyButtonStyle (Sub): aplica estilo visual ON/OFF ao botao Git LOG.
' - Painel_Click_Registar (Sub): rotina publica do modulo.
' - Painel_Click_SetDefault (Sub): rotina publica do modulo.
' - Painel_Click_CriarMapa (Sub): rotina publica do modulo.
' - Painel_StatusBar_SetPhase (Private Sub): resolve indice/total da fase interna e atualiza a status bar.
' - Painel_BuildInternalPhasePlan (Private Function): constroi lista dinâmica de fases do passo atual.
' - Painel_PhasePlanIndex (Private Function): obtém posição de uma fase no plano dinâmico.
' - Painel_HasFileOutputIntent (Private Function): heurística de intenção de output para incluir fase condicional.
' - Painel_LogStepStage (Private Sub): breadcrumb de fase no DEBUG para troubleshooting pre-API.
' - Painel_EspelharDebugNoCatalogo (Private Sub): grava snapshot TSV do DEBUG no bloco da prompt executada.
' - Painel_RegistarFalhaNoSeguimento (Private Sub): fallback de auditoria no Seguimento para erros inesperados.
' - Painel_GitLog_RegisterStepExecution (Private Sub): envia resumo normalizado da prompt concluida para a folha GIT LOG.
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
'   3) Status bar: "(hh:mm) Step: x of y  |  Retry: z  |  Row n of z (pipeline)"
'   4) Check/diagnostico de FILES (3 checks) + logging util
' ============================================================

Private Const SHEET_PAINEL As String = "PAINEL"
Private Const SHEET_CONFIG As String = "Config"
Private Const SHEET_SEGUIMENTO As String = "Seguimento"
Private Const SHEET_DEBUG As String = "DEBUG"

Private Const LIST_START_ROW As Long = 10
Private Const LIST_MAX_ROWS As Long = 40

Private Const PIPELINES As Long = 10

Private mStepLastStage As String

' Prefixos de nomes de botoes (Shapes) criados no PAINEL
Private Const BTN_PREFIX As String = "BTN_"
Private Const BTN_INICIAR As String = "BTN_INICIAR_"
Private Const BTN_REGISTAR As String = "BTN_REGISTAR_"
Private Const BTN_SETDEFAULT As String = "BTN_SETDEFAULT_"
Private Const BTN_MAPA As String = "BTN_MAPA_"
Private Const BTN_GITLOG As String = "BTN_GITLOG_"

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
        Call Painel_CriarBotaoEmCelula(ws, ws.Cells(9, colIniciar), BTN_GITLOG & Format$(i, "00"), Painel_GitLog_Label(i), "Painel_Click_GitLog")
        Call Painel_GitLog_RefreshButtonCaption(i)
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



Public Sub Painel_Click_GitLog()
    Dim idx As Long
    idx = Painel_ExtrairIndiceDoCaller(BTN_GITLOG)
    If idx = 0 Then Exit Sub

    Call Painel_GitLog_SetEnabled(idx, Not Painel_GitLog_IsEnabled(idx))
    Call Painel_GitLog_RefreshButtonCaption(idx)
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


Private Function Painel_GitLog_Name(ByVal pipelineIndex As Long) As String
    Painel_GitLog_Name = "PIPE_GITLOG_" & Format$(pipelineIndex, "00")
End Function

Private Function Painel_GitLog_Label(ByVal pipelineIndex As Long) As String
    If Painel_GitLog_IsEnabled(pipelineIndex) Then
        Painel_GitLog_Label = "Git LOG ON"
    Else
        Painel_GitLog_Label = "Git LOG OFF"
    End If
End Function

Private Function Painel_GitLog_IsEnabled(ByVal pipelineIndex As Long) As Boolean
    On Error GoTo Fallback

    Dim nm As Name
    Set nm = ThisWorkbook.Names(Painel_GitLog_Name(pipelineIndex))

    Dim raw As String
    raw = UCase$(Replace(CStr(nm.RefersTo), "=", ""))
    raw = Replace(raw, """", "")
    raw = Trim$(raw)

    Painel_GitLog_IsEnabled = (raw = "ON" Or raw = "TRUE" Or raw = "1")
    Exit Function

Fallback:
    Painel_GitLog_IsEnabled = False
End Function

Private Sub Painel_GitLog_SetEnabled(ByVal pipelineIndex As Long, ByVal enabled As Boolean)
    On Error Resume Next

    Dim refValue As String
    If enabled Then
        refValue = "=""ON"""
    Else
        refValue = "=""OFF"""
    End If

    Dim keyName As String
    keyName = Painel_GitLog_Name(pipelineIndex)

    ThisWorkbook.Names(keyName).Delete
    ThisWorkbook.Names.Add Name:=keyName, RefersTo:=refValue

    On Error GoTo 0
End Sub

Private Sub Painel_GitLog_RefreshButtonCaption(ByVal pipelineIndex As Long)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PAINEL)

    Dim shp As Shape
    Set shp = ws.Shapes(BTN_GITLOG & Format$(pipelineIndex, "00"))

    If shp Is Nothing Then Exit Sub

    shp.TextFrame.Characters.text = Painel_GitLog_Label(pipelineIndex)
    Call Painel_GitLog_ApplyButtonStyle(shp, Painel_GitLog_IsEnabled(pipelineIndex))

    On Error GoTo 0
End Sub

Private Sub Painel_GitLog_ApplyButtonStyle(ByVal shp As Shape, ByVal isOn As Boolean)
    On Error Resume Next

    Const COLOR_ON As Long = 14737632   ' RGB(224,236,255) - azul claro acinzentado
    Const COLOR_OFF As Long = 15658734  ' RGB(238,238,238) - cinza claro padrao

    shp.Fill.Visible = -1

    If isOn Then
        shp.Fill.ForeColor.RGB = COLOR_ON
        shp.TextFrame.Characters.Font.Bold = True
    Else
        shp.Fill.ForeColor.RGB = COLOR_OFF
        shp.TextFrame.Characters.Font.Bold = False
    End If

    On Error GoTo 0
End Sub

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
        Call Painel_ValidarConsistenciaNextConfig(nextPrompt, nextDefault, nextAllowed, 0, atual)

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

    ' ---- FILE OUTPUT (declaracoes) ----
    Dim fo_outputKind As String, fo_processMode As String, fo_autoSave As String, fo_overwriteMode As String
    Dim fo_prefixTmpl As String, fo_subfolderTmpl As String, fo_structuredMode As String
    Dim fo_pptxMode As String, fo_xlsxMode As String, fo_pdfMode As String, fo_imageMode As String
    Dim modosEfetivo As String, extraFragmentFO As String
    Dim fo_filesUsedOut As String, fo_filesOpsOut As String, fo_logSeguimento As String
    Dim textoSeguimento As String, filesUsedResumo As String, filesOpsResumo As String
    Dim runExecutouPassos As Boolean


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

    If Painel_GitLog_IsEnabled(pipelineIndex) Then
        painelAutoSave = "debug"

        Dim wsGitLog As Worksheet
        Set wsGitLog = GitLog_EnsureSheet()
        If wsGitLog Is Nothing Then
            Call Debug_Registar(0, pipelineNome, "ALERTA", "", "GIT LOG", _
                "Nao foi possivel inicializar a folha GIT LOG; o run vai continuar sem bootstrap da folha.", _
                "Sugestao: verificar permissao de escrita/estado do workbook e repetir a execucao.")
        End If
    End If

    Dim runDumpFolder As String
    runDumpFolder = outputFolderBase
    If Trim$(runDumpFolder) = "" Then runDumpFolder = Environ$("TEMP")
    If Right$(runDumpFolder, 1) = "\" Then runDumpFolder = Left$(runDumpFolder, Len(runDumpFolder) - 1)
    runDumpFolder = runDumpFolder & "\DEBUG_PAYLOAD_DUMPS\" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Replace$(Replace$(pipelineNome, " ", "_"), "/", "_")
    Call M05_SetRunDumpFolder(runDumpFolder, pipelineNome)

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

    ' Garantir foco no DEBUG (monitorizacao em tempo real)
    Call Painel_FocarDebugA1

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

    Dim retryCountTotal As Long
    retryCountTotal = 0

    mStepLastStage = "pipeline_start"

    ' Execucao
    Dim atual As String
    atual = startId

    Dim cursorRow As Long
    cursorRow = LIST_START_ROW

    Dim passo As Long
    Dim passoCtx As Long
    Dim promptCtx As String
    Dim stepStartAt As Date, stepEndAt As Date
    Dim filesPrepareMs As Long, apiCallMs As Long, directiveParseMs As Long
    Dim tMark As Double
    passoCtx = 0
    promptCtx = ""

    For passo = 1 To maxSteps
        runExecutouPassos = True
        passoCtx = passo
        promptCtx = atual
        stepStartAt = Now
        filesPrepareMs = 0
        apiCallMs = 0
        directiveParseMs = 0
        wsPainel.Cells(cursorRow, colIniciar).value = atual

        Dim rowPos As Long
        Dim rowTotal As Long
        Dim hasDiagContract As Boolean
        Dim hasOutputIntent As Boolean
        rowPos = Painel_PosicaoPromptPlaneado(wsPainel, colIniciar, cursorRow)
        rowTotal = Painel_ContarPromptsPlaneados(wsPainel, colIniciar)
        If rowTotal < rowPos Then rowTotal = rowPos

        Call Painel_LogStepStage(passo, atual, "enter_step", "row=" & CStr(rowPos) & "/" & CStr(rowTotal))
        Call Painel_LogStepStage(passo, atual, "before_context_inject", "")
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
        Call Painel_LogStepStage(passo, atual, "catalog_loaded", "lenPrompt=" & CStr(Len(prompt.textoPrompt)))

        hasDiagContract = Painel_HasDiagnosticContractConfigured(prompt.ConfigExtra)
        hasOutputIntent = Painel_HasFileOutputIntent(prompt.textoPrompt, prompt.ConfigExtra, prompt.modos)

        Call Painel_DeterminarFlagsFiles(atual, promptTemFiles, promptTemRequiredFiles, linhaFilesLista)
        Call Painel_StatusBar_SetPhase(inicioHHMM, "prepare", retryCountTotal, rowPos, rowTotal, atual, promptTemFiles, hasDiagContract, hasOutputIntent, "A preparar passo")

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
        Call Painel_LogStepStage(passo, prompt.Id, "after_context_inject", "injectOk=" & IIf(injectOk, "SIM", "NAO"))

        Call Painel_StatusBar_SetPhase(inicioHHMM, "context", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "A montar contexto")

        If injectOk = False Then
            Call Debug_Registar(passo, prompt.Id, "ERRO", "", "CONTEXT_KV", injectErro, "")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            GoTo SaidaLimpa
        End If

        ' INPUTS declarados no catalogo (incluindo FILES/FICHEIROS) seguem para o modelo.
        ' O bloco textual e informativo; o anexo tecnico de ficheiros continua em M09.
        Call Painel_LogStepStage(passo, prompt.Id, "before_inputs_attach", "")
        Call Painel_AnexarInputsTextuaisAoPrompt(prompt.Id, promptTextFinal)
        Call Painel_LogStepStage(passo, prompt.Id, "after_inputs_attach", "lenPrompt=" & CStr(Len(promptTextFinal)))

        ' Converter Config extra (amigavel) -> JSON (audit) / input override / extra fragment
        Dim auditJson As String, inputJsonLiteral As String, extraFragment As String
        Call Painel_LogStepStage(passo, prompt.Id, "config_parse_start", "lenConfigExtra=" & CStr(Len(prompt.ConfigExtra)))
        Call ConfigExtra_Converter(prompt.ConfigExtra, promptTextFinal, passo, prompt.Id, auditJson, inputJsonLiteral, extraFragment)
        Call Painel_LogStepStage(passo, prompt.Id, "config_parsed", _
            "lenAudit=" & CStr(Len(auditJson)) & "|lenInputJson=" & CStr(Len(inputJsonLiteral)))

        ' Encadear previous_response_id, apenas se o config extra nao tiver conversation/previous_response_id
        If prevResponseId <> "" Then
            If InStr(1, auditJson, """conversation""", vbTextCompare) = 0 And _
               InStr(1, auditJson, """previous_response_id""", vbTextCompare) = 0 Then
                extraFragment = Painel_AdicionarCampoJson(extraFragment, "previous_response_id", prevResponseId)
            End If
        End If

        If Painel_HasCsvExecuteIntent(prompt.textoPrompt, prompt.ConfigExtra) And _
           (Not Painel_HasDiagnosticContractConfigured(prompt.ConfigExtra)) Then
            Dim processModeHint As String
            processModeHint = Painel_ConfigExtraGetValue(prompt.ConfigExtra, "process_mode")
            Call Debug_Registar(passo, prompt.Id, "ALERTA", "", "CONFIG_LINT_CONTRACT", _
                "Prompt com intencao CSV/EXECUTE sem contrato diagnostico ativo no Config extra." & _
                IIf(Trim$(processModeHint) <> "", " process_mode=" & processModeHint & ".", ""), _
                "Sugestao: adicionar diagnostic_contract: ci_csv_v1 para ativar gate de prova minima antes de EXECUTE.")
        End If

        ' -------------------------------
        ' FILES MANAGEMENT (M09)
        ' -------------------------------
        Dim inputJsonFinal As String
        Dim filesUsed As String, filesOps As String, fileIds As String
        Dim falhaCriticaFiles As Boolean, erroFiles As String


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
        tMark = Timer
        If promptTemFiles Then
            Call Painel_StatusBar_SetPhase(inicioHHMM, "files_prepare", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "Uploading file")
            Call Painel_LogStepStage(passo, prompt.Id, "files_prepare_start", "temFiles=SIM")
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
            Call Painel_LogStepStage(passo, prompt.Id, "files_prepare_skip", "temFiles=NAO")
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

        filesPrepareMs = CLng((Timer - tMark) * 1000)
        If filesPrepareMs < 0 Then filesPrepareMs = 0

        Call Debug_Registar(passo, prompt.Id, "INFO", "", "REQ_INPUT_JSON", _
            "len=" & Len(inputJsonFinal) & _
            " | has_input_file=" & IIf(InStr(1, inputJsonFinal, """type"":""input_file""", vbTextCompare) > 0, "SIM", "NAO") & _
            " | has_input_image=" & IIf(InStr(1, inputJsonFinal, """type"":""input_image""", vbTextCompare) > 0, "SIM", "NAO") & _
            " | has_text_embed=" & IIf(InStr(1, inputJsonFinal, "----- BEGIN FILE:", vbTextCompare) > 0, "SIM", "NAO") & _
            " | preview=" & Left$(inputJsonFinal, 350), _
            "Input final construído; este retrato confirma se anexos seguiram como file/image e/ou text_embed. Se esperava text_embed, validar has_text_embed=SIM e blocos BEGIN/END FILE no payload dump.")

        Call Painel_LogStepStage(passo, prompt.Id, "before_api", "lenInputJsonFinal=" & CStr(Len(inputJsonFinal)))
        Call Painel_StatusBar_SetPhase(inicioHHMM, "request_ready", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "Pedido pronto")

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

        Dim ciModoAtivo As Boolean
        Dim cfgExtraL As String
        Dim fileOutputIntent As Boolean
        ciModoAtivo = (InStr(1, modosEfetivo, "Code Interpreter", vbTextCompare) > 0)
        cfgExtraL = LCase$(prompt.ConfigExtra)
        fileOutputIntent = (LCase$(Trim$(fo_outputKind)) = "file" Or _
                            InStr(1, cfgExtraL, "output_kind", vbTextCompare) > 0 Or _
                            InStr(1, cfgExtraL, "process_mode", vbTextCompare) > 0)

        If ciModoAtivo And fileOutputIntent And (LCase$(Trim$(fo_outputKind)) <> "file" Or LCase$(Trim$(fo_processMode)) <> "code_interpreter") Then
            Call Debug_Registar(passo, prompt.Id, "ALERTA", "", "M07_FILEOUTPUT_MODE_MISMATCH", _
                "PROBLEMA=Code Interpreter ativo, mas modo efetivo=" & LCase$(Trim$(fo_outputKind)) & "/" & LCase$(Trim$(fo_processMode)) & _
                " | IMPACTO=contrato de output pode cair para texto/metadata | ACAO=validar output_kind:file + process_mode:code_interpreter no catalogo/Config extra | DETALHE=linhas invalidas sem chave:valor (ex.: True) impedem aplicacao.", _
                "Conferir M05_PAYLOAD_CHECK (mode=...) e corrigir Config extra com uma linha por chave: valor.")
        End If

        If fileOutputIntent And LCase$(Trim$(fo_outputKind)) = "text" And LCase$(Trim$(fo_processMode)) = "metadata" Then
            Call Debug_Registar(passo, prompt.Id, "ALERTA", "", "M07_FILEOUTPUT_PARSE_GUARD", _
                "PROBLEMA=Config extra/output intent para File Output, mas efetivo caiu em text/metadata | IMPACTO=pos-processamento M10 pode nao gerar ficheiro | ACAO=normalizar sintaxe do Config extra | DETALHE=usar output_kind: file e process_mode: code_interpreter em linhas parseaveis.", _
                "Regra: uma linha = chave: valor; linhas sem ':' sao ignoradas com alerta.")
        End If

        Call Painel_LogStepStage(passo, prompt.Id, "api_call_start", "model=" & modeloUsado)
        Call Painel_StatusBar_SetPhase(inicioHHMM, "api_call", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "A executar prompt")
        DoEvents

        Dim debugFingerprintSeed As String
        debugFingerprintSeed = "pipeline=" & pipelineNome & "|step=" & CStr(passo) & "|prompt=" & prompt.Id & "|mode=" & LCase$(Trim$(fo_outputKind)) & "/" & LCase$(Trim$(fo_processMode))

        Dim ciIntentResolved As Boolean
        ciIntentResolved = (LCase$(Trim$(fo_outputKind)) = "file" And LCase$(Trim$(fo_processMode)) = "code_interpreter")

        tMark = Timer
        resultado = OpenAI_Executar(apiKey, modeloUsado, promptTextFinal, temperaturaDefault, maxTokensDefault, _
                                    modosEfetivo, prompt.storage, inputJsonFinal, extraFragmentFO, prompt.Id, debugFingerprintSeed, ciIntentResolved)
        apiCallMs = CLng((Timer - tMark) * 1000)
        If apiCallMs < 0 Then apiCallMs = 0

        retryCountTotal = retryCountTotal + resultado.retryCount
        Call Painel_StatusBar_SetPhase(inicioHHMM, "response", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "Resposta recebida")
        DoEvents

                ' -------------------------------
        ' FILE OUTPUT (pos-response): guardar raw + ficheiros (metadata / code_interpreter)
        ' -------------------------------
        fo_filesUsedOut = "": fo_filesOpsOut = "": fo_logSeguimento = ""
        fo_logSeguimento = FileOutput_ProcessAfterResponse(apiKey, outputFolderBase, pipelineNome, pipelineIndex, passo, prompt.Id, resultado, _
            fo_outputKind, fo_processMode, fo_autoSave, fo_overwriteMode, fo_prefixTmpl, fo_subfolderTmpl, _
            fo_pptxMode, fo_xlsxMode, fo_pdfMode, fo_imageMode, fo_filesUsedOut, fo_filesOpsOut)

        ' -------------------------------
        ' CONTRATO DIAGNOSTICO (M19) - tri-state por passo
        ' -------------------------------
        Dim ctHasContract As Boolean, ctMode As String, ctState As String, ctRule As String
        Dim ctProblem As String, ctSuggestion As String, ctDetail As String
        Call ContractDiag_EvaluateStep(passo, prompt.Id, prompt.ConfigExtra, resultado.outputText, resultado.rawResponseJson, _
            ctHasContract, ctMode, ctState, ctRule, ctProblem, ctSuggestion, ctDetail)
        If ctHasContract Then
            Call Painel_StatusBar_SetPhase(inicioHHMM, "contract_gate", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "Validar contrato")
        End If

        Dim fo_executeOpsLog As String
        fo_executeOpsLog = ""
        Dim fo_executeM10Signals As String
        fo_executeM10Signals = "filesUsed=" & fo_filesUsedOut & " | filesOps=" & fo_filesOpsOut
        tMark = Timer
        If Trim$(resultado.Erro) = "" And resultado.httpStatus >= 200 And resultado.httpStatus < 300 Then
            If ctHasContract And (UCase$(ctState) = "BLOCKED" Or UCase$(ctState) = "FAIL") Then
                Call Debug_Registar(passo, prompt.Id, "ERRO", "", "CONTRACT_GATE_BLOCK", _
                    "Gate bloqueou execução de OUTPUT EXECUTE. [Estado=" & ctState & "] [Regra=" & ctRule & "]", _
                    ctSuggestion)
                resultado.Erro = "Contrato do passo bloqueou continuidade: estado=" & ctState & " regra=" & ctRule
            Else
                If hasOutputIntent Then
                    Call Painel_StatusBar_SetPhase(inicioHHMM, "output_execute", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "Executar output")
                End If
                fo_executeOpsLog = OutputOrders_TryExecute(passo, prompt.Id, resultado.responseId, resultado.outputText, outputFolderBase, fo_filesOpsOut, fo_executeM10Signals)
                If Trim$(fo_executeOpsLog) <> "" Then
                    If Trim$(fo_filesOpsOut) <> "" Then
                        fo_filesOpsOut = fo_filesOpsOut & " | " & fo_executeOpsLog
                    Else
                        fo_filesOpsOut = fo_executeOpsLog
                    End If
                End If
            End If
        End If
        directiveParseMs = CLng((Timer - tMark) * 1000)
        If directiveParseMs < 0 Then directiveParseMs = 0

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

        stepEndAt = Now
        Call DebugDiag_RecordStep(pipelineNome, pipelineIndex, passo, prompt.Id, modeloUsado, fo_outputKind, fo_processMode, _
            stepStartAt, stepEndAt, filesPrepareMs, apiCallMs, directiveParseMs, inputJsonFinal, filesUsedResumo, filesOpsResumo, fileIds, linhaFilesLista, _
            resultado, promptTextFinal, outputFolderBase, prompt.ConfigExtra)

        Call Seguimento_Registar(passo, prompt, modeloUsado, auditJson, resultado.httpStatus, resultado.responseId, _
            textoSeguimento, pipelineNome, "", filesUsedResumo, filesOpsResumo, fileIds)
        Call Painel_GitLog_RegisterStepExecution(pipelineNome, prompt.Id, resultado, textoSeguimento, runToken)
        Call Painel_LogStepStage(passo, prompt.Id, "step_completed", "http=" & CStr(resultado.httpStatus) & " | response_id=" & Left$(Trim$(resultado.responseId), 24))

        ' ================================
        ' CONTEXTKV - REGISTAR + CAPTURAR
        ' ================================
        On Error Resume Next
        Call Painel_StatusBar_SetPhase(inicioHHMM, "context_capture", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "Capturar contexto")
        Call ContextKV_WriteInjectedVars(pipelineNome, passo, prompt.Id, injectedVarsJson, outputFolderBase, runToken)
        Call ContextKV_CaptureRow(pipelineNome, passo, prompt.Id, outputFolderBase, runToken)
        If Err.Number <> 0 Then
            Call Debug_Registar(passo, prompt.Id, "ALERTA", "", "CONTEXT_KV", _
                "Erro em WriteInjectedVars/CaptureRow: " & Err.Description, "")
            Err.Clear
        End If
        On Error GoTo TrataErro

        Call Painel_StatusBar_SetPhase(inicioHHMM, "completed", retryCountTotal, rowPos, rowTotal, prompt.Id, promptTemFiles, hasDiagContract, hasOutputIntent, "Passo concluido")

        If Trim$(resultado.Erro) <> "" Then
            Call Debug_Registar(passo, atual, "ERRO", "", "API", _
                resultado.Erro, _
                "Sugestao: verifique modelo, quota, payload e configuracao.")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
            GoTo SaidaLimpa
        End If

        prevResponseId = resultado.responseId

        ' Ler Next config
        Dim nextPrompt As String, nextDefault As String, nextAllowed As String
        Call Catalogo_LerNextConfig(atual, nextPrompt, nextDefault, nextAllowed)
        Call Painel_ValidarConsistenciaNextConfig(nextPrompt, nextDefault, nextAllowed, passo, atual)

        ' Resolver proximo esperado com output (AUTO tenta extrair; senao default)
        Dim proximoEsperado As String
        proximoEsperado = Painel_ResolverNextComOutput(nextPrompt, nextDefault, resultado.outputText)
        If UCase$(Trim$(nextPrompt)) = "AUTO" Then
            Dim autoExtracted As String
            autoExtracted = Painel_ExtrairNextPromptIdDoOutput(resultado.outputText)
            If Trim$(autoExtracted) = "" Then
                Call Debug_Registar(passo, atual, "ALERTA", "", "NEXT_PROMPT_ID", _
                    "Next PROMPT=AUTO sem NEXT_PROMPT_ID no output; aplicado fallback para default/STOP.", _
                    "Sugestao: devolver linha isolada NEXT_PROMPT_ID: <ID|STOP> e rever instrucoes de output do prompt.")
            End If
        End If
        If proximoEsperado = "" Then proximoEsperado = nextDefault

        If proximoEsperado = "" Or Painel_EhSTOP(proximoEsperado) Then
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
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
                    Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
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
                Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
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
            Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
            GoTo SaidaLimpa
        End If

        ' Detecao ABAB
        Call Painel_AtualizarUltimos4(ultimos4, atual)
        If Painel_DetetouAlternanciaABAB(ultimos4) Then
            Call Debug_Registar(passo, atual, "ALERTA", "", "Ciclos", _
                "Detetada alternancia A-B-A-B. Pipeline interrompida.", _
                "Sugestao: adicione condicao de saida, restrinja allowed, ou introduza STOP.")
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
            GoTo SaidaLimpa
        End If

        Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)

        atual = proximoFinal
        cursorRow = nextCursorRow

        If Painel_EhSTOP(atual) Then
            wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"
            Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
            GoTo SaidaLimpa
        End If

        If cursorRow > LIST_START_ROW + LIST_MAX_ROWS - 2 Then
            Call Debug_Registar(passo, atual, "ALERTA", "", "LimiteLista", _
                "A lista INICIAR excedeu o limite de linhas do PAINEL.", _
                "Sugestao: aumente LIST_MAX_ROWS no codigo ou reduza a pipeline.")
            wsPainel.Cells(LIST_START_ROW + LIST_MAX_ROWS - 1, colIniciar).value = "STOP"
            Call Painel_EspelharDebugNoCatalogo(passo, prompt.Id)
            GoTo SaidaLimpa
        End If

    Next passo

    Call Debug_Registar(maxSteps, startId, "ALERTA", "", "MaxSteps", _
        "Max Steps atingido. Pipeline terminou por limite.", _
        "Sugestao: aumente Max Steps no PAINEL ou defina STOP.")
    wsPainel.Cells(cursorRow + 1, colIniciar).value = "STOP"

SaidaLimpa:
    If runExecutouPassos Then
        Call PipelineGitDebug_ExportIfEnabled(pipelineIndex, pipelineNome, painelAutoSave)
    End If

    Call M05_ClearRunDumpFolder

    Application.StatusBar = False
    Application.DisplayStatusBar = oldDisplayStatusBar
    Application.EnableEvents = oldEnableEvents
    Exit Sub

TrataErro:
    Dim errResumo As String
    errResumo = "Erro inesperado: #" & CStr(Err.Number) & " | " & Err.Description & " | stage=" & mStepLastStage

    Call Debug_Registar(0, "PIPELINE_" & CStr(pipelineIndex), "ERRO", "", "VBA", _
        errResumo, _
        "Sugestao: verifique IDs, folhas e referencias. Compile o VBAProject.")

    If passoCtx > 0 Then
        Call Painel_RegistarFalhaNoSeguimento(passoCtx, promptCtx, modeloUsado, pipelineNome, errResumo)
    End If

    Resume SaidaLimpa
End Sub


Private Sub Painel_GitLog_RegisterStepExecution(ByVal pipelineNome As String, ByVal promptId As String, ByRef resultado As ApiResultado, ByVal textoSeguimento As String, ByVal runId As String)
    On Error GoTo Falha

    Dim versao As String
    versao = Painel_ExtrairVersaoDoPromptId(promptId)

    Dim successRaw As String
    successRaw = Painel_GitLog_ExtractField(resultado.outputText, "Success")
    If Trim$(successRaw) = "" Then successRaw = Painel_GitLog_ExtractField(resultado.outputText, "Sucesso")

    Dim successNorm As String
    successNorm = Painel_GitLog_NormalizeSuccess(successRaw)
    If Trim$(successNorm) = "Outro" Then
        If Trim$(resultado.Erro) <> "" Or resultado.httpStatus < 200 Or resultado.httpStatus >= 300 Then
            successNorm = "Não"
        End If
    End If

    Dim newVersionRaw As String
    newVersionRaw = Painel_GitLog_ExtractField(resultado.outputText, "New version")
    If Trim$(newVersionRaw) = "" Then newVersionRaw = Painel_GitLog_ExtractField(resultado.outputText, "Nova versao")

    Dim newVersionNorm As String
    newVersionNorm = Painel_GitLog_NormalizeYesNo(newVersionRaw)

    Dim analysisLink As String
    analysisLink = Painel_GitLog_ExtractField(resultado.outputText, "Analysis Link")
    If Trim$(analysisLink) = "" Then analysisLink = Painel_GitLog_ExtractField(resultado.outputText, "Analysis")

    Dim newPromptLink As String
    newPromptLink = Painel_GitLog_ExtractField(resultado.outputText, "New Prompt Link")
    If Trim$(newPromptLink) = "" Then newPromptLink = Painel_GitLog_ExtractField(resultado.outputText, "Prompt Link")

    Dim summaryRaw As String
    summaryRaw = Painel_GitLog_ExtractField(resultado.outputText, "Summary")
    If Trim$(summaryRaw) = "" Then summaryRaw = textoSeguimento

    Dim summaryFinal As String
    summaryFinal = Painel_GitLog_NormalizeSummary(summaryRaw, 4)

    Dim status As String
    status = "Success=" & successNorm & " | New version=" & newVersionNorm

    Application.Run "GitLog_RegisterPromptExecution", pipelineNome, promptId, versao, status, analysisLink, newPromptLink, summaryFinal, runId
    Exit Sub

Falha:
    Call Debug_Registar(0, promptId, "ALERTA", "", "GIT_LOG", _
        "Nao foi possivel chamar GitLog_RegisterPromptExecution: " & Err.Description, _
        "Sugestao: confirme se a macro GitLog_RegisterPromptExecution existe e aceita 8 argumentos.")
End Sub

Private Function Painel_ExtrairVersaoDoPromptId(ByVal promptId As String) As String
    Dim arr() As String
    arr = Split(Trim$(promptId), "/")
    If UBound(arr) >= 3 Then
        Painel_ExtrairVersaoDoPromptId = Trim$(arr(UBound(arr)))
    Else
        Painel_ExtrairVersaoDoPromptId = ""
    End If
End Function

Private Function Painel_GitLog_NormalizeSuccess(ByVal rawValue As String) As String
    Dim s As String
    s = UCase$(Trim$(rawValue))

    Select Case s
        Case "SIM", "YES", "TRUE", "Y", "OK", "SUCESSO"
            Painel_GitLog_NormalizeSuccess = "Sim"
        Case "NAO", "NÃO", "NO", "FALSE", "N", "FAIL", "FALHA"
            Painel_GitLog_NormalizeSuccess = "Não"
        Case "CONDICIONADO", "CONDITIONAL", "PARTIAL", "PARCIAL", "DEPENDE"
            Painel_GitLog_NormalizeSuccess = "Condicionado"
        Case Else
            Painel_GitLog_NormalizeSuccess = "Outro"
    End Select
End Function

Private Function Painel_GitLog_NormalizeYesNo(ByVal rawValue As String) As String
    Dim s As String
    s = UCase$(Trim$(rawValue))

    Select Case s
        Case "SIM", "YES", "TRUE", "Y", "1"
            Painel_GitLog_NormalizeYesNo = "Sim"
        Case Else
            Painel_GitLog_NormalizeYesNo = "Não"
    End Select
End Function

Private Function Painel_GitLog_ExtractField(ByVal outputText As String, ByVal label As String) As String
    Dim txt As String
    txt = Replace$(Replace$(CStr(outputText), vbCrLf, vbLf), vbCr, vbLf)

    Dim linhas() As String
    linhas = Split(txt, vbLf)

    Dim i As Long
    For i = LBound(linhas) To UBound(linhas)
        Dim ln As String
        ln = Trim$(linhas(i))

        If UCase$(Left$(ln, Len(label) + 1)) = UCase$(label & ":") Then
            Painel_GitLog_ExtractField = Trim$(Mid$(ln, Len(label) + 2))
            Exit Function
        End If
    Next i
End Function

Private Function Painel_GitLog_NormalizeSummary(ByVal summaryText As String, ByVal maxLines As Long) As String
    Dim txt As String
    txt = Replace$(Replace$(CStr(summaryText), vbCrLf, vbLf), vbCr, vbLf)

    Dim linhas() As String
    linhas = Split(txt, vbLf)

    Dim resultado As String
    Dim i As Long
    Dim n As Long
    Dim lastBlank As Boolean
    resultado = ""
    n = 0
    lastBlank = True

    For i = LBound(linhas) To UBound(linhas)
        Dim ln As String
        ln = Trim$(linhas(i))

        If ln = "" Then
            If lastBlank Then GoTo NextLine
            lastBlank = True
        Else
            lastBlank = False
        End If

        n = n + 1
        If n > maxLines Then Exit For

        If resultado = "" Then
            resultado = ln
        Else
            resultado = resultado & vbLf & ln
        End If

NextLine:
    Next i

    Painel_GitLog_NormalizeSummary = resultado
End Function


' ============================================================
' 4.x) Ajudas pedidas: foco/DEBUG/status bar/checks de FILES
' ============================================================

Private Sub Painel_FocarDebugA1()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEBUG)

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

Private Sub Painel_StatusBar_SetPhase(ByVal inicioHHMM As String, ByVal phaseCode As String, ByVal retryCount As Long, ByVal rowPos As Long, ByVal rowTotal As Long, ByVal promptId As String, ByVal promptTemFiles As Boolean, ByVal hasDiagContract As Boolean, ByVal hasOutputIntent As Boolean, ByVal detalhe As String)
    Dim phasePlan As Collection
    Set phasePlan = Painel_BuildInternalPhasePlan(promptTemFiles, hasDiagContract, hasOutputIntent)

    Dim idx As Long
    idx = Painel_PhasePlanIndex(phasePlan, phaseCode)
    If idx <= 0 Then idx = 1

    Call Painel_StatusBar_Set(inicioHHMM, idx, phasePlan.Count, retryCount, detalhe, rowPos, rowTotal, promptId)
End Sub

Private Function Painel_BuildInternalPhasePlan(ByVal promptTemFiles As Boolean, ByVal hasDiagContract As Boolean, ByVal hasOutputIntent As Boolean) As Collection
    Dim plan As Collection
    Set plan = New Collection

    plan.Add "prepare"
    plan.Add "context"
    If promptTemFiles Then plan.Add "files_prepare"
    plan.Add "request_ready"
    plan.Add "api_call"
    plan.Add "response"
    If hasDiagContract Then plan.Add "contract_gate"
    If hasOutputIntent Then plan.Add "output_execute"
    plan.Add "context_capture"
    plan.Add "completed"

    Set Painel_BuildInternalPhasePlan = plan
End Function

Private Function Painel_PhasePlanIndex(ByVal plan As Collection, ByVal phaseCode As String) As Long
    Dim i As Long
    For i = 1 To plan.Count
        If StrComp(CStr(plan(i)), phaseCode, vbTextCompare) = 0 Then
            Painel_PhasePlanIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function Painel_HasFileOutputIntent(ByVal promptText As String, ByVal configExtraText As String, ByVal modos As String) As Boolean
    Dim cfgLower As String
    cfgLower = LCase$(Trim$(configExtraText))

    Painel_HasFileOutputIntent = (InStr(1, cfgLower, "output_kind", vbTextCompare) > 0 Or _
                                  InStr(1, cfgLower, "process_mode", vbTextCompare) > 0 Or _
                                  InStr(1, LCase$(modos), "code interpreter", vbTextCompare) > 0 Or _
                                  InStr(1, LCase$(promptText), "ci_output_file:", vbTextCompare) > 0)
End Function

Private Sub Painel_StatusBar_Set(ByVal inicioHHMM As String, ByVal internalStep As Long, ByVal internalTotal As Long, ByVal retryCount As Long, Optional ByVal detalhe As String = "", Optional ByVal rowPos As Long = 0, Optional ByVal rowTotal As Long = 0, Optional ByVal promptId As String = "")
    On Error Resume Next

    Dim passoTxt As String
    If internalTotal > 10 Then
        passoTxt = Format$(internalStep, "00")
    Else
        passoTxt = CStr(internalStep)
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
        rowLabel = "  |  Row " & CStr(rowPos) & " of " & CStr(rowTotal) & " (pipeline)"
    End If

    Application.StatusBar = "(" & inicioHHMM & ") Step: " & passoTxt & " of " & CStr(internalTotal) & "  |  Retry: " & CStr(retryCount) & _
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
    ' D1) Validador bloqueante de payload para anexos declarados em FILES
    ' ------------------------------------------------------------
    Dim expectedNames As Collection
    Set expectedNames = Painel_Files_ExpectedNamesFromLista(listaFiles)

    If expectedNames.Count > 0 Then
        Dim missingNames As Collection
        Set missingNames = Painel_Files_FindMissingExpected(expectedNames, filesUsed, inputFolder)

        If missingNames.Count > 0 Then
            Dim gotInputFile As Long
            gotInputFile = Painel_CountOccurrences(inputJsonFinal, """type"":""input_file""")

            Dim gotInputImage As Long
            gotInputImage = Painel_CountOccurrences(inputJsonFinal, """type"":""input_image""")

            Dim gotTextEmbed As Long
            gotTextEmbed = Painel_CountOccurrences(inputJsonFinal, "----- BEGIN FILE:")

            Call Debug_Registar(passo, promptId, "ERRO", "", "INPUTFILES_MISSING", _
                "expected=" & CStr(expectedNames.Count) & _
                " got_input_file=" & CStr(gotInputFile) & _
                " got_input_image=" & CStr(gotInputImage) & _
                " got_text_embed=" & CStr(gotTextEmbed) & _
                " missing=[" & Painel_JoinCollection(missingNames, ", ") & "]", _
                "Bloqueado antes do envio: rever INPUTS: FILES e anexacao M09 para garantir contrato minimo do payload.")
            Painel_Files_Checks_Debug = False
            Exit Function
        End If
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
    End If

    If temTextEmbed Then
        Dim countTextEmbed As Long
        countTextEmbed = Painel_CountOccurrences(inputJsonFinal, "----- BEGIN FILE:")

        Call Debug_Registar(passo, promptId, "INFO", "", "FILES", _
            "Anexacao OK via text_embed. blocos_text_embed=" & CStr(countTextEmbed), _
            "Nota: text_embed coexiste com input_file/input_image quando ha anexos mistos; neste modo nao existe file_id para os ficheiros textuais.")
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

Private Sub Painel_ValidarConsistenciaNextConfig(ByVal nextPrompt As String, ByVal nextDefault As String, ByVal nextAllowed As String, ByVal passo As Long, ByVal promptAtual As String)
    Dim d As String
    d = Trim$(nextDefault)
    If d = "" Then
        Call Debug_Registar(passo, promptAtual, "ALERTA", "", "Next PROMPT default", _
            "Next PROMPT default vazio; o fallback pode terminar em STOP inesperado.", _
            "Sugestao: preencher Next PROMPT default com ID valido ou STOP explicito.")
        Exit Sub
    End If

    If UCase$(d) <> "STOP" Then
        If Not Painel_AllowedContem(nextAllowed, d) Then
            Call Debug_Registar(passo, promptAtual, "ALERTA", "", "Next allowed", _
                "Inconsistencia de configuracao: default nao pertence a allowed.", _
                "Sugestao: incluir o default em Next PROMPT allowed para manter fallback deterministico.")
        End If
    End If
End Sub

Private Function Painel_HasCsvExecuteIntent(ByVal promptText As String, ByVal configExtraText As String) As Boolean
    Dim allTxt As String
    allTxt = UCase$(Painel_Nz(promptText) & vbLf & Painel_Nz(configExtraText))

    Dim hasExecute As Boolean
    hasExecute = (InStr(1, allTxt, "LOAD_CSV", vbTextCompare) > 0) Or _
                 (InStr(1, allTxt, "EXECUTE:", vbTextCompare) > 0) Or _
                 (InStr(1, allTxt, "ACTION=LOAD_CSV", vbTextCompare) > 0)

    Dim hasCsvProofContract As Boolean
    hasCsvProofContract = (InStr(1, allTxt, "EXPORT_OK_CSV", vbTextCompare) > 0) Or _
                          (InStr(1, allTxt, "CSV_EXISTE_EM_MNT_DATA", vbTextCompare) > 0) Or _
                          (InStr(1, allTxt, "FILE_CSV:", vbTextCompare) > 0)

    Painel_HasCsvExecuteIntent = hasExecute Or hasCsvProofContract
End Function

Private Function Painel_HasDiagnosticContractConfigured(ByVal configExtraText As String) As Boolean
    Dim v As String
    v = Painel_ConfigExtraGetValue(configExtraText, "diagnostic_contract")
    If Trim$(v) = "" Then v = Painel_ConfigExtraGetValue(configExtraText, "contract_mode")
    If Trim$(v) = "" Then v = Painel_ConfigExtraGetValue(configExtraText, "diagnostic-contract")
    If Trim$(v) = "" Then v = Painel_ConfigExtraGetValue(configExtraText, "diagnostic contract")
    If Trim$(v) = "" Then v = Painel_ConfigExtraGetValue(configExtraText, "diagnostic_contract_mode")

    Painel_HasDiagnosticContractConfigured = (Trim$(v) <> "")
End Function

Private Function Painel_ConfigExtraGetValue(ByVal configExtraText As String, ByVal keyName As String) As String
    Dim txt As String
    txt = Replace(configExtraText, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Dim lines() As String
    lines = Split(txt, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(CStr(lines(i)))
        If lineText <> "" Then
            Dim p As Long
            p = InStr(1, lineText, ":", vbTextCompare)
            If p > 0 Then
                Dim k As String
                k = Trim$(Left$(lineText, p - 1))
                If StrComp(k, keyName, vbTextCompare) = 0 Then
                    Painel_ConfigExtraGetValue = Trim$(Mid$(lineText, p + 1))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Sub Painel_LogStepStage(ByVal passo As Long, ByVal promptId As String, ByVal stageName As String, ByVal detail As String)
    mStepLastStage = CStr(stageName)
    On Error Resume Next
    Call Debug_Registar(passo, promptId, "INFO", "", "STEP_STAGE", _
        "stage=" & stageName & IIf(Trim$(detail) <> "", " | " & detail, ""), _
        "Use o ultimo stage para localizar o ponto em que a execucao parou antes do Seguimento.")
    On Error GoTo 0
End Sub

Private Sub Painel_RegistarFalhaNoSeguimento(ByVal passo As Long, ByVal promptId As String, ByVal modelo As String, ByVal pipelineNome As String, ByVal erroResumo As String)
    On Error Resume Next
    Dim p As PromptDefinicao
    p.Id = promptId
    p.textoPrompt = ""
    p.ConfigExtra = ""
    Call Seguimento_Registar(passo, p, modelo, "", 0, "", "[ERRO VBA] " & erroResumo, pipelineNome, "", "", "", "")
    On Error GoTo 0
End Sub

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


Private Sub Painel_EspelharDebugNoCatalogo(ByVal passo As Long, ByVal promptId As String)
    On Error GoTo TrataErro

    Dim pid As String
    pid = Trim$(promptId)
    If pid = "" Then
        Call Debug_Registar(passo, "[SEM_PROMPT_ID]", "ALERTA", "", "DEBUG_SNAPSHOT", _
            "Nao foi possivel gravar o espelho do DEBUG no catalogo porque o Prompt ID esta vazio.", _
            "Sugestao: confirme se o catalogo tem IDs validos na coluna A.")
        Exit Sub
    End If

    Dim wsCatalogo As Worksheet
    Dim wsDebug As Worksheet
    Set wsCatalogo = Painel_ObterFolhaCatalogoPorPromptId(pid)
    If wsCatalogo Is Nothing Then
        Call Debug_Registar(passo, pid, "ALERTA", "", "DEBUG_SNAPSHOT", _
            "Nao foi possivel localizar a folha de catalogo para guardar o DEBUG desta prompt.", _
            "Sugestao: confirme se o prefixo do ID corresponde exatamente ao nome da folha.")
        Exit Sub
    End If

    Set wsDebug = ThisWorkbook.Worksheets(SHEET_DEBUG)

    Dim rowId As Long
    rowId = Painel_LocalizarLinhaPromptNoCatalogo(wsCatalogo, pid)
    If rowId = 0 Then
        Call Debug_Registar(passo, pid, "ALERTA", "", "DEBUG_SNAPSHOT", _
            "A prompt foi executada, mas nao foi encontrada no catalogo para escrever o espelho do DEBUG.", _
            "Sugestao: valide se o ID existe na coluna A da folha de catalogo.")
        Exit Sub
    End If

    Dim headerCell As Range
    Dim bodyCell As Range
    Dim bodyOverflowCell As Range
    Set headerCell = wsCatalogo.Cells(rowId + 1, 10)
    Set bodyCell = wsCatalogo.Cells(rowId + 2, 10)
    Set bodyOverflowCell = wsCatalogo.Cells(rowId + 3, 10)

    Dim snapshotTsv As String
    snapshotTsv = Painel_DebugSheetToTsv(wsDebug, passo, pid)
    If Trim$(snapshotTsv) = "" Then
        snapshotTsv = "[Sem linhas no DEBUG para a prompt " & pid & "]"
        Call Debug_Registar(passo, pid, "ALERTA", "", "DEBUG_SNAPSHOT", _
            "Nao foram encontradas linhas no DEBUG para a prompt executada; foi gravado um marcador informativo no catalogo.", _
            "Sugestao: confirme se o Prompt ID do DEBUG coincide com o ID do catalogo.")
    End If

    Dim bodyLinha3 As String
    Dim bodyLinha4 As String
    Dim truncado As Boolean
    Call Painel_SplitSnapshotEmDuasCelulas(snapshotTsv, 32767, bodyLinha3, bodyLinha4, truncado)

    headerCell.Value = "DEBUG [" & Format$(Now, "dd-mm-yyyy hh:mm") & "]"
    headerCell.Font.Bold = True
    headerCell.WrapText = False

    bodyCell.Value = bodyLinha3
    bodyCell.Font.Bold = False
    bodyCell.WrapText = False

    bodyOverflowCell.Value = bodyLinha4
    bodyOverflowCell.Font.Bold = False
    bodyOverflowCell.WrapText = False

    Dim salmonLight As Long
    salmonLight = RGB(255, 204, 153)
    headerCell.Interior.Color = salmonLight
    bodyCell.Interior.Color = salmonLight
    bodyOverflowCell.Interior.Color = salmonLight

    Call Debug_Registar(passo, pid, "INFO", "", "DEBUG_SNAPSHOT", _
        "DEBUG consolidado no catalogo com sucesso para consulta rapida.", _
        "Foi atualizado o bloco de Notas para desenvolvimento (linhas 3 e 4 da prompt), substituindo conteudo anterior.")

    If truncado Then
        Call Debug_Registar(passo, pid, "ALERTA", "", "DEBUG_SNAPSHOT", _
            "O espelho do DEBUG excedeu o limite de 2 celulas (linhas 3 e 4) e foi truncado no catalogo.", _
            "Sugestao: reduzir verbosidade do DEBUG ou consultar a folha DEBUG completa para detalhes adicionais.")
    End If
    Exit Sub

TrataErro:
    Call Debug_Registar(passo, pid, "ERRO", "", "DEBUG_SNAPSHOT", _
        "Falha ao gravar o espelho do DEBUG no catalogo: " & Err.Description, _
        "Sugestao: verifique se a folha de catalogo nao esta protegida e se a coluna J esta editavel.")
End Sub

Private Function Painel_ObterFolhaCatalogoPorPromptId(ByVal promptId As String) As Worksheet
    Dim prefixo As String
    Dim posSep As Long
    posSep = InStr(1, Trim$(promptId), "/")

    If posSep > 1 Then
        prefixo = Left$(Trim$(promptId), posSep - 1)
    Else
        prefixo = Trim$(promptId)
    End If

    If prefixo = "" Then Exit Function

    On Error Resume Next
    Set Painel_ObterFolhaCatalogoPorPromptId = ThisWorkbook.Worksheets(prefixo)
    On Error GoTo 0
End Function

Private Function Painel_LocalizarLinhaPromptNoCatalogo(ByVal wsCatalogo As Worksheet, ByVal promptId As String) As Long
    Dim lastRow As Long
    lastRow = wsCatalogo.Cells(wsCatalogo.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim i As Long
    For i = 2 To lastRow
        If Trim$(CStr(wsCatalogo.Cells(i, 1).Value)) = Trim$(promptId) Then
            Painel_LocalizarLinhaPromptNoCatalogo = i
            Exit Function
        End If
    Next i
End Function

Private Function Painel_DebugSheetToTsv(ByVal wsDebug As Worksheet, ByVal passo As Long, ByVal promptId As String) As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim colPrompt As Long
    Dim colPasso As Long
    Dim r As Long
    Dim c As Long
    Dim lineTxt As String
    Dim acc As String
    Dim pidKey As String
    Dim rowPidKey As String
    Dim rowPasso As String

    lastRow = wsDebug.Cells(wsDebug.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Function

    lastCol = wsDebug.Cells(1, wsDebug.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function

    colPrompt = Painel_FindHeaderColumn(wsDebug, "Prompt ID")
    colPasso = Painel_FindHeaderColumn(wsDebug, "Passo")
    pidKey = Painel_NormalizarPromptIdKey(promptId)

    For r = 1 To lastRow
        If r > 1 Then
            If colPrompt > 0 Then
                rowPidKey = Painel_NormalizarPromptIdKey(CStr(wsDebug.Cells(r, colPrompt).Value))
            Else
                rowPidKey = ""
            End If

            If colPasso > 0 Then
                rowPasso = Trim$(CStr(wsDebug.Cells(r, colPasso).Value))
            Else
                rowPasso = ""
            End If

            If (pidKey <> "" And rowPidKey = pidKey) Then
                ' inclui por Prompt ID
            ElseIf (passo > 0 And rowPasso = CStr(passo)) Then
                ' fallback: inclui por Passo para capturar linhas sem Prompt ID
            Else
                GoTo NextRow
            End If
        End If

        lineTxt = CStr(r)
        For c = 1 To lastCol
            lineTxt = lineTxt & vbTab
            lineTxt = lineTxt & Painel_DebugSanitizarCampoTsv(CStr(wsDebug.Cells(r, c).Value))
        Next c

        If Len(acc) > 0 Then acc = acc & vbLf
        acc = acc & lineTxt
NextRow:
    Next r

    Painel_DebugSheetToTsv = acc
End Function

Private Function Painel_DebugSanitizarCampoTsv(ByVal valor As String) As String
    Dim t As String
    t = CStr(valor)
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCrLf, " | ")
    t = Replace(t, vbCr, " | ")
    t = Replace(t, vbLf, " | ")
    Painel_DebugSanitizarCampoTsv = t
End Function

Private Function Painel_FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function

    Dim alvo As String
    alvo = Painel_NormalizeHeaderToken(headerName)

    Dim c As Long
    Dim atual As String
    For c = 1 To lastCol
        atual = Painel_NormalizeHeaderToken(CStr(ws.Cells(1, c).Value))
        If atual = alvo Then
            Painel_FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function Painel_NormalizeHeaderToken(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, "á", "a")
    t = Replace(t, "à", "a")
    t = Replace(t, "â", "a")
    t = Replace(t, "ã", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "ê", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ô", "o")
    t = Replace(t, "õ", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ç", "c")
    Painel_NormalizeHeaderToken = t
End Function

Private Sub Painel_SplitSnapshotEmDuasCelulas(ByVal snapshotTsv As String, ByVal maxChars As Long, ByRef parte1 As String, ByRef parte2 As String, ByRef truncado As Boolean)
    parte1 = ""
    parte2 = ""
    truncado = False

    If snapshotTsv = "" Then Exit Sub

    parte1 = Left$(snapshotTsv, maxChars)
    If Len(snapshotTsv) > maxChars Then
        parte2 = Mid$(snapshotTsv, maxChars + 1, maxChars)
    End If

    truncado = (Len(snapshotTsv) > (maxChars * 2))
End Sub

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

Private Function Painel_Files_ExpectedNamesFromLista(ByVal listaFiles As String) As Collection
    Set Painel_Files_ExpectedNamesFromLista = New Collection
    On Error GoTo Falha
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim raw As String
    raw = Painel_Nz(listaFiles)
    If Trim$(raw) = "" Then Exit Function
    Dim items() As String
    items = Split(raw, ";")
    Dim i As Long
    For i = LBound(items) To UBound(items)
        Dim nm As String
        nm = Painel_Files_NormalizeExpectedName(CStr(items(i)))
        If nm <> "" Then
            If Not seen.exists(LCase$(nm)) Then
                seen.Add LCase$(nm), True
                Painel_Files_ExpectedNamesFromLista.Add nm
            End If
        End If
    Next i
    Exit Function
Falha:
End Function

Private Function Painel_Files_NormalizeExpectedName(ByVal item As String) As String
    Dim t As String
    t = Trim$(CStr(item))
    If t = "" Then Exit Function
    Dim p As Long
    p = InStr(1, t, "(", vbTextCompare)
    If p > 0 Then t = Trim$(Left$(t, p - 1))
    If InStrRev(t, "\") > 0 Then t = Mid$(t, InStrRev(t, "\") + 1)
    If InStrRev(t, "/") > 0 Then t = Mid$(t, InStrRev(t, "/") + 1)
    Painel_Files_NormalizeExpectedName = Trim$(t)
End Function

Private Function Painel_Nz(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    Painel_Nz = CStr(v)
End Function

Private Function Painel_Files_FindMissingExpected(ByVal expectedNames As Collection, ByVal filesUsed As String, ByVal inputFolder As String) As Collection
    Set Painel_Files_FindMissingExpected = New Collection
    On Error GoTo Falha
    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare
    Dim rawUsed As String
    rawUsed = Painel_Nz(filesUsed)
    If Trim$(rawUsed) <> "" Then
        Dim parts() As String
        parts = Split(rawUsed, ";")
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            Dim token As String
            token = Trim$(CStr(parts(i)))
            If InStr(1, token, "(", vbTextCompare) > 0 Then token = Trim$(Left$(token, InStr(1, token, "(", vbTextCompare) - 1))
            If token <> "" Then used(LCase$(token)) = True
        Next i
    End If

    Dim j As Long
    For j = 1 To expectedNames.Count
        Dim expName As String
        expName = CStr(expectedNames(j))

        Dim resolvedExpected As String
        resolvedExpected = Painel_Files_ResolveExpectedName(expName, inputFolder)

        If InStr(1, expName, "*", vbTextCompare) > 0 Then
            If Not Painel_Files_AnyUsedMatchesWildcard(used, expName) Then
                If resolvedExpected <> "" Then
                    If Not used.exists(LCase$(resolvedExpected)) Then Painel_Files_FindMissingExpected.Add resolvedExpected
                Else
                    Painel_Files_FindMissingExpected.Add expName
                End If
            End If
        Else
            If resolvedExpected <> "" Then expName = resolvedExpected
            If Not used.exists(LCase$(expName)) Then Painel_Files_FindMissingExpected.Add expName
        End If
    Next j
    Exit Function
Falha:
End Function

Private Function Painel_Files_ResolveExpectedName(ByVal expectedToken As String, ByVal inputFolder As String) As String
    On Error GoTo Falha
    Dim token As String
    token = Trim$(expectedToken)
    If token = "" Then Exit Function

    Dim hasLatest As Boolean
    hasLatest = (InStr(1, token, "(latest)", vbTextCompare) > 0) Or _
                (InStr(1, token, "(mais recente)", vbTextCompare) > 0) Or _
                (InStr(1, token, "(mais_recente)", vbTextCompare) > 0)

    token = Painel_Files_NormalizeExpectedName(token)
    If token = "" Then Exit Function

    If InStr(1, token, "*", vbTextCompare) = 0 Then
        Painel_Files_ResolveExpectedName = token
        Exit Function
    End If

    If Trim$(inputFolder) = "" Or Dir$(inputFolder, vbDirectory) = "" Then Exit Function

    Dim bestName As String
    Dim bestDate As Date
    Dim f As String
    f = Dir$(inputFolder & "\" & token)

    Do While f <> ""
        If hasLatest Then
            Dim dt As Date
            On Error Resume Next
            dt = FileDateTime(inputFolder & "\" & f)
            On Error GoTo Falha
            If bestName = "" Or dt > bestDate Then
                bestName = f
                bestDate = dt
            End If
        Else
            If bestName = "" Then bestName = f
        End If
        f = Dir$()
    Loop

    Painel_Files_ResolveExpectedName = bestName
    Exit Function
Falha:
    Painel_Files_ResolveExpectedName = ""
End Function

Private Function Painel_Files_AnyUsedMatchesWildcard(ByVal used As Object, ByVal wildcardPattern As String) As Boolean
    On Error GoTo Falha
    Dim normPattern As String
    normPattern = LCase$(Painel_Files_NormalizeExpectedName(wildcardPattern))
    If normPattern = "" Then Exit Function

    Dim k As Variant
    For Each k In used.Keys
        If LCase$(CStr(k)) Like normPattern Then
            Painel_Files_AnyUsedMatchesWildcard = True
            Exit Function
        End If
    Next k
    Exit Function
Falha:
    Painel_Files_AnyUsedMatchesWildcard = False
End Function

Private Function Painel_CountOccurrences(ByVal hay As String, ByVal needle As String) As Long
    Dim p As Long
    p = InStr(1, hay, needle, vbTextCompare)
    Do While p > 0
        Painel_CountOccurrences = Painel_CountOccurrences + 1
        p = InStr(p + Len(needle), hay, needle, vbTextCompare)
    Loop
End Function

Private Function Painel_JoinCollection(ByVal coll As Collection, ByVal sep As String) As String
    Dim i As Long
    For i = 1 To coll.Count
        If i > 1 Then Painel_JoinCollection = Painel_JoinCollection & sep
        Painel_JoinCollection = Painel_JoinCollection & CStr(coll(i))
    Next i
End Function
