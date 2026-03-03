Attribute VB_Name = "M02_Logger_DEBUG_e_Seguimento"
Option Explicit

' =============================================================================
' Módulo: M02_Logger_DEBUG_e_Seguimento
' Propósito:
' - Registar auditoria operacional nas folhas DEBUG e Seguimento.
' - Manter escrita resiliente a reordenação de colunas e apoiar arquivamento/limpeza de logs.
'
' Atualizações:
' - 2026-03-03 | Codex | Inclui CI_PROOF_MNT_DATA_MISSING no mapeamento operacional
'   - Classifica o novo parametro nas acoes de OUTPUT_EXECUTE/diagnostico CI na coluna Funcionalidade.
'   - Mantem coerencia de troubleshooting para eventos de artefacto em falta no fluxo CI.
' - 2026-03-03 | Codex | Classificacao de trace de text_embed no DEBUG
'   - Classifica parametro TEXT_EMBED_TRACE no bloco de gestao de anexos/text_embed.
'   - Acrescenta acao dedicada para leitura de name/len_chars/hash_short no troubleshooting.
' - 2026-03-03 | Codex | Mapeia trace padronizado de override de modo FILES
'   - Inclui `FILES_MODE_OVERRIDE_TRACE` nas ações deduzidas da coluna Funcionalidade.
'   - Amplia extração de contexto operacional com requested/resolved/raw_mode/effective_mode/reason.
' - 2026-03-03 | Codex | Mapeia eventos de lint do Output Orders
'   - Adiciona cobertura para EXECUTE_LINT_MULTIPLE e EXECUTE_LINT_IN_CODEBLOCK na deducao de acao em curso.
' - 2026-03-03 | Codex | Mapeia novo alerta CI_PROOF_MNT_DATA_MISSING na coluna de acao
'   - Evita descricao generica para diagnostico de ausencia de artefacto CSV com sinais M10.
' - 2026-03-03 | Codex | Expande contexto M10 extraido para CI_PROOF_MNT_DATA_MISSING
'   - Adiciona chaves de falha de download/listagem para manter acao em curso especifica e auditavel.
' - 2026-03-03 | Codex | Cobertura ampliada de acoes especificas no DEBUG
'   - Reforca mapeamento por sinais de parametro/contexto para suportar combinacoes de acoes no mesmo registo.
'   - Amplia extracao de contexto com pares chave=valor e chave:valor para diagnostico mais objetivo.
' - 2026-03-02 | Codex | Enriquecimento da coluna Funcionalidade com acao em curso por evento
'   - Acrescenta segunda linha padrao "ACAO EM CURSO" com detalhe operacional por registo.
'   - Aplica negrito apenas à linha de acao e tenta inferir detalhe especifico (ficheiro/etapa/endpoint).
' - 2026-03-02 | Codex | Completa mapeamento da coluna Funcionalidade para eventos DEBUG sem contexto claro
'   - Cobre codigos de CI/container (M10_* e M05_CI_*), catalogo/fluxo e sinalizacao INFO/ALERTA.
'   - Reduz quedas para descricao generica em eventos de erro/alerta/info com parametro tecnico curto.
' - 2026-03-02 | Codex | Adiciona coluna Funcionalidade no DEBUG com descricao leiga
'   - Garante criacao automatica da coluna entre Parametro e Problema quando em falta.
'   - Passa a preencher Funcionalidade em todas as entradas via mapeamento por parametro.
' - 2026-03-02 | Codex | Torna pausa de render no DEBUG configuravel pela folha Config
'   - Le o parametro DEBUG_RENDER_PAUSE_MS (coluna A/B) com fallback interno para 3 ms.
'   - Mantem compatibilidade retroativa quando a chave nao existe ou e invalida.
' - 2026-03-02 | Codex | Ajusta auto-scroll do DEBUG para manter contexto visual
'   - Adiciona pausa curta (~3 ms) apos cada nova linha para favorecer refresh no ecra.
'   - Mantem a linha mais recente visivel no limite inferior da janela (sem saltar para o topo).
' - 2026-03-02 | Codex | Formatacao visual automatica no DEBUG
'   - Aplica negrito/cor por severidade (ERRO=vermelho, ALERTA=azul).
'   - Destaca STEP_STAGE stage=step_completed em verde e mantem a ultima linha visivel/ativa.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - Debug_Registar (Sub): rotina pública do módulo.
' - Seguimento_Registar (Sub): rotina pública do módulo.
' - Seguimento_ArquivarLimpar (Sub): rotina pública do módulo.
' - Debug_GetRenderPauseSeconds (Function): helper de leitura da Config para pausa de render no DEBUG.
' - Debug_DeduzirFuncionalidade (Function): descreve em linguagem simples o processo associado ao parametro.
' - Debug_DeduzirAcaoEmCurso (Function): descreve de forma especifica a acao operacional corrente.
' - Debug_AcaoAdd (Sub): agrega acoes sem duplicar texto na mesma celula.
' - Debug_AplicarAcoesPorSinal (Sub): acrescenta acoes por sinais de parametro/contexto.
' - Debug_ExtrairDetalheOperacional (Function): recolhe detalhes chave=valor para contexto da acao.
' =============================================================================

Private Const SHEET_DEBUG As String = "DEBUG"
Private Const SHEET_SEGUIMENTO As String = "Seguimento"
Private Const SHEET_HISTORICO As String = "HISTÓRICO"
Private Const DEFAULT_DEBUG_RENDER_PAUSE_S As Double = 0.003
Private Const CFG_DEBUG_RENDER_PAUSE_MS As String = "DEBUG_RENDER_PAUSE_MS"

' ============================================================================
' Debug_Registar (robusto)
' - Procura sempre as colunas pelos nomes do cabecalho (linha 1)
' - Resistente a reordenacao de colunas
' - Tolerante a espacos, maiusculas/minusculas e acentos no cabecalho
' - Se algum cabecalho nao existir, ignora essa escrita (nao gera erro)
' ============================================================================
Public Sub Debug_Registar( _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal severidade As String, _
    ByVal linhaConfigExtra As Variant, _
    ByVal parametro As String, _
    ByVal problema As String, _
    ByVal sugestao As String, _
    Optional ByVal funcionalidade As String = "" _
)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEBUG)

    Dim mapa As Object
    Set mapa = Debug_MapaCabecalhos(ws)

    Call Debug_GarantirColunaFuncionalidade(ws, mapa)

    Dim novaLinha As Long
    novaLinha = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1

    Dim funcionalidadeBase As String
    If Len(Trim$(funcionalidade)) = 0 Then
        funcionalidadeBase = Debug_DeduzirFuncionalidade(parametro)
    Else
        funcionalidadeBase = Trim$(funcionalidade)
    End If

    Dim acaoEmCurso As String
    acaoEmCurso = Debug_DeduzirAcaoEmCurso(parametro, problema, sugestao)

    funcionalidade = Debug_MontarFuncionalidade(funcionalidadeBase, acaoEmCurso)

    Debug_SetValue ws, mapa, novaLinha, "Timestamp", Now
    Debug_SetValue ws, mapa, novaLinha, "Passo", passo
    Debug_SetValue ws, mapa, novaLinha, "Prompt ID", promptId
    Debug_SetValue ws, mapa, novaLinha, "Severidade", severidade
    Debug_SetValue ws, mapa, novaLinha, "Linha (Config extra)", linhaConfigExtra
    Debug_SetValue ws, mapa, novaLinha, "Parametro", parametro          ' aceita "Parâmetro" no Excel
    Debug_SetValue ws, mapa, novaLinha, "Funcionalidade", funcionalidade
    Debug_SetValue ws, mapa, novaLinha, "Problema", problema
    Debug_SetValue ws, mapa, novaLinha, "Sugestao", sugestao            ' aceita "Sugestão" no Excel

    Call Debug_AplicarEstiloLinha(ws, mapa, novaLinha, severidade, parametro, problema)
    Call Debug_FormatarAcaoEmCurso(ws, mapa, novaLinha, funcionalidade)
    Call Debug_FocarUltimaLinha(ws, novaLinha)
End Sub

Private Sub Debug_GarantirColunaFuncionalidade(ByVal ws As Worksheet, ByRef mapa As Object)
    On Error GoTo Fim

    Dim keyFunc As String
    keyFunc = Debug_NormalizarCabecalho("Funcionalidade")
    If mapa.exists(keyFunc) Then Exit Sub

    Dim keyParam As String
    keyParam = Debug_NormalizarCabecalho("Parametro")

    Dim keyProb As String
    keyProb = Debug_NormalizarCabecalho("Problema")

    Dim colInsert As Long
    colInsert = 0

    If mapa.exists(keyProb) Then
        colInsert = CLng(mapa(keyProb))
    ElseIf mapa.exists(keyParam) Then
        colInsert = CLng(mapa(keyParam)) + 1
    End If

    If colInsert <= 0 Then Exit Sub

    ws.Columns(colInsert).Insert Shift:=xlToRight
    ws.Cells(1, colInsert).value = "Funcionalidade"

    Set mapa = Debug_MapaCabecalhos(ws)

Fim:
    On Error GoTo 0
End Sub

Private Function Debug_DeduzirFuncionalidade(ByVal parametro As String) As String
    Dim p As String
    p = UCase$(Trim$(parametro))

    If p = "" Then
        Debug_DeduzirFuncionalidade = "Registo tecnico do passo da pipeline."
        Exit Function
    End If

    If p = "INFO" Or p = "ALERTA" Then
        Debug_DeduzirFuncionalidade = "Sinalizacao operacional para acompanhamento rapido da execucao."
        Exit Function
    End If

    If Left$(p, 4) = "M10_" Or InStr(1, p, "_CI_", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Download, validacao e contrato dos ficheiros finais gerados pelo Code Interpreter."
        Exit Function
    End If

    If Left$(p, 4) = "M07_" Or Left$(p, 15) = "OUTPUT_EXECUTE_" Or Left$(p, 13) = "EXECUTE_LINT_" Then
        Debug_DeduzirFuncionalidade = "Aplicacao do plano de output e validacao dos artefactos guardados no OUTPUT Folder."
        Exit Function
    End If

    If InStr(1, p, "CATALOGO", vbTextCompare) > 0 Or InStr(1, p, "CRIAR MAPA", vbTextCompare) > 0 Or InStr(1, p, "SET DEFAULT", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Leitura do catalogo de prompts e preparacao dos defaults estruturais de execucao."
        Exit Function
    End If

    If InStr(1, p, "NEXTDEFAULT", vbTextCompare) > 0 Or InStr(1, p, "NEXT ALLOWED", vbTextCompare) > 0 Or InStr(1, p, "CICLOS", vbTextCompare) > 0 Or InStr(1, p, "LIMITELISTA", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Validacao de transicoes de Next PROMPT e protecao contra ciclos inesperados."
        Exit Function
    End If

    If InStr(1, p, "INPUT FOLDER", vbTextCompare) > 0 Or InStr(1, p, "INPUT_TEXT_SIZE", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Preparacao dos inputs do passo (pasta de entrada e volume de contexto enviado)."
        Exit Function
    End If

    If InStr(1, p, "FILE", vbTextCompare) > 0 Or InStr(1, p, "UPLOAD", vbTextCompare) > 0 Or InStr(1, p, "PDF", vbTextCompare) > 0 Or InStr(1, p, "DOCX", vbTextCompare) > 0 Or InStr(1, p, "TEXT_EMBED", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Gestao de anexos e transformacao de ficheiros da pipeline."
        Exit Function
    End If

    If InStr(1, p, "OUTPUT", vbTextCompare) > 0 Or InStr(1, p, "CHAIN", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Geracao e validacao dos ficheiros de saida."
        Exit Function
    End If

    If InStr(1, p, "API", vbTextCompare) > 0 Or InStr(1, p, "HTTP", vbTextCompare) > 0 Or InStr(1, p, "JSON", vbTextCompare) > 0 Or InStr(1, p, "UTF8", vbTextCompare) > 0 Or InStr(1, p, "PAYLOAD", vbTextCompare) > 0 Or InStr(1, p, "TIMEOUT", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Comunicacao com a API e validacao tecnica do pedido."
        Exit Function
    End If

    If InStr(1, p, "NEXT PROMPT", vbTextCompare) > 0 Or InStr(1, p, "MAX", vbTextCompare) > 0 Or InStr(1, p, "STEP", vbTextCompare) > 0 Or InStr(1, p, "PIPELINE", vbTextCompare) > 0 Or InStr(1, p, "STARTID", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Controlo do fluxo e dos limites de execucao da pipeline."
        Exit Function
    End If

    If InStr(1, p, "CONFIG", vbTextCompare) > 0 Or InStr(1, p, "PARAM", vbTextCompare) > 0 Or InStr(1, p, "OPENAI_API_KEY", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Leitura e validacao das configuracoes necessarias para correr a pipeline."
        Exit Function
    End If

    If InStr(1, p, "CONTEXT", vbTextCompare) > 0 Or InStr(1, p, "KV", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Partilha de contexto entre passos para manter continuidade."
        Exit Function
    End If

    If InStr(1, p, "SELFTEST", vbTextCompare) > 0 Or InStr(1, p, "DEBUG_DIAG", vbTextCompare) > 0 Then
        Debug_DeduzirFuncionalidade = "Autotestes e diagnostico interno para detetar regressões."
        Exit Function
    End If

    Debug_DeduzirFuncionalidade = "Monitorizacao tecnica de uma etapa da execucao."
End Function

Private Function Debug_DeduzirAcaoEmCurso(ByVal parametro As String, ByVal problema As String, ByVal sugestao As String) As String
    Dim p As String
    Dim ctx As String
    Dim acoes As String
    Dim detalhe As String

    p = UCase$(Trim$(parametro))
    ctx = UCase$(Trim$(problema & " | " & sugestao))

    If p = "STEP_STAGE" Then
        Call Debug_AcaoAdd(acoes, "Atualizacao do stage de execucao")
        If InStr(1, ctx, "STAGE=BEFORE_API", vbTextCompare) > 0 Then Call Debug_AcaoAdd(acoes, "Preparacao da chamada /v1/responses")
        If InStr(1, ctx, "STAGE=API_CALL_START", vbTextCompare) > 0 Then Call Debug_AcaoAdd(acoes, "Disparo do pedido HTTP para a API")
        If InStr(1, ctx, "STAGE=STEP_COMPLETED", vbTextCompare) > 0 Then Call Debug_AcaoAdd(acoes, "Fecho tecnico do passo com logging final")
    End If

    Select Case p
        Case "INFO", "ALERTA"
            Call Debug_AcaoAdd(acoes, "Sinalizacao operacional para acompanhamento da execucao")

        Case "REQ_INPUT_JSON"
            Call Debug_AcaoAdd(acoes, "Montagem do bloco input do request")
            Call Debug_AcaoAdd(acoes, "Validacao de anexos (input_file/input_image/text_embed)")

        Case "INPUTFILES_MISSING"
            Call Debug_AcaoAdd(acoes, "Validacao bloqueante do contrato FILES declarado")
            Call Debug_AcaoAdd(acoes, "Interrupcao preventiva por anexo obrigatorio em falta")

        Case "FILES_ITEM_TRACE"
            Call Debug_AcaoAdd(acoes, "Resolucao de item FILES declarado no INPUTS")
            Call Debug_AcaoAdd(acoes, "Aplicacao de flags required/latest/as_pdf/as_is/text_embed")

        Case "FILES", "FILES UPLOAD", "FILES REUSE", "FILES IA", "FILES_DIAG", "DOCX_INPUTFILE_OVERRIDDEN", "PDF_CACHE_HIT", "PDF_CACHE_MISS_CONVERTED", "TEXT_EMBED_EMPTY", "TEXT_EMBED_TOO_LARGE", "AS_PDF", "TEXT_EMBED"
            Call Debug_AcaoAdd(acoes, "Preparacao e transformacao de anexos")

        Case "M05_PAYLOAD_CHECK", "M05_JSON_PREFLIGHT", "M05_UTF8_ROUNDTRIP", "M05_PAYLOAD_DUMP", "M05_PAYLOAD_DUMP_FAIL", "M05_TIMEOUT_DECISION", "M05_HTTP_TIMEOUTS", "M05_HTTP_TIMEOUT_INVALID", "M05_HTTP_TIMEOUT_ERROR", "M05_HTTP_RESULT", "API", "API_RETRY_5XX", "API_CONTEXT_LENGTH_ACTION", "API_CONTEXT_LENGTH_EXCEEDED"
            Call Debug_AcaoAdd(acoes, "Montagem e validacao tecnica do payload")
            Call Debug_AcaoAdd(acoes, "Execucao/monitorizacao da chamada HTTP")

        Case "M05_CI_AUTO_SUPPRESS", "M05_CI_INTENT_EVAL"
            Call Debug_AcaoAdd(acoes, "Avaliacao da intencao explicita de Code Interpreter")
            Call Debug_AcaoAdd(acoes, "Gate da auto-injecao da tool code_interpreter")

        Case "M07_FILEOUTPUT_MODE_MISMATCH", "M07_FILEOUTPUT_PARSE_GUARD", "OUTPUT_CHAIN_RESOLVE", "OUTPUT_CHAIN_NOT_FOUND", "OUTPUT_CHAIN_AMBIGUOUS", "OUTPUT_CONTRACT_FAIL", "OUTPUT_FILE_MISSING_ON_DISK"
            Call Debug_AcaoAdd(acoes, "Validacao do contrato de File Output")
            Call Debug_AcaoAdd(acoes, "Resolucao do ficheiro esperado no OUTPUT Folder")

        Case "OUTPUT_EXECUTE_FOUND", "OUTPUT_EXECUTE_PARSED", "OUTPUT_EXECUTE_UNKNOWN_CMD", "OUTPUT_EXECUTE_INVALID_FILENAME", "OUTPUT_EXECUTE_FILE_NOT_FOUND", "OUTPUT_EXECUTE_CSV_PRECHECK", "OUTPUT_EXECUTE_CSV_BOM_FAIL", "OUTPUT_EXECUTE_CSV_CRLF_IN_FIELDS", "OUTPUT_EXECUTE_SHEET_CREATED", "OUTPUT_EXECUTE_IMPORT_FAIL", "OUTPUT_EXECUTE_CSV_IMPORTED", "OUTPUT_EXECUTE_VERIFIED"
            Call Debug_AcaoAdd(acoes, "Execucao da diretiva OUTPUT_EXECUTE")
            Call Debug_AcaoAdd(acoes, "Importacao/validacao de CSV em worksheet dedicada")

        Case "M10_CI_NO_CITATION", "M10_CI_NO_CONTAINER_ID", "M10_CI_CONTAINER_EMPTY", "M10_CI_CONTAINER_LIST", "M10_CI_CONTAINER_SELECT_DIAG", "M10_CI_CONTAINER_INPUT_LIKE", "M10_CI_TEXT_FILENAME_HINTS", "M10_CI_MARKER_NOT_FOUND", "M10_CI_AMBIGUOUS_MARKER", "M10_CI_AMBIGUOUS_FALLBACK", "M10_CI_RESOLVE_RULE", "M10_CI_CONTRACT_STATUS", "M10_CI_DOWNLOAD_NOFILE", "M10_CI_DOWNLOAD_FAIL", "M10_CI_LIST_FAIL", "M10_CI_RAW_MISSING", "M10_CI_ZERO_BYTES", "M10_RAWFOLDER", "M10_RUNFOLDER", "M10_RAW_WRITE_FAIL", "M10_SCHEMA_SUMMARY", "M10_SCHEMA_INVALID", "M10_SCHEMA_DIAG_FAIL", "M10_META_PATH_TOO_LONG", "M10_PATH_TOO_LONG", "M10_FOLDER_CREATE_FAIL"
            Call Debug_AcaoAdd(acoes, "Inspecao de artefactos de output no container do Code Interpreter")
            Call Debug_AcaoAdd(acoes, "Selecao e download do artefacto final")
            Call Debug_AcaoAdd(acoes, "Persistencia local e validacao do ficheiro descarregado")

        Case "NEXT PROMPT", "NEXTDEFAULT", "NEXT ALLOWED", "CICLOS", "MAXSTEPS", "MAXREPETITIONS", "LIMITELISTA", "PIPELINE"
            Call Debug_AcaoAdd(acoes, "Validacao de transicao de Next PROMPT")
            Call Debug_AcaoAdd(acoes, "Aplicacao de limites de execucao da pipeline")

        Case "CATALOGO", "CRIAR MAPA", "SET DEFAULT", "CONFIG_OUTPUT_CHAIN_MODE_INVALID", "OPENAI_API_KEY", "CONTEXT_KV", "INPUT FOLDER", "INPUT_TEXT_SIZE"
            Call Debug_AcaoAdd(acoes, "Leitura de configuracao e metadados do passo")
            If p = "CONTEXT_KV" Then Call Debug_AcaoAdd(acoes, "Captura/injecao de variaveis partilhadas entre prompts")
            If p = "INPUT FOLDER" Then Call Debug_AcaoAdd(acoes, "Resolucao da pasta de entrada e politica de seguranca")

        Case "SELFTEST_SUMMARY", "SELFTEST_CRASH", "DEBUG_DIAG", "CONFIG_EXTRA_CASE_OK", "CONFIG_EXTRA_CASE_FAIL"
            Call Debug_AcaoAdd(acoes, "Execucao de autoteste/diagnostico interno")
            Call Debug_AcaoAdd(acoes, "Registo de resultado PASS/FAIL com recomendacoes")
    End Select

    Call Debug_AplicarAcoesPorSinal(acoes, p, ctx)

    If acoes = "" Then acoes = "Registo tecnico da operacao em curso"

    detalhe = Debug_ExtrairDetalheOperacional(problema, sugestao)
    If detalhe <> "" Then
        Debug_DeduzirAcaoEmCurso = acoes & " | contexto: " & detalhe
    Else
        Debug_DeduzirAcaoEmCurso = acoes
    End If
End Function

Private Sub Debug_AcaoAdd(ByRef lista As String, ByVal item As String)
    Dim txt As String
    txt = Trim$(item)
    If txt = "" Then Exit Sub

    If lista = "" Then
        lista = txt
    ElseIf InStr(1, "|" & lista & "|", "|" & txt & "|", vbTextCompare) = 0 Then
        lista = lista & " | " & txt
    End If
End Sub

Private Sub Debug_AplicarAcoesPorSinal(ByRef acoes As String, ByVal p As String, ByVal ctx As String)
    Select Case p
        Case "FILES REUSE"
            Call Debug_AcaoAdd(acoes, "Reutilizacao de upload existente por hash/nome")
        Case "FILES IA"
            Call Debug_AcaoAdd(acoes, "Desambiguacao assistida para escolha do ficheiro candidato")
        Case "DOCX_INPUTFILE_OVERRIDDEN"
            Call Debug_AcaoAdd(acoes, "Override de modo de anexo Office para formato suportado")
        Case "PDF_CACHE_HIT", "PDF_CACHE_MISS_CONVERTED"
            Call Debug_AcaoAdd(acoes, "Gestao da cache de conversao PDF para anexos Office")
        Case "TEXT_EMBED_EMPTY"
            Call Debug_AcaoAdd(acoes, "Detecao de extracao vazia em text_embed")
        Case "TEXT_EMBED_TOO_LARGE"
            Call Debug_AcaoAdd(acoes, "Aplicacao de politica de overflow de text_embed")
        Case "M10_CI_DOWNLOAD_FAIL", "M10_CI_DOWNLOAD_NOFILE"
            Call Debug_AcaoAdd(acoes, "Tratamento de falha no download de artefacto do container")
        Case "M10_CI_AMBIGUOUS_MARKER", "M10_CI_AMBIGUOUS_FALLBACK"
            Call Debug_AcaoAdd(acoes, "Resolucao de ambiguidade na selecao do artefacto final")
        Case "M10_META_PATH_TOO_LONG", "M10_PATH_TOO_LONG"
            Call Debug_AcaoAdd(acoes, "Validacao de limite de caminho no sistema de ficheiros")
        Case "API_CONTEXT_LENGTH_EXCEEDED", "API_CONTEXT_LENGTH_ACTION"
            Call Debug_AcaoAdd(acoes, "Mitigacao de excesso de contexto antes de reenviar pedido")
        Case "OUTPUT_CONTRACT_FAIL"
            Call Debug_AcaoAdd(acoes, "Bloqueio por violacao do contrato de output esperado")
        Case "M10_CI_CONTAINER_LIST"
            Call Debug_AcaoAdd(acoes, "Listagem de ficheiros no container para eleicao de candidato")
        Case "M10_CI_CONTAINER_SELECT_DIAG"
            Call Debug_AcaoAdd(acoes, "Avaliacao de elegibilidade por ficheiro (motivo SIM/NAO)")
        Case "M10_CI_NO_CITATION"
            Call Debug_AcaoAdd(acoes, "Fallback para resolucao sem citation via contrato textual")
        Case "M10_CI_RESOLVE_RULE"
            Call Debug_AcaoAdd(acoes, "Aplicacao da prioridade citation > CI_OUTPUT_FILE > fallback")
        Case "M10_CI_RAW_MISSING"
            Call Debug_AcaoAdd(acoes, "Detecao de ausencia de raw de resposta para extracao de artefacto")
        Case "M10_CI_MARKER_NOT_FOUND"
            Call Debug_AcaoAdd(acoes, "Detecao de marcador de ficheiro ausente no output textual")
        Case "M10_CI_TEXT_FILENAME_HINTS"
            Call Debug_AcaoAdd(acoes, "Inferencia de candidato por pistas textuais de filename")
        Case "M10_CI_CONTRACT_STATUS"
            Call Debug_AcaoAdd(acoes, "Validacao de conformidade do artefacto com contrato de output")
        Case "M10_CI_LIST_FAIL"
            Call Debug_AcaoAdd(acoes, "Tratamento de erro ao listar ficheiros do container")
        Case "M10_FOLDER_CREATE_FAIL", "M10_RAW_WRITE_FAIL"
            Call Debug_AcaoAdd(acoes, "Tratamento de erro de escrita/criacao em disco local")
        Case "M10_RUNFOLDER", "M10_RAWFOLDER"
            Call Debug_AcaoAdd(acoes, "Preparacao de pastas de run/raw para auditoria")
        Case "M10_SCHEMA_INVALID", "M10_SCHEMA_DIAG_FAIL"
            Call Debug_AcaoAdd(acoes, "Validacao de schema esperado e classificacao de desvio")
        Case "M05_HTTP_TIMEOUT_ERROR"
            Call Debug_AcaoAdd(acoes, "Classificacao de fase de timeout (resolve/connect/send/receive)")
        Case "M05_TIMEOUT_DECISION"
            Call Debug_AcaoAdd(acoes, "Decisao operacional de retry/backoff apos timeout")
        Case "M05_PAYLOAD_CHECK"
            Call Debug_AcaoAdd(acoes, "Checklist pre-envio do payload (tools/input/tamanho)")
        Case "M05_JSON_PREFLIGHT"
            Call Debug_AcaoAdd(acoes, "Preflight de validade JSON antes do HTTP")
        Case "M05_PAYLOAD_DUMP", "M05_PAYLOAD_DUMP_FAIL"
            Call Debug_AcaoAdd(acoes, "Persistencia diagnostica do payload final de request")
        Case "M05_CI_INTENT_EVAL", "M05_CI_AUTO_SUPPRESS"
            Call Debug_AcaoAdd(acoes, "Avaliacao de intencao explicita de Code Interpreter no passo")
        Case "M07_FILEOUTPUT_MODE_MISMATCH", "M07_FILEOUTPUT_PARSE_GUARD"
            Call Debug_AcaoAdd(acoes, "Validacao de coerencia entre modo efetivo e contrato de File Output")
        Case "OUTPUT_EXECUTE_CSV_PRECHECK"
            Call Debug_AcaoAdd(acoes, "Pre-validacao de encoding/BOM e estrutura CSV")
        Case "OUTPUT_EXECUTE_VERIFIED"
            Call Debug_AcaoAdd(acoes, "Confirmacao final do artefacto importado e rastreabilidade")
        Case "OUTPUT_EXECUTE_IMPORT_FAIL"
            Call Debug_AcaoAdd(acoes, "Tratamento de falha de importacao CSV para worksheet")
        Case "OUTPUT_EXECUTE_INVALID_FILENAME"
            Call Debug_AcaoAdd(acoes, "Bloqueio de filename invalido por regras de seguranca")
        Case "INPUTFILES_MISSING"
            Call Debug_AcaoAdd(acoes, "Comparacao entre FILES declarados e anexos efetivos no payload")
        Case "NEXTDEFAULT", "NEXT ALLOWED", "NEXT PROMPT"
            Call Debug_AcaoAdd(acoes, "Aplicacao de regras default/allowed na transicao de prompt")
        Case "CICLOS", "MAXSTEPS", "MAXREPETITIONS", "LIMITELISTA"
            Call Debug_AcaoAdd(acoes, "Aplicacao de guardas anti-loop e limites de execucao")
        Case "CATALOGO", "CRIAR MAPA"
            Call Debug_AcaoAdd(acoes, "Resolucao de bloco do prompt no catalogo de origem")
        Case "CONTEXT_KV"
            Call Debug_AcaoAdd(acoes, "Injecao/captura de variaveis de contexto entre passos")
        Case "OPENAI_API_KEY"
            Call Debug_AcaoAdd(acoes, "Validacao de disponibilidade da chave API em Config")
        Case "OUTPUT_CHAIN_NOT_FOUND", "OUTPUT_CHAIN_AMBIGUOUS"
            Call Debug_AcaoAdd(acoes, "Resolucao da cadeia de output com fallback/controlos de ambiguidade")
    End Select

    If Left$(p, 4) = "M10_" Or InStr(1, p, "_CI_", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Inspecao de artefactos de output no container do Code Interpreter")
    End If

    If InStr(1, p, "DOWNLOAD", vbTextCompare) > 0 Or InStr(1, ctx, "DOWNLOAD", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Download de artefacto para staging/destino final")
    End If

    If InStr(1, p, "UPLOAD", vbTextCompare) > 0 Or InStr(1, p, "FILES", vbTextCompare) > 0 Or InStr(1, ctx, "/V1/FILES", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Upload/ligacao de anexos no pedido")
    End If

    If InStr(1, p, "TIMEOUT", vbTextCompare) > 0 Or InStr(1, ctx, "TIMEOUT", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Calculo e aplicacao de timeouts efetivos")
    End If

    If InStr(1, p, "RETRY", vbTextCompare) > 0 Or InStr(1, ctx, "RETRY", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Gestao de retry por erro transitorio")
    End If

    If InStr(1, p, "OUTPUT", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Validacao de output esperado e artefacto final")
    End If

    If InStr(1, p, "NEXT", vbTextCompare) > 0 Or InStr(1, p, "PIPELINE", vbTextCompare) > 0 Or InStr(1, p, "MAX", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Controlo de transicao/limites da pipeline")
    End If

    If InStr(1, p, "CONFIG", vbTextCompare) > 0 Or InStr(1, p, "CATALOGO", vbTextCompare) > 0 Or InStr(1, p, "PARAM", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Leitura e validacao de configuracao")
    End If

    If Left$(p, 9) = "SELFTEST_" Or InStr(1, p, "DEBUG_DIAG", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Execucao de autoteste e consolidacao diagnostica")
    End If

    If InStr(1, ctx, "FILE_ID", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Associacao de ficheiro por file_id no payload")
    End If

    If InStr(1, ctx, "CONTAINER_ID", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Rastreio de container para resolucao de artefactos")
    End If

    If InStr(1, ctx, "PROCESS_MODE=CODE_INTERPRETER", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Correlacao do modo code_interpreter com o contrato de output")
    End If

    If InStr(1, ctx, "HTTP STATUS", vbTextCompare) > 0 Or InStr(1, ctx, "HTTP_STATUS", vbTextCompare) > 0 Then
        Call Debug_AcaoAdd(acoes, "Correlacao de status HTTP com acao de recuperacao")
    End If
End Sub

Private Function Debug_ExtrairDetalheOperacional(ByVal problema As String, ByVal sugestao As String) As String
    Dim fonte As String
    Dim detalhe As String
    Dim detalhes As String
    Dim keys As Variant
    Dim i As Long

    fonte = Trim$(problema & " | " & sugestao)
    keys = Array("filename", "file", "full_path", "resolvedPath", "resolved_path", "inputFolder", "input_folder", "outputFolder", "output_folder", "stage", "endpoint", "prompt", "promptId", "pipeline", "pipeline_name", "container_id", "file_id", "status", "http_status", "httpStatus", "engine", "profile", "effective_mode", "mode_effective", "mode", "bytes", "created_at", "elapsed_ms", "payload_len", "response_id", "cause_hint", "confidence", "dlErr", "retry_outcome")

    For i = LBound(keys) To UBound(keys)
        detalhe = Debug_ExtrairDetalhePorChave(fonte, CStr(keys(i)))
        If detalhe <> "" Then Call Debug_AcaoAdd(detalhes, detalhe)
    Next i

    Debug_ExtrairDetalheOperacional = detalhes
End Function

Private Function Debug_ExtrairDetalhePorChave(ByVal fonte As String, ByVal chave As String) As String
    Dim pos As Long
    Dim ini As Long
    Dim i As Long
    Dim ch As String
    Dim valor As String

    pos = InStr(1, fonte, chave & "=", vbTextCompare)
    If pos > 0 Then
        ini = pos + Len(chave) + 1
    Else
        pos = InStr(1, fonte, chave & ":", vbTextCompare)
        If pos <= 0 Then Exit Function
        ini = pos + Len(chave) + 1
    End If
    For i = ini To Len(fonte)
        ch = Mid$(fonte, i, 1)
        If ch = ";" Or ch = "," Or ch = "|" Or ch = vbCr Or ch = vbLf Then Exit For
        valor = valor & ch
    Next i

    valor = Trim$(Replace(Replace(Replace(valor, Chr$(34), ""), "'", ""), "`", ""))
    If valor <> "" Then Debug_ExtrairDetalhePorChave = chave & "=" & valor
End Function

Private Function Debug_MontarFuncionalidade(ByVal funcionalidadeBase As String, ByVal acaoEmCurso As String) As String
    Dim baseTxt As String
    Dim acaoTxt As String

    baseTxt = Trim$(funcionalidadeBase)
    If baseTxt = "" Then baseTxt = "Monitorizacao tecnica de uma etapa da execucao."

    acaoTxt = Trim$(acaoEmCurso)
    If acaoTxt = "" Then acaoTxt = "Registo tecnico da operacao em curso."

    Debug_MontarFuncionalidade = baseTxt & vbLf & "ACAO EM CURSO: " & acaoTxt
End Function

Private Sub Debug_FormatarAcaoEmCurso(ByVal ws As Worksheet, ByVal mapa As Object, ByVal linha As Long, ByVal funcionalidadeTexto As String)
    On Error GoTo Fim

    Dim keyFunc As String
    Dim colFunc As Long
    keyFunc = Debug_NormalizarCabecalho("Funcionalidade")
    If Not mapa.exists(keyFunc) Then Exit Sub

    colFunc = CLng(mapa(keyFunc))

    Dim posQuebra As Long
    posQuebra = InStr(1, funcionalidadeTexto, vbLf, vbBinaryCompare)
    If posQuebra <= 0 Then Exit Sub

    Dim startBold As Long
    startBold = posQuebra + 1

    With ws.Cells(linha, colFunc)
        .WrapText = True
        .Characters(startBold, Len(funcionalidadeTexto) - startBold + 1).Font.Bold = True
    End With

Fim:
    On Error GoTo 0
End Sub

' ============================================================================
' Helpers (DEBUG)
' ============================================================================
Private Sub Debug_SetValue(ByVal ws As Worksheet, ByVal mapa As Object, ByVal linha As Long, ByVal cabecalho As String, ByVal valor As Variant)
    Dim chave As String
    chave = Debug_NormalizarCabecalho(cabecalho)

    If mapa.exists(chave) Then
        ws.Cells(linha, CLng(mapa(chave))).value = valor
    End If
End Sub



Private Sub Debug_AplicarEstiloLinha(ByVal ws As Worksheet, ByVal mapa As Object, ByVal linha As Long, ByVal severidade As String, ByVal parametro As String, ByVal problema As String)
    On Error Resume Next

    Dim ultimaColuna As Long
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ultimaColuna <= 0 Then Exit Sub

    With ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultimaColuna)).Font
        .Bold = False
        .ColorIndex = xlColorIndexAutomatic
    End With

    Dim sev As String
    sev = UCase$(Trim$(severidade))

    If sev = "ERRO" Then
        ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultimaColuna)).Font.Bold = True
        ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultimaColuna)).Font.Color = RGB(192, 0, 0)
        Exit Sub
    End If

    If sev = "ALERTA" Then
        ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultimaColuna)).Font.Bold = True
        ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultimaColuna)).Font.Color = RGB(0, 102, 204)
        Exit Sub
    End If

    If UCase$(Trim$(parametro)) = "STEP_STAGE" And InStr(1, LCase$(problema), "stage=step_completed", vbTextCompare) > 0 Then
        ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultimaColuna)).Font.Bold = True
        ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultimaColuna)).Font.Color = RGB(0, 128, 0)
    End If

    On Error GoTo 0
End Sub

Private Sub Debug_FocarUltimaLinha(ByVal ws As Worksheet, ByVal linha As Long)
    On Error Resume Next

    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = True
    ws.Activate

    Dim visRows As Long
    visRows = 0
    If Not ActiveWindow Is Nothing Then
        visRows = ActiveWindow.VisibleRange.Rows.Count
    End If
    If visRows <= 0 Then visRows = 1

    If Not ActiveWindow Is Nothing Then
        ActiveWindow.ScrollRow = Application.Max(1, linha - visRows + 1)
        ActiveWindow.ScrollColumn = 1
    End If

    ws.Cells(linha, 1).Select

    Call Debug_PausaCurta(Debug_GetRenderPauseSeconds())

    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
End Sub


Private Function Debug_GetRenderPauseSeconds() As Double
    On Error GoTo Falha

    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Worksheets("Config")

    Dim lr As Long
    lr = wsCfg.Cells(wsCfg.rowS.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 1 To lr
        If StrComp(Trim$(CStr(wsCfg.Cells(r, 1).value)), CFG_DEBUG_RENDER_PAUSE_MS, vbTextCompare) = 0 Then
            Dim ms As Double
            ms = CDbl(Val(Replace(Trim$(CStr(wsCfg.Cells(r, 2).value)), ",", ".")))
            If ms < 0 Then ms = 0
            Debug_GetRenderPauseSeconds = ms / 1000#
            Exit Function
        End If
    Next r

Falha:
    Debug_GetRenderPauseSeconds = DEFAULT_DEBUG_RENDER_PAUSE_S
End Function

Private Sub Debug_PausaCurta(ByVal segundos As Double)
    On Error Resume Next

    If segundos <= 0 Then Exit Sub

    Dim limite As Double
    limite = Timer + segundos

    If limite >= 86400# Then
        limite = limite - 86400#
        Do While Timer >= 0 And Timer < limite
            DoEvents
        Loop
        Exit Sub
    End If

    Do While Timer < limite
        DoEvents
    Loop
End Sub

Private Function Debug_MapaCabecalhos(ByVal ws As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim ultimaColuna As Long
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To ultimaColuna
        Dim texto As String
        texto = Trim$(CStr(ws.Cells(1, c).value))
        If texto <> "" Then
            Dim k As String
            k = Debug_NormalizarCabecalho(texto)
            If Not d.exists(k) Then
                d.Add k, c
            End If
        End If
    Next c

    Set Debug_MapaCabecalhos = d
End Function


Private Function Debug_NormalizarCabecalho(ByVal s As String) As String
    ' Normalizacao tolerante:
    ' - trim + minusculas
    ' - colapsa multiplos espacos
    ' - remove acentos (usando ChrW para evitar problemas de codificacao em .bas)
    s = LCase$(Trim$(s))

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    s = Debug_RemoverAcentos(s)
    Debug_NormalizarCabecalho = s
End Function


Private Function Debug_RemoverAcentos(ByVal s As String) As String
    ' Substitui letras acentuadas comuns por equivalentes sem acento.
    ' Usa ChrW(...) para evitar caracteres "esquisitos" quando se importa .bas.

    ' a
    s = Replace(s, ChrW(225), "a") ' á
    s = Replace(s, ChrW(224), "a") ' à
    s = Replace(s, ChrW(226), "a") ' â
    s = Replace(s, ChrW(227), "a") ' ã
    s = Replace(s, ChrW(228), "a") ' ä

    ' e
    s = Replace(s, ChrW(233), "e") ' é
    s = Replace(s, ChrW(232), "e") ' è
    s = Replace(s, ChrW(234), "e") ' ê
    s = Replace(s, ChrW(235), "e") ' ë

    ' i
    s = Replace(s, ChrW(237), "i") ' í
    s = Replace(s, ChrW(236), "i") ' ì
    s = Replace(s, ChrW(238), "i") ' î
    s = Replace(s, ChrW(239), "i") ' ï

    ' o
    s = Replace(s, ChrW(243), "o") ' ó
    s = Replace(s, ChrW(242), "o") ' ò
    s = Replace(s, ChrW(244), "o") ' ô
    s = Replace(s, ChrW(245), "o") ' õ
    s = Replace(s, ChrW(246), "o") ' ö

    ' u
    s = Replace(s, ChrW(250), "u") ' ú
    s = Replace(s, ChrW(249), "u") ' ù
    s = Replace(s, ChrW(251), "u") ' û
    s = Replace(s, ChrW(252), "u") ' ü

    ' c
    s = Replace(s, ChrW(231), "c") ' ç

    Debug_RemoverAcentos = s
End Function




' ============================================================================
' Seguimento_Registar (robusta)
' - Procura sempre as colunas pelos nomes do cabecalho (linha 1)
' - Resistente a reordenacao de colunas
' - Se algum cabecalho nao existir, ignora essa escrita (nao gera erro)
' ============================================================================
Public Sub Seguimento_Registar( _
    ByVal passo As Long, _
    ByRef prompt As PromptDefinicao, _
    ByVal modeloUsado As String, _
    ByVal auditJson As String, _
    ByVal httpStatus As Long, _
    ByVal responseId As String, _
    ByVal outputOuErro As String, _
    Optional ByVal pipelineNome As String = "", _
    Optional ByVal nextPromptDecidido As String = "", _
    Optional ByVal filesUsedResumo As String = "", _
    Optional ByVal filesOpsResumo As String = "", _
    Optional ByVal fileIdsUsed As String = "" _
)
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Seguimento")

    Dim linha As Long
    linha = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1
    If linha < 2 Then linha = 2

    ' Mapear cabecalhos (linha 1)
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")

    Dim c As Long
    For c = 1 To 200
        Dim h As String
        h = Trim$(CStr(ws.Cells(1, c).value))
        If h <> "" Then mapa(h) = c
    Next c

    ' Helper local: escrever so se existir coluna
    Dim txt As String

    If mapa.exists("Timestamp") Then ws.Cells(linha, mapa("Timestamp")).value = Now
    If mapa.exists("Passo") Then ws.Cells(linha, mapa("Passo")).value = passo
    If mapa.exists("Prompt ID") Then ws.Cells(linha, mapa("Prompt ID")).value = prompt.Id

    If mapa.exists("Texto da prompt") Then ws.Cells(linha, mapa("Texto da prompt")).value = prompt.textoPrompt
    If mapa.exists("Modelo") Then ws.Cells(linha, mapa("Modelo")).value = modeloUsado
    If mapa.exists("Modos") Then ws.Cells(linha, mapa("Modos")).value = prompt.modos
    If mapa.exists("Storage") Then ws.Cells(linha, mapa("Storage")).value = IIf(prompt.storage, "TRUE", "FALSE")

    If mapa.exists("Config extra (amigável)") Then
        ws.Cells(linha, mapa("Config extra (amigável)")).value = prompt.ConfigExtra
    ElseIf mapa.exists("Config extra (amigavel)") Then
        ws.Cells(linha, mapa("Config extra (amigavel)")).value = prompt.ConfigExtra
    ElseIf mapa.exists("Config extra (amigÃ¡vel)") Then
        ws.Cells(linha, mapa("Config extra (amigÃ¡vel)")).value = prompt.ConfigExtra
    End If
    If mapa.exists("Config extra (JSON convertido)") Then ws.Cells(linha, mapa("Config extra (JSON convertido)")).value = auditJson

    If mapa.exists("HTTP Status") Then
        ws.Cells(linha, mapa("HTTP Status")).value = httpStatus
        SegHTTPNota ws.Cells(linha, mapa("HTTP Status")), httpStatus
    End If
    If mapa.exists("Response ID") Then ws.Cells(linha, mapa("Response ID")).value = responseId

    txt = CStr(outputOuErro)
    Call Seguimento_EscreverOutputSemTruncagem(ws, linha, mapa, txt)

    ' Alinhamento final: usar apenas pipeline_name (nao escrever "Pipeline")
    If mapa.exists("pipeline_name") Then ws.Cells(linha, mapa("pipeline_name")).value = pipelineNome

    If mapa.exists("Next prompt decidido") Then ws.Cells(linha, mapa("Next prompt decidido")).value = nextPromptDecidido

    If mapa.exists("files_used") Then ws.Cells(linha, mapa("files_used")).value = filesUsedResumo
    If mapa.exists("files_ops_log") Then ws.Cells(linha, mapa("files_ops_log")).value = filesOpsResumo
    If mapa.exists("file_ids_used") Then ws.Cells(linha, mapa("file_ids_used")).value = fileIdsUsed

    ' Acompanhar a linha recem-criada (apenas se o utilizador estiver a ver o Seguimento)
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        If ActiveSheet.name = ws.name Then
            Application.GoTo ws.Cells(linha, 1), True   ' coluna A da linha nova
            If Application.ScreenUpdating Then DoEvents
        End If
    End If
    On Error GoTo TrataErro

    Exit Sub

TrataErro:
    ' Nao interromper pipeline por falha de logging
End Sub




' ============================================================================
' Helpers
' ============================================================================
Private Sub Seguimento_SetValue(ByVal ws As Worksheet, ByVal mapa As Object, ByVal linha As Long, ByVal cabecalho As String, ByVal valor As Variant)
    Dim chave As String
    chave = Seguimento_NormalizarCabecalho(cabecalho)

    If mapa.exists(chave) Then
        ws.Cells(linha, CLng(mapa(chave))).value = valor
    End If
End Sub


Private Function Seguimento_MapaCabecalhos(ByVal ws As Worksheet) As Object
    ' Cria um dicionario: cabecalho_normalizado -> numero da coluna
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim ultimaColuna As Long
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To ultimaColuna
        Dim texto As String
        texto = Trim$(CStr(ws.Cells(1, c).value))
        If texto <> "" Then
            Dim k As String
            k = Seguimento_NormalizarCabecalho(texto)
            If Not d.exists(k) Then
                d.Add k, c
            End If
        End If
    Next c

    Set Seguimento_MapaCabecalhos = d
End Function


Private Function Seguimento_NormalizarCabecalho(ByVal s As String) As String
    ' Normalizacao simples para ser tolerante a espacos e maiusculas/minusculas
    s = LCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Seguimento_NormalizarCabecalho = s
End Function
Private Sub SegHTTPNota(ByVal cel As Range, ByVal httpStatus As Variant)
    On Error Resume Next
    If cel Is Nothing Then Exit Sub
    If IsEmpty(httpStatus) Or httpStatus = "" Then Exit Sub
    Dim code As Long: code = CLng(httpStatus)
    Dim t As String: t = SegHTTPDesc(code)
    If t = "" Then Exit Sub
    cel.ClearComments
    cel.AddComment t
    cel.Comment.Visible = False
End Sub

Private Function SegHTTPDesc(ByVal code As Long) As String
    Select Case code
        Case 200: SegHTTPDesc = "200 OK — Sucesso"
        Case 201: SegHTTPDesc = "201 — Criado"
        Case 400: SegHTTPDesc = "400 — Pedido inválido"
        Case 401: SegHTTPDesc = "401 — Não autorizado"
        Case 403: SegHTTPDesc = "403 — Proibido"
        Case 404: SegHTTPDesc = "404 — Não encontrado"
        Case 429: SegHTTPDesc = "429 — Limite atingido"
        Case 500: SegHTTPDesc = "500 — Erro interno"
        Case 502: SegHTTPDesc = "502 — Bad gateway"
        Case 503: SegHTTPDesc = "503 — Indisponível"
        Case 504: SegHTTPDesc = "504 — Timeout"
        Case Else
            If code >= 200 And code < 300 Then SegHTTPDesc = CStr(code) & " — Sucesso (2xx)": Exit Function
            If code >= 300 And code < 400 Then SegHTTPDesc = CStr(code) & " — Redirecionamento (3xx)": Exit Function
            If code >= 400 And code < 500 Then SegHTTPDesc = CStr(code) & " — Erro do pedido (4xx)": Exit Function
            If code >= 500 And code < 600 Then SegHTTPDesc = CStr(code) & " — Erro do servidor (5xx)": Exit Function
            SegHTTPDesc = "HTTP " & CStr(code)
    End Select
End Function


' =============================================================================
' Seguimento_ArquivarLimpar (versão híbrida por HEADERS — alinhada aos teus nomes reais)
'
' Seguimento (headers):
'   Timestamp | Passo | Prompt ID | Texto da prompt | Modelo | Modos | Storage |
'   Config extra (amigável) | Config extra (JSON convertido) | HTTP Status | Response ID |
'   Output (texto) | pipeline_name | Next prompt decidido | files_used | files_ops_log |
'   file_ids_used | captured_vars | captured_vars_meta | injected_vars
'
' HISTÓRICO (headers):
'   Timestamp | Nome do Pipeline | Passo | Prompt ID | Texto da prompt | Output (texto) |
'   Modelo | Modos | Storage | Config extra (amigável) | Config extra (JSON convertido) |
'   HTTP Status | Response ID | Next prompt decidido | files_used | files_ops_log | file_ids_used |
'   captured_vars | captured_vars_meta | injected_vars
'
' Mantém a lógica do original:
' - Insere no topo (linha 2 do HISTÓRICO)
' - Cria linha separadora preta (altura 6) abaixo do bloco novo
' - Limpa Seguimento (ClearContents + ClearComments)
' - AutoFit às linhas do Seguimento
' - Em erro: tenta Debug_Registar (se existir), e fallback MsgBox
'
' Implementação:
' - Mapeia por nome de cabeçalho (case-insensitive; tolerante a acentos e espaços)
' - Não depende de posições fixas nem de "20 colunas"
' =============================================================================
Public Sub Seguimento_ArquivarLimpar()
    On Error GoTo EH
    
    Const AUTO_CREATE_MISSING_HEADERS As Boolean = True
    Const HIST_TOP_ROW As Long = 2
    Const SEPARATOR_HEIGHT As Double = 6
    
    Dim wsS As Worksheet, wsH As Worksheet
    Set wsS = ThisWorkbook.Worksheets("Seguimento")
    Set wsH = ThisWorkbook.Worksheets("HISTÓRICO")
    
    ' Ordem canónica no HISTÓRICO (exactamente como o teu cabeçalho)
    Dim histHeaders As Variant
    histHeaders = Array( _
        "Timestamp", _
        "Nome do Pipeline", _
        "Passo", _
        "Prompt ID", _
        "Texto da prompt", _
        "Output (texto)", _
        "Modelo", _
        "Modos", _
        "Storage", _
        "Config extra (amigável)", _
        "Config extra (JSON convertido)", _
        "HTTP Status", _
        "Response ID", _
        "Next prompt decidido", _
        "files_used", _
        "files_ops_log", _
        "file_ids_used", _
        "captured_vars", _
        "captured_vars_meta", _
        "injected_vars" _
    )
    
    ' Mapeamento Seguimento -> Histórico por nome (sem taxonomias; só headers)
    ' Nota: no Seguimento o pipeline chama-se pipeline_name; no Histórico chama-se Nome do Pipeline.
    Dim srcForHist As Object
    Set srcForHist = CreateObject("Scripting.Dictionary")
    srcForHist.CompareMode = 1 ' TextCompare
    
    srcForHist("Timestamp") = "Timestamp"
    srcForHist("Nome do Pipeline") = "pipeline_name"
    srcForHist("Passo") = "Passo"
    srcForHist("Prompt ID") = "Prompt ID"
    srcForHist("Texto da prompt") = "Texto da prompt"
    srcForHist("Output (texto)") = "Output (texto)"
    srcForHist("Modelo") = "Modelo"
    srcForHist("Modos") = "Modos"
    srcForHist("Storage") = "Storage"
    srcForHist("Config extra (amigável)") = "Config extra (amigável)"
    srcForHist("Config extra (JSON convertido)") = "Config extra (JSON convertido)"
    srcForHist("HTTP Status") = "HTTP Status"
    srcForHist("Response ID") = "Response ID"
    srcForHist("Next prompt decidido") = "Next prompt decidido"
    srcForHist("files_used") = "files_used"
    srcForHist("files_ops_log") = "files_ops_log"
    srcForHist("file_ids_used") = "file_ids_used"
    srcForHist("captured_vars") = "captured_vars"
    srcForHist("captured_vars_meta") = "captured_vars_meta"
    srcForHist("injected_vars") = "injected_vars"
    
    ' Mapas de headers -> coluna
    Dim mapS As Object, mapH As Object
    Set mapS = HeaderMap_ByName(wsS)
    Set mapH = HeaderMap_ByName(wsH)
    
    ' Garantir headers no HISTÓRICO (e, opcionalmente, no Seguimento para as 3 novas colunas)
    EnsureHeader wsS, mapS, "captured_vars", AUTO_CREATE_MISSING_HEADERS
    EnsureHeader wsS, mapS, "captured_vars_meta", AUTO_CREATE_MISSING_HEADERS
    EnsureHeader wsS, mapS, "injected_vars", AUTO_CREATE_MISSING_HEADERS
    Set mapS = HeaderMap_ByName(wsS)
    
    Dim i As Long
    For i = LBound(histHeaders) To UBound(histHeaders)
        EnsureHeader wsH, mapH, CStr(histHeaders(i)), AUTO_CREATE_MISSING_HEADERS
    Next i
    Set mapH = HeaderMap_ByName(wsH)
    
    ' Determinar a última linha com dados no Seguimento (usando colunas chave)
    Dim lastRowS As Long
    lastRowS = LastDataRow_Seguimento(wsS, mapS)
    If lastRowS < 2 Then Exit Sub
    
    Dim nLin As Long
    nLin = lastRowS - 1 ' linhas 2..lastRowS
    
    ' Inserir espaço no topo do HISTÓRICO: nLin + 1 (linha separadora)
    wsH.rowS(HIST_TOP_ROW).Resize(nLin + 1).Insert Shift:=xlDown
    
    ' Copiar linha a linha / coluna a coluna (por headers)
    Dim r As Long, c As Long
    For r = 1 To nLin
        Dim srcRow As Long
        srcRow = r + 1 ' começa na linha 2
        
        For c = LBound(histHeaders) To UBound(histHeaders)
            Dim hDest As String
            hDest = CStr(histHeaders(c))
            
            Dim hSrc As String
            hSrc = CStr(srcForHist(hDest))
            
            Dim colS As Long, colH As Long
            colS = GetCol(mapS, hSrc)
            colH = GetCol(mapH, hDest)
            
            If colH > 0 Then
                If colS > 0 Then
                    wsH.Cells(HIST_TOP_ROW + (r - 1), colH).value = wsS.Cells(srcRow, colS).value
                Else
                    wsH.Cells(HIST_TOP_ROW + (r - 1), colH).value = vbNullString
                    Log_Debug "ALERTA", "Seguimento_ArquivarLimpar", "Header em falta no Seguimento: '" & hSrc & "' (para preencher '" & hDest & "')."
                End If
            End If
        Next c
    Next r
    
    ' Linha separadora preta imediatamente abaixo do bloco novo
    Dim sepRow As Long
    sepRow = HIST_TOP_ROW + nLin
    
    With wsH.rowS(sepRow)
        .RowHeight = SEPARATOR_HEIGHT
        .Interior.Color = vbBlack
        .Font.Color = vbWhite
    End With
    
    ' Evitar wrap no bloco novo (mantém legibilidade)
    Dim firstDataRowH As Long, lastDataRowH As Long
    firstDataRowH = HIST_TOP_ROW
    lastDataRowH = HIST_TOP_ROW + nLin - 1
    
    For c = LBound(histHeaders) To UBound(histHeaders)
        Dim colHH As Long
        colHH = GetCol(mapH, CStr(histHeaders(c)))
        If colHH > 0 Then
            wsH.Range(wsH.Cells(firstDataRowH, colHH), wsH.Cells(lastDataRowH, colHH)).WrapText = False
        End If
    Next c
    
    ' Limpar Seguimento (dados + comentários) e AutoFit (como no original)
    Dim lastColS As Long
    lastColS = wsS.Cells(1, wsS.Columns.Count).End(xlToLeft).Column
    If lastColS < 1 Then lastColS = 1
    
    With wsS.Range(wsS.Cells(2, 1), wsS.Cells(lastRowS, lastColS))
        .ClearContents
        .ClearComments
    End With
    
    wsS.rowS("2:" & CStr(Application.Max(2, lastRowS))).AutoFit
    
    Exit Sub
    
EH:
    Log_Debug "ERRO", "Seguimento_ArquivarLimpar", "Erro: " & Err.Description
    MsgBox "Erro em Seguimento_ArquivarLimpar: " & Err.Description, vbCritical
End Sub

' =============================================================================
' Helpers
' =============================================================================

Private Function HeaderMap_ByName(ByVal ws As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then
        Set HeaderMap_ByName = d
        Exit Function
    End If
    
    Dim c As Long
    For c = 1 To lastCol
        Dim h As String
        h = Trim$(CStr(ws.Cells(1, c).value))
        If h <> "" Then
            Dim k As String
            k = NormalizeKey(h)
            If Not d.exists(k) Then d.Add k, c
        End If
    Next c
    
    Set HeaderMap_ByName = d
End Function

Private Function NormalizeKey(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = RemoveDiacritics(s)
    NormalizeKey = s
End Function

Private Function RemoveDiacritics(ByVal s As String) As String
    Dim a As Variant, b As Variant, i As Long
    a = Array("á", "à", "ã", "â", "ä", "é", "è", "ê", "ë", "í", "ì", "î", "ï", "ó", "ò", "õ", "ô", "ö", "ú", "ù", "û", "ü", "ç", _
              "Á", "À", "Ã", "Â", "Ä", "É", "È", "Ê", "Ë", "Í", "Ì", "Î", "Ï", "Ó", "Ò", "Õ", "Ô", "Ö", "Ú", "Ù", "Û", "Ü", "Ç")
    b = Array("a", "a", "a", "a", "a", "e", "e", "e", "e", "i", "i", "i", "i", "o", "o", "o", "o", "o", "u", "u", "u", "u", "c", _
              "a", "a", "a", "a", "a", "e", "e", "e", "e", "i", "i", "i", "i", "o", "o", "o", "o", "o", "u", "u", "u", "u", "c")
    For i = LBound(a) To UBound(a)
        s = Replace(s, CStr(a(i)), CStr(b(i)))
    Next i
    RemoveDiacritics = s
End Function

Private Function GetCol(ByVal map As Object, ByVal headerName As String) As Long
    Dim k As String
    k = NormalizeKey(headerName)
    If map.exists(k) Then
        GetCol = CLng(map(k))
    Else
        GetCol = 0
    End If
End Function

Private Sub EnsureHeader(ByVal ws As Worksheet, ByVal map As Object, ByVal headerName As String, ByVal autoCreate As Boolean)
    Dim k As String
    k = NormalizeKey(headerName)
    If map.exists(k) Then Exit Sub
    
    If Not autoCreate Then
        Log_Debug "ALERTA", "Seguimento_ArquivarLimpar", "Header em falta (não criado): " & headerName & " na folha " & ws.name
        Exit Sub
    End If
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    
    ws.Cells(1, lastCol + 1).value = headerName
End Sub

Private Function LastDataRow_Seguimento(ByVal wsS As Worksheet, ByVal mapS As Object) As Long
    Dim colPasso As Long, colTs As Long, colPid As Long
    colPasso = GetCol(mapS, "Passo")
    colTs = GetCol(mapS, "Timestamp")
    colPid = GetCol(mapS, "Prompt ID")
    
    Dim r1 As Long, r2 As Long, r3 As Long
    r1 = 0: r2 = 0: r3 = 0
    
    If colPasso > 0 Then r1 = wsS.Cells(wsS.rowS.Count, colPasso).End(xlUp).Row
    If colTs > 0 Then r2 = wsS.Cells(wsS.rowS.Count, colTs).End(xlUp).Row
    If colPid > 0 Then r3 = wsS.Cells(wsS.rowS.Count, colPid).End(xlUp).Row
    
    LastDataRow_Seguimento = Application.Max(r1, r2, r3)
End Function

Private Sub Log_Debug(ByVal severidade As String, ByVal parametro As String, ByVal problema As String)
    ' Tenta registar no DEBUG via Debug_Registar (se existir); fallback silencioso.
    On Error Resume Next
    Call Debug_Registar(0, parametro, severidade, "", parametro, problema, "")
    On Error GoTo 0
End Sub


' ============================================================
' Seguimento - escrita segura (sem truncagem)
'   - SAFE_LIMIT lido de Config (SEGUIMENTO_SAFE_LIMIT), default 32000
'   - se exceder: guarda output completo em disco e divide por múltiplas linhas
' ============================================================

Private Function Seguimento_GetSafeLimit() As Long
    On Error GoTo Falha
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Worksheets("Config")
    Dim lr As Long: lr = wsCfg.Cells(wsCfg.rowS.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 1 To lr
        If Trim$(CStr(wsCfg.Cells(r, 1).value)) = "SEGUIMENTO_SAFE_LIMIT" Then
            Seguimento_GetSafeLimit = CLng(val(wsCfg.Cells(r, 2).value))
            If Seguimento_GetSafeLimit <= 0 Then Seguimento_GetSafeLimit = 32000
            Exit Function
        End If
    Next r
Falha:
    Seguimento_GetSafeLimit = 32000
End Function

Private Function Seguimento_LerOutputFolderBase() As String
    On Error GoTo Falha
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets("PAINEL")
    ' Por compatibilidade: OUTPUT Folder está em B3 no layout observado
    Seguimento_LerOutputFolderBase = Trim$(CStr(wsP.Range("B3").value))
    If Right$(Seguimento_LerOutputFolderBase, 1) = "\" Then
        Seguimento_LerOutputFolderBase = Left$(Seguimento_LerOutputFolderBase, Len(Seguimento_LerOutputFolderBase) - 1)
    End If
    Exit Function
Falha:
    Seguimento_LerOutputFolderBase = ""
End Function

Private Sub Seguimento_EscreverTextoUTF8(ByVal caminho As String, ByVal conteudo As String)
    On Error GoTo Falha
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText conteudo
    st.Position = 0
    st.SaveToFile caminho, 2 ' overwrite
    st.Close
    Exit Sub
Falha:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Sub

Private Sub Seguimento_EscreverOutputSemTruncagem(ByVal ws As Worksheet, ByVal linhaBase As Long, ByVal mapa As Object, ByVal txtCompleto As String)
    On Error GoTo Falha

    If Not mapa.exists("Output (texto)") Then Exit Sub

    Dim safeLimit As Long
    safeLimit = Seguimento_GetSafeLimit()
    If safeLimit <= 1000 Then safeLimit = 32000

    Dim colOut As Long
    colOut = CLng(mapa("Output (texto)"))

    If Len(txtCompleto) <= safeLimit Then
        ws.Cells(linhaBase, colOut).value = txtCompleto
        Exit Sub
    End If

    ' 1) Guardar output completo em disco
    Dim baseOut As String
    baseOut = Seguimento_LerOutputFolderBase()
    Dim pasta As String
    If baseOut <> "" Then
        pasta = baseOut & "\_raw_overflow"
    Else
        pasta = ThisWorkbook.path & "\_raw_overflow"
    End If
    On Error Resume Next
    If Dir(pasta, vbDirectory) = "" Then MkDir pasta
    On Error GoTo Falha

    Dim nomeF As String
    nomeF = "Seguimento_Output_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    Dim caminhoFull As String
    caminhoFull = pasta & "\" & nomeF
    Call Seguimento_EscreverTextoUTF8(caminhoFull, txtCompleto)

    ' 2) Dividir em N segmentos
    Dim chunkBase As Long
    chunkBase = safeLimit - 80
    If chunkBase < 1000 Then chunkBase = safeLimit - 20
    Dim n As Long
    n = CLng((Len(txtCompleto) + chunkBase - 1) / chunkBase)
    If n < 1 Then n = 1

    ' Inserir linhas extra
    If n > 1 Then
        ws.rowS(linhaBase + 1).Resize(n - 1).Insert Shift:=xlDown
    End If

    Dim i As Long, startPos As Long
    startPos = 1

    For i = 1 To n
        Dim prefixo As String
        prefixo = "(CONTINUAÇÃO " & CStr(i) & "/" & CStr(n) & ") "

        Dim cabecalho As String
        cabecalho = ""
        If i = 1 Then
            cabecalho = "[OUTPUT COMPLETO EM DISCO] " & caminhoFull & vbCrLf
        End If

        Dim maxLen As Long
        maxLen = safeLimit - Len(prefixo) - Len(cabecalho)
        If maxLen < 1 Then maxLen = 1

        Dim parte As String
        parte = Mid$(txtCompleto, startPos, maxLen)
        startPos = startPos + Len(parte)

        ws.Cells(linhaBase + (i - 1), colOut).value = prefixo & cabecalho & parte
    Next i

    Exit Sub
Falha:
    ' fallback: não truncar - escreve pelo menos a indicação de erro
    On Error Resume Next
    ws.Cells(linhaBase, colOut).value = "[ERRO] Seguimento_EscreverOutputSemTruncagem falhou."
End Sub
