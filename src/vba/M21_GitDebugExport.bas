Attribute VB_Name = "M21_GitDebugExport"
Option Explicit

' =============================================================================
' Modulo: M21_GitDebugExport
' Proposito:
' - Manter facade/entrypoint de compatibilidade para exportacao Git debug.
' - Gerar artefactos DEBUG/Seguimento/catalogo/painel para upload.
' - Delegar configuracao, HTTP, blobs, tree/commit e logging aos modulos GH dedicados.
'
' Atualizacoes:
' - 2026-03-11 | Codex | Normaliza filtro de pipeline no Seguimento para export por prompt
'   - Evita queda indevida para fallback da primeira prompt quando `pipeline_name` no Seguimento tem espacos invisiveis/variacoes.
'   - Mantem deteccao de Prompt IDs por ordem de execucao e melhora rastreabilidade em runs multi-step.
' - 2026-03-11 | Codex | Export Git com pasta propria por prompt executada
'   - Passa a gerar/upload para cada Prompt ID observado no Seguimento da pipeline (mantendo run stamp comum).
'   - Mantem fallback para o primeiro Prompt ID do PAINEL quando o Seguimento nao tem linhas elegiveis.
' - 2026-03-10 | Codex | Inclui dumps raw request/response no Git LOG da run
'   - Agrega ficheiros JSON/TXT produzidos pelo M05 em subpasta payload_dumps da pasta remota da run.
'   - Regista INFO/ALERTA no DEBUG quando adiciona (ou falha a adicionar) dumps ao pacote de upload.
' - 2026-03-09 | Codex | Remove prefixo do indice no PROMPT_NAME do run folder Git
'   - Altera derivacao para <ordem>_<nomeCurto> (ex.: 01_WF_PROMPT_AUDIT), sem prefixo do slot da pipeline.
'   - Mantem fallback por nome da pipeline apenas para resolver Prompt ID, sem impactar o nome final da prompt.
' - 2026-03-09 | Codex | Tolera linhas vazias/comentarios na lista INICIAR ao derivar Prompt ID
'   - Deixa de parar na primeira linha vazia e passa a varrer a janela completa da lista para encontrar o primeiro Prompt ID valido.
'   - Ignora entradas nao-ID (sem formato <folha>/<ordem>/<nome>/<versao>) para reduzir quedas indevidas em PROMPT_DESCONHECIDO.
' - 2026-03-09 | Codex | Endurece resolucao de Prompt ID no PAINEL para pasta remota Git
'   - Resolve coluna/slot da pipeline com normalizacao de texto (CR/LF/TAB/NBSP/espacos) e extrai prompt mesmo com variacoes de nome.
'   - Usa pipelineIndex efetivo apenas para resolver coluna da pipeline no PAINEL e regista alerta diagnostico quando nao encontra Prompt ID.
' - 2026-03-08 | Codex | Corrige derivacao de PROMPT_NAME/VERSION para pasta remota Git
'   - Passa a montar PROMPT_NAME no formato <ordem>_<nomeCurto> (ex.: 01_WF_PROMPT_AUDIT).
'   - Adiciona fallback por nome da pipeline para resolver o primeiro Prompt ID quando o indice nao estiver disponivel.
' - 2026-03-09 | Codex | Tolera linhas vazias/comentarios na lista INICIAR ao derivar Prompt ID
'   - Deixa de parar na primeira linha vazia e passa a varrer a janela completa da lista para encontrar o primeiro Prompt ID valido.
'   - Ignora entradas nao-ID (sem formato <folha>/<ordem>/<nome>/<versao>) para reduzir quedas indevidas em PROMPT_DESCONHECIDO.
' - 2026-03-09 | Codex | Corrige erro de compilacao em BuildExecutedCatalogCsv
'   - Declara colecao `orderedIds` e usa dictionary local `seen` para deduplicar Prompt IDs sem `Variable not defined`.
'   - Fecha a funcao com `End Function` e preserva ordem de execucao ao montar `catalogo_prompts_executadas.csv`.
' - 2026-03-09 | Codex | Endurece resolucao de Prompt ID no PAINEL para pasta remota Git
'   - Resolve coluna/slot da pipeline com normalizacao de texto (CR/LF/TAB/NBSP/espacos) e extrai prompt mesmo com variacoes de nome.
'   - Passa a usar pipelineIndex efetivo (resolvido pelo PAINEL) na derivacao de PROMPT_NAME e regista alerta diagnostico quando nao encontra Prompt ID.
' - 2026-03-08 | Codex | Corrige derivacao de PROMPT_NAME/VERSION para pasta remota Git
'   - Passa a montar PROMPT_NAME no formato <pipelineIndex><ordem>_<nomeCurto> (ex.: 701_WF_PROMPT_AUDIT).
'   - Adiciona fallback por nome da pipeline para resolver o primeiro Prompt ID quando o indice nao estiver disponivel.
' - 2026-03-09 | Codex | Corrige deteccao de colunas no Seguimento para export do catalogo
'   - Remove dependencia de helper inexistente (`HeaderColByName`) e usa mapa de cabecalhos local com aliases robustos.
'   - Garante fallback para devolver sempre cabecalho CSV (evita ficheiro vazio/BOM-only no GitHub).
' - 2026-03-09 | Codex | Corrige conflito de merge e helper de cabecalho no export Git
'   - Remove marcadores `<<<<<<< ======= >>>>>>>` remanescentes no modulo para restaurar compilacao.
'   - Reintroduz helper local `HeaderColByName` usando `HeaderMap/MapGet` para evitar `Sub or Function not defined`.
' - 2026-03-09 | Codex | Publica DEBUG.csv final apos GH_UPLOAD_DONE para reduzir drift
'   - Reenvia apenas DEBUG.csv no fim da rotina para aproximar o artefacto remoto ao estado final da folha DEBUG.
'   - Mantem o upload principal inalterado e trata falha da republicacao final como ALERTA nao bloqueante.
' - 2026-03-08 | Codex | Alinha export do catalogo ao layout fisico de blocos (5 linhas por prompt)
'   - Gera `catalogo_prompts_executadas.csv` com colunas A:K (cabecalho do catalogo) e 5 linhas por prompt (ID + Next/default/allowed + linha em branco).
'   - Preserva a estrutura visual do catalogo para auditoria 1:1 entre Excel e CSV exportado.
' - 2026-03-08 | Codex | Corrige derivacao de PROMPT_NAME/VERSION para pasta remota Git
'   - Passa a montar PROMPT_NAME no formato <pipelineIndex><ordem>_<nomeCurto> (ex.: 701_WF_PROMPT_AUDIT).
'   - Adiciona fallback por nome da pipeline para resolver o primeiro Prompt ID quando o indice nao estiver disponivel.
' - 2026-03-08 | Codex | Alinha export do catalogo ao layout fisico de blocos (5 linhas por prompt)
'   - Gera `catalogo_prompts_executadas.csv` com colunas A:K (cabecalho do catalogo) e 5 linhas por prompt (ID + Next/default/allowed + linha em branco).
'   - Preserva a estrutura visual do catalogo para auditoria 1:1 entre Excel e CSV exportado.
' - 2026-03-08 | Codex | Corrige export do catalogo para refletir bloco completo do prompt
'   - Substitui CSV reduzido (7 colunas) por export completo com colunas A:K e campos Next/INPUTS/OUTPUTS.
'   - Faz lookup robusto da linha do prompt por ID com normalizacao (CR/LF/TAB/NBSP) para evitar falhas por caracteres invisiveis.
' - 2026-03-08 | Codex | Ajusta template default da pasta GitHub para hierarquia pipeline/prompt/versao
'   - Define default de run folder como {{PIPELINE_NAME}}/{{PROMPT_NAME}}/{{VERSION}}/{{RUN_STAMP}}.
'   - Extrai prompt/version a partir do primeiro Prompt ID da lista da pipeline no PAINEL (com fallback seguro).
'   - Adiciona placeholder {{YYYY-MM-DD HHDD}} por compatibilidade com templates legados de timestamp.
' - 2026-03-08 | Codex | Alerta explicito quando GH_API_VERSION e normalizado
'   - Emite GH_CONFIG (ALERTA) quando valor em Config nao estiver no formato canonico yyyy-mm-dd.
'   - Mostra raw/normalizado no detalhe para diagnostico rapido sem interromper o fluxo.
' - 2026-03-08 | Codex | Enriquecimento de contexto em GH_MODE_SELECTED
'   - Inclui owner/repo/branch/remote_folder/total_files/token_source para auditoria imediata no DEBUG.
' - 2026-03-08 | Codex | Ativa dispatch por GH_UPLOAD_MODE no runtime principal
'   - Adiciona selecao explicita de modo (tree_commit|contents_api) com default e erro para valor invalido.
'   - Regista eventos de inicio/fim com resumo (sucessos/falhas/retries) sem expor segredos.
' - 2026-03-07 | Codex | Corrige erro de compilacao por dependencia JsonPick ausente
'   - Adiciona helper local `JsonPick` para extrair valores string de chaves top-level em respostas GitHub.
'   - Elimina `Compile error: Sub or Function not defined` reportado no VBE durante compile.
' - 2026-03-07 | Codex | Corrige erro de compilacao por helper de config inexistente
'   - Reintroduz wrapper local `GitCfg_Get` como compatibilidade para call-sites legados do modulo.
'   - Evita `Compile error: Sub or Function not defined` quando o projeto referencia `GitCfg_Get`.
' - 2026-03-07 | Codex | Auditoria da origem do token e ficheiros enviados
'   - Regista no DEBUG a fonte de resolucao do token GitHub em cada execucao de export.
'   - Regista path remoto e nome de cada ficheiro preparado para upload no run.
' - 2026-03-05 | Codex | Pasta remota em logs com template configuravel por run
'   - Passa a compor pasta alvo com GH_BASE_PATH/GH_LOG_FOLDER e nome por template (com fallback retrocompativel).
'   - Suporta placeholders {{YYYY}}, {{MM}}, {{DD}}, {{HHMM}} e {{PIPELINE_NAME}} para naming do run.
' - 2026-03-04 | Codex | Endurece instalacao de parametros GH_* na folha Config
'   - Garante cabecalhos Key/Value/Explicacao/Default/Valores na linha 8 e dados apenas em linhas >= 9.
'   - Mantem politica de overwrite seletivo em B:E e regista falhas no DEBUG com codigo estavel.
' - 2026-03-04 | Codex | Macro de instalacao guiada dos parametros GH_* no Config
'   - Adiciona rotina para criar/atualizar chaves GH_* com default e explicacoes.
'
' Funcoes e procedimentos:
' - GitDebug_BuildFilesForUpload(pipelineIndex As Long, pipelineNome As String, remoteFolders As Collection, cfg As Object) As Collection (Private Function)
'   - Prepara artefactos CSV/TXT e agrega dumps raw request/response para upload Git em cada pasta de prompt.
' - GitDebug_BuildRemoteFoldersForPipeline(cfg As Object, pipelineNome As String, pipelineIndex As Long, runStampHhdd As String) As Collection (Private Function)
'   - Resolve uma pasta remota por prompt executada, com fallback seguro para o primeiro Prompt ID do PAINEL.
' - GitDebug_CollectExecutedPromptIds(wsSeg As Worksheet, pipelineNome As String) As Collection (Private Function)
'   - Recolhe Prompt IDs unicos da pipeline no Seguimento, preservando ordem de execucao.
' - PipelineGitDebug_ExportIfEnabled(pipelineIndex As Long, pipelineNome As String, painelAutoSave As String)
'   - Entry point chamado no fim da pipeline para export opcional de debug para GitHub.
' - GitDebug_Config_InstalarParametros(Optional sobrescreverValores As Boolean = False)
'   - Preenche/atualiza chaves GH_* na folha Config sem quebra de retrocompatibilidade.
' - GitDebug_Config_InstalarMinimos()
'   - Macro rapida para instalar parametros minimos GH_* com defaults e explicacao para leigos.
' - BuildExecutedCatalogCsv(wsSeg As Worksheet, pipelineNome As String) As String (Private Function)
'   - Exporta para CSV os blocos completos (5 linhas) de cada prompt executada na pipeline.
' - SheetToCsv(ws As Worksheet, Optional includeRowNumber As Boolean = False) As String (Private Function)
'   - Converte folha para CSV; opcionalmente prefixa `row_number` com o indice da linha.
' - GitDebug_LogFilesForUpload(pipelineNome As String, remoteFolder As String, files As Collection) (Private Sub)
'   - Regista path remoto e nome dos ficheiros preparados para upload GitHub.
' - GitDebug_RunUploadByMode(cfg As Object, files As Collection, pipelineNome As String, uploadMode As String, reason As String, successCount As Long, failCount As Long, retryCount As Long) As Boolean
'   - Faz dispatch operacional entre tree_commit e contents_api com rastreabilidade.
' - GitDebug_NormalizeApiVersionForDiag(rawValue As String) As String (Private Function)
'   - Normaliza GH_API_VERSION para diagnostico/log sem alterar compatibilidade do Config.
' - GitDebug_BuildRunFolder(cfg As Object, pipelineNome As String, pipelineIndex As Long) As String (Private Function)
'   - Resolve pasta remota por run e aplica estrutura canonica obrigatoria pipeline/prompt/versao/data.
' - GitDebug_FirstPromptIdFromPainel(pipelineIndex As Long, pipelineNome As String, resolvedPipelineIndex As Long) As String (Private Function)
'   - Resolve o primeiro Prompt ID ativo e devolve pipelineIndex efetivo por indice ou fallback por nome no PAINEL.
' - GitDebug_IsLikelyPromptId(promptId As String) As Boolean (Private Function)
'   - Valida formato minimo de Prompt ID para ignorar linhas decorativas/comentarios na lista do PAINEL.
' - JsonPick(body As String, keyName As String) As String (Private Function)
'   - Extrai valor string de chave JSON simples para compatibilidade de parsing em M21.
' =============================================================================

Private Const SHEET_DEBUG As String = "DEBUG"
Private Const SHEET_SEGUIMENTO As String = "Seguimento"
Private Const SHEET_HIST As String = "HISTORICO"
Private Const GH_CONFIG_HEADER_ROW As Long = 8
Private Const GH_CONFIG_FIRST_DATA_ROW As Long = 9

Public Sub PipelineGitDebug_ExportIfEnabled(ByVal pipelineIndex As Long, ByVal pipelineNome As String, ByVal painelAutoSave As String)
    On Error GoTo EH

    Dim cfg As Object
    Set cfg = GH_Config_Load(painelAutoSave)

    If Not GH_Config_GetBoolean(cfg, "enabled", False) Then Exit Sub

    Dim reason As String
    If Not GH_Config_Validate(cfg, reason) Then
        Call GH_LogWarn(0, pipelineNome, GH_EVT_CONFIG, "Configuracao GitHub invalida.", reason)
        Exit Sub
    End If

    Call GH_LogInfo(0, pipelineNome, GH_EVT_CONFIG, "Fonte do token GitHub resolvida.", "token_source=" & GH_Config_GetString(cfg, "token_source", "desconhecida"))

    Dim apiVersionRaw As String
    Dim apiVersionNormalized As String
    apiVersionRaw = GH_Config_GetString(cfg, "api_version", "")
    apiVersionNormalized = GitDebug_NormalizeApiVersionForDiag(apiVersionRaw)
    If Trim$(apiVersionRaw) <> "" And StrComp(Trim$(apiVersionRaw), apiVersionNormalized, vbTextCompare) <> 0 Then
        Call GH_LogWarn(0, pipelineNome, GH_EVT_CONFIG, "GH_API_VERSION normalizado para formato canonico.", "raw=" & Trim$(apiVersionRaw) & " | normalized=" & apiVersionNormalized & " | action=Preferir YYYY-MM-DD na Config.")
    End If

    Dim runStampHhdd As String
    runStampHhdd = Format$(Now, "yyyy-mm-dd") & " " & Format$(Now, "hhdd")

    Dim remoteFolders As Collection
    Set remoteFolders = GitDebug_BuildRemoteFoldersForPipeline(cfg, pipelineNome, pipelineIndex, runStampHhdd)
    If remoteFolders Is Nothing Or remoteFolders.Count = 0 Then
        Call GH_LogError(0, pipelineNome, GH_EVT_CONFIG, "Falha a resolver remote_folder para upload.", "[ACTION] Confirme Prompt IDs executados e GH_BASE_PATH/GH_LOG_FOLDER.")
        Exit Sub
    End If

    Dim primaryRemoteFolder As String
    primaryRemoteFolder = CStr(remoteFolders(1))
    Call GH_LogInfo(0, pipelineNome, GH_EVT_CONFIG, "Pasta remota principal resolvida.", "remote_folder=" & primaryRemoteFolder & " | total_prompt_folders=" & CStr(remoteFolders.Count))

    Dim files As Collection
    Set files = GitDebug_BuildFilesForUpload(pipelineIndex, pipelineNome, remoteFolders, cfg)
    If files Is Nothing Or files.Count = 0 Then Exit Sub
    Call GitDebug_LogFilesForUpload(pipelineNome, primaryRemoteFolder, files)

    Dim uploadModeReason As String
    Dim uploadModeDefaulted As Boolean
    Dim uploadMode As String
    uploadMode = GH_Config_ResolveUploadMode(cfg, uploadModeReason, uploadModeDefaulted)
    If uploadMode = "" Then
        Call GH_LogError(0, pipelineNome, GH_EVT_UPLOAD_MODE_INVALID, "Modo de upload invalido.", uploadModeReason)
        Call GH_LogError(0, pipelineNome, GH_EVT_UPLOAD_FAILED, "Falha no auto-upload de debug.", "upload_mode_invalido")
        Exit Sub
    End If

    If uploadModeDefaulted Then
        Call GH_LogWarn(0, pipelineNome, GH_EVT_UPLOAD_MODE_DEFAULTED, "GH_UPLOAD_MODE vazio; aplicado default.", "upload_mode=tree_commit")
    End If

    Call GH_LogInfo(0, pipelineNome, GH_EVT_UPLOAD_START, "Inicio do upload GitHub.", "upload_mode=" & uploadMode & " | remote_folder=" & primaryRemoteFolder & " | total_prompt_folders=" & CStr(remoteFolders.Count) & " | total_files=" & CStr(files.Count) & " | token_source=" & GH_Config_GetString(cfg, "token_source", "desconhecida"))
    Call GH_LogInfo(0, pipelineNome, GH_EVT_MODE_SELECTED, "Modo de upload selecionado.", "upload_mode=" & uploadMode & " | owner=" & GH_Config_GetString(cfg, "owner") & " | repo=" & GH_Config_GetString(cfg, "repo") & " | branch=" & GH_Config_GetString(cfg, "branch") & " | remote_folder=" & primaryRemoteFolder & " | total_prompt_folders=" & CStr(remoteFolders.Count) & " | total_files=" & CStr(files.Count) & " | token_source=" & GH_Config_GetString(cfg, "token_source", "desconhecida"))

    Dim successCount As Long
    Dim failCount As Long
    Dim retryCount As Long
    If Not GitDebug_RunUploadByMode(cfg, files, pipelineNome, uploadMode, reason, successCount, failCount, retryCount) Then
        Call GH_LogError(0, pipelineNome, GH_EVT_UPLOAD_FAILED, "Falha no auto-upload de debug.", reason & " | upload_mode=" & uploadMode & " | success=" & CStr(successCount) & " | fail=" & CStr(failCount) & " | retries=" & CStr(retryCount))
        Exit Sub
    End If

    Dim webUrl As String
    webUrl = GH_TreeCommit_BuildWebFolderUrl(cfg, primaryRemoteFolder)

    Call GitDebug_WriteLinkToSeguimento(pipelineNome, webUrl)
    Call GitDebug_WriteLinkToHistorico(pipelineNome, webUrl)
    Call GH_LogInfo(0, pipelineNome, GH_EVT_CONFIG, "Link registado em Seguimento/HISTORICO.", webUrl)

    Call GH_LogInfo(0, pipelineNome, GH_EVT_UPLOAD_DONE, "Debug export publicado no GitHub.", "upload_mode=" & uploadMode & " | success=" & CStr(successCount) & " | fail=" & CStr(failCount) & " | retries=" & CStr(retryCount) & " | " & webUrl)

    Dim finalRefreshReason As String
    If Not GitDebug_RefreshDebugCsvFinal(cfg, pipelineNome, remoteFolders, uploadMode, finalRefreshReason) Then
        Call GH_LogWarn(0, pipelineNome, GH_EVT_UPLOAD, "Falha ao republicar DEBUG.csv final no GitHub.", finalRefreshReason)
    End If
    Exit Sub

EH:
    Call GH_LogError(0, pipelineNome, GH_EVT_UPLOAD, "Falha no auto-upload de debug: " & Err.Description, "Validar parametros GH_* e conectividade com api.github.com.")
End Sub

Private Function GitDebug_RefreshDebugCsvFinal( _
    ByVal cfg As Object, _
    ByVal pipelineNome As String, _
    ByVal remoteFolders As Collection, _
    ByVal uploadMode As String, _
    ByRef reason As String) As Boolean

    On Error GoTo EH

    reason = ""

    Dim wsDebug As Worksheet
    Set wsDebug = ThisWorkbook.Worksheets(SHEET_DEBUG)

    Dim csvDebugFinal As String
    csvDebugFinal = SheetToCsv(wsDebug, True)

    Dim files As New Collection

    Dim idx As Long
    For idx = 1 To remoteFolders.Count
        files.Add GitFileItem(CStr(remoteFolders(idx)) & "/DEBUG.csv", csvDebugFinal)
    Next idx

    Dim successCount As Long
    Dim failCount As Long
    Dim retryCount As Long

    GitDebug_RefreshDebugCsvFinal = GitDebug_RunUploadByMode(cfg, files, pipelineNome, uploadMode, reason, successCount, failCount, retryCount)

    If Not GitDebug_RefreshDebugCsvFinal Then
        reason = Trim$(reason & " | upload_mode=" & uploadMode & " | success=" & CStr(successCount) & " | fail=" & CStr(failCount) & " | retries=" & CStr(retryCount))
    End If

    Exit Function

EH:
    reason = "err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 180)
    GitDebug_RefreshDebugCsvFinal = False
End Function

Private Function GitDebug_NormalizeApiVersionForDiag(ByVal rawValue As String) As String
    Dim valueText As String
    valueText = Trim$(rawValue)

    If valueText = "" Then
        GitDebug_NormalizeApiVersionForDiag = "2022-11-28"
        Exit Function
    End If

    If valueText Like "####-##-##" Then
        GitDebug_NormalizeApiVersionForDiag = valueText
        Exit Function
    End If

    If valueText Like "##/##/####" Then
        GitDebug_NormalizeApiVersionForDiag = Right$(valueText, 4) & "-" & Mid$(valueText, 4, 2) & "-" & Left$(valueText, 2)
        Exit Function
    End If

    GitDebug_NormalizeApiVersionForDiag = "2022-11-28"
End Function

Private Function GitDebug_BuildRunFolder(ByVal cfg As Object, ByVal pipelineNome As String, ByVal pipelineIndex As Long) As String
    Dim ignored As String
    ignored = GH_Config_GetString(cfg, "run_folder_template", "")

    Dim effectivePipelineIndex As Long
    Dim firstPromptId As String
    firstPromptId = GitDebug_FirstPromptIdFromPainel(pipelineIndex, pipelineNome, effectivePipelineIndex)

    If Trim$(firstPromptId) = "" Then
        Call GH_LogWarn(0, pipelineNome, GH_EVT_CONFIG, "Nao foi possivel derivar Prompt ID para run_folder.", "pipeline_index_in=" & CStr(pipelineIndex) & " | pipeline_index_effective=" & CStr(effectivePipelineIndex) & " | action=Verifique PAINEL (linha 1 nome; lista INICIAR a partir da linha 9).")
    End If

    Dim safePipeline As String
    safePipeline = GitDebug_SanitizePathPart(pipelineNome)

    Dim safePromptName As String
    safePromptName = GitDebug_SanitizePathPart(GitDebug_PromptNameFromId(firstPromptId, effectivePipelineIndex))

    Dim safeVersion As String
    safeVersion = GitDebug_SanitizePathPart(GitDebug_PromptVersionFromId(firstPromptId))

    Dim runStampHhdd As String
    runStampHhdd = Format$(Now, "yyyy-mm-dd") & " " & Format$(Now, "hhdd")

    GitDebug_BuildRunFolder = GitDebug_SanitizePathTemplate( _
        safePipeline & "/" & safePromptName & "/" & safeVersion & "/" & runStampHhdd)
End Function

Private Function GitDebug_BuildRunFolderFromPromptId(ByVal pipelineNome As String, ByVal promptId As String, ByVal runStampHhdd As String) As String
    Dim safePipeline As String
    safePipeline = GitDebug_SanitizePathPart(pipelineNome)

    Dim safePromptName As String
    safePromptName = GitDebug_SanitizePathPart(GitDebug_PromptNameFromId(promptId, 0))

    Dim safeVersion As String
    safeVersion = GitDebug_SanitizePathPart(GitDebug_PromptVersionFromId(promptId))

    GitDebug_BuildRunFolderFromPromptId = GitDebug_SanitizePathTemplate( _
        safePipeline & "/" & safePromptName & "/" & safeVersion & "/" & runStampHhdd)
End Function

Private Function GitDebug_BuildRemoteFoldersForPipeline(ByVal cfg As Object, ByVal pipelineNome As String, ByVal pipelineIndex As Long, ByVal runStampHhdd As String) As Collection
    On Error GoTo EH

    Dim wsSeg As Worksheet
    Set wsSeg = ThisWorkbook.Worksheets(SHEET_SEGUIMENTO)

    Dim promptIds As Collection
    Set promptIds = GitDebug_CollectExecutedPromptIds(wsSeg, pipelineNome)

    If promptIds.Count = 0 Then
        Dim resolvedPipelineIndex As Long
        Dim firstPromptId As String
        firstPromptId = GitDebug_FirstPromptIdFromPainel(pipelineIndex, pipelineNome, resolvedPipelineIndex)
        If Trim$(firstPromptId) <> "" Then
            promptIds.Add firstPromptId
            Call GH_LogWarn(0, pipelineNome, GH_EVT_CONFIG, "Seguimento sem Prompt IDs elegiveis para export Git.", "fallback_prompt_id=" & firstPromptId & " | action=validar se Seguimento regista Prompt ID em todas as execucoes.")
        End If
    End If

    Dim folders As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To promptIds.Count
        Dim promptId As String
        promptId = Trim$(CStr(promptIds(i)))
        If promptId <> "" Then
            Dim runFolder As String
            runFolder = GitDebug_BuildRunFolderFromPromptId(pipelineNome, promptId, runStampHhdd)

            Dim remoteFolder As String
            remoteFolder = GitDebug_BuildRemoteFolder(cfg, runFolder)

            If Trim$(remoteFolder) <> "" Then
                If Not seen.Exists(remoteFolder) Then
                    seen(remoteFolder) = True
                    folders.Add remoteFolder
                    Call GH_LogInfo(0, pipelineNome, GH_EVT_CONFIG, "Run folder resolvida por prompt.", "prompt_id=" & promptId & " | remote_folder=" & remoteFolder)
                End If
            End If
        End If
    Next i

    Set GitDebug_BuildRemoteFoldersForPipeline = folders
    Exit Function
EH:
    Set GitDebug_BuildRemoteFoldersForPipeline = Nothing
End Function

Private Function GitDebug_SanitizePathTemplate(ByVal templatePath As String) As String
    Dim normalized As String
    normalized = Replace(Trim$(templatePath), "\", "/")

    Dim parts() As String
    parts = Split(normalized, "/")

    Dim i As Long
    Dim out As String
    For i = LBound(parts) To UBound(parts)
        Dim part As String
        part = GitDebug_SanitizePathPart(parts(i))
        If part <> "" Then
            If out <> "" Then out = out & "/"
            out = out & part
        End If
    Next i

    GitDebug_SanitizePathTemplate = out
End Function

Private Function GitDebug_FirstPromptIdFromPainel(ByVal pipelineIndex As Long, ByVal pipelineNome As String, ByRef resolvedPipelineIndex As Long) As String
    On Error GoTo EH

    Const LIST_START_ROW As Long = 9

    Dim wsPainel As Worksheet
    Set wsPainel = ThisWorkbook.Worksheets("PAINEL")

    Dim colIniciar As Long
    colIniciar = GitDebug_ResolvePainelStartColumn(wsPainel, pipelineIndex, pipelineNome, resolvedPipelineIndex)
    If colIniciar = 0 Then Exit Function

    Dim r As Long
    For r = LIST_START_ROW To LIST_START_ROW + 400
        Dim promptId As String
        promptId = Trim$(CStr(wsPainel.Cells(r, colIniciar).Value))

        If promptId <> "" Then
            If UCase$(promptId) <> "STOP" Then
                If GitDebug_IsLikelyPromptId(promptId) Then
                    GitDebug_FirstPromptIdFromPainel = promptId
                    Exit Function
                End If
            End If
        End If
    Next r

    Exit Function
EH:
    GitDebug_FirstPromptIdFromPainel = ""
End Function

Private Function GitDebug_ResolvePainelStartColumn(ByVal wsPainel As Worksheet, ByVal pipelineIndex As Long, ByVal pipelineNome As String, ByRef resolvedPipelineIndex As Long) As Long
    resolvedPipelineIndex = 0

    If pipelineIndex >= 1 And pipelineIndex <= 10 Then
        resolvedPipelineIndex = pipelineIndex
        GitDebug_ResolvePainelStartColumn = 2 + (pipelineIndex - 1) * 2
        Exit Function
    End If

    Dim targetName As String
    targetName = GitDebug_NormalizeCompareToken(pipelineNome)
    If targetName = "" Then Exit Function

    Dim idx As Long
    For idx = 1 To 10
        Dim candidateCol As Long
        candidateCol = 2 + (idx - 1) * 2
        If GitDebug_NormalizeCompareToken(CStr(wsPainel.Cells(1, candidateCol).Value)) = targetName Then
            resolvedPipelineIndex = idx
            GitDebug_ResolvePainelStartColumn = candidateCol
            Exit Function
        End If
    Next idx
End Function

Private Function GitDebug_NormalizeCompareToken(ByVal rawText As String) As String
    Dim s As String
    s = Trim$(rawText)
    If s = "" Then Exit Function

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW$(160), " ")

    Dim hasDouble As Boolean
    Do
        hasDouble = (InStr(1, s, "  ", vbBinaryCompare) > 0)
        If hasDouble Then s = Replace(s, "  ", " ")
    Loop While hasDouble

    GitDebug_NormalizeCompareToken = UCase$(Trim$(s))
End Function

Private Function GitDebug_IsLikelyPromptId(ByVal promptId As String) As Boolean
    Dim cleaned As String
    cleaned = Trim$(promptId)
    If cleaned = "" Then Exit Function

    Dim parts() As String
    parts = Split(cleaned, "/")
    If UBound(parts) < 3 Then Exit Function

    If Trim$(parts(0)) = "" Then Exit Function
    If Trim$(parts(2)) = "" Then Exit Function
    If Trim$(parts(3)) = "" Then Exit Function

    GitDebug_IsLikelyPromptId = True
End Function

Private Function GitDebug_PromptNameFromId(ByVal promptId As String, ByVal pipelineIndex As Long) As String
    Dim parts() As String
    parts = Split(Trim$(promptId), "/")

    Dim promptLabel As String
    If UBound(parts) >= 2 Then
        promptLabel = Trim$(parts(2))
    End If

    Dim stepCode As String
    If UBound(parts) >= 1 Then
        stepCode = GitDebug_KeepDigitsOnly(Trim$(parts(1)))
    End If

    If stepCode <> "" And promptLabel <> "" Then
        GitDebug_PromptNameFromId = stepCode & "_" & promptLabel
    Else
        GitDebug_PromptNameFromId = promptLabel
    End If

    If GitDebug_PromptNameFromId = "" Then GitDebug_PromptNameFromId = "PROMPT_DESCONHECIDO"
End Function

Private Function GitDebug_KeepDigitsOnly(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch >= "0" And ch <= "9" Then
            GitDebug_KeepDigitsOnly = GitDebug_KeepDigitsOnly & ch
        End If
    Next i
End Function

Private Function GitDebug_PromptVersionFromId(ByVal promptId As String) As String
    Dim parts() As String
    parts = Split(Trim$(promptId), "/")

    If UBound(parts) >= 3 Then
        GitDebug_PromptVersionFromId = Trim$(parts(3))
    End If

    If GitDebug_PromptVersionFromId = "" Then GitDebug_PromptVersionFromId = "VERSAO_DESCONHECIDA"
End Function

Private Function GitDebug_BuildRemoteFolder(ByVal cfg As Object, ByVal ghFolder As String) As String
    Dim cfgBase As String
    cfgBase = Trim$(GH_Config_GetString(cfg, "base_path", "pipeliner_runs"))

    Dim logFolder As String
    logFolder = Trim$(GH_Config_GetString(cfg, "log_folder", "logs"))

    Dim fullPath As String
    fullPath = ""
    If cfgBase <> "" Then fullPath = cfgBase
    If logFolder <> "" Then
        If fullPath <> "" Then
            fullPath = fullPath & "/" & logFolder
        Else
            fullPath = logFolder
        End If
    End If

    If fullPath <> "" Then
        GitDebug_BuildRemoteFolder = fullPath & "/" & ghFolder
    Else
        GitDebug_BuildRemoteFolder = ghFolder
    End If
End Function

Private Function GitDebug_BuildFilesForUpload(ByVal pipelineIndex As Long, ByVal pipelineNome As String, ByVal remoteFolders As Collection, ByVal cfg As Object) As Collection
    On Error GoTo EH

    Dim wsDebug As Worksheet
    Dim wsSeg As Worksheet
    Set wsDebug = ThisWorkbook.Worksheets(SHEET_DEBUG)
    Set wsSeg = ThisWorkbook.Worksheets(SHEET_SEGUIMENTO)

    Dim csvDebug As String
    csvDebug = SheetToCsv(wsDebug, True)

    Dim csvSeg As String
    csvSeg = SheetToCsv(wsSeg)

    Dim csvCatalogo As String
    csvCatalogo = BuildExecutedCatalogCsv(wsSeg, pipelineNome)

    Dim txtPainel As String
    txtPainel = BuildPainelPipelineInfo(pipelineIndex)

    Dim files As New Collection
    Dim folderIdx As Long
    For folderIdx = 1 To remoteFolders.Count
        Dim remoteFolder As String
        remoteFolder = CStr(remoteFolders(folderIdx))

        files.Add GitFileItem(remoteFolder & "/DEBUG.csv", csvDebug)
        files.Add GitFileItem(remoteFolder & "/catalogo_prompts_executadas.csv", csvCatalogo)
        files.Add GitFileItem(remoteFolder & "/Seguimento.csv", csvSeg)
        files.Add GitFileItem(remoteFolder & "/painel_pipeline.txt", txtPainel)
    Next folderIdx

    For folderIdx = 1 To remoteFolders.Count
        Dim dumpFiles As Collection
        Set dumpFiles = M05_ListRunDumpFileItems(CStr(remoteFolders(folderIdx)) & "/payload_dumps", pipelineNome)

        If Not dumpFiles Is Nothing Then
            Dim dumpIdx As Long
            For dumpIdx = 1 To dumpFiles.Count
                files.Add dumpFiles(dumpIdx)
            Next dumpIdx

            If dumpFiles.Count > 0 Then
                Call GH_LogInfo(0, pipelineNome, GH_EVT_UPLOAD, "Dumps raw adicionados ao pacote Git.", "remote_folder=" & CStr(remoteFolders(folderIdx)) & " | payload_dump_files=" & CStr(dumpFiles.Count))
            Else
                Call GH_LogInfo(0, pipelineNome, GH_EVT_UPLOAD, "Sem dumps raw para adicionar ao pacote Git.", "remote_folder=" & CStr(remoteFolders(folderIdx)) & " | payload_dump_files=0")
            End If
        Else
            Call GH_LogWarn(0, pipelineNome, GH_EVT_UPLOAD, "Nao foi possivel listar dumps raw da run.", "remote_folder=" & CStr(remoteFolders(folderIdx)) & " | action=Verifique eventos M05_RUN_DUMP no DEBUG.")
        End If
    Next folderIdx

    Set GitDebug_BuildFilesForUpload = files
    Exit Function

EH:
    Set GitDebug_BuildFilesForUpload = Nothing
End Function

Private Function GitDebug_CollectExecutedPromptIds(ByVal wsSeg As Worksheet, ByVal pipelineNome As String) As Collection
    On Error GoTo EH

    Dim ids As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim map As Object
    Set map = HeaderMap(wsSeg)

    Dim cPipe As Long
    Dim cPid As Long
    cPipe = MapGetFirst(map, Array("Pipeline", "pipeline_name", "Nome do Pipeline"))
    cPid = MapGetFirst(map, Array("Prompt ID", "Prompt Id", "ID Prompt", "prompt_id"))

    If cPipe = 0 Or cPid = 0 Then
        Set GitDebug_CollectExecutedPromptIds = ids
        Exit Function
    End If

    Dim lr As Long
    lr = wsSeg.Cells(wsSeg.Rows.Count, cPipe).End(xlUp).Row

    Dim targetPipelineToken As String
    targetPipelineToken = GitDebug_NormalizeCompareToken(pipelineNome)

    Dim r As Long
    For r = 2 To lr
        If GitDebug_NormalizeCompareToken(CStr(wsSeg.Cells(r, cPipe).Value)) = targetPipelineToken Then
            Dim promptId As String
            promptId = Trim$(CStr(wsSeg.Cells(r, cPid).Value))

            If promptId <> "" And UCase$(promptId) <> "STOP" And GitDebug_IsLikelyPromptId(promptId) Then
                If Not seen.Exists(promptId) Then
                    seen(promptId) = True
                    ids.Add promptId
                End If
            End If
        End If
    Next r

    Set GitDebug_CollectExecutedPromptIds = ids
    Exit Function
EH:
    Set GitDebug_CollectExecutedPromptIds = New Collection
End Function

Private Sub GitDebug_LogFilesForUpload(ByVal pipelineNome As String, ByVal remoteFolder As String, ByVal files As Collection)
    On Error GoTo EH

    Call GH_LogInfo(0, pipelineNome, GH_EVT_UPLOAD, "Ficheiros preparados para upload Git.", "remote_folder=" & remoteFolder & " | total=" & CStr(files.Count))

    Dim i As Long
    For i = 1 To files.Count
        Dim item As Object
        Set item = files(i)
        If Not item Is Nothing Then
            Call GH_LogInfo(0, pipelineNome, GH_EVT_UPLOAD, "Upload file", "path=" & CStr(item("path")))
        End If
    Next i
    Exit Sub
EH:
    Call GH_LogWarn(0, pipelineNome, GH_EVT_UPLOAD, "Falha ao listar ficheiros de upload no DEBUG.", "err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 180))
End Sub

Private Function GitDebug_RunUploadByMode( _
    ByVal cfg As Object, _
    ByVal files As Collection, _
    ByVal pipelineNome As String, _
    ByVal uploadMode As String, _
    ByRef reason As String, _
    ByRef successCount As Long, _
    ByRef failCount As Long, _
    ByRef retryCount As Long) As Boolean

    reason = ""
    successCount = 0
    failCount = 0
    retryCount = 0

    Select Case LCase$(Trim$(uploadMode))
        Case "tree_commit"
            Dim commitSha As String
            If Not GH_TreeCommit_CommitFiles(cfg, files, pipelineNome, commitSha, reason, retryCount) Then
                failCount = files.Count
                Exit Function
            End If
            successCount = files.Count
            GitDebug_RunUploadByMode = True

        Case "contents_api"
            GitDebug_RunUploadByMode = GH_ContentsApi_UploadFiles(cfg, files, pipelineNome, successCount, failCount, retryCount, reason)

        Case Else
            reason = "Modo de upload nao suportado: " & uploadMode
            GitDebug_RunUploadByMode = False
    End Select
End Function

Private Function BuildPainelPipelineInfo(ByVal pipelineIndex As Long) As String
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("PAINEL")

    Dim colIniciar As Long
    colIniciar = 2 + (pipelineIndex - 1) * 2

    Dim colReg As Long
    colReg = colIniciar + 1

    Dim txt As String
    txt = "Pipeline Index: " & CStr(pipelineIndex) & vbCrLf
    txt = txt & "Nome: " & CStr(ws.Cells(1, colIniciar).Value) & vbCrLf
    txt = txt & "INPUT Folder: " & CStr(ws.Cells(2, colIniciar).Value) & vbCrLf
    txt = txt & "OUTPUT Folder: " & CStr(ws.Cells(3, colIniciar).Value) & vbCrLf
    txt = txt & "Auto-guardar ficheiros: " & CStr(ws.Cells(4, colIniciar).Value) & vbCrLf
    txt = txt & "Max Steps: " & CStr(ws.Cells(5, colIniciar).Value) & vbCrLf
    txt = txt & "Max Repetitions: " & CStr(ws.Cells(6, colIniciar).Value) & vbCrLf
    txt = txt & "Primeiros IDs (INICIAR):" & vbCrLf

    Dim r As Long
    For r = 10 To 20
        txt = txt & "- " & CStr(ws.Cells(r, colIniciar).Value) & vbCrLf
    Next r

    txt = txt & "Primeiros IDs (REGISTAR):" & vbCrLf
    For r = 10 To 20
        txt = txt & "- " & CStr(ws.Cells(r, colReg).Value) & vbCrLf
    Next r

    BuildPainelPipelineInfo = txt
    Exit Function

EH:
    BuildPainelPipelineInfo = ""
End Function

Private Function BuildExecutedCatalogCsv(ByVal wsSeg As Worksheet, ByVal pipelineNome As String) As String
    Dim out As String
    out = CsvRow(Array("ID", "Nome curto", "Nome descritivo", "Texto prompt", "Modelo", "Modos", "Storage", "Config extra", "Comentarios", "Notas para desenvolvimento", "Historico de versoes")) & vbCrLf

    Dim map As Object
    Set map = HeaderMap(wsSeg)

    Dim cPipe As Long
    Dim cPid As Long
    cPipe = MapGetFirst(map, Array("Pipeline", "pipeline_name", "Nome do Pipeline"))
    cPid = MapGetFirst(map, Array("Prompt ID", "Prompt Id", "ID Prompt", "prompt_id"))

    If cPipe = 0 Or cPid = 0 Then
        BuildExecutedCatalogCsv = out
        Exit Function
    End If

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim orderedIds As Collection
    Set orderedIds = New Collection

    Dim lr As Long
    lr = wsSeg.Cells(wsSeg.Rows.Count, cPipe).End(xlUp).Row

    Dim r As Long
    For r = 2 To lr
        If StrComp(Trim$(CStr(wsSeg.Cells(r, cPipe).Value)), Trim$(pipelineNome), vbTextCompare) = 0 Then
            Dim pid As String
            pid = Trim$(CStr(wsSeg.Cells(r, cPid).Value))
            If pid <> "" And UCase$(pid) <> "STOP" Then
                If Not seen.Exists(pid) Then
                    seen(pid) = True
                    orderedIds.Add pid
                End If
            End If
        End If
    Next r

    Dim k As Variant
    For Each k In orderedIds
        out = out & BuildExecutedCatalogCsvBlock(CStr(k))
    Next k

    BuildExecutedCatalogCsv = out
End Function

Private Function ResolvePromptSheet(ByVal promptId As String) As Worksheet
    On Error GoTo EH

    Dim sheetName As String
    sheetName = PrefixFromId(promptId)
    If sheetName = "" Then Exit Function

    Set ResolvePromptSheet = ThisWorkbook.Worksheets(sheetName)
    Exit Function
EH:
    Set ResolvePromptSheet = Nothing
End Function

Private Function BuildExecutedCatalogCsvBlock(ByVal promptId As String) As String
    Dim p As PromptDefinicao
    p = Catalogo_ObterPromptPorID(promptId)

    Dim nextPrompt As String
    Dim nextPromptDefault As String
    Dim nextPromptAllowed As String
    Dim descricaoTextual As String
    Dim inputsText As String
    Dim outputsText As String

    Call Catalogo_ReadBlockMetadata(promptId, nextPrompt, nextPromptDefault, nextPromptAllowed, descricaoTextual, inputsText, outputsText)

    Dim out As String
    out = ""

    out = out & CsvRow(Array( _
        promptId, _
        p.NomeCurto, _
        p.NomeDescritivo, _
        p.textoPrompt, _
        p.modelo, _
        p.modos, _
        CStr(p.storage), _
        p.ConfigExtra, _
        p.Comentarios, _
        p.NotasDev, _
        p.HistoricoVersoes)) & vbCrLf

    out = out & CsvRow(Array( _
        "", _
        "Next PROMPT: " & nextPrompt, _
        "Descricao textual:", _
        descricaoTextual, _
        "", "", "", "", "", "", "")) & vbCrLf

    out = out & CsvRow(Array( _
        "", _
        "Next PROMPT default: " & nextPromptDefault, _
        "INPUTS:", _
        inputsText, _
        "", "", "", "", "", "", "")) & vbCrLf

    out = out & CsvRow(Array( _
        "", _
        "Next PROMPT allowed: " & nextPromptAllowed, _
        "OUTPUTS:", _
        outputsText, _
        "", "", "", "", "", "", "")) & vbCrLf

    out = out & CsvRow(Array("", "", "", "", "", "", "", "", "", "", "")) & vbCrLf

    BuildExecutedCatalogCsvBlock = out
End Function

Private Sub Catalogo_ReadBlockMetadata(ByVal promptId As String, ByRef nextPrompt As String, ByRef nextPromptDefault As String, ByRef nextPromptAllowed As String, ByRef descricaoTextual As String, ByRef inputsText As String, ByRef outputsText As String)
    On Error GoTo EH

    Dim ws As Worksheet
    Dim rowPrompt As Long
    Set ws = Catalogo_FindPromptRow(promptId, rowPrompt)
    If ws Is Nothing Or rowPrompt <= 0 Then Exit Sub

    nextPrompt = Catalogo_ValueAfterLabel(CStr(ws.Cells(rowPrompt + 1, 2).Value), "Next PROMPT:")
    nextPromptDefault = Catalogo_ValueAfterLabel(CStr(ws.Cells(rowPrompt + 2, 2).Value), "Next PROMPT default:")
    nextPromptAllowed = Catalogo_ValueAfterLabel(CStr(ws.Cells(rowPrompt + 3, 2).Value), "Next PROMPT allowed:")

    descricaoTextual = Catalogo_ValueAfterLabel(CStr(ws.Cells(rowPrompt + 1, 3).Value), "Descricao textual:")
    If descricaoTextual = "" Then descricaoTextual = Catalogo_ValueAfterLabel(CStr(ws.Cells(rowPrompt + 1, 3).Value), "Descriï¿½ï¿½o textual:")

    inputsText = Catalogo_ValueAfterLabel(CStr(ws.Cells(rowPrompt + 2, 3).Value), "INPUTS:")
    outputsText = Catalogo_ValueAfterLabel(CStr(ws.Cells(rowPrompt + 3, 3).Value), "OUTPUTS:")

    If descricaoTextual = "" Then descricaoTextual = Trim$(CStr(ws.Cells(rowPrompt + 1, 4).Value))
    If inputsText = "" Then inputsText = Trim$(CStr(ws.Cells(rowPrompt + 2, 4).Value))
    If outputsText = "" Then outputsText = Trim$(CStr(ws.Cells(rowPrompt + 3, 4).Value))
    Exit Sub

EH:
    nextPrompt = ""
    nextPromptDefault = ""
    nextPromptAllowed = ""
    descricaoTextual = ""
    inputsText = ""
    outputsText = ""
End Sub

Private Function Catalogo_FindPromptRow(ByVal promptId As String, ByRef rowPrompt As Long) As Worksheet
    rowPrompt = 0

    Dim folha As String
    folha = PrefixFromId(promptId)
    If folha = "" Then Exit Function

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(folha)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim lookupRaw As String
    lookupRaw = Trim$(promptId)
    Dim lookupNorm As String
    lookupNorm = NormalizePromptIdForLookup(lookupRaw)

    Dim r As Long
    For r = 2 To lastRow
        Dim idRaw As String
        idRaw = Trim$(CStr(ws.Cells(r, 1).Value))
        If StrComp(idRaw, lookupRaw, vbTextCompare) = 0 Then
            rowPrompt = r
            Set Catalogo_FindPromptRow = ws
            Exit Function
        End If

        If lookupNorm <> "" Then
            If StrComp(NormalizePromptIdForLookup(idRaw), lookupNorm, vbTextCompare) = 0 Then
                rowPrompt = r
                Set Catalogo_FindPromptRow = ws
                Exit Function
            End If
        End If
    Next r
End Function

Private Function NormalizePromptIdForLookup(ByVal textValue As String) As String
    Dim s As String
    s = Trim$(CStr(textValue))
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, ChrW$(160), "")
    NormalizePromptIdForLookup = Trim$(s)
End Function

Private Function Catalogo_ValueAfterLabel(ByVal cellText As String, ByVal labelText As String) As String
    Dim raw As String
    raw = Trim$(cellText)
    If raw = "" Then Exit Function

    If LCase$(Left$(raw, Len(labelText))) = LCase$(labelText) Then
        Catalogo_ValueAfterLabel = Trim$(Mid$(raw, Len(labelText) + 1))
    Else
        Catalogo_ValueAfterLabel = ""
    End If
End Function

Private Function PrefixFromId(ByVal promptId As String) As String
    Dim p As Long
    p = InStr(1, promptId, "/")
    If p > 1 Then
        PrefixFromId = Left$(promptId, p - 1)
    Else
        PrefixFromId = ""
    End If
End Function

Private Function SheetToCsv(ByVal ws As Worksheet, Optional ByVal includeRowNumber As Boolean = False) As String
    Dim lr As Long
    Dim lc As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim r As Long
    Dim c As Long
    Dim startRow As Long
    Dim csvLine As String
    Dim out As String

    If includeRowNumber Then
        out = CsvEscape("row_number")
        For c = 1 To lc
            out = out & "," & CsvEscape(CStr(ws.Cells(1, c).Value))
        Next c
        out = out & vbCrLf
        startRow = 2
    Else
        startRow = 1
    End If

    For r = startRow To lr
        csvLine = ""
        If includeRowNumber Then
            csvLine = CsvEscape(CStr(r))
        End If
        For c = 1 To lc
            If c > 1 Or includeRowNumber Then csvLine = csvLine & ","
            csvLine = csvLine & CsvEscape(CStr(ws.Cells(r, c).Value))
        Next c
        out = out & csvLine & vbCrLf
    Next r

    SheetToCsv = out
End Function

Private Function CsvRow(ByVal vals As Variant) As String
    Dim i As Long
    Dim s As String

    For i = LBound(vals) To UBound(vals)
        If i > LBound(vals) Then s = s & ","
        s = s & CsvEscape(CStr(vals(i)))
    Next i

    CsvRow = s
End Function

Private Function CsvEscape(ByVal s As String) As String
    s = Replace(s, """", """""")
    CsvEscape = """" & s & """"
End Function

Private Function GitFileItem(ByVal path As String, ByVal content As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = path
    d("content") = content
    Set GitFileItem = d
End Function

Private Function GitData_CommitFiles(ByVal owner As String, ByVal repo As String, ByVal branch As String, ByVal token As String, ByVal files As Collection, ByVal pipelineNome As String) As String
    On Error GoTo EH

    Dim apiBase As String: apiBase = GitCfg_Get("GH_API_BASE", "https://api.github.com")
    Dim headRefUrl As String
    headRefUrl = apiBase & "/repos/" & owner & "/" & repo & "/git/ref/heads/" & branch

    Dim refBody As String
    refBody = Git_Http("GET", headRefUrl, token, "")
    Dim headSha As String: headSha = JsonPick(refBody, "sha")
    If headSha = "" Then Exit Function
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_REF_OK", "HEAD obtido.", "sha=" & Left$(headSha, 10))

    Dim commitBody As String
    commitBody = Git_Http("GET", apiBase & "/repos/" & owner & "/" & repo & "/git/commits/" & headSha, token, "")
    Dim baseTreeSha As String: baseTreeSha = JsonPickTreeSha(commitBody)
    If baseTreeSha = "" Then baseTreeSha = JsonPick(commitBody, "sha")

    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_BASE_TREE_OK", "Base tree resolvida.", "tree_sha=" & Left$(baseTreeSha, 10))

    Dim treeItems As String
    Dim i As Long
    For i = 1 To files.Count
        Dim f As Object: Set f = files(i)
        Dim blobSha As String
        blobSha = Git_CreateBlob(apiBase, owner, repo, token, CStr(f("content")))
        If blobSha = "" Then Exit Function
        If treeItems <> "" Then treeItems = treeItems & ","
        treeItems = treeItems & "{""path"":""" & Json_EscapeString(CStr(f("path")) ) & """,""mode"":""100644"",""type"":""blob"",""sha"":""" & blobSha & """}"
    Next i
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_BLOBS_CREATED", "Blobs criados.", "n=" & CStr(files.Count))

    Dim treeReq As String
    treeReq = "{""base_tree"":""" & baseTreeSha & """,""tree"": [" & treeItems & "]}"
    Dim treeResp As String
    treeResp = Git_Http("POST", apiBase & "/repos/" & owner & "/" & repo & "/git/trees", token, treeReq)
    Dim treeSha As String: treeSha = JsonPick(treeResp, "sha")
    If treeSha = "" Then Exit Function
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_TREE_CREATED", "Tree criada.", "sha=" & Left$(treeSha, 10))

    Dim commitMsg As String
    commitMsg = Replace(GitCfg_Get("GH_COMMIT_MESSAGE_TEMPLATE", "PIPELINER run {{RUN_ID}}"), "{{RUN_ID}}", Format$(Now, "yyyymmdd_hhnnss"))

    Dim newCommitReq As String
    newCommitReq = "{""message"":""" & Json_EscapeString(commitMsg) & """,""tree"":""" & treeSha & """,""parents"": [""" & headSha & """]}"
    Dim newCommitResp As String
    newCommitResp = Git_Http("POST", apiBase & "/repos/" & owner & "/" & repo & "/git/commits", token, newCommitReq)
    Dim newCommitSha As String: newCommitSha = JsonPick(newCommitResp, "sha")
    If newCommitSha = "" Then Exit Function
    Call Debug_Registar(0, pipelineNome, "INFO", "", "GH_COMMIT_CREATED", "Commit criado.", "sha=" & Left$(newCommitSha, 10))

    Dim updReq As String
    updReq = "{""sha"":""" & newCommitSha & """,""force"":" & LCase$(GitCfg_Get("GH_FORCE_UPDATE", "false")) & "}"
    Call Git_Http("PATCH", headRefUrl, token, updReq)

    GitData_CommitFiles = newCommitSha
    Exit Function
EH:
    GitData_CommitFiles = ""
End Function

Private Function Git_CreateBlob(ByVal apiBase As String, ByVal owner As String, ByVal repo As String, ByVal token As String, ByVal content As String) As String
    Dim req As String
    req = "{""content"":""" & Json_EscapeString(content) & """,""encoding"":""utf-8""}"
    Dim resp As String
    resp = Git_Http("POST", apiBase & "/repos/" & owner & "/" & repo & "/git/blobs", token, req)
    Git_CreateBlob = JsonPick(resp, "sha")
End Function

Private Function Git_Http(ByVal method As String, ByVal url As String, ByVal token As String, ByVal body As String) As String
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Accept", GitCfg_Get("GH_ACCEPT_HEADER", "application/vnd.github+json")
    http.SetRequestHeader "X-GitHub-Api-Version", GitCfg_Get("GH_API_VERSION", "2022-11-28")
    http.SetRequestHeader "User-Agent", GitCfg_Get("GH_USER_AGENT", "PIPELINER-VBA")
    If body <> "" Then
        http.SetRequestHeader "Content-Type", "application/json"
        http.Send body
    Else
        http.Send
    End If

    Git_Http = CStr(http.ResponseText)
End Function

Private Function GitCfg_Get(ByVal keyName As String, ByVal defaultValue As String) As String
    GitCfg_Get = GH_Config_Get(keyName, defaultValue)
End Function

Private Function JsonPick(ByVal body As String, ByVal keyName As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """" & keyName & """\s*:\s*""([^""]+)"""
    If re.Test(body) Then JsonPick = re.Execute(body)(0).SubMatches(0)
End Function

Private Function JsonPickTreeSha(ByVal body As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """tree""\s*:\s*\{\s*""sha""\s*:\s*""([^""]+)"""
    If re.Test(body) Then JsonPickTreeSha = re.Execute(body)(0).SubMatches(0)
End Function

Private Sub GitDebug_WriteLinkToSeguimento(ByVal pipelineNome As String, ByVal link As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEGUIMENTO)

    Dim map As Object
    Set map = HeaderMap(ws)

    Dim cPipe As Long
    cPipe = MapGet(map, "pipeline_name")

    Dim cGit As Long
    cGit = MapGet(map, "GIT_DEBUG")
    If cGit = 0 Then
        cGit = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, cGit).Value = "GIT_DEBUG"
    End If

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, cPipe).Value)) = pipelineNome Then ws.Cells(r, cGit).Value = link
    Next r
End Sub

Private Sub GitDebug_WriteLinkToHistorico(ByVal pipelineNome As String, ByVal link As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = GitDebug_GetHistoricoSheet()
    If ws Is Nothing Then Exit Sub

    Dim map As Object
    Set map = HeaderMap(ws)

    Dim cPipe As Long
    cPipe = MapGet(map, "Nome do Pipeline")

    Dim cGit As Long
    cGit = MapGet(map, "GIT_DEBUG")
    If cGit = 0 Then
        cGit = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, cGit).Value = "GIT_DEBUG"
    End If

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, cPipe).Value)) = pipelineNome And Trim$(CStr(ws.Cells(r, cGit).Value)) = "" Then
            ws.Cells(r, cGit).Value = link
        End If
    Next r
End Sub


Private Function GitDebug_GetHistoricoSheet() As Worksheet
    On Error Resume Next
    Set GitDebug_GetHistoricoSheet = ThisWorkbook.Worksheets(SHEET_HIST)
    If GitDebug_GetHistoricoSheet Is Nothing Then Set GitDebug_GetHistoricoSheet = ThisWorkbook.Worksheets("HIST" & ChrW$(&HD3) & "RICO")
    On Error GoTo 0
End Function

Private Function HeaderMap(ByVal ws As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim lc As Long
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lc
        d(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c

    Set HeaderMap = d
End Function

Private Function MapGet(ByVal d As Object, ByVal keyName As String) As Long
    If d.exists(keyName) Then
        MapGet = CLng(d(keyName))
    Else
        MapGet = 0
    End If
End Function

Private Function MapGetFirst(ByVal d As Object, ByVal keys As Variant) As Long
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        MapGetFirst = MapGet(d, CStr(keys(i)))
        If MapGetFirst > 0 Then Exit Function
    Next i
    MapGetFirst = 0
End Function


Private Function GitDebug_SanitizePathPart(ByVal s As String) As String
    Dim out As String
    out = Trim$(s)
    out = Replace(out, "\", "-")
    out = Replace(out, "/", "-")
    out = Replace(out, ":", "-")
    out = Replace(out, "*", "-")
    out = Replace(out, "?", "-")
    out = Replace(out, """", "-")
    out = Replace(out, "<", "-")
    out = Replace(out, ">", "-")
    out = Replace(out, "|", "-")
    If out = "" Then out = "pipeline"
    GitDebug_SanitizePathPart = out
End Function

Public Sub GitDebug_Config_InstalarParametros(Optional ByVal sobrescreverValores As Boolean = False)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")

    Call GitDebug_Config_EnsureGuideHeaders(ws)

    Dim defs As Collection
    Set defs = GitDebug_Config_Definitions()

    Dim i As Long
    Dim createdCount As Long
    Dim updatedCount As Long

    For i = 1 To defs.Count
        Dim d As Object
        Set d = defs(i)

        Dim rowKey As Long
        rowKey = GitDebug_Config_FindKeyRow(ws, CStr(d("key")))

        If rowKey = 0 Then
            rowKey = GitDebug_Config_NextRow(ws)
            ws.Cells(rowKey, 1).Value = CStr(d("key"))
            createdCount = createdCount + 1
        Else
            updatedCount = updatedCount + 1
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 2).Value)) = "" Then
            ws.Cells(rowKey, 2).Value = CStr(d("default"))
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 3).Value)) = "" Then
            ws.Cells(rowKey, 3).Value = CStr(d("help"))
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 4).Value)) = "" Then
            ws.Cells(rowKey, 4).Value = CStr(d("default"))
        End If

        If sobrescreverValores Or Trim$(CStr(ws.Cells(rowKey, 5).Value)) = "" Then
            ws.Cells(rowKey, 5).Value = CStr(d("allowed"))
        End If
    Next i

    MsgBox "Parametros GH_* preparados no Config." & vbCrLf & _
           "Criados: " & CStr(createdCount) & " | Atualizados/validados: " & CStr(updatedCount), vbInformation
    Exit Sub

EH:
    Call Debug_Registar(0, "", "ERRO", "", "GH_CONFIG_INSTALL_FAIL", "Falha ao instalar parametros GH_* no Config.", "err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 180))
    MsgBox "Erro em GitDebug_Config_InstalarParametros: " & Err.Description, vbExclamation
End Sub

Public Sub GitDebug_Config_InstalarMinimos()
    On Error GoTo EH

    Call GitDebug_Config_InstalarParametros(False)
    Call Debug_Registar(0, "", "INFO", "", "GH_CONFIG_INSTALL_MIN", _
        "Parametros minimos GH_* preparados na folha Config.", _
        "[ACTION] Rever GH_OWNER/GH_REPO/GH_BRANCH/GH_TOKEN_ENV ou GH_TOKEN_CONFIG/GH_BASE_PATH antes de executar.")
    Exit Sub
EH:
    Call Debug_Registar(0, "", "ERRO", "", "GH_CONFIG_INSTALL_MIN_FAIL", _
        "Falha ao preparar parametros minimos GH_*.", _
        "err=" & CStr(Err.Number) & " | " & Left$(Err.Description, 180))
End Sub

Private Sub GitDebug_Config_EnsureGuideHeaders(ByVal ws As Worksheet)
    ws.Cells(GH_CONFIG_HEADER_ROW, 1).Value = "Key"
    ws.Cells(GH_CONFIG_HEADER_ROW, 2).Value = "Value"
    ws.Cells(GH_CONFIG_HEADER_ROW, 3).Value = "Explicacao (leigos)"
    ws.Cells(GH_CONFIG_HEADER_ROW, 4).Value = "Default"
    ws.Cells(GH_CONFIG_HEADER_ROW, 5).Value = "Valores possiveis / intervalo"
End Sub

Private Function GitDebug_Config_Definitions() As Collection
    Dim defs As New Collection

    Call GitDebug_Config_Add(defs, "GH_UPLOAD_MODE", "tree_commit", "Modo global do upload para GitHub no PIPELINER.", "contents_api | tree_commit")
    Call GitDebug_Config_Add(defs, "GH_OWNER", "cpsa-org", "Dono do repositorio (organizacao ou utilizador).", "texto nao vazio")
    Call GitDebug_Config_Add(defs, "GH_REPO", "pipeliner-data", "Nome do repositorio onde guardar os debug runs.", "texto nao vazio")
    Call GitDebug_Config_Add(defs, "GH_BRANCH", "main", "Branch alvo para criar commits de debug.", "branch existente")
    Call GitDebug_Config_Add(defs, "GH_API_BASE", "https://api.github.com", "URL base da API GitHub (ou GitHub Enterprise).", "URL valida")

    Call GitDebug_Config_Add(defs, "GH_AUTH_MODE", "PAT", "Modo de autenticacao. Hoje o fluxo usa token (PAT).", "PAT | GITHUB_APP")
    Call GitDebug_Config_Add(defs, "GH_TOKEN_ENV", "GITHUB_TOKEN", "Nome da variavel de ambiente que guarda o token.", "nome de variavel de ambiente")
    Call GitDebug_Config_Add(defs, "GH_TOKEN_CONFIG", "", "Fallback local para token quando ENV estiver vazio (evitar em producao).", "string vazia ou token")

    Call GitDebug_Config_Add(defs, "GH_COMMIT_PREFIX", "PIPELINER", "Prefixo visual para identificar commits automaticos.", "texto curto")
    Call GitDebug_Config_Add(defs, "GH_COMMIT_AUTHOR_NAME", "PIPELINER Bot", "Nome de autor para auditoria nos commits.", "texto")
    Call GitDebug_Config_Add(defs, "GH_COMMIT_AUTHOR_EMAIL", "bot@cpsa.pt", "Email de autor para auditoria nos commits.", "email")
    Call GitDebug_Config_Add(defs, "GH_COMMIT_MESSAGE_TEMPLATE", "PIPELINER run {{RUN_ID}}", "Template da mensagem de commit. {{RUN_ID}} e substituido no runtime.", "template com placeholders")

    Call GitDebug_Config_Add(defs, "GH_BATCH_MODE", "tree_commit", "Modo de upload em batch para este modulo.", "tree_commit")
    Call GitDebug_Config_Add(defs, "GH_CONTENTS_BATCH_POLICY", "fail_fast", "Politica de lote para contents_api (aborta no 1o erro ou continua).", "fail_fast | best_effort")
    Call GitDebug_Config_Add(defs, "GH_MAX_FILES", "200", "Numero maximo de ficheiros por commit (protecao).", "1..1000")
    Call GitDebug_Config_Add(defs, "GH_MAX_FILE_MB", "50", "Tamanho maximo por ficheiro (MB).", "1..200")
    Call GitDebug_Config_Add(defs, "GH_ENCODING_TEXT", "utf-8", "Encoding dos ficheiros de texto enviados para blobs.", "utf-8")
    Call GitDebug_Config_Add(defs, "GH_BINARY_MODE", "base64", "Encoding recomendado para ficheiros binarios.", "base64")

    Call GitDebug_Config_Add(defs, "GH_BASE_PATH", "pipeliner_runs", "Pasta base no repo para agrupar execucoes.", "path relativo sem / inicial")
    Call GitDebug_Config_Add(defs, "GH_RUN_FOLDER_TEMPLATE", "{{PIPELINE_NAME}}/{{PROMPT_NAME}}/{{VERSION}}/{{YYYY-MM-DD HHDD}}", "Template de referencia da subpasta do run (estrutura canonica obrigatoria).", "formato fixo: {{PIPELINE_NAME}}/{{PROMPT_NAME}}/{{VERSION}}/{{YYYY-MM-DD HHDD}}")
    Call GitDebug_Config_Add(defs, "GH_LOG_FOLDER", "logs", "Subpasta para logs complementares (quando aplicavel).", "path relativo")

    Call GitDebug_Config_Add(defs, "GH_RETRY_ON_CONFLICT", "true", "Se true, tenta novamente quando o HEAD muda durante commit.", "true | false")
    Call GitDebug_Config_Add(defs, "GH_MAX_RETRIES", "3", "Numero maximo de tentativas em conflito 409 ao atualizar refs.", "inteiro >= 1")
    Call GitDebug_Config_Add(defs, "GH_FORCE_UPDATE", "false", "Se true, faz update forcado da ref (nao recomendado).", "true | false")

    Call GitDebug_Config_Add(defs, "GH_DEBUG_MODE", "true", "Liga registos de troubleshooting GH_* no DEBUG.", "true | false")
    Call GitDebug_Config_Add(defs, "GH_LOG_HTTP", "false", "Se true, regista requests/responses HTTP resumidos no DEBUG.", "true | false")
    Call GitDebug_Config_Add(defs, "GH_LOG_BLOB_SHA", "true", "Se true, mostra SHA curto dos blobs criados no DEBUG.", "true | false")

    Call GitDebug_Config_Add(defs, "GH_API_VERSION", "2022-11-28", "Versao da API GitHub enviada em header.", "YYYY-MM-DD")
    Call GitDebug_Config_Add(defs, "GH_ACCEPT_HEADER", "application/vnd.github+json", "Header Accept enviado para a API GitHub.", "media type HTTP valido")
    Call GitDebug_Config_Add(defs, "GH_USER_AGENT", "PIPELINER-VBA", "User-Agent usado nas chamadas a API.", "texto sem vazio")
    Call GitDebug_Config_Add(defs, "GH_HEADERS_EXTRA_JSON", "", "Headers extra opcionais em JSON simples (ex.: {""X-Trace"":""abc""}).", "JSON objeto ou vazio")

    Set GitDebug_Config_Definitions = defs
End Function

Private Sub GitDebug_Config_Add(ByRef defs As Collection, ByVal keyName As String, ByVal defaultValue As String, ByVal helpText As String, ByVal allowed As String)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("key") = keyName
    d("default") = defaultValue
    d("help") = helpText
    d("allowed") = allowed
    defs.Add d
End Sub

Private Function GitDebug_Config_FindKeyRow(ByVal ws As Worksheet, ByVal keyName As String) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < GH_CONFIG_FIRST_DATA_ROW Then
        GitDebug_Config_FindKeyRow = 0
        Exit Function
    End If

    Dim r As Long
    For r = GH_CONFIG_FIRST_DATA_ROW To lr
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), keyName, vbTextCompare) = 0 Then
            GitDebug_Config_FindKeyRow = r
            Exit Function
        End If
    Next r

    GitDebug_Config_FindKeyRow = 0
End Function

Private Function GitDebug_Config_NextRow(ByVal ws As Worksheet) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < GH_CONFIG_FIRST_DATA_ROW Then lr = GH_CONFIG_HEADER_ROW
    GitDebug_Config_NextRow = lr + 1
End Function
