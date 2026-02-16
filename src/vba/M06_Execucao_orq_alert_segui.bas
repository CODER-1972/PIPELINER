Attribute VB_Name = "M06_Execucao_orq_alert_segui"
Option Explicit

' =============================================================================
' Módulo: M06_Execucao_orq_alert_segui
' Propósito:
' - Disponibilizar entry point simples para execução de prompt por ID em contexto local.
' - Encaminhar execução para a orquestração principal preservando compatibilidade.
'
' Atualizações:
' - 2026-02-16 | Codex | Resolução de API key com prioridade para OPENAI_API_KEY
'   - Usa resolver central (M14_ConfigApiKey) em vez de leitura direta de Config!B1.
'   - Emite ALERTA/ERRO no DEBUG para origem/falhas de credencial sem expor segredos.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - ExecutarPrompt_PorID (Sub): rotina pública do módulo.
' =============================================================================


Public Sub ExecutarPrompt_PorID()
    Dim promptId As String
    promptId = InputBox("Introduza o Prompt ID (ex.: AvalCap/01/Mapa/A):", "Executar Prompt")
    If Trim$(promptId) = "" Then Exit Sub

    ' Ler config base
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Worksheets("Config")

    Dim apiKey As String
    Dim apiKeySource As String
    Dim apiKeyAlert As String
    Dim apiKeyError As String

    If Not Config_ResolveOpenAIApiKey(apiKey, apiKeySource, apiKeyAlert, apiKeyError) Then
        Call Debug_Registar(0, promptId, "ERRO", "", "OPENAI_API_KEY", apiKeyError, "Defina OPENAI_API_KEY no ambiente ou fallback válido em Config!B1.")
        MsgBox "OPENAI_API_KEY ausente. Verifique a variável de ambiente OPENAI_API_KEY ou o fallback em Config!B1.", vbExclamation
        Exit Sub
    End If

    If Trim$(apiKeyAlert) <> "" Then
        Call Debug_Registar(0, promptId, "ALERTA", "", "OPENAI_API_KEY", apiKeyAlert, "Sem ação imediata; recomendado migrar para variável de ambiente.")
    End If

    Dim modeloDefault As String
    modeloDefault = Trim$(CStr(wsCfg.Range("B2").value))
    If modeloDefault = "" Then modeloDefault = "gpt-4.1"

    Dim temperaturaDefault As Double
    temperaturaDefault = 0.7
    On Error Resume Next
    If Trim$(CStr(wsCfg.Range("B3").value)) <> "" Then temperaturaDefault = CDbl(wsCfg.Range("B3").value)
    On Error GoTo 0

    Dim maxTokensDefault As Long
    maxTokensDefault = 250
    On Error Resume Next
    If Trim$(CStr(wsCfg.Range("B4").value)) <> "" Then maxTokensDefault = CLng(wsCfg.Range("B4").value)
    On Error GoTo 0

    ' Ler prompt por ID
    Dim prompt As PromptDefinicao
    prompt = Catalogo_ObterPromptPorID(promptId)
    If Trim$(prompt.Id) = "" Then
        MsgBox "Prompt ID não encontrado: " & promptId, vbExclamation
        Exit Sub
    End If

    Dim modeloUsado As String
    modeloUsado = Trim$(prompt.modelo)
    If modeloUsado = "" Then modeloUsado = modeloDefault

    Dim passo As Long
    passo = 1

    ' Converter Config extra (amigável) -> JSON
    Dim auditJson As String, inputJsonLiteral As String, extraFragment As String
    auditJson = ""
    inputJsonLiteral = ""
    extraFragment = ""
    Call ConfigExtra_Converter(prompt.ConfigExtra, prompt.textoPrompt, passo, prompt.Id, auditJson, inputJsonLiteral, extraFragment)

    ' PAINEL: OUTPUT Folder + toggle auto-save (coluna B, por compatibilidade)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets("PAINEL")

    Dim outputFolderBase As String
    outputFolderBase = Trim$(CStr(wsP.Range("B3").value))
    If Right$(outputFolderBase, 1) = "\" Then outputFolderBase = Left$(outputFolderBase, Len(outputFolderBase) - 1)

    Dim painelAutoSave As String
    painelAutoSave = Trim$(CStr(wsP.Range("B4").value))
    If painelAutoSave = "" Then painelAutoSave = "Sim"

    ' FILE OUTPUT (pre-request): resolver config + preparar request
    Dim pipelineNome As String
    pipelineNome = "MANUAL"

    Dim fo_outputKind As String, fo_processMode As String, fo_autoSave As String, fo_overwriteMode As String
    Dim fo_prefixTmpl As String, fo_subfolderTmpl As String, fo_structuredMode As String
    Dim fo_pptxMode As String, fo_xlsxMode As String, fo_pdfMode As String, fo_imageMode As String

    Call FileOutput_ResolveEffectiveConfig(0, pipelineNome, prompt.Id, painelAutoSave, _
        fo_outputKind, fo_processMode, fo_autoSave, fo_overwriteMode, fo_prefixTmpl, fo_subfolderTmpl, _
        fo_structuredMode, fo_pptxMode, fo_xlsxMode, fo_pdfMode, fo_imageMode, prompt.ConfigExtra)

    Dim modosEfetivo As String
    modosEfetivo = prompt.modos

    Dim extraFragmentFO As String
    extraFragmentFO = extraFragment

    Call FileOutput_PrepareRequest(fo_outputKind, fo_processMode, fo_structuredMode, modosEfetivo, extraFragmentFO)

    ' Executar chamada
    Dim resultado As ApiResultado
    resultado = OpenAI_Executar( _
        apiKey, _
        modeloUsado, _
        prompt.textoPrompt, _
        temperaturaDefault, _
        maxTokensDefault, _
        modosEfetivo, _
        prompt.storage, _
        inputJsonLiteral, _
        extraFragmentFO, _
        prompt.Id _
    )

    ' FILE OUTPUT (pos-response): guardar raw + ficheiros
    Dim fo_filesUsedOut As String, fo_filesOpsOut As String, fo_logSeguimento As String
    fo_filesUsedOut = "": fo_filesOpsOut = "": fo_logSeguimento = ""

    fo_logSeguimento = FileOutput_ProcessAfterResponse(apiKey, outputFolderBase, pipelineNome, 0, passo, prompt.Id, resultado, _
        fo_outputKind, fo_processMode, fo_autoSave, fo_overwriteMode, fo_prefixTmpl, fo_subfolderTmpl, _
        fo_pptxMode, fo_xlsxMode, fo_pdfMode, fo_imageMode, fo_filesUsedOut, fo_filesOpsOut)

    Dim textoSeguimento As String
    If Trim$(resultado.Erro) <> "" Then
        textoSeguimento = "[ERRO] " & resultado.Erro
    ElseIf Trim$(fo_logSeguimento) <> "" Then
        textoSeguimento = fo_logSeguimento
    Else
        textoSeguimento = resultado.outputText
    End If

    Call Seguimento_Registar( _
        passo, prompt, modeloUsado, auditJson, resultado.httpStatus, resultado.responseId, _
        textoSeguimento, pipelineNome, "", fo_filesUsedOut, fo_filesOpsOut, "" _
    )

    If Trim$(resultado.Erro) <> "" Then
        MsgBox "Falhou (ver DEBUG e Seguimento).", vbCritical
    Else
        MsgBox "Concluído. Ver Seguimento e FILES_MANAGEMENT.", vbInformation
    End If
End Sub
