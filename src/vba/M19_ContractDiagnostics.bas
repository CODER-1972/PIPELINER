Attribute VB_Name = "M19_ContractDiagnostics"
Option Explicit

' =============================================================================
' Módulo: M19_ContractDiagnostics
' Propósito:
' - Avaliar contrato diagnóstico por passo (opt-in via Config extra) com estado tri-state.
' - Emitir eventos canónicos no DEBUG com metadados legíveis ([RunID], [Passo], [Estado], ...).
' - Fornecer payload compacto (DetailJsonCompact) com orçamento configurável na folha Config.
'
' Atualizações:
' - 2026-03-04 | Codex | Hardening de contrato com prova minima textual obrigatoria
'   - Torna parsing de diagnostic_contract resiliente a aliases comuns (underscore/hifen/espaco).
'   - Exige marcadores minimos de prova (CSV_EXISTE_EM_MNT_DATA, FILE_CSV, MNT_DATA_LIST) quando ha intencao CSV/EXECUTE.
'   - Reforca CONTRACT_SUGGEST_ENABLE com contexto (modo/processo) e sugestao prescritiva.
'   - Valida consistencia de estado (CSV_EXISTE_EM_MNT_DATA afirmativo quando EXPORT_OK_CSV/LOAD_CSV indicam sucesso).
' - 2026-03-04 | Codex | Diff determinístico PROVA_CI e extração pública do bloco
'   - Adiciona diff formal expected-vs-prova no contrato (R5_PROVA_EXPECTED_DIFF).
'   - Expõe ContractDiag_ExtractProvaCiBlock para reutilização no bundle de diagnósticos.
'   - Expande selftests para validar cenário de diff em falta (FAIL inequívoco).
' - 2026-03-04 | Codex | Regras hierárquicas e observabilidade para SEM_CONTRATO
'   - Emite CONTRACT_STATE_DECISION/CONTRACT_NEXT_ACTION também quando não há contrato ativo.
'   - Ajusta regra de citation: WARN (sem bloqueio) quando há PROVA_CI válida com FLOW_TEMPLATE.csv.
'   - Expande selftests para cenários de fallback válido e bloqueio sem prova equivalente.
' - 2026-03-04 | Codex | Hardening de regras e eventos canónicos por regra
'   - Adiciona emissão explícita de CONTRACT_RULE_RESULT (PASS/WARN/FAIL) para auditoria por regra.
'   - Reforça validação de PROVA_CI para priorizar bloco delimitado (PROVA_CI_START/PROVA_CI_END).
'   - Inclui SelfTest_ContractDiagnostics_RunAll para cenários DoD críticos.
' - 2026-03-03 | Codex | Implementação inicial do contrato ci_csv_v1
'   - Adiciona avaliação de marcadores mínimos (PROVA_CI/FOUND/EXPORT/citation/EXECUTE).
'   - Emite eventos CONTRACT_* em DEBUG e devolve decisão OK/FAIL/BLOCKED para gate no M07.
'   - Suporta budget configurável DEBUG_DETAIL_JSON_MAX_CHARS para detalhe compacto.
'
' Funções e procedimentos:
' - ContractDiag_EvaluateStep(...)
'   - Avalia contrato do passo e devolve decisão + detalhe + sugestão.
' - ContractDiag_ExtractProvaCiBlock(outputText As String) As String
'   - Extrai o bloco PROVA_CI para auditoria e artefactos de suporte.
' - ContractDiag_GetBundleMode(...)
'   - Resolve modo de bundle por precedência (prompt > Config global > default).
' - ContractDiag_GetDiagnosticsSubfolder(...)
'   - Resolve subpasta de diagnósticos por precedência (prompt > Config global > default).
' - SelfTest_ContractDiagnostics_RunAll()
'   - Executa cenários mínimos do contrato (sem contrato, marcador ausente, inconsistência e caso OK).
' =============================================================================

Private Const CFG_DEBUG_DETAIL_JSON_MAX_CHARS As String = "DEBUG_DETAIL_JSON_MAX_CHARS"
Private Const CFG_DIAG_BUNDLE_MODE As String = "DIAG_BUNDLE_MODE"
Private Const CFG_DIAGNOSTICS_SUBFOLDER As String = "DIAGNOSTICS_SUBFOLDER"
Private Const EXPECTED_FLOW_CSV As String = "FLOW_TEMPLATE.csv"

Public Sub ContractDiag_EvaluateStep( _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal configExtraText As String, _
    ByVal outputText As String, _
    ByVal rawResponseJson As String, _
    ByRef outHasContract As Boolean, _
    ByRef outContractMode As String, _
    ByRef outStepState As String, _
    ByRef outRuleCode As String, _
    ByRef outProblem As String, _
    ByRef outSuggestion As String, _
    ByRef outDetailJsonCompact As String)

    Dim runId As String
    runId = ContractDiag_RunId()

    outHasContract = False
    outContractMode = "SEM_CONTRATO"
    outStepState = "OK"
    outRuleCode = "C0_NO_CONTRACT"
    outProblem = "Passo sem contrato explícito; execução segue com observação diagnóstica."
    outSuggestion = "Sem ação obrigatória. Se houver falhas recorrentes, ativar diagnostic_contract: ci_csv_v1 no passo."
    outDetailJsonCompact = "{}"

    Dim contractMode As String
    contractMode = LCase$(Trim$(ContractDiag_GetRequestedContractMode(configExtraText)))

    If contractMode <> "ci_csv_v1" Then
        Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_EVAL_START", runId, "SEM_CONTRATO", "OK", "C0_NO_CONTRACT", _
            "Validação de consistência desativada neste passo.", _
            "Se houver intermitência de ficheiros, considerar ativar diagnostic_contract: ci_csv_v1.", _
            "{""contract"":""SEM_CONTRATO""}")

        If InStr(1, outputText, EXPECTED_FLOW_CSV, vbTextCompare) > 0 Or _
           InStr(1, outputText, "LOAD_CSV", vbTextCompare) > 0 Or _
           InStr(1, outputText, "EXECUTE:", vbTextCompare) > 0 Then
            Dim lintContext As String
            lintContext = "{" & _
                """hint"":""csv_or_execute_without_contract""," & _
                """hasFlowTemplateCsvMention"":" & LCase$(CStr(InStr(1, outputText, EXPECTED_FLOW_CSV, vbTextCompare) > 0)) & "," & _
                """hasExecute"":" & LCase$(CStr(InStr(1, outputText, "EXECUTE:", vbTextCompare) > 0)) & "," & _
                """hasLoadCsv"":" & LCase$(CStr(InStr(1, outputText, "LOAD_CSV", vbTextCompare) > 0)) & "," & _
                """processModeHint"":""" & ContractDiag_JsonEscape(ContractDiag_ConfigExtraGet(configExtraText, "process_mode")) & """" & _
                "}"
            Call ContractDiag_LogEvent(passo, promptId, "ALERTA", "CONTRACT_SUGGEST_ENABLE", runId, "SEM_CONTRATO", "OK", "C0_SUGGEST", _
                "Este passo menciona CSV/LOAD_CSV sem contrato ativo.", _
                "Sugestão: ativar diagnostic_contract: ci_csv_v1 e exigir prova textual mínima (CSV_EXISTE_EM_MNT_DATA, FILE_CSV, MNT_DATA_LIST).", _
                lintContext)
        End If

        Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_STATE_DECISION", runId, "SEM_CONTRATO", "OK", "C0_NO_CONTRACT", _
            outProblem, outSuggestion, "{""contract"":""SEM_CONTRATO""}")
        Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_NEXT_ACTION", runId, "SEM_CONTRATO", "OK", "C0_NO_CONTRACT", _
            "Próxima ação recomendada definida.", outSuggestion, "{""action"":""continue_pipeline""}")
        Exit Sub
    End If

    outHasContract = True
    outContractMode = "ci_csv_v1"
    outRuleCode = "C1_PARSE"

    Dim hasProva As Boolean, hasFound As Boolean, hasExport As Boolean
    Dim foundCsv As String, exportOk As String
    Dim hasCitation As Boolean, hasExecuteLoadCsv As Boolean, provaHasCsv As Boolean
    Dim hasCsvInMntData As Boolean, csvInMntData As String
    Dim hasFileCsv As Boolean, fileCsvName As String
    Dim hasMntDataList As Boolean, mntDataListText As String
    Dim hasCsvIntent As Boolean

    hasProva = ContractDiag_HasValidProva(outputText)
    hasFound = ContractDiag_TryReadBoolMarker(outputText, "FOUND_FLOW_TEMPLATE_CSV", foundCsv)
    hasExport = ContractDiag_TryReadBoolMarker(outputText, "EXPORT_OK_CSV", exportOk)
    hasCitation = (InStr(1, outputText, "container_file_citation", vbTextCompare) > 0) Or _
                  (InStr(1, rawResponseJson, """type"":""container_file_citation""", vbTextCompare) > 0)
    hasExecuteLoadCsv = (InStr(1, outputText, "EXECUTE:", vbTextCompare) > 0 And InStr(1, outputText, "LOAD_CSV", vbTextCompare) > 0)
    provaHasCsv = ContractDiag_ProvaHasFlowTemplateCsv(outputText)
    hasCsvInMntData = ContractDiag_TryReadMarkerText(outputText, "CSV_EXISTE_EM_MNT_DATA", csvInMntData)
    hasFileCsv = ContractDiag_TryReadMarkerText(outputText, "FILE_CSV", fileCsvName)
    hasMntDataList = ContractDiag_TryReadMarkerText(outputText, "MNT_DATA_LIST", mntDataListText)
    hasCsvIntent = hasExecuteLoadCsv Or (LCase$(foundCsv) = "true") Or (LCase$(exportOk) = "true") Or _
        (InStr(1, outputText, "FILE_CSV:", vbTextCompare) > 0) Or (InStr(1, outputText, "CSV_EXISTE_EM_MNT_DATA", vbTextCompare) > 0)

    Dim csvExistsAffirmative As Boolean
    csvExistsAffirmative = ContractDiag_IsAffirmativeMarker(csvInMntData)

    outDetailJsonCompact = "{" & _
        """hasProvaCi"":" & LCase$(CStr(hasProva)) & "," & _
        """foundCsv"":" & JsonBool(foundCsv) & "," & _
        """exportOkCsv"":" & JsonBool(exportOk) & "," & _
        """hasCitation"":" & LCase$(CStr(hasCitation)) & "," & _
        """hasExecuteLoadCsv"":" & LCase$(CStr(hasExecuteLoadCsv)) & "," & _
        """provaHasCsv"":" & LCase$(CStr(provaHasCsv)) & "," & _
        """hasCsvInMntData"":" & LCase$(CStr(hasCsvInMntData)) & "," & _
        """csvExistsAffirmative"":" & LCase$(CStr(csvExistsAffirmative)) & "," & _
        """csvInMntData"":""" & ContractDiag_JsonEscape(csvInMntData) & """," & _
        """hasFileCsv"":" & LCase$(CStr(hasFileCsv)) & "," & _
        """fileCsv"":""" & ContractDiag_JsonEscape(fileCsvName) & """," & _
        """hasMntDataList"":" & LCase$(CStr(hasMntDataList)) & "," & _
        """mntDataListLen"":" & CStr(Len(mntDataListText)) & "," & _
        """hasCsvIntent"":" & LCase$(CStr(hasCsvIntent)) & "}"
    outDetailJsonCompact = ContractDiag_ApplyDetailBudget(outDetailJsonCompact)

    Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_MARKERS_PARSED", runId, outContractMode, "EM_ANALISE", "C1_PARSE", _
        "Marcadores do contrato avaliados.", "Seguir validação das regras de consistência.", outDetailJsonCompact)

    If (Not hasProva) Or (Not hasFound) Or (Not hasExport) Then
        outStepState = "BLOCKED"
        outRuleCode = "C1_MISSING_MARKER"
        outProblem = "Faltam marcadores obrigatórios do contrato (PROVA_CI/FOUND/EXPORT)."
        outSuggestion = "Pedir resposta com todos os marcadores obrigatórios e repetir o passo."
        GoTo Finalize
    End If

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R1B_MIN_PROOF_MARKERS", _
        IIf((Not hasCsvIntent) Or (hasCsvInMntData And hasFileCsv And hasMntDataList), "PASS", "FAIL"), _
        "Valida prova textual mínima (CSV_EXISTE_EM_MNT_DATA, FILE_CSV, MNT_DATA_LIST) quando há intenção CSV.", outDetailJsonCompact)

    If hasCsvIntent Then
        If (Not hasCsvInMntData) Or (Not hasFileCsv) Or (Not hasMntDataList) Then
            outStepState = "BLOCKED"
            outRuleCode = "C1B_MIN_PROOF_MARKERS_MISSING"
            outProblem = "Faltam marcadores mínimos de prova textual para output CSV em CI."
            outSuggestion = "Incluir no output final: CSV_EXISTE_EM_MNT_DATA: SIM/NAO; FILE_CSV: <basename.csv>; MNT_DATA_LIST: <nome(bytes=...) ; ...>."
            GoTo Finalize
        End If
    End If

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R1C_STATE_CONSISTENCY", _
        IIf(((LCase$(exportOk) = "true") Or hasExecuteLoadCsv) And (Not csvExistsAffirmative), "FAIL", "PASS"), _
        "Valida consistência entre sucesso CSV e marcador CSV_EXISTE_EM_MNT_DATA.", outDetailJsonCompact)

    If ((LCase$(exportOk) = "true") Or hasExecuteLoadCsv) And (Not csvExistsAffirmative) Then
        outStepState = "FAIL"
        outRuleCode = "C1C_CSV_STATE_INCONSISTENT"
        outProblem = "Output reporta sucesso CSV/EXECUTE, mas CSV_EXISTE_EM_MNT_DATA não confirma estado afirmativo."
        outSuggestion = "Confirmar existência real do CSV em /mnt/data e corrigir marcador para SIM antes de emitir EXECUTE/LOAD_CSV."
        GoTo Finalize
    End If

    Dim expectedFiles As Collection
    Dim provaFiles As Collection
    Set expectedFiles = ContractDiag_BuildExpectedFiles(foundCsv, exportOk, hasExecuteLoadCsv)
    Set provaFiles = ContractDiag_ExtractProvaFileNames(outputText)

    Dim missingCsv As String, unexpectedCsv As String, matchedCsv As String
    Call ContractDiag_DiffCollections(expectedFiles, provaFiles, missingCsv, unexpectedCsv, matchedCsv)

    Dim diffDetail As String
    diffDetail = "{" & _
        """expectedFiles"":""" & ContractDiag_JsonEscape(ContractDiag_JoinCollection(expectedFiles, ";")) & """," & _
        """provaFiles"":""" & ContractDiag_JsonEscape(ContractDiag_JoinCollection(provaFiles, ";")) & """," & _
        """missingExpected"":""" & ContractDiag_JsonEscape(missingCsv) & """," & _
        """unexpectedInProva"":""" & ContractDiag_JsonEscape(unexpectedCsv) & """," & _
        """matched"":""" & ContractDiag_JsonEscape(matchedCsv) & """}"
    diffDetail = ContractDiag_ApplyDetailBudget(diffDetail)

    Call ContractDiag_LogEvent(passo, promptId, IIf(Trim$(missingCsv) = "", "INFO", "ALERTA"), "CONTRACT_PROVA_DIFF", runId, outContractMode, "EM_ANALISE", "R5_PROVA_DIFF", _
        "Comparação expected vs PROVA_CI concluída.", "Rever ficheiros em falta antes de avançar para passos dependentes de CSV.", diffDetail)

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R5_PROVA_EXPECTED_DIFF", _
        IIf(Trim$(missingCsv) = "", "PASS", "FAIL"), _
        "Valida que os ficheiros esperados estejam listados no PROVA_CI.", diffDetail)

    If Trim$(missingCsv) <> "" Then
        outStepState = "FAIL"
        outRuleCode = "C5_PROVA_EXPECTED_MISSING"
        outProblem = "PROVA_CI não comprovou ficheiro(s) esperado(s): " & missingCsv
        outSuggestion = "Reexecutar o passo e garantir listagem PROVA_CI com path completo dos ficheiros esperados."
        outDetailJsonCompact = diffDetail
        GoTo Finalize
    End If

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R2_EXECUTE_REQUIRES_FOUND", _
        IIf(hasExecuteLoadCsv And LCase$(foundCsv) = "false", "FAIL", "PASS"), _
        "Valida que LOAD_CSV só ocorre quando FOUND_FLOW_TEMPLATE_CSV=true.", outDetailJsonCompact)

    If hasExecuteLoadCsv And LCase$(foundCsv) = "false" Then
        outStepState = "FAIL"
        outRuleCode = "C2_EXECUTE_WITH_FOUND_FALSE"
        outProblem = "Inconsistência: LOAD_CSV foi solicitado mas FOUND_FLOW_TEMPLATE_CSV=false."
        outSuggestion = "Corrigir decisão do modelo ou garantir presença real do CSV antes de executar LOAD_CSV."
        GoTo Finalize
    End If

    Dim r3State As String
    r3State = "PASS"
    If LCase$(foundCsv) = "true" And (Not hasCitation) Then
        If provaHasCsv Then
            r3State = "WARN"
        Else
            r3State = "FAIL"
        End If
    End If

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R3_FOUND_REQUIRES_CITATION", _
        r3State, "Valida citação quando FOUND_FLOW_TEMPLATE_CSV=true.", outDetailJsonCompact)

    If LCase$(foundCsv) = "true" And (Not hasCitation) Then
        If provaHasCsv Then
            outStepState = "OK"
            outRuleCode = "C3_FOUND_WITHOUT_CITATION_WARN"
            outProblem = "CSV foi reportado sem container_file_citation, mas PROVA_CI confirmou FLOW_TEMPLATE.csv."
            outSuggestion = "Pode avançar, mas recomenda-se incluir citation para facilitar recolha automática em runs futuros."
        Else
            outStepState = "BLOCKED"
            outRuleCode = "C3_FOUND_WITHOUT_CITATION"
            outProblem = "CSV foi reportado como encontrado mas sem container_file_citation e sem prova equivalente."
            outSuggestion = "Pedir citation explícita ou evidência equivalente antes de avançar."
            GoTo Finalize
        End If
    End If

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R4_EXPORT_REQUIRES_PROOF", _
        IIf(LCase$(exportOk) = "true" And (Not provaHasCsv), "FAIL", "PASS"), _
        "Valida que EXPORT_OK_CSV=true esteja comprovado em PROVA_CI.", outDetailJsonCompact)

    If LCase$(exportOk) = "true" And (Not provaHasCsv) Then
        outStepState = "FAIL"
        outRuleCode = "C4_EXPORT_NOT_PROVEN"
        outProblem = "EXPORT_OK_CSV=true, mas PROVA_CI não comprova FLOW_TEMPLATE.csv."
        outSuggestion = "Reexecutar com prova de ficheiros no /mnt/data e validar listagem antes do próximo passo."
        GoTo Finalize
    End If

    outStepState = "OK"
    outRuleCode = "C9_OK"
    outProblem = "Contrato validado com sucesso para este passo."
    outSuggestion = "Pode avançar para o próximo passo da pipeline."

Finalize:
    Dim sev As String
    sev = "INFO"
    If outStepState = "BLOCKED" Then sev = "ERRO"
    If outStepState = "FAIL" Then sev = "ERRO"

    Call ContractDiag_LogEvent(passo, promptId, sev, "CONTRACT_STATE_DECISION", runId, outContractMode, outStepState, outRuleCode, outProblem, outSuggestion, outDetailJsonCompact)
    Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_NEXT_ACTION", runId, outContractMode, outStepState, outRuleCode, "Próxima ação recomendada definida.", outSuggestion, outDetailJsonCompact)
End Sub

Public Function ContractDiag_ExtractProvaCiBlock(ByVal outputText As String) As String
    Dim txt As String
    txt = Replace(outputText, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Dim a As Long, b As Long
    a = InStr(1, txt, "PROVA_CI_START", vbTextCompare)
    b = InStr(1, txt, "PROVA_CI_END", vbTextCompare)
    If a > 0 And b > a Then
        ContractDiag_ExtractProvaCiBlock = Mid$(txt, a, b - a + Len("PROVA_CI_END"))
        Exit Function
    End If

    Dim p As Long
    p = InStr(1, txt, "PROVA_CI:", vbTextCompare)
    If p > 0 Then
        ContractDiag_ExtractProvaCiBlock = Mid$(txt, p)
        Exit Function
    End If
End Function

Public Function ContractDiag_GetBundleMode(ByVal promptConfigExtra As String) As String
    Dim v As String
    v = LCase$(Trim$(ContractDiag_ConfigExtraGet(promptConfigExtra, "diag_bundle_mode")))
    If v = "" Then v = LCase$(Trim$(ContractDiag_ConfigGet(CFG_DIAG_BUNDLE_MODE, "local_only")))
    Select Case v
        Case "local_only", "zip_only", "local_and_zip"
            ContractDiag_GetBundleMode = v
        Case Else
            ContractDiag_GetBundleMode = "local_only"
    End Select
End Function

Public Function ContractDiag_GetDiagnosticsSubfolder(ByVal promptConfigExtra As String) As String
    Dim v As String
    v = Trim$(ContractDiag_ConfigExtraGet(promptConfigExtra, "diagnostics_subfolder"))
    If v = "" Then v = Trim$(ContractDiag_ConfigGet(CFG_DIAGNOSTICS_SUBFOLDER, "DEBUG_BUNDLE"))
    If v = "" Then v = "DEBUG_BUNDLE"
    ContractDiag_GetDiagnosticsSubfolder = v
End Function

Private Sub ContractDiag_LogEvent(ByVal passo As Long, ByVal promptId As String, ByVal severidade As String, ByVal parametro As String, ByVal runId As String, ByVal contractMode As String, ByVal stepState As String, ByVal ruleId As String, ByVal problemaTxt As String, ByVal sugestaoTxt As String, ByVal detailJson As String)
    Dim prefix As String
    prefix = "[RunID: " & runId & "]" & vbLf & _
             "[Passo: " & CStr(passo) & "]" & vbLf & _
             "[PromptID: " & promptId & "]" & vbLf & _
             "[Contrato: " & contractMode & "]" & vbLf & _
             "[Estado: " & stepState & "]" & vbLf & _
             "[Regra: " & ruleId & "]" & vbLf & _
             "[Severidade: " & severidade & "]"

    Dim sug As String
    sug = sugestaoTxt
    If Trim$(detailJson) <> "" Then
        sug = sug & vbLf & "[DetailJsonCompact]" & vbLf & ContractDiag_ApplyDetailBudget(detailJson)
    End If

    Call Debug_Registar(passo, promptId, severidade, "", parametro, prefix & vbLf & problemaTxt, sug)
End Sub

Private Function ContractDiag_ApplyDetailBudget(ByVal txt As String) As String
    Dim lim As Long
    lim = ContractDiag_GetDetailBudget()
    If lim < 256 Then lim = 256
    If Len(txt) <= lim Then
        ContractDiag_ApplyDetailBudget = txt
    Else
        ContractDiag_ApplyDetailBudget = Left$(txt, lim) & "...[TRUNCATED:" & CStr(Len(txt) - lim) & "]"
    End If
End Function

Private Function ContractDiag_GetDetailBudget() As Long
    Dim raw As String
    raw = Trim$(ContractDiag_ConfigGet(CFG_DEBUG_DETAIL_JSON_MAX_CHARS, "1536"))
    ContractDiag_GetDetailBudget = CLng(Val(raw))
    If ContractDiag_GetDetailBudget <= 0 Then ContractDiag_GetDetailBudget = 1536
End Function

Private Function ContractDiag_ConfigGet(ByVal keyName As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")

    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lr
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), keyName, vbTextCompare) = 0 Then
            ContractDiag_ConfigGet = Trim$(CStr(ws.Cells(i, 2).Value))
            If ContractDiag_ConfigGet = "" Then ContractDiag_ConfigGet = defaultValue
            Exit Function
        End If
    Next i
Fallback:
    ContractDiag_ConfigGet = defaultValue
End Function

Private Function ContractDiag_ConfigExtraGet(ByVal configExtraText As String, ByVal keyName As String) As String
    Dim txt As String
    txt = Replace(configExtraText, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Dim lines() As String
    lines = Split(txt, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(CStr(lines(i)))
        If line <> "" Then
            Dim p As Long
            p = InStr(1, line, ":", vbTextCompare)
            If p > 0 Then
                Dim k As String
                k = Trim$(Left$(line, p - 1))
                If StrComp(k, keyName, vbTextCompare) = 0 Then
                    ContractDiag_ConfigExtraGet = Trim$(Mid$(line, p + 1))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function ContractDiag_GetRequestedContractMode(ByVal configExtraText As String) As String
    Dim candidates As Variant
    candidates = Array("diagnostic_contract", "contract_mode", "diagnostic-contract", "diagnostic contract", "diagnostic_contract_mode")

    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        Dim v As String
        v = Trim$(ContractDiag_ConfigExtraGet(configExtraText, CStr(candidates(i))))
        If v <> "" Then
            ContractDiag_GetRequestedContractMode = v
            Exit Function
        End If
    Next i
End Function

Private Function ContractDiag_IsAffirmativeMarker(ByVal rawValue As String) As Boolean
    Dim v As String
    v = UCase$(Trim$(rawValue))
    ContractDiag_IsAffirmativeMarker = (v = "SIM" Or v = "TRUE" Or v = "YES" Or v = "1" Or v = "OK" Or v = "Y" Or v = "S")
End Function

Private Function ContractDiag_TryReadMarkerText(ByVal outputText As String, ByVal marker As String, ByRef outVal As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(^|\\n)\\s*" & marker & "\\s*:\\s*([^\\n]+)"

    If re.Test(outputText) Then
        Dim m As Object
        Set m = re.Execute(outputText)(0)
        outVal = Trim$(CStr(m.SubMatches(1)))
        ContractDiag_TryReadMarkerText = True
    End If
End Function

Private Function ContractDiag_RunId() As String
    Static sRun As String
    If Trim$(sRun) = "" Then
        Randomize
        sRun = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(1000 + Rnd() * 8999), "0000")
    End If
    ContractDiag_RunId = sRun
End Function

Private Function ContractDiag_TryReadBoolMarker(ByVal outputText As String, ByVal marker As String, ByRef outVal As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(^|\\n)\\s*" & marker & "\\s*:\\s*(true|false)"

    Dim m As Object
    If re.Test(outputText) Then
        Set m = re.Execute(outputText)(0)
        outVal = LCase$(Trim$(CStr(m.SubMatches(1))))
        ContractDiag_TryReadBoolMarker = True
    End If
End Function

Private Function ContractDiag_ProvaHasFlowTemplateCsv(ByVal outputText As String) As Boolean
    If InStr(1, outputText, EXPECTED_FLOW_CSV, vbTextCompare) = 0 Then Exit Function

    Dim s As String
    s = UCase$(outputText)
    If InStr(1, s, "PROVA_CI_START", vbTextCompare) > 0 And InStr(1, s, "PROVA_CI_END", vbTextCompare) > 0 Then
        Dim a As Long, b As Long
        a = InStr(1, s, "PROVA_CI_START", vbTextCompare)
        b = InStr(a + 1, s, "PROVA_CI_END", vbTextCompare)
        If b > a Then
            ContractDiag_ProvaHasFlowTemplateCsv = (InStr(1, Mid$(outputText, a, b - a + 1), EXPECTED_FLOW_CSV, vbTextCompare) > 0)
            Exit Function
        End If
    End If

    If InStr(1, outputText, "PROVA_CI", vbTextCompare) > 0 Then
        ContractDiag_ProvaHasFlowTemplateCsv = True
    End If
End Function

Private Sub ContractDiag_LogRuleResult(ByVal passo As Long, ByVal promptId As String, ByVal runId As String, ByVal contractMode As String, ByVal ruleCode As String, ByVal resultState As String, ByVal desc As String, ByVal detailJson As String)
    Dim sev As String
    If UCase$(resultState) = "FAIL" Then
        sev = "ERRO"
    ElseIf UCase$(resultState) = "WARN" Then
        sev = "ALERTA"
    Else
        sev = "INFO"
    End If

    Call ContractDiag_LogEvent(passo, promptId, sev, "CONTRACT_RULE_RESULT", runId, contractMode, "EM_ANALISE", ruleCode, _
        "Resultado da regra: " & resultState & ". " & desc, _
        "Se FAIL, rever marcadores/evidências deste passo.", detailJson)
End Sub

Private Function ContractDiag_HasValidProva(ByVal outputText As String) As Boolean
    Dim s As String
    s = UCase$(outputText)

    Dim a As Long, b As Long
    a = InStr(1, s, "PROVA_CI_START", vbTextCompare)
    b = InStr(1, s, "PROVA_CI_END", vbTextCompare)

    If a > 0 And b > a Then
        ContractDiag_HasValidProva = True
        Exit Function
    End If

    If InStr(1, s, "PROVA_CI:", vbTextCompare) > 0 Then
        ContractDiag_HasValidProva = True
    End If
End Function

Private Function ContractDiag_BuildExpectedFiles(ByVal foundCsv As String, ByVal exportOk As String, ByVal hasExecuteLoadCsv As Boolean) As Collection
    Dim c As Collection
    Set c = New Collection

    If LCase$(Trim$(foundCsv)) = "true" Or _
       LCase$(Trim$(exportOk)) = "true" Or _
       hasExecuteLoadCsv Then
        Call ContractDiag_CollectionAddUnique(c, EXPECTED_FLOW_CSV)
    End If

    Set ContractDiag_BuildExpectedFiles = c
End Function

Private Function ContractDiag_ExtractProvaFileNames(ByVal outputText As String) As Collection
    Dim c As Collection
    Set c = New Collection

    Dim bloco As String
    bloco = ContractDiag_ExtractProvaCiBlock(outputText)
    If Trim$(bloco) = "" Then
        Set ContractDiag_ExtractProvaFileNames = c
        Exit Function
    End If

    Dim txt As String
    txt = Replace(bloco, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Dim lines() As String
    lines = Split(txt, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(CStr(lines(i)))
        If line <> "" Then
            Dim p As Long
            p = InStr(1, line, "FILE:", vbTextCompare)
            If p > 0 Then
                Dim token As String
                token = Trim$(Mid$(line, p + 5))
                Dim pSep As Long
                pSep = InStr(1, token, "|", vbTextCompare)
                If pSep > 0 Then token = Trim$(Left$(token, pSep - 1))
                token = ContractDiag_BaseName(token)
                If token <> "" Then Call ContractDiag_CollectionAddUnique(c, token)
            End If
        End If
    Next i

    If c.Count = 0 Then
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        re.Pattern = "([A-Za-z0-9_\-\.]+\.(csv|pdf|xlsx|docx|pptx|txt|json))"

        Dim mc As Object, m As Object
        Set mc = re.Execute(bloco)
        For Each m In mc
            Call ContractDiag_CollectionAddUnique(c, ContractDiag_BaseName(CStr(m.SubMatches(0))))
        Next m
    End If

    Set ContractDiag_ExtractProvaFileNames = c
End Function

Private Sub ContractDiag_DiffCollections(ByVal expected As Collection, ByVal found As Collection, ByRef outMissingCsv As String, ByRef outUnexpectedCsv As String, ByRef outMatchedCsv As String)
    Dim i As Long, v As String

    For i = 1 To expected.Count
        v = CStr(expected(i))
        If ContractDiag_CollectionContains(found, v) Then
            outMatchedCsv = ContractDiag_AppendCsv(outMatchedCsv, v)
        Else
            outMissingCsv = ContractDiag_AppendCsv(outMissingCsv, v)
        End If
    Next i

    For i = 1 To found.Count
        v = CStr(found(i))
        If Not ContractDiag_CollectionContains(expected, v) Then
            outUnexpectedCsv = ContractDiag_AppendCsv(outUnexpectedCsv, v)
        End If
    Next i
End Sub

Private Function ContractDiag_CollectionContains(ByVal c As Collection, ByVal itemText As String) As Boolean
    Dim i As Long
    For i = 1 To c.Count
        If StrComp(CStr(c(i)), itemText, vbTextCompare) = 0 Then
            ContractDiag_CollectionContains = True
            Exit Function
        End If
    Next i
End Function

Private Sub ContractDiag_CollectionAddUnique(ByRef c As Collection, ByVal itemText As String)
    If Trim$(itemText) = "" Then Exit Sub
    If Not ContractDiag_CollectionContains(c, itemText) Then c.Add itemText
End Sub

Private Function ContractDiag_JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long
    For i = 1 To c.Count
        If ContractDiag_JoinCollection <> "" Then ContractDiag_JoinCollection = ContractDiag_JoinCollection & sep
        ContractDiag_JoinCollection = ContractDiag_JoinCollection & CStr(c(i))
    Next i
End Function

Private Function ContractDiag_BaseName(ByVal pathLike As String) As String
    Dim s As String
    s = Trim$(pathLike)
    s = Replace(s, "\", "/")
    If Right$(s, 1) = "/" Then s = Left$(s, Len(s) - 1)

    Dim p As Long
    p = InStrRev(s, "/")
    If p > 0 Then
        ContractDiag_BaseName = Mid$(s, p + 1)
    Else
        ContractDiag_BaseName = s
    End If
End Function

Private Function ContractDiag_AppendCsv(ByVal csv As String, ByVal itemText As String) As String
    If Trim$(itemText) = "" Then
        ContractDiag_AppendCsv = csv
    ElseIf Trim$(csv) = "" Then
        ContractDiag_AppendCsv = itemText
    Else
        ContractDiag_AppendCsv = csv & "," & itemText
    End If
End Function

Private Function ContractDiag_JsonEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\" & Chr$(34))
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    ContractDiag_JsonEscape = t
End Function

Public Sub SelfTest_ContractDiagnostics_RunAll()
    On Error GoTo EH

    Dim hasContract As Boolean, cMode As String, st As String, rule As String
    Dim pr As String, sg As String, dj As String

    Call ContractDiag_EvaluateStep(1, "SELFTEST/A", "", "texto sem contrato", "", hasContract, cMode, st, rule, pr, sg, dj)
    If hasContract Then Err.Raise 5, , "T1 fail: sem contrato deveria hasContract=False"
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T1 fail: sem contrato deveria manter OK"

    Call ContractDiag_EvaluateStep(2, "SELFTEST/B", "diagnostic_contract: ci_csv_v1", "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "BLOCKED" Then Err.Raise 5, , "T2 fail: faltando PROVA_CI deveria BLOCKED"

    Call ContractDiag_EvaluateStep(3, "SELFTEST/C", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: false" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "CSV_EXISTE_EM_MNT_DATA: SIM" & vbLf & "FILE_CSV: FLOW_TEMPLATE.csv" & vbLf & "MNT_DATA_LIST: FLOW_TEMPLATE.csv(bytes=10)" & vbLf & "EXECUTE: LOAD_CSV(FLOW_TEMPLATE.csv)", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "FAIL" Then Err.Raise 5, , "T3 fail: execute com FOUND=false deveria FAIL"

    Call ContractDiag_EvaluateStep(4, "SELFTEST/D", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "CSV_EXISTE_EM_MNT_DATA: SIM" & vbLf & "FILE_CSV: FLOW_TEMPLATE.csv" & vbLf & "MNT_DATA_LIST: FLOW_TEMPLATE.csv(bytes=10)" & vbLf & "container_file_citation: sandbox:/mnt/data/FLOW_TEMPLATE.csv", "{""type"":""container_file_citation""}", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T4 fail: cenário consistente deveria OK"

    Call ContractDiag_EvaluateStep(5, "SELFTEST/E", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "CSV_EXISTE_EM_MNT_DATA: SIM" & vbLf & "FILE_CSV: FLOW_TEMPLATE.csv" & vbLf & "MNT_DATA_LIST: FLOW_TEMPLATE.csv(bytes=10)", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T5 fail: sem citation mas com prova válida deveria OK (warn)"
    If UCase$(rule) <> "C3_FOUND_WITHOUT_CITATION_WARN" Then Err.Raise 5, , "T5 fail: regra esperada C3_FOUND_WITHOUT_CITATION_WARN"

    Call ContractDiag_EvaluateStep(6, "SELFTEST/F", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/out.pdf" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "CSV_EXISTE_EM_MNT_DATA: SIM" & vbLf & "FILE_CSV: FLOW_TEMPLATE.csv" & vbLf & "MNT_DATA_LIST: out.pdf(bytes=10)", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "BLOCKED" Then Err.Raise 5, , "T6 fail: found=true sem citation e sem prova csv deveria BLOCKED"

    Call ContractDiag_EvaluateStep(7, "SELFTEST/G", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/out.pdf" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: false" & vbLf & "EXPORT_OK_CSV: false" & vbLf & "CSV_EXISTE_EM_MNT_DATA: NAO" & vbLf & "FILE_CSV: [n/a]" & vbLf & "MNT_DATA_LIST: out.pdf(bytes=10)", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T7 fail: sem expected files no contrato deveria manter OK"

    Call ContractDiag_EvaluateStep(8, "SELFTEST/H", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/out.pdf" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: false" & vbLf & "CSV_EXISTE_EM_MNT_DATA: SIM" & vbLf & "FILE_CSV: FLOW_TEMPLATE.csv" & vbLf & "MNT_DATA_LIST: out.pdf(bytes=10)", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "FAIL" Then Err.Raise 5, , "T8 fail: diff expected-vs-prova deveria FAIL"
    If UCase$(rule) <> "C5_PROVA_EXPECTED_MISSING" Then Err.Raise 5, , "T8 fail: regra esperada C5_PROVA_EXPECTED_MISSING"

    Call ContractDiag_EvaluateStep(9, "SELFTEST/I", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "CSV_EXISTE_EM_MNT_DATA: NAO" & vbLf & "FILE_CSV: FLOW_TEMPLATE.csv" & vbLf & "MNT_DATA_LIST: FLOW_TEMPLATE.csv(bytes=10)" & vbLf & "EXECUTE: LOAD_CSV(FLOW_TEMPLATE.csv)", "{""type"":""container_file_citation""}", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "FAIL" Then Err.Raise 5, , "T9 fail: estado inconsistente CSV_EXISTE_EM_MNT_DATA deveria FAIL"
    If UCase$(rule) <> "C1C_CSV_STATE_INCONSISTENT" Then Err.Raise 5, , "T9 fail: regra esperada C1C_CSV_STATE_INCONSISTENT"

    Call ContractDiag_EvaluateStep(10, "SELFTEST/J", "contract_mode: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "CSV_EXISTE_EM_MNT_DATA: OK" & vbLf & "FILE_CSV: FLOW_TEMPLATE.csv" & vbLf & "MNT_DATA_LIST: FLOW_TEMPLATE.csv(bytes=10)", "", hasContract, cMode, st, rule, pr, sg, dj)
    If (Not hasContract) Then Err.Raise 5, , "T10 fail: alias contract_mode deveria ativar contrato"
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T10 fail: CSV_EXISTE_EM_MNT_DATA=OK com prova valida deveria OK"

    Call Debug_Registar(0, "SELFTEST_CONTRACT", "INFO", "", "SELFTEST_CONTRACT", "PASS", "SelfTest_ContractDiagnostics_RunAll concluído com sucesso.")
    MsgBox "SelfTest_ContractDiagnostics_RunAll PASS", vbInformation
    Exit Sub
EH:
    Call Debug_Registar(0, "SELFTEST_CONTRACT", "ERRO", "", "SELFTEST_CONTRACT", "FAIL: " & Err.Description, "Rever regras do M19_ContractDiagnostics.")
    MsgBox "SelfTest_ContractDiagnostics_RunAll FAIL: " & Err.Description, vbExclamation
End Sub

Private Function JsonBool(ByVal v As String) As String
    If LCase$(Trim$(v)) = "true" Then
        JsonBool = "true"
    ElseIf LCase$(Trim$(v)) = "false" Then
        JsonBool = "false"
    Else
        JsonBool = "null"
    End If
End Function
