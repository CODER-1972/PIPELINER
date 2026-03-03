Attribute VB_Name = "M19_ContractDiagnostics"
Option Explicit

' =============================================================================
' MÃ³dulo: M19_ContractDiagnostics
' PropÃ³sito:
' - Avaliar contrato diagnÃ³stico por passo (opt-in via Config extra) com estado tri-state.
' - Emitir eventos canÃ³nicos no DEBUG com metadados legÃ­veis ([RunID], [Passo], [Estado], ...).
' - Fornecer payload compacto (DetailJsonCompact) com orÃ§amento configurÃ¡vel na folha Config.
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | Diff determinÃ­stico PROVA_CI e extraÃ§Ã£o pÃºblica do bloco
'   - Adiciona diff formal expected-vs-prova no contrato (R5_PROVA_EXPECTED_DIFF).
'   - ExpÃµe ContractDiag_ExtractProvaCiBlock para reutilizaÃ§Ã£o no bundle de diagnÃ³sticos.
'   - Expande selftests para validar cenÃ¡rio de diff em falta (FAIL inequÃ­voco).
' - 2026-03-04 | Codex | Regras hierÃ¡rquicas e observabilidade para SEM_CONTRATO
'   - Emite CONTRACT_STATE_DECISION/CONTRACT_NEXT_ACTION tambÃ©m quando nÃ£o hÃ¡ contrato ativo.
'   - Ajusta regra de citation: WARN (sem bloqueio) quando hÃ¡ PROVA_CI vÃ¡lida com FLOW_TEMPLATE.csv.
'   - Expande selftests para cenÃ¡rios de fallback vÃ¡lido e bloqueio sem prova equivalente.
' - 2026-03-04 | Codex | Hardening de regras e eventos canÃ³nicos por regra
'   - Adiciona emissÃ£o explÃ­cita de CONTRACT_RULE_RESULT (PASS/WARN/FAIL) para auditoria por regra.
'   - ReforÃ§a validaÃ§Ã£o de PROVA_CI para priorizar bloco delimitado (PROVA_CI_START/PROVA_CI_END).
'   - Inclui SelfTest_ContractDiagnostics_RunAll para cenÃ¡rios DoD crÃ­ticos.
' - 2026-03-03 | Codex | ImplementaÃ§Ã£o inicial do contrato ci_csv_v1
'   - Adiciona avaliaÃ§Ã£o de marcadores mÃ­nimos (PROVA_CI/FOUND/EXPORT/citation/EXECUTE).
'   - Emite eventos CONTRACT_* em DEBUG e devolve decisÃ£o OK/FAIL/BLOCKED para gate no M07.
'   - Suporta budget configurÃ¡vel DEBUG_DETAIL_JSON_MAX_CHARS para detalhe compacto.
'
' FunÃ§Ãµes e procedimentos:
' - ContractDiag_EvaluateStep(...)
'   - Avalia contrato do passo e devolve decisÃ£o + detalhe + sugestÃ£o.
' - ContractDiag_ExtractProvaCiBlock(outputText As String) As String
'   - Extrai o bloco PROVA_CI para auditoria e artefactos de suporte.
' - ContractDiag_GetBundleMode(...)
'   - Resolve modo de bundle por precedÃªncia (prompt > Config global > default).
' - ContractDiag_GetDiagnosticsSubfolder(...)
'   - Resolve subpasta de diagnÃ³sticos por precedÃªncia (prompt > Config global > default).
' - SelfTest_ContractDiagnostics_RunAll()
'   - Executa cenÃ¡rios mÃ­nimos do contrato (sem contrato, marcador ausente, inconsistÃªncia e caso OK).
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
    outProblem = "Passo sem contrato explÃ­cito; execuÃ§Ã£o segue com observaÃ§Ã£o diagnÃ³stica."
    outSuggestion = "Sem aÃ§Ã£o obrigatÃ³ria. Se houver falhas recorrentes, ativar diagnostic_contract: ci_csv_v1 no passo."
    outDetailJsonCompact = "{}"

    Dim contractMode As String
    contractMode = LCase$(Trim$(ContractDiag_ConfigExtraGet(configExtraText, "diagnostic_contract")))
    If contractMode = "" Then
        contractMode = LCase$(Trim$(ContractDiag_ConfigExtraGet(configExtraText, "contract_mode")))
    End If

    If contractMode <> "ci_csv_v1" Then
        Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_EVAL_START", runId, "SEM_CONTRATO", "OK", "C0_NO_CONTRACT", _
            "ValidaÃ§Ã£o de consistÃªncia desativada neste passo.", _
            "Se houver intermitÃªncia de ficheiros, considerar ativar diagnostic_contract: ci_csv_v1.", _
            "{""contract"":""SEM_CONTRATO""}")

        If InStr(1, outputText, EXPECTED_FLOW_CSV, vbTextCompare) > 0 Or _
           InStr(1, outputText, "LOAD_CSV", vbTextCompare) > 0 Then
            Call ContractDiag_LogEvent(passo, promptId, "ALERTA", "CONTRACT_SUGGEST_ENABLE", runId, "SEM_CONTRATO", "OK", "C0_SUGGEST", _
                "Este passo menciona CSV/LOAD_CSV sem contrato ativo.", _
                "SugestÃ£o: ativar diagnostic_contract: ci_csv_v1 para reduzir risco de falso sucesso.", _
                "{""hint"":""csv_or_execute_without_contract""}")
        End If

        Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_STATE_DECISION", runId, "SEM_CONTRATO", "OK", "C0_NO_CONTRACT", _
            outProblem, outSuggestion, "{""contract"":""SEM_CONTRATO""}")
        Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_NEXT_ACTION", runId, "SEM_CONTRATO", "OK", "C0_NO_CONTRACT", _
            "PrÃ³xima aÃ§Ã£o recomendada definida.", outSuggestion, "{""action"":""continue_pipeline""}")
        Exit Sub
    End If

    outHasContract = True
    outContractMode = "ci_csv_v1"
    outRuleCode = "C1_PARSE"

    Dim hasProva As Boolean, hasFound As Boolean, hasExport As Boolean
    Dim foundCsv As String, exportOk As String
    Dim hasCitation As Boolean, hasExecuteLoadCsv As Boolean, provaHasCsv As Boolean

    hasProva = ContractDiag_HasValidProva(outputText)
    hasFound = ContractDiag_TryReadBoolMarker(outputText, "FOUND_FLOW_TEMPLATE_CSV", foundCsv)
    hasExport = ContractDiag_TryReadBoolMarker(outputText, "EXPORT_OK_CSV", exportOk)
    hasCitation = (InStr(1, outputText, "container_file_citation", vbTextCompare) > 0) Or _
                  (InStr(1, rawResponseJson, """type"":""container_file_citation""", vbTextCompare) > 0)
    hasExecuteLoadCsv = (InStr(1, outputText, "EXECUTE:", vbTextCompare) > 0 And InStr(1, outputText, "LOAD_CSV", vbTextCompare) > 0)
    provaHasCsv = ContractDiag_ProvaHasFlowTemplateCsv(outputText)

    outDetailJsonCompact = "{" & _
        """hasProvaCi"":" & LCase$(CStr(hasProva)) & "," & _
        """foundCsv"":" & JsonBool(foundCsv) & "," & _
        """exportOkCsv"":" & JsonBool(exportOk) & "," & _
        """hasCitation"":" & LCase$(CStr(hasCitation)) & "," & _
        """hasExecuteLoadCsv"":" & LCase$(CStr(hasExecuteLoadCsv)) & "," & _
        """provaHasCsv"":" & LCase$(CStr(provaHasCsv)) & "}"
    outDetailJsonCompact = ContractDiag_ApplyDetailBudget(outDetailJsonCompact)

    Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_MARKERS_PARSED", runId, outContractMode, "EM_ANALISE", "C1_PARSE", _
        "Marcadores do contrato avaliados.", "Seguir validaÃ§Ã£o das regras de consistÃªncia.", outDetailJsonCompact)

    If (Not hasProva) Or (Not hasFound) Or (Not hasExport) Then
        outStepState = "BLOCKED"
        outRuleCode = "C1_MISSING_MARKER"
        outProblem = "Faltam marcadores obrigatÃ³rios do contrato (PROVA_CI/FOUND/EXPORT)."
        outSuggestion = "Pedir resposta com todos os marcadores obrigatÃ³rios e repetir o passo."
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
        "ComparaÃ§Ã£o expected vs PROVA_CI concluÃ­da.", "Rever ficheiros em falta antes de avanÃ§ar para passos dependentes de CSV.", diffDetail)

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R5_PROVA_EXPECTED_DIFF", _
        IIf(Trim$(missingCsv) = "", "PASS", "FAIL"), _
        "Valida que os ficheiros esperados estejam listados no PROVA_CI.", diffDetail)

    If Trim$(missingCsv) <> "" Then
        outStepState = "FAIL"
        outRuleCode = "C5_PROVA_EXPECTED_MISSING"
        outProblem = "PROVA_CI nÃ£o comprovou ficheiro(s) esperado(s): " & missingCsv
        outSuggestion = "Reexecutar o passo e garantir listagem PROVA_CI com path completo dos ficheiros esperados."
        outDetailJsonCompact = diffDetail
        GoTo Finalize
    End If

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R2_EXECUTE_REQUIRES_FOUND", _
        IIf(hasExecuteLoadCsv And LCase$(foundCsv) = "false", "FAIL", "PASS"), _
        "Valida que LOAD_CSV sÃ³ ocorre quando FOUND_FLOW_TEMPLATE_CSV=true.", outDetailJsonCompact)

    If hasExecuteLoadCsv And LCase$(foundCsv) = "false" Then
        outStepState = "FAIL"
        outRuleCode = "C2_EXECUTE_WITH_FOUND_FALSE"
        outProblem = "InconsistÃªncia: LOAD_CSV foi solicitado mas FOUND_FLOW_TEMPLATE_CSV=false."
        outSuggestion = "Corrigir decisÃ£o do modelo ou garantir presenÃ§a real do CSV antes de executar LOAD_CSV."
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
        r3State, "Valida citaÃ§Ã£o quando FOUND_FLOW_TEMPLATE_CSV=true.", outDetailJsonCompact)

    If LCase$(foundCsv) = "true" And (Not hasCitation) Then
        If provaHasCsv Then
            outStepState = "OK"
            outRuleCode = "C3_FOUND_WITHOUT_CITATION_WARN"
            outProblem = "CSV foi reportado sem container_file_citation, mas PROVA_CI confirmou FLOW_TEMPLATE.csv."
            outSuggestion = "Pode avanÃ§ar, mas recomenda-se incluir citation para facilitar recolha automÃ¡tica em runs futuros."
        Else
            outStepState = "BLOCKED"
            outRuleCode = "C3_FOUND_WITHOUT_CITATION"
            outProblem = "CSV foi reportado como encontrado mas sem container_file_citation e sem prova equivalente."
            outSuggestion = "Pedir citation explÃ­cita ou evidÃªncia equivalente antes de avanÃ§ar."
            GoTo Finalize
        End If
    End If

    Call ContractDiag_LogRuleResult(passo, promptId, runId, outContractMode, "R4_EXPORT_REQUIRES_PROOF", _
        IIf(LCase$(exportOk) = "true" And (Not provaHasCsv), "FAIL", "PASS"), _
        "Valida que EXPORT_OK_CSV=true esteja comprovado em PROVA_CI.", outDetailJsonCompact)

    If LCase$(exportOk) = "true" And (Not provaHasCsv) Then
        outStepState = "FAIL"
        outRuleCode = "C4_EXPORT_NOT_PROVEN"
        outProblem = "EXPORT_OK_CSV=true, mas PROVA_CI nÃ£o comprova FLOW_TEMPLATE.csv."
        outSuggestion = "Reexecutar com prova de ficheiros no /mnt/data e validar listagem antes do prÃ³ximo passo."
        GoTo Finalize
    End If

    outStepState = "OK"
    outRuleCode = "C9_OK"
    outProblem = "Contrato validado com sucesso para este passo."
    outSuggestion = "Pode avanÃ§ar para o prÃ³ximo passo da pipeline."

Finalize:
    Dim sev As String
    sev = "INFO"
    If outStepState = "BLOCKED" Then sev = "ERRO"
    If outStepState = "FAIL" Then sev = "ERRO"

    Call ContractDiag_LogEvent(passo, promptId, sev, "CONTRACT_STATE_DECISION", runId, outContractMode, outStepState, outRuleCode, outProblem, outSuggestion, outDetailJsonCompact)
    Call ContractDiag_LogEvent(passo, promptId, "INFO", "CONTRACT_NEXT_ACTION", runId, outContractMode, outStepState, outRuleCode, "PrÃ³xima aÃ§Ã£o recomendada definida.", outSuggestion, outDetailJsonCompact)
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
        "Se FAIL, rever marcadores/evidÃªncias deste passo.", detailJson)
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

    Call ContractDiag_EvaluateStep(3, "SELFTEST/C", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: false" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "EXECUTE: LOAD_CSV(FLOW_TEMPLATE.csv)", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "FAIL" Then Err.Raise 5, , "T3 fail: execute com FOUND=false deveria FAIL"

    Call ContractDiag_EvaluateStep(4, "SELFTEST/D", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true" & vbLf & "container_file_citation: sandbox:/mnt/data/FLOW_TEMPLATE.csv", "{""type"":""container_file_citation""}", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T4 fail: cenÃ¡rio consistente deveria OK"

    Call ContractDiag_EvaluateStep(5, "SELFTEST/E", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/FLOW_TEMPLATE.csv" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T5 fail: sem citation mas com prova vÃ¡lida deveria OK (warn)"
    If UCase$(rule) <> "C3_FOUND_WITHOUT_CITATION_WARN" Then Err.Raise 5, , "T5 fail: regra esperada C3_FOUND_WITHOUT_CITATION_WARN"

    Call ContractDiag_EvaluateStep(6, "SELFTEST/F", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/out.pdf" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: true", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "BLOCKED" Then Err.Raise 5, , "T6 fail: found=true sem citation e sem prova csv deveria BLOCKED"

    Call ContractDiag_EvaluateStep(7, "SELFTEST/G", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/out.pdf" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: false" & vbLf & "EXPORT_OK_CSV: false", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "OK" Then Err.Raise 5, , "T7 fail: sem expected files no contrato deveria manter OK"

    Call ContractDiag_EvaluateStep(8, "SELFTEST/H", "diagnostic_contract: ci_csv_v1", "PROVA_CI_START" & vbLf & "FILE: /mnt/data/out.pdf" & vbLf & "PROVA_CI_END" & vbLf & "FOUND_FLOW_TEMPLATE_CSV: true" & vbLf & "EXPORT_OK_CSV: false", "", hasContract, cMode, st, rule, pr, sg, dj)
    If UCase$(st) <> "FAIL" Then Err.Raise 5, , "T8 fail: diff expected-vs-prova deveria FAIL"
    If UCase$(rule) <> "C5_PROVA_EXPECTED_MISSING" Then Err.Raise 5, , "T8 fail: regra esperada C5_PROVA_EXPECTED_MISSING"

    Call Debug_Registar(0, "SELFTEST_CONTRACT", "INFO", "", "SELFTEST_CONTRACT", "PASS", "SelfTest_ContractDiagnostics_RunAll concluÃ­do com sucesso.")
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
