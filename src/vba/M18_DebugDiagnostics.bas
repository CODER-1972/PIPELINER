Attribute VB_Name = "M18_DebugDiagnostics"
Option Explicit

' =============================================================================
' MÃ³dulo: M18_DebugDiagnostics
' PropÃ³sito:
' - Produzir diagnÃ³stico avanÃ§ado em relatÃ³rio paralelo DEBUG_DIAG sem quebrar DEBUG legado.
' - Ler nÃ­veis de diagnÃ³stico (BASE/DIAG/TRACE), classificar causa provÃ¡vel e gerar bundle opcional.
'
' AtualizaÃ§Ãµes:
' - 2026-03-01 | Codex | Introduz DEBUG_SCHEMA_VERSION=2 com DEBUG_LEVEL/DEBUG_BUNDLE
'   - Cria escrita aditiva em DEBUG_DIAG com campos de fingerprint, FILES, CI/container e contrato EXECUTE.
'   - Implementa classificador simples de causa raiz com confidence e aÃ§Ã£o sugerida.
'   - Adiciona bundle opcional por execuÃ§Ã£o com payload/response/extratos sem segredos.
'
' FunÃ§Ãµes e procedimentos:
' - DebugDiag_RecordStep(...)
'   - Regista uma linha em DEBUG_DIAG (apenas quando DEBUG_LEVEL>=DIAG).
' - DebugDiag_ClassifyForSelfTest(...)
'   - Exponibiliza classificador para SelfTests idempotentes.
' =============================================================================

Private Const CFG_DEBUG_LEVEL As String = "DEBUG_LEVEL"
Private Const CFG_DEBUG_BUNDLE As String = "DEBUG_BUNDLE"
Private Const DEBUG_SCHEMA_VERSION As String = "2"
Private Const DEFAULT_DEBUG_LEVEL As String = "BASE"
Private Const DEFAULT_DEBUG_BUNDLE As String = "FALSE"
Private Const DIAG_PREVIEW_LEN As Long = 280

Public Sub DebugDiag_RecordStep( _
    ByVal pipelineName As String, ByVal pipelineIndex As Long, ByVal passo As Long, _
    ByVal promptId As String, ByVal modelName As String, ByVal outputKind As String, _
    ByVal processMode As String, ByVal stageStart As Date, ByVal stageEnd As Date, _
    ByVal elapsedFilesPrepareMs As Long, ByVal elapsedApiCallMs As Long, _
    ByVal elapsedDirectiveParseMs As Long, ByVal inputJsonFinal As String, _
    ByVal filesUsedResumo As String, ByVal filesOpsResumo As String, ByVal fileIds As String, _
    ByVal filesRequestedRaw As String, ByVal resultado As ApiResultado, _
    ByVal promptTextFinal As String, ByVal outputFolderBase As String)

    On Error GoTo EH

    Dim debugLevel As String
    debugLevel = DebugDiag_GetDebugLevel()
    If debugLevel = "BASE" Then Exit Sub

    Dim ws As Worksheet
    Set ws = DebugDiag_EnsureDiagSheet()

    Dim headers As Variant
    headers = DebugDiag_Headers()

    Dim rowMap As Object
    Set rowMap = CreateObject("Scripting.Dictionary")

    Dim fp As String
    fp = pipelineName & "|" & CStr(passo) & "|" & promptId & "|" & Left$(resultado.responseId, 24) & "|" & modelName & "|" & LCase$(Trim$(outputKind)) & "/" & LCase$(Trim$(processMode))

    Dim outputPreview As String
    outputPreview = DebugDiag_Preview(resultado.outputText, DIAG_PREVIEW_LEN)

    Dim filesResolved As String
    filesResolved = Trim$(filesUsedResumo)

    Dim filesEffectiveModes As String
    filesEffectiveModes = DebugDiag_ExtractModesFromOps(filesOpsResumo, inputJsonFinal)

    Dim hasInputFile As String, hasTextEmbed As String, inputFileCount As Long, textEmbedBlocks As Long
    hasInputFile = IIf(InStr(1, inputJsonFinal, """type"":""input_file""", vbTextCompare) > 0 Or InStr(1, inputJsonFinal, """type"":""input_image""", vbTextCompare) > 0, "SIM", "NAO")
    hasTextEmbed = IIf(InStr(1, inputJsonFinal, "----- BEGIN FILE:", vbTextCompare) > 0, "SIM", "NAO")
    inputFileCount = DebugDiag_CountToken(inputJsonFinal, """type"":""input_file""") + DebugDiag_CountToken(inputJsonFinal, """type"":""input_image""")
    textEmbedBlocks = DebugDiag_CountToken(inputJsonFinal, "----- BEGIN FILE:")

    Dim ciExpected As String, ciObserved As String
    ciExpected = IIf(LCase$(Trim$(processMode)) = "code_interpreter", "SIM", "NAO")
    ciObserved = IIf(InStr(1, resultado.rawResponseJson, """type"":""code_interpreter_call""", vbTextCompare) > 0, "SIM", "NAO")

    Dim containerId As String
    containerId = DebugDiag_FirstJsonString(resultado.rawResponseJson, "container_id")

    Dim cTotal As String, cEligible As String, cMatched As String, cFiles As String
    cTotal = "": cEligible = "": cMatched = "": cFiles = ""
    Call DebugDiag_ReadContainerSignals(passo, promptId, cTotal, cEligible, cMatched, cFiles)

    Dim executeCount As Long, executeLines As String
    executeCount = DebugDiag_ExtractExecuteLines(resultado.outputText, executeLines)

    Dim manifestoDetected As String
    manifestoDetected = IIf(DebugDiag_LooksLikeManifest(resultado.outputText), "SIM", "NAO")

    Dim manifestoFields As String
    manifestoFields = DebugDiag_ExtractManifestFields(resultado.outputText)

    Dim outputExpected As String, outputFound As String
    outputExpected = DebugDiag_ExpectedOutputNames(promptTextFinal)
    outputFound = DebugDiag_OutputFoundFromOps(filesOpsResumo)

    Dim warnReqMissing As String
    warnReqMissing = IIf(InStr(1, filesOpsResumo, "missing_required", vbTextCompare) > 0 Or InStr(1, filesOpsResumo, "required", vbTextCompare) > 0 And InStr(1, filesUsedResumo, "NOT_FOUND", vbTextCompare) > 0, "SIM", "NAO")

    Dim rcCode As String, rcSummary As String, rcFix As String, rcConfidence As Long
    Call DebugDiag_ClassifyCore(LCase$(Trim$(processMode)), ciObserved, cTotal, cFiles, filesEffectiveModes, promptTextFinal, executeCount, filesOpsResumo, resultado.rawResponseJson, rcCode, rcSummary, rcFix, rcConfidence)

    rowMap("debug_schema_version") = DEBUG_SCHEMA_VERSION
    rowMap("debug_level") = debugLevel
    rowMap("timestamp") = Now
    rowMap("pipeline_name") = pipelineName
    rowMap("pipeline_index") = pipelineIndex
    rowMap("passo") = passo
    rowMap("prompt_id") = promptId
    rowMap("response_id") = resultado.responseId
    rowMap("http_status") = resultado.httpStatus
    rowMap("fp") = fp
    rowMap("stage_start") = stageStart
    rowMap("stage_end") = stageEnd
    rowMap("elapsed_ms") = CLng((stageEnd - stageStart) * 86400000#)
    rowMap("elapsed_files_prepare_ms") = elapsedFilesPrepareMs
    rowMap("elapsed_api_call_ms") = elapsedApiCallMs
    rowMap("elapsed_ci_retrieve_ms") = 0
    rowMap("elapsed_download_ms") = 0
    rowMap("elapsed_directive_parse_ms") = elapsedDirectiveParseMs
    rowMap("files_requested") = DebugDiag_Preview(filesRequestedRaw, 380)
    rowMap("files_resolved") = DebugDiag_Preview(filesResolved, 380)
    rowMap("files_effective_modes") = filesEffectiveModes
    rowMap("files_required_flags") = DebugDiag_ExtractRequiredFlags(filesRequestedRaw)
    rowMap("files_hash") = ""
    rowMap("size_bytes") = ""
    rowMap("has_input_file") = hasInputFile
    rowMap("has_text_embed") = hasTextEmbed
    rowMap("text_embed_blocks") = textEmbedBlocks
    rowMap("input_file_count") = inputFileCount
    rowMap("file_ids_used") = DebugDiag_Preview(fileIds, 220)
    rowMap("warning_required_unresolved") = warnReqMissing
    rowMap("ci_expected") = ciExpected
    rowMap("ci_observed") = ciObserved
    rowMap("container_id") = containerId
    rowMap("container_list_total") = cTotal
    rowMap("container_list_eligible") = cEligible
    rowMap("container_list_matched") = cMatched
    rowMap("container_files") = DebugDiag_Preview(cFiles, 420)
    rowMap("output_files_expected") = DebugDiag_Preview(outputExpected, 200)
    rowMap("output_files_found") = DebugDiag_Preview(outputFound, 200)
    rowMap("warning_only_inputs_in_container") = IIf(rcCode = "RC_ONLY_INPUTS_IN_CONTAINER" Or rcCode = "RC_NO_CITATION_FALLBACK_GRABBED_INPUT", "SIM", "NAO")
    rowMap("output_text_preview") = outputPreview
    rowMap("manifesto_detected") = manifestoDetected
    rowMap("manifesto_fields") = DebugDiag_Preview(manifestoFields, 300)
    rowMap("execute_directives_found") = executeCount
    rowMap("execute_lines") = DebugDiag_Preview(executeLines, 240)
    rowMap("warning_execute_missing") = IIf(executeCount = 0 And DebugDiag_PromptRequiresExecute(promptTextFinal), "SIM", "NAO")
    rowMap("root_cause_code") = rcCode
    rowMap("root_cause_summary") = rcSummary
    rowMap("suggested_fix") = rcFix
    rowMap("confidence") = rcConfidence

    Call DebugDiag_WriteRow(ws, headers, rowMap)

    If DebugDiag_IsBundleEnabled() Then
        Call DebugDiag_WriteBundle(pipelineName, passo, promptId, resultado.responseId, outputFolderBase, rowMap, resultado)
    End If

    Exit Sub
EH:
    Call Debug_Registar(passo, promptId, "ALERTA", "", "DEBUG_DIAG", "Falha ao registar DEBUG_DIAG: " & Err.Description, "Rever M18_DebugDiagnostics; o DEBUG legado manteve-se operacional.")
End Sub

Public Sub DebugDiag_ClassifyForSelfTest( _
    ByVal processMode As String, ByVal ciObserved As String, ByVal containerTotal As String, _
    ByVal containerFiles As String, ByVal filesEffectiveModes As String, ByVal promptText As String, _
    ByVal executeCount As Long, ByVal filesOpsResumo As String, ByVal rawJson As String, _
    ByRef outCode As String, ByRef outSummary As String, ByRef outFix As String, ByRef outConfidence As Long)

    Call DebugDiag_ClassifyCore(processMode, ciObserved, containerTotal, containerFiles, filesEffectiveModes, promptText, executeCount, filesOpsResumo, rawJson, outCode, outSummary, outFix, outConfidence)
End Sub

Private Sub DebugDiag_ClassifyCore( _
    ByVal processMode As String, ByVal ciObserved As String, ByVal containerTotal As String, _
    ByVal containerFiles As String, ByVal filesEffectiveModes As String, ByVal promptText As String, _
    ByVal executeCount As Long, ByVal filesOpsResumo As String, ByVal rawJson As String, _
    ByRef outCode As String, ByRef outSummary As String, ByRef outFix As String, ByRef outConfidence As Long)

    outCode = ""
    outSummary = "Sem anomalia crÃ­tica detetada no classificador mÃ­nimo."
    outFix = "Sem aÃ§Ã£o imediata."
    outConfidence = 35

    If InStr(1, filesOpsResumo, "PATH_TOO_LONG", vbTextCompare) > 0 Or InStr(1, rawJson, "PATH_TOO_LONG", vbTextCompare) > 0 Then
        outCode = "RC_PATH_TOO_LONG"
        outSummary = "Falha de path demasiado longo no fluxo de output/download."
        outFix = "Reduza comprimento de pasta/nome e mantenha FILE_MAX_PATH_SAFE conservador."
        outConfidence = 95
        Exit Sub
    End If

    If InStr(1, filesOpsResumo, "CONFIG_PARSE", vbTextCompare) > 0 Or InStr(1, filesOpsResumo, "config extra", vbTextCompare) > 0 And InStr(1, filesOpsResumo, "ignorada", vbTextCompare) > 0 Then
        outCode = "RC_CONFIG_PARSE_ERROR"
        outSummary = "Config extra com parse invÃ¡lido afetou o modo efetivo."
        outFix = "Normalizar Config extra para uma linha por chave:valor e evitar literais ambÃ­guos."
        outConfidence = 80
        Exit Sub
    End If

    If LCase$(Trim$(processMode)) = "code_interpreter" And UCase$(Trim$(ciObserved)) <> "SIM" And InStr(1, rawJson, "code_interpreter_call", vbTextCompare) = 0 Then
        outCode = "RC_NO_CI_CALL"
        outSummary = "O passo esperava CI mas a resposta nÃ£o trouxe code_interpreter_call."
        outFix = "ForÃ§ar process_mode: code_interpreter e validar tool_choice/intenÃ§Ã£o explÃ­cita no Config extra."
        outConfidence = 90
        Exit Sub
    End If

    If CLng(Val(containerTotal)) = 1 And DebugDiag_ContainerLooksOnlyInput(containerFiles) Then
        outCode = "RC_ONLY_INPUTS_IN_CONTAINER"
        outSummary = "Fallback do container devolveu apenas ficheiro de input."
        outFix = "ReforÃ§ar contrato de output (manifest/nomes esperados) e padrÃ£o forte de seleÃ§Ã£o."
        outConfidence = 88
        Exit Sub
    End If

    If InStr(1, rawJson, "M10_CI_NO_CITATION", vbTextCompare) > 0 And DebugDiag_ContainerLooksOnlyInput(containerFiles) Then
        outCode = "RC_NO_CITATION_FALLBACK_GRABBED_INPUT"
        outSummary = "Sem citation; fallback escolheu input por ausÃªncia de artefacto novo."
        outFix = "Pedir output determinÃ­stico + EXECUTE explÃ­cito e validar container_file_citation."
        outConfidence = 84
        Exit Sub
    End If

    If InStr(1, LCase$(filesEffectiveModes), "text_embed") > 0 And DebugDiag_PromptTriesOpenFile(promptText) Then
        outCode = "RC_TEXT_EMBED_FILE_NOT_FOUND_RISK"
        outSummary = "Prompt tenta abrir ficheiro local mas anexo foi enviado como text_embed."
        outFix = "Usar input_file/pdf_upload para ficheiros que o CI precisa abrir em /mnt/data."
        outConfidence = 90
        Exit Sub
    End If

    If executeCount = 0 And DebugDiag_PromptRequiresExecute(promptText) Then
        outCode = "RC_EXECUTE_MISSING"
        outSummary = "Contrato pede EXECUTE mas nenhuma diretiva foi detetada no output."
        outFix = "Exigir linha EXECUTE Ãºnica e validÃ¡vel (ex.: EXECUTE: LOAD_CSV <basename.csv>)."
        outConfidence = 87
        Exit Sub
    End If

    If InStr(1, LCase$(filesOpsResumo), "nome", vbTextCompare) > 0 And InStr(1, LCase$(filesOpsResumo), "mismatch", vbTextCompare) > 0 Then
        outCode = "RC_FILENAME_MISMATCH"
        outSummary = "Artefactos gerados com nome fora do padrÃ£o esperado."
        outFix = "Alinhar naming no prompt/manifest e regex de seleÃ§Ã£o de output."
        outConfidence = 72
        Exit Sub
    End If

    If InStr(1, LCase$(filesOpsResumo), "0 ficheiro", vbTextCompare) > 0 And LCase$(Trim$(processMode)) = "code_interpreter" Then
        outCode = "RC_NO_OUTPUT_FILES"
        outSummary = "Passo CI terminou sem artefactos de output elegÃ­veis."
        outFix = "Pedir explicitamente gravaÃ§Ã£o de ficheiros e validar citaÃ§Ãµes/container list."
        outConfidence = 65
        Exit Sub
    End If
End Sub

Private Function DebugDiag_EnsureDiagSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DEBUG_DIAG")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "DEBUG_DIAG"
    End If

    Dim headers As Variant
    headers = DebugDiag_Headers()

    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then
        Dim i As Long
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
    End If

    Set DebugDiag_EnsureDiagSheet = ws
End Function

Private Function DebugDiag_Headers() As Variant
    DebugDiag_Headers = Array( _
        "debug_schema_version", "debug_level", "timestamp", "pipeline_name", "pipeline_index", "passo", "prompt_id", "response_id", "http_status", "fp", _
        "stage_start", "stage_end", "elapsed_ms", "elapsed_files_prepare_ms", "elapsed_api_call_ms", "elapsed_ci_retrieve_ms", "elapsed_download_ms", "elapsed_directive_parse_ms", _
        "files_requested", "files_resolved", "files_effective_modes", "files_required_flags", "files_hash", "size_bytes", "has_input_file", "has_text_embed", "text_embed_blocks", "input_file_count", "file_ids_used", "warning_required_unresolved", _
        "ci_expected", "ci_observed", "container_id", "container_list_total", "container_list_eligible", "container_list_matched", "container_files", "output_files_expected", "output_files_found", "warning_only_inputs_in_container", _
        "output_text_preview", "manifesto_detected", "manifesto_fields", "execute_directives_found", "execute_lines", "warning_execute_missing", _
        "root_cause_code", "root_cause_summary", "suggested_fix", "confidence")
End Function

Private Sub DebugDiag_WriteRow(ByVal ws As Worksheet, ByVal headers As Variant, ByVal valuesMap As Object)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    Dim i As Long, h As String
    For i = LBound(headers) To UBound(headers)
        h = CStr(headers(i))
        If valuesMap.Exists(h) Then ws.Cells(r, i + 1).Value = valuesMap(h)
    Next i
End Sub

Private Function DebugDiag_GetDebugLevel() As String
    Dim raw As String
    raw = UCase$(Trim$(DebugDiag_ConfigGet(CFG_DEBUG_LEVEL, DEFAULT_DEBUG_LEVEL)))
    Select Case raw
        Case "TRACE", "DIAG", "BASE"
            DebugDiag_GetDebugLevel = raw
        Case Else
            DebugDiag_GetDebugLevel = DEFAULT_DEBUG_LEVEL
    End Select
End Function

Private Function DebugDiag_IsBundleEnabled() As Boolean
    Dim raw As String
    raw = UCase$(Trim$(DebugDiag_ConfigGet(CFG_DEBUG_BUNDLE, DEFAULT_DEBUG_BUNDLE)))
    DebugDiag_IsBundleEnabled = (raw = "TRUE" Or raw = "SIM" Or raw = "YES" Or raw = "1")
End Function

Private Function DebugDiag_ConfigGet(ByVal keyName As String, ByVal defaultValue As String) As String
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")

    Dim lastRow As Long, i As Long, k As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        k = Trim$(CStr(ws.Cells(i, 1).Value))
        If StrComp(k, keyName, vbTextCompare) = 0 Then
            DebugDiag_ConfigGet = Trim$(CStr(ws.Cells(i, 2).Value))
            If DebugDiag_ConfigGet = "" Then DebugDiag_ConfigGet = defaultValue
            Exit Function
        End If
    Next i
EH:
    DebugDiag_ConfigGet = defaultValue
End Function

Private Sub DebugDiag_ReadContainerSignals(ByVal passo As Long, ByVal promptId As String, ByRef outTotal As String, ByRef outEligible As String, ByRef outMatched As String, ByRef outFiles As String)
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DEBUG")

    Dim map As Object
    Set map = DebugDiag_MapHeaders(ws)
    If Not map.Exists("passo") Or Not map.Exists("prompt id") Or Not map.Exists("parametro") Or Not map.Exists("problema") Then Exit Sub

    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = lastR To 2 Step -1
        If CLng(Val(ws.Cells(r, map("passo")).Value)) = passo Then
            If StrComp(Trim$(CStr(ws.Cells(r, map("prompt id")).Value)), promptId, vbTextCompare) = 0 Then
                If StrComp(Trim$(CStr(ws.Cells(r, map("parametro")).Value)), "M10_CI_CONTAINER_LIST", vbTextCompare) = 0 Then
                    Dim p As String
                    p = CStr(ws.Cells(r, map("problema")).Value)
                    outTotal = DebugDiag_KeyValueFromText(p, "total")
                    outEligible = DebugDiag_KeyValueFromText(p, "elegÃ­veis")
                    If outEligible = "" Then outEligible = DebugDiag_KeyValueFromText(p, "elegiveis")
                    outMatched = DebugDiag_KeyValueFromText(p, "matched")
                    outFiles = DebugDiag_Preview(p, 400)
                    Exit For
                End If
            End If
        End If
    Next r
    Exit Sub
EH:
End Sub

Private Function DebugDiag_MapHeaders(ByVal ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long, c As Long, h As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = LCase$(Trim$(CStr(ws.Cells(1, c).Value)))
        If h <> "" Then d(h) = c
    Next c
    Set DebugDiag_MapHeaders = d
End Function

Private Function DebugDiag_KeyValueFromText(ByVal s As String, ByVal keyName As String) As String
    Dim p As Long
    p = InStr(1, s, keyName & "=", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(keyName) + 1
    Dim i As Long, ch As String
    For i = p To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "|" Or ch = ";" Then Exit For
        DebugDiag_KeyValueFromText = DebugDiag_KeyValueFromText & ch
    Next i
    DebugDiag_KeyValueFromText = Trim$(DebugDiag_KeyValueFromText)
End Function

Private Function DebugDiag_CountToken(ByVal txt As String, ByVal token As String) As Long
    Dim p As Long, n As Long
    p = 1
    Do
        p = InStr(p, txt, token, vbTextCompare)
        If p = 0 Then Exit Do
        n = n + 1
        p = p + Len(token)
    Loop
    DebugDiag_CountToken = n
End Function

Private Function DebugDiag_FirstJsonString(ByVal json As String, ByVal keyName As String) As String
    Dim pattern As String
    pattern = """" & keyName & """\s*:\s*""([^""]+)"""
    DebugDiag_FirstJsonString = DebugDiag_RegexFirstGroup(json, pattern)
End Function

Private Function DebugDiag_RegexFirstGroup(ByVal textIn As String, ByVal pattern As String) As String
    On Error GoTo EH
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    If re.Test(textIn) Then
        DebugDiag_RegexFirstGroup = CStr(re.Execute(textIn)(0).SubMatches(0))
    End If
    Exit Function
EH:
    DebugDiag_RegexFirstGroup = ""
End Function

Private Function DebugDiag_ExtractExecuteLines(ByVal outputText As String, ByRef outLines As String) As Long
    Dim arr() As String, i As Long, line As String
    If Trim$(outputText) = "" Then Exit Function
    arr = Split(Replace(outputText, vbCrLf, vbLf), vbLf)
    For i = LBound(arr) To UBound(arr)
        line = Trim$(arr(i))
        If UCase$(Left$(line, 8)) = "EXECUTE:" Then
            If outLines <> "" Then outLines = outLines & " || "
            outLines = outLines & DebugDiag_SanitizeExecuteLine(line)
            DebugDiag_ExtractExecuteLines = DebugDiag_ExtractExecuteLines + 1
        End If
    Next i
End Function

Private Function DebugDiag_SanitizeExecuteLine(ByVal line As String) As String
    Dim p As Long
    p = InStrRev(line, "\")
    If p > 0 Then line = Left$(line, p) & "<basename>"
    DebugDiag_SanitizeExecuteLine = line
End Function

Private Function DebugDiag_LooksLikeManifest(ByVal outputText As String) As Boolean
    Dim s As String
    s = LCase$(outputText)
    DebugDiag_LooksLikeManifest = (InStr(1, s, """output_kind""", vbTextCompare) > 0 Or InStr(1, s, "manifest", vbTextCompare) > 0)
End Function

Private Function DebugDiag_ExtractManifestFields(ByVal outputText As String) As String
    Dim keys As Variant
    keys = Array("export_ok_csv", "export_ok_word", "file_csv", "file_word", "delimiter", "cols", "rows")
    Dim i As Long, v As String
    For i = LBound(keys) To UBound(keys)
        v = DebugDiag_FirstJsonString(outputText, CStr(keys(i)))
        If v <> "" Then
            If DebugDiag_ExtractManifestFields <> "" Then DebugDiag_ExtractManifestFields = DebugDiag_ExtractManifestFields & " | "
            DebugDiag_ExtractManifestFields = DebugDiag_ExtractManifestFields & CStr(keys(i)) & "=" & v
        End If
    Next i
End Function

Private Function DebugDiag_ExtractModesFromOps(ByVal filesOps As String, ByVal inputJson As String) As String
    Dim s As String
    s = LCase$(filesOps & "|" & inputJson)
    If InStr(1, s, "text_embed", vbTextCompare) > 0 Then DebugDiag_ExtractModesFromOps = DebugDiag_AppendCSV(DebugDiag_ExtractModesFromOps, "text_embed")
    If InStr(1, s, "as_pdf", vbTextCompare) > 0 Then DebugDiag_ExtractModesFromOps = DebugDiag_AppendCSV(DebugDiag_ExtractModesFromOps, "as_pdf")
    If InStr(1, s, "pdf_upload", vbTextCompare) > 0 Then DebugDiag_ExtractModesFromOps = DebugDiag_AppendCSV(DebugDiag_ExtractModesFromOps, "pdf_upload")
    If InStr(1, s, """type"":""input_file""", vbTextCompare) > 0 Then DebugDiag_ExtractModesFromOps = DebugDiag_AppendCSV(DebugDiag_ExtractModesFromOps, "input_file")
    If DebugDiag_ExtractModesFromOps = "" Then DebugDiag_ExtractModesFromOps = "as_is"
End Function

Private Function DebugDiag_ExtractRequiredFlags(ByVal filesRequestedRaw As String) As String
    Dim s As String
    s = LCase$(filesRequestedRaw)
    If InStr(1, s, "required", vbTextCompare) > 0 Or InStr(1, s, "obrigatorio", vbTextCompare) > 0 Or InStr(1, s, "obrigatoria", vbTextCompare) > 0 Then
        DebugDiag_ExtractRequiredFlags = DebugDiag_AppendCSV(DebugDiag_ExtractRequiredFlags, "required")
    End If
    If InStr(1, s, "latest", vbTextCompare) > 0 Or InStr(1, s, "mais recente", vbTextCompare) > 0 Or InStr(1, s, "mais_recente", vbTextCompare) > 0 Then
        DebugDiag_ExtractRequiredFlags = DebugDiag_AppendCSV(DebugDiag_ExtractRequiredFlags, "latest")
    End If
End Function

Private Function DebugDiag_AppendCSV(ByVal baseCsv As String, ByVal token As String) As String
    If Trim$(baseCsv) = "" Then
        DebugDiag_AppendCSV = token
    ElseIf InStr(1, "," & baseCsv & ",", "," & token & ",", vbTextCompare) > 0 Then
        DebugDiag_AppendCSV = baseCsv
    Else
        DebugDiag_AppendCSV = baseCsv & "," & token
    End If
End Function

Private Function DebugDiag_PromptRequiresExecute(ByVal promptText As String) As Boolean
    DebugDiag_PromptRequiresExecute = (InStr(1, promptText, "EXECUTE:", vbTextCompare) > 0 Or InStr(1, promptText, "LOAD_CSV", vbTextCompare) > 0)
End Function

Private Function DebugDiag_PromptTriesOpenFile(ByVal promptText As String) As Boolean
    Dim s As String
    s = LCase$(promptText)
    DebugDiag_PromptTriesOpenFile = (InStr(1, s, "open(", vbTextCompare) > 0 Or InStr(1, s, "/mnt/data", vbTextCompare) > 0 Or InStr(1, s, "read_csv", vbTextCompare) > 0)
End Function

Private Function DebugDiag_ContainerLooksOnlyInput(ByVal containerFiles As String) As Boolean
    Dim s As String
    s = LCase$(containerFiles)
    If Trim$(s) = "" Then Exit Function
    If InStr(1, s, ".pdf", vbTextCompare) > 0 And InStr(1, s, ".csv", vbTextCompare) = 0 And InStr(1, s, ".docx", vbTextCompare) = 0 Then
        DebugDiag_ContainerLooksOnlyInput = True
    End If
End Function

Private Function DebugDiag_ExpectedOutputNames(ByVal promptText As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([A-Za-z0-9_\-]+\.(csv|docx|xlsx|pptx|pdf))"
    re.Global = True
    re.IgnoreCase = True
    Dim mc As Object, m As Object
    Set mc = re.Execute(promptText)
    For Each m In mc
        DebugDiag_ExpectedOutputNames = DebugDiag_AppendCSV(DebugDiag_ExpectedOutputNames, CStr(m.SubMatches(0)))
    Next m
End Function

Private Function DebugDiag_OutputFoundFromOps(ByVal filesOpsResumo As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([A-Za-z0-9_\-]+\.(csv|docx|xlsx|pptx|pdf))"
    re.Global = True
    re.IgnoreCase = True
    Dim mc As Object, m As Object
    Set mc = re.Execute(filesOpsResumo)
    For Each m In mc
        DebugDiag_OutputFoundFromOps = DebugDiag_AppendCSV(DebugDiag_OutputFoundFromOps, CStr(m.SubMatches(0)))
    Next m
End Function

Private Function DebugDiag_Preview(ByVal txt As String, ByVal maxLen As Long) As String
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    If Len(txt) <= maxLen Then
        DebugDiag_Preview = txt
    Else
        DebugDiag_Preview = Left$(txt, maxLen) & "..."
    End If
End Function

Private Sub DebugDiag_WriteBundle(ByVal pipelineName As String, ByVal passo As Long, ByVal promptId As String, ByVal responseId As String, ByVal outputFolderBase As String, ByVal rowMap As Object, ByVal resultado As ApiResultado)
    On Error GoTo EH

    Dim baseFolder As String
    baseFolder = outputFolderBase
    If Trim$(baseFolder) = "" Then baseFolder = Environ$("TEMP")
    If Right$(baseFolder, 1) = "\" Then baseFolder = Left$(baseFolder, Len(baseFolder) - 1)

    Dim runFolder As String
    runFolder = baseFolder & "\DEBUG_BUNDLE\" & Format$(Now, "yyyymmdd_hhnnss") & "_" & DebugDiag_SafeToken(promptId) & "_" & DebugDiag_SafeToken(Left$(responseId, 14))
    Call DebugDiag_EnsurePath(runFolder)

    Dim payloadPath As String
    payloadPath = "C:\Temp\payload.json"
    If Dir(payloadPath) <> "" Then FileCopy payloadPath, runFolder & "\payload.json"

    Call DebugDiag_WriteText(runFolder & "\response.json", resultado.rawResponseJson)
    Call DebugDiag_WriteText(runFolder & "\extracted_manifest.json", CStr(rowMap("manifesto_fields")))
    Call DebugDiag_WriteText(runFolder & "\extracted_execute.txt", CStr(rowMap("execute_lines")))
    Call DebugDiag_WriteText(runFolder & "\debug_diag_row.tsv", DebugDiag_MapToTsv(rowMap))
    Exit Sub
EH:
End Sub

Private Function DebugDiag_MapToTsv(ByVal rowMap As Object) As String
    Dim headers As Variant
    headers = DebugDiag_Headers()
    Dim i As Long, h As String
    For i = LBound(headers) To UBound(headers)
        h = CStr(headers(i))
        If DebugDiag_MapToTsv <> "" Then DebugDiag_MapToTsv = DebugDiag_MapToTsv & vbTab
        If rowMap.Exists(h) Then
            DebugDiag_MapToTsv = DebugDiag_MapToTsv & Replace(CStr(rowMap(h)), vbTab, " ")
        Else
            DebugDiag_MapToTsv = DebugDiag_MapToTsv & ""
        End If
    Next i
End Function

Private Sub DebugDiag_EnsurePath(ByVal folderPath As String)
    Dim parts() As String
    parts = Split(folderPath, "\")
    Dim i As Long, cur As String
    If InStr(folderPath, ":") > 0 Then
        cur = parts(0) & "\"
        i = 1
    Else
        cur = ""
        i = 0
    End If
    For i = i To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            If Right$(cur, 1) <> "\" Then cur = cur & "\"
            cur = cur & parts(i)
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        End If
    Next i
End Sub

Private Sub DebugDiag_WriteText(ByVal filePath As String, ByVal txt As String)
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, txt
    Close #ff
End Sub

Private Function DebugDiag_SafeToken(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then
            DebugDiag_SafeToken = DebugDiag_SafeToken & ch
        ElseIf ch = "/" Then
            DebugDiag_SafeToken = DebugDiag_SafeToken & "_"
        End If
        If Len(DebugDiag_SafeToken) >= 32 Then Exit For
    Next i
    If DebugDiag_SafeToken = "" Then DebugDiag_SafeToken = "na"
End Function
