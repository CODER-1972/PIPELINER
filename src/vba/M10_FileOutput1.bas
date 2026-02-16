Attribute VB_Name = "M10_FileOutput1"
Option Explicit

' =============================================================================
' Módulo: M10_FileOutput1
' Propósito:
' - Gerir registo e resolução de ficheiros de output produzidos por prompts.
' - Suportar cadeia output->input e escrita de eventos de output no histórico de ficheiros.
'
' Atualizações:
' - 2026-02-16 | Codex | Test macro alinhada com resolução central de API key
'   - Test_FileOutput passa a usar Config_ResolveOpenAIApiKey para evitar dependência direta de Config!B1.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - FileOutput_ResolveEffectiveConfig (Sub): rotina pública do módulo.
' - FileOutput_PrepareRequest (Sub): rotina pública do módulo.
' - FileOutput_ProcessAfterResponse (Function): rotina pública do módulo.
' - Test_FileOutput (Sub): rotina pública do módulo.
' =============================================================================

' ============================================================
' M10_FileOutput
'   Gestão de outputs em ficheiro (metadata + code_interpreter)
'   - Resolve config (FLOW_TEMPLATE > PAINEL > Config)
'   - Prepara request (tools + Structured Outputs quando aplicável)
'   - Grava raw outputs em disco (por step)
'   - Processa manifest metadata/files[] e grava ficheiros
'   - Detecta container_file_citation e descarrega bytes (CI)
'   - Cria sidecar .meta.json por ficheiro
'   - Regista em FILES_MANAGEMENT (via wrapper em M09) e devolve resumo p/ Seguimento
' ============================================================

Private Const SAFE_LIMIT As Long = 32000
Private Const MAX_PATH_SAFE_DEFAULT As Long = 240 ' [POR CONFIRMAR] depende da política de Long Paths no Windows

' Cache simples por pipeline (0..10) para Run_ID
Private gRunId(0 To 10) As String

' -------------------------------
' PUBLIC - Resolver Config (precedência)
' -------------------------------
Public Sub FileOutput_ResolveEffectiveConfig( _
    ByVal pipelineIndex As Long, _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByVal painelAutoSave As String, _
    ByRef out_outputKind As String, _
    ByRef out_processMode As String, _
    ByRef out_autoSave As String, _
    ByRef out_overwriteMode As String, _
    ByRef out_prefixTemplate As String, _
    ByRef out_subfolderTemplate As String, _
    ByRef out_structuredMode As String, _
    ByRef out_pptxMode As String, _
    ByRef out_xlsxMode As String, _
    ByRef out_pdfMode As String, _
    ByRef out_imageMode As String, _
    Optional ByVal promptConfigExtraTexto As String = "" _
)
    On Error GoTo Falha

    ' Defaults globais (Config)
    Dim def_process As String: def_process = LCase$(Config_Get("FILE_OUTPUT_PROCESS_MODE_DEFAULT", "metadata"))
    Dim def_auto As String: def_auto = Config_Get("FILE_AUTO_SAVE_DEFAULT", "Sim")
    Dim def_over As String: def_over = LCase$(Config_Get("FILE_OVERWRITE_MODE_DEFAULT", "suffix"))
    Dim def_pptx As String: def_pptx = LCase$(Config_Get("FILE_PPTX_MODE_DEFAULT", "structure"))
    Dim def_xlsx As String: def_xlsx = LCase$(Config_Get("FILE_XLSX_MODE_DEFAULT", "structure"))
    Dim def_pdf As String: def_pdf = LCase$(Config_Get("FILE_PDF_MODE_DEFAULT", "export"))
    Dim def_img_meta As String: def_img_meta = LCase$(Config_Get("FILE_IMAGE_MODE_METADATA_DEFAULT", "base64"))
    Dim def_img_ci As String: def_img_ci = LCase$(Config_Get("FILE_IMAGE_MODE_CI_DEFAULT", "container_download"))

    ' PAINEL (pipeline-level): apenas auto_save
    Dim painel_auto As String
    painel_auto = painelAutoSave
    If Trim$(painel_auto) = "" Then painel_auto = def_auto

    ' FLOW_TEMPLATE (prompt-level)
    Dim ft As Object
    Set ft = FlowTemplate_GetPromptRow(promptId)

    Dim p_outputKind As String: p_outputKind = ""
    Dim p_processMode As String: p_processMode = ""
    Dim p_autoSave As String: p_autoSave = ""
    Dim p_overwrite As String: p_overwrite = ""
    Dim p_prefix As String: p_prefix = ""
    Dim p_subfolder As String: p_subfolder = ""
    Dim p_structured As String: p_structured = ""
    Dim p_pptx As String: p_pptx = ""
    Dim p_xlsx As String: p_xlsx = ""
    Dim p_pdf As String: p_pdf = ""
    Dim p_img As String: p_img = ""

    If Not ft Is Nothing Then
        p_outputKind = LCase$(CStr(ft("output_kind")))
        p_processMode = LCase$(CStr(ft("process_mode")))
        p_autoSave = CStr(ft("auto_save"))
        p_overwrite = LCase$(CStr(ft("overwrite_mode")))
        p_prefix = CStr(ft("file_name_prefix_template"))
        p_subfolder = CStr(ft("subfolder_template"))
        p_structured = LCase$(CStr(ft("structured_outputs_mode")))
        p_pptx = LCase$(CStr(ft("pptx_mode")))
        p_xlsx = LCase$(CStr(ft("xlsx_mode")))
        p_pdf = LCase$(CStr(ft("pdf_mode")))
        p_img = LCase$(CStr(ft("image_mode")))
    End If

    ' Overrides por Config extra do prompt (prompt-level) - prevalece sobre FLOW_TEMPLATE
    Dim ov As Object
    Set ov = FileOutput_ParseFileOutputKeysFromConfigExtra(promptConfigExtraTexto)

    If Not ov Is Nothing Then
        If ov.exists("output_kind") Then p_outputKind = LCase$(CStr(ov("output_kind")))
        If ov.exists("process_mode") Then p_processMode = LCase$(CStr(ov("process_mode")))
        If ov.exists("auto_save") Then p_autoSave = CStr(ov("auto_save"))
        If ov.exists("overwrite_mode") Then p_overwrite = LCase$(CStr(ov("overwrite_mode")))
        If ov.exists("file_name_prefix_template") Then p_prefix = CStr(ov("file_name_prefix_template"))
        If ov.exists("subfolder_template") Then p_subfolder = CStr(ov("subfolder_template"))
        If ov.exists("structured_outputs_mode") Then p_structured = LCase$(CStr(ov("structured_outputs_mode")))
        If ov.exists("pptx_mode") Then p_pptx = LCase$(CStr(ov("pptx_mode")))
        If ov.exists("xlsx_mode") Then p_xlsx = LCase$(CStr(ov("xlsx_mode")))
        If ov.exists("pdf_mode") Then p_pdf = LCase$(CStr(ov("pdf_mode")))
        If ov.exists("image_mode") Then p_img = LCase$(CStr(ov("image_mode")))
    End If

    ' output_kind: default "text" (inherit -> assume text)
    out_outputKind = "text"
    If p_outputKind = "file" Then out_outputKind = "file"
    If p_outputKind = "text" Then out_outputKind = "text"

    ' process_mode
    out_processMode = def_process
    If p_processMode <> "" And p_processMode <> "inherit" Then out_processMode = p_processMode

    ' auto_save
    out_autoSave = def_auto
    If Trim$(painel_auto) <> "" Then out_autoSave = painel_auto
    If Trim$(p_autoSave) <> "" And LCase$(Trim$(p_autoSave)) <> "inherit" Then out_autoSave = p_autoSave

    ' overwrite_mode
    out_overwriteMode = def_over
    If p_overwrite <> "" And p_overwrite <> "inherit" Then out_overwriteMode = p_overwrite

    out_prefixTemplate = p_prefix
    out_subfolderTemplate = p_subfolder

    out_structuredMode = p_structured
    If out_structuredMode = "" Or out_structuredMode = "inherit" Then out_structuredMode = "off"

    out_pptxMode = def_pptx
    If p_pptx <> "" And p_pptx <> "inherit" Then out_pptxMode = p_pptx

    out_xlsxMode = def_xlsx
    If p_xlsx <> "" And p_xlsx <> "inherit" Then out_xlsxMode = p_xlsx

    out_pdfMode = def_pdf
    If p_pdf <> "" And p_pdf <> "inherit" Then out_pdfMode = p_pdf

    ' image_mode depende do processo
    If out_processMode = "code_interpreter" Then
        out_imageMode = def_img_ci
    Else
        out_imageMode = def_img_meta
    End If
    If p_img <> "" And p_img <> "inherit" Then out_imageMode = p_img

    Exit Sub

Falha:
    out_outputKind = "text"
    out_processMode = "metadata"
    out_autoSave = "Sim"
    out_overwriteMode = "suffix"
    out_prefixTemplate = ""
    out_subfolderTemplate = ""
    out_structuredMode = "off"
    out_pptxMode = "structure"
    out_xlsxMode = "structure"
    out_pdfMode = "export"
    out_imageMode = "base64"
End Sub

' -------------------------------
' PUBLIC - Preparar request (tools + Structured Outputs)
' -------------------------------
Public Sub FileOutput_PrepareRequest( _
    ByVal outputKind As String, _
    ByVal processMode As String, _
    ByVal structuredMode As String, _
    ByRef modos As String, _
    ByRef extraFragment As String _
)
    ' Tools: code_interpreter (M05 trata a injecção via "modos")
    If LCase$(Trim$(processMode)) = "code_interpreter" Then
        If InStr(1, modos, "Code Interpreter", vbTextCompare) = 0 Then
            If Trim$(modos) <> "" Then
                modos = modos & " + Code Interpreter"
            Else
                modos = "Code Interpreter"
            End If
        End If
    End If

    ' Structured Outputs: apenas quando output_kind=file e process_mode=metadata
    If LCase$(Trim$(outputKind)) = "file" And LCase$(Trim$(processMode)) = "metadata" Then
        If LCase$(Trim$(structuredMode)) = "json_schema" Then
            Call ExtraFragment_Append(extraFragment, FileOutput_TextFormat_JsonSchema())
        ElseIf LCase$(Trim$(structuredMode)) = "json_object" Then
            ' [POR CONFIRMAR] suporte exacto de json_object em Responses.text.format
            Call ExtraFragment_Append(extraFragment, FileOutput_TextFormat_JsonObject())
        End If
    End If
End Sub

' -------------------------------
' PUBLIC - Pós-resposta (gravar ficheiros + logs)
'   Devolve texto curto para Seguimento (sem colar JSON/base64 gigantes).
' -------------------------------
Public Function FileOutput_ProcessAfterResponse( _
    ByVal apiKey As String, _
    ByVal outputFolderBase As String, _
    ByVal pipelineNome As String, _
    ByVal pipelineIndex As Long, _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByRef resultado As ApiResultado, _
    ByVal outputKind As String, _
    ByVal processMode As String, _
    ByVal autoSave As String, _
    ByVal overwriteMode As String, _
    ByVal prefixTemplate As String, _
    ByVal subfolderTemplate As String, _
    ByVal pptxMode As String, _
    ByVal xlsxMode As String, _
    ByVal pdfMode As String, _
    ByVal imageMode As String, _
    ByRef outFilesUsed As String, _
    ByRef outFilesOps As String _
) As String
    On Error GoTo Falha
    outFilesUsed = ""
    outFilesOps = ""
    Dim maxPath As Long
    maxPath = FileOutput_MaxPathSafe()
    Dim runId As String
    runId = FileOutput_GetRunId(pipelineIndex)
    Dim runFolder As String
    runFolder = FileOutput_BuildRunFolder(outputFolderBase, pipelineNome, runId)
    Call EnsureFolder(runFolder)
    If Dir(runFolder, vbDirectory) = "" Then
        Call Debug_Registar(passo, promptId, "ERRO", "", "M10_RUNFOLDER", _
            "Não foi possível criar/aceder à pasta de execução: " & runFolder, _
            "Verifique permissões, OneDrive/Sync, e OUTPUT Folder no PAINEL.")
        FileOutput_ProcessAfterResponse = "[ERRO] Não foi possível criar/aceder à pasta de execução: " & runFolder
        Exit Function
    End If
    Dim rawFolder As String
    rawFolder = runFolder & "\_raw"
    Call EnsureFolder(rawFolder)
    If Dir(rawFolder, vbDirectory) = "" Then
        Call Debug_Registar(passo, promptId, "ERRO", "", "M10_RAWFOLDER", _
            "Não foi possível criar/aceder à pasta _raw: " & rawFolder, _
            "Verifique permissões e path demasiado longo.")
        FileOutput_ProcessAfterResponse = "[ERRO] Não foi possível criar/aceder à pasta _raw: " & rawFolder
        Exit Function
    End If
    ' Guardar raw response JSON (sempre, para auditoria)
    Dim rawPath As String, msgRaw As String
    rawPath = rawFolder & "\" & FileOutput_SafeFileName("step_" & Format$(passo, "00") & "_" & Replace(promptId, "/", "_") & "_response.json")
    If FileOutput_PathLenOK(rawPath, maxPath, msgRaw) Then
        Call WriteTextUTF8(rawPath, Nz(resultado.rawResponseJson))
        If Dir(rawPath) = "" Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "M10_RAW_WRITE_FAIL", _
                "Tentativa de guardar rawResponseJson falhou (ficheiro não encontrado após SaveToFile): " & rawPath, _
                "Verifique permissões e/ou path longo.")
        End If
    Else
        Call Debug_Registar(passo, promptId, "ERRO", "", "M10_PATH_TOO_LONG", msgRaw, _
            "Encurtar OUTPUT Folder no PAINEL e/ou reduzir pipeline_name.")
    End If
    If LCase$(Trim$(outputKind)) <> "file" Then
        FileOutput_ProcessAfterResponse = ""
        Exit Function
    End If
    Dim aAuto As String
    aAuto = LCase$(Trim$(autoSave))
    If aAuto = "no" Or aAuto = "nao" Or aAuto = "não" Then
        FileOutput_ProcessAfterResponse = "[FILE OUTPUT] auto_save=Não (config) - raw guardado: " & rawPath
        Exit Function
    End If
    Dim summary As String
    summary = ""
    If LCase$(Trim$(processMode)) = "code_interpreter" Then
        summary = Process_CodeInterpreter(apiKey, runFolder, rawFolder, pipelineNome, promptId, resultado, overwriteMode, prefixTemplate, subfolderTemplate, runId, passo, outFilesUsed, outFilesOps)
        FileOutput_ProcessAfterResponse = summary
        Exit Function
    End If
    ' default: metadata
    summary = Process_Metadata(runFolder, rawFolder, pipelineNome, promptId, resultado, overwriteMode, prefixTemplate, subfolderTemplate, runId, passo, pptxMode, xlsxMode, pdfMode, imageMode, outFilesUsed, outFilesOps)
    FileOutput_ProcessAfterResponse = summary
    Exit Function
Falha:
    On Error Resume Next
    FileOutput_ProcessAfterResponse = "[ERRO] FileOutput_ProcessAfterResponse: " & Err.Description
End Function

' ============================================================
' Implementação - METADATA (manifest JSON)
' ============================================================
Private Function Process_Metadata( _
    ByVal runFolder As String, _
    ByVal rawFolder As String, _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByRef resultado As ApiResultado, _
    ByVal overwriteMode As String, _
    ByVal prefixTemplate As String, _
    ByVal subfolderTemplate As String, _
    ByVal runId As String, _
    ByVal passo As Long, _
    ByVal pptxMode As String, _
    ByVal xlsxMode As String, _
    ByVal pdfMode As String, _
    ByVal imageMode As String, _
    ByRef outFilesUsed As String, _
    ByRef outFilesOps As String _
) As String
    On Error GoTo Falha

    Dim manifestJson As String
    manifestJson = Trim$(Nz(resultado.outputText))

    If manifestJson = "" Then
        Process_Metadata = "[FILE OUTPUT/metadata] Sem output_text."
        Exit Function
    End If

    ' Guardar manifest bruto
    Dim manifestPath As String
    manifestPath = rawFolder & "\" & FileOutput_SafeFileName("step_" & Format$(passo, "00") & "_" & Replace(promptId, "/", "_") & "_manifest.json")
    Call WriteTextUTF8(manifestPath, manifestJson)

    ' Validar output_kind
    Dim okKind As String
    okKind = LCase$(Json_GetString(manifestJson, "output_kind"))
    If okKind <> "file" Then
        Process_Metadata = "[FILE OUTPUT/metadata] output_kind!=" & okKind & " (raw: " & manifestPath & ")"
        Exit Function
    End If

    Dim filesArr As String
    filesArr = Json_GetArrayRaw(manifestJson, "files")
    If Trim$(filesArr) = "" Then
        Process_Metadata = "[FILE OUTPUT/metadata] Sem files[]. (raw: " & manifestPath & ")"
        Exit Function
    End If

    Dim fileObjs As Collection
    Set fileObjs = Json_SplitArrayObjects(filesArr)

    Dim i As Long, n As Long
    n = fileObjs.Count
    If n = 0 Then
        Process_Metadata = "[FILE OUTPUT/metadata] files[] vazio. (raw: " & manifestPath & ")"
        Exit Function
    End If

    Dim savedCount As Long
    savedCount = 0

    For i = 1 To n
        Dim obj As String
        obj = fileObjs(i)

        Dim file_name As String, file_type As String, subFolder As String, payload_kind As String, payload As String
        file_name = Json_GetString(obj, "file_name")
        file_type = LCase$(Json_GetString(obj, "file_type"))
        subFolder = Json_GetString(obj, "subfolder")
        payload_kind = LCase$(Json_GetString(obj, "payload_kind"))
        payload = Json_GetString(obj, "payload") ' já unescaped

        If Trim$(file_name) = "" Then file_name = "output_" & Format$(i, "00")
        file_name = FileOutput_SafeFileName(file_name)

        Dim ext As String
        ext = ""
        If InStr(1, file_name, ".", vbTextCompare) > 0 Then ext = LCase$(Mid$(file_name, InStrRev(file_name, ".") + 1))
        If file_type = "" Then file_type = ext

        Dim folderAbs As String
        folderAbs = FileOutput_ResolveSubfolder(runFolder, pipelineNome, promptId, passo, runId, subFolder, subfolderTemplate)

        Dim prefix As String
        prefix = FileOutput_ResolvePrefix(pipelineNome, promptId, passo, runId, prefixTemplate)
        If Trim$(prefix) <> "" Then
            file_name = FileOutput_SafeFileName(prefix & "__" & file_name)
        End If

        Dim fullPath As String
        fullPath = folderAbs & "\" & file_name
        fullPath = FileOutput_ResolveCollision(fullPath, overwriteMode)

        Dim ok As Boolean
        ok = False

        If file_type = "txt" Or file_type = "md" Or file_type = "json" Then
            Call WriteTextUTF8(fullPath, payload)
            ok = True

        ElseIf file_type = "docx" Then
            ok = CreateDocx_FromText(fullPath, payload)

        ElseIf file_type = "pdf" Then
            If payload_kind = "base64" Then
                Call WriteBinaryFromBase64(fullPath, payload)
                ok = True
            Else
                ok = ExportPdf_FromText(fullPath, payload)
            End If

        ElseIf file_type = "png" Or file_type = "jpg" Or file_type = "jpeg" Or file_type = "gif" Or file_type = "webp" Then
            If payload_kind = "base64" Then
                Call WriteBinaryFromBase64(fullPath, payload)
                ok = True
            End If

        ElseIf file_type = "pptx" Or file_type = "xlsx" Then
            ' [POR CONFIRMAR] estrutura exacta de "structure". Suporta base64 como fallback.
            If payload_kind = "base64" Then
                Call WriteBinaryFromBase64(fullPath, payload)
                ok = True
            End If

        Else
            Call WriteTextUTF8(fullPath, payload)
            ok = True
        End If

        If ok Then
            savedCount = savedCount + 1

            ' sidecar .meta.json
            Call WriteMetaJson(fullPath, pipelineNome, promptId, resultado.responseId, "metadata", overwriteMode, runFolder, runId, passo, "")

            ' FILES_MANAGEMENT (wrapper em M09; chamada best-effort via Application.Run)
            Call Try_Files_LogEventOutput(pipelineNome, promptId, runFolder, fullPath, "output(metadata)", "", "process=metadata", resultado.responseId, runId, passo, i - 1, "OUTPUT")

            ' outputs para Seguimento
            Call AppendList(outFilesUsed, "OUT:" & Replace(fullPath, runFolder & "\", ""))
            Call AppendList(outFilesOps, "SAVE:" & Replace(fullPath, runFolder & "\", ""))
        End If
    Next i

    Process_Metadata = "[FILE OUTPUT/metadata] " & CStr(savedCount) & " ficheiro(s) gravado(s) em " & runFolder & " | raw: " & manifestPath
    Exit Function

Falha:
    Process_Metadata = "[ERRO] Process_Metadata: " & Err.Description
End Function

' ============================================================
' Implementação - CODE_INTERPRETER (container_file_citation)
' ============================================================
Private Function Process_CodeInterpreter( _
    ByVal apiKey As String, _
    ByVal runFolder As String, _
    ByVal rawFolder As String, _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByRef resultado As ApiResultado, _
    ByVal overwriteMode As String, _
    ByVal prefixTemplate As String, _
    ByVal subfolderTemplate As String, _
    ByVal runId As String, _
    ByVal passo As Long, _
    ByRef outFilesUsed As String, _
    ByRef outFilesOps As String _
) As String
    On Error GoTo Falha
    Dim maxPath As Long
    maxPath = FileOutput_MaxPathSafe()
    Dim rawJson As String
    rawJson = Nz(resultado.rawResponseJson)
    If Trim$(rawJson) = "" Then
        Call Debug_Registar(passo, promptId, "ALERTA", "", "M10_CI_RAW_MISSING", _
            "process_mode=code_interpreter mas rawResponseJson está vazio.", _
            "Confirme se M05 está a guardar resultado.rawResponseJson.")
        Process_CodeInterpreter = "[FILE OUTPUT/CI] Sem rawResponseJson."
        Exit Function
    End If
    Dim ciList As Collection
    Set ciList = CI_ExtractCitations(rawJson)
    Dim usedFallback As Boolean
    usedFallback = False
    ' ------------------------------------------------------------
    ' FALLBACK robusto:
    ' - Quando não há container_file_citation, tenta obter container_id
    '   a partir de code_interpreter_call e listar ficheiros no container.
    ' ------------------------------------------------------------
    If ciList.Count = 0 Then
        Call Debug_Registar(passo, promptId, "INFO", "", "M10_CI_NO_CITATION", _
            "Sem container_file_citation; a tentar fallback por listagem do container.", _
            "Se isto acontecer frequentemente, reforce o prompt para anexar ficheiros e/ou force tool_choice.")
        Dim containerFromCall As String
        containerFromCall = CI_ExtractContainerIdFromCall(rawJson)
        If Trim$(containerFromCall) = "" Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "M10_CI_NO_CONTAINER_ID", _
                "Sem container_file_citation e sem container_id em code_interpreter_call.", _
                "Isto sugere que o Code Interpreter não foi realmente executado, ou o formato mudou.")
            Process_CodeInterpreter = "[FILE OUTPUT/CI] Sem container_file_citation e sem container_id (code_interpreter_call)."
            Exit Function
        End If
        Dim listStatus As Long, listJson As String
        Dim files As Collection
        Set files = CI_ListContainerFiles(apiKey, containerFromCall, listStatus, listJson)
        ' Auditoria: guardar listagem do container
        Dim listRawPath As String, msgPath As String
        listRawPath = rawFolder & "\" & FileOutput_SafeFileName("step_" & Format$(passo, "00") & "_" & Replace(promptId, "/", "_") & "_ci_container_files.json")
        If FileOutput_PathLenOK(listRawPath, maxPath, msgPath) Then
            Call WriteTextUTF8(listRawPath, Nz(listJson))
        Else
            Call Debug_Registar(passo, promptId, "ALERTA", "", "M10_PATH_TOO_LONG", msgPath, _
                "Encurtar OUTPUT Folder no PAINEL e/ou reduzir prefix/subfolder.")
        End If
        If listStatus < 200 Or listStatus >= 300 Then
            Call Debug_Registar(passo, promptId, "ERRO", "", "M10_CI_LIST_FAIL", _
                "Falha ao listar ficheiros no container (HTTP " & CStr(listStatus) & "). container_id=" & containerFromCall, _
                "Verifique API key/permissões e se o endpoint /containers/{id}/files está disponível.")
            Process_CodeInterpreter = "[FILE OUTPUT/CI] Fallback list container falhou (HTTP " & CStr(listStatus) & ")."
            Exit Function
        End If
        Dim eligible As Long
        Dim ciList2 As Collection
        Set ciList2 = CI_BuildCitationsFromContainerList(containerFromCall, files, eligible)
        Call Debug_Registar(passo, promptId, "INFO", "", "M10_CI_CONTAINER_LIST", _
            "container_id=" & containerFromCall & " | total=" & CStr(files.Count) & " | elegíveis=" & CStr(eligible), _
            "Elegíveis = extensões típicas (docx/xlsx/pptx/pdf/imagens/txt/csv/...).")
        If ciList2.Count = 0 Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "M10_CI_CONTAINER_EMPTY", _
                "Listagem do container devolveu 0 ficheiros elegíveis para download.", _
                "Confirme se o CI gravou o ficheiro em disco (ex.: /mnt/data/...).")
            Process_CodeInterpreter = "[FILE OUTPUT/CI] Sem container_file_citation; container_id=" & containerFromCall & "; 0 ficheiros elegíveis."
            Exit Function
        End If
        Set ciList = ciList2
        usedFallback = True
    End If
    Dim savedCount As Long
    savedCount = 0
    Dim i As Long
    For i = 1 To ciList.Count
        Dim it As Object
        Set it = ciList(i)
        Dim container_id As String, file_id As String, fileName As String
        container_id = CStr(it("container_id"))
        file_id = CStr(it("file_id"))
        fileName = CStr(it("filename"))
        If fileName = "" Then fileName = "container_file_" & Format$(i, "00")
        fileName = FileOutput_SafeFileName(fileName)
        Dim prefix As String
        prefix = FileOutput_ResolvePrefix(pipelineNome, promptId, passo, runId, prefixTemplate)
        If Trim$(prefix) <> "" Then
            fileName = FileOutput_SafeFileName(prefix & "__" & fileName)
        End If
        Dim folderAbs As String
        folderAbs = FileOutput_ResolveSubfolder(runFolder, pipelineNome, promptId, passo, runId, "", subfolderTemplate)
        If Dir(folderAbs, vbDirectory) = "" Then
            Call Debug_Registar(passo, promptId, "ERRO", "", "M10_FOLDER_CREATE_FAIL", _
                "Pasta de output não existe (falha MkDir/Permissões?): " & folderAbs, _
                "Verifique permissões e OUTPUT Folder no PAINEL.")
            GoTo ProximoFicheiro
        End If
        Dim fullPath As String
        fullPath = folderAbs & "\" & fileName
        fullPath = FileOutput_ResolveCollision(fullPath, overwriteMode)
        Dim msg As String
        If Not FileOutput_PathLenOK(fullPath, maxPath, msg) Then
            Call Debug_Registar(passo, promptId, "ERRO", "", "M10_PATH_TOO_LONG", msg, _
                "Encurtar OUTPUT Folder no PAINEL e/ou reduzir prefix/subfolder.")
            GoTo ProximoFicheiro
        End If
        Dim dlStatus As Long, dlErr As String
        Dim ok As Boolean
        ok = DownloadContainerFileEx(apiKey, container_id, file_id, fullPath, dlStatus, dlErr)
        If Not ok Then
            Call Debug_Registar(passo, promptId, "ERRO", "", "M10_CI_DOWNLOAD_FAIL", _
                "Falha download (HTTP " & CStr(dlStatus) & ") " & container_id & ":" & file_id & " -> " & fullPath & IIf(Trim$(dlErr) <> "", " | " & dlErr, ""), _
                "Verifique API key, permissões, conectividade e path local.")
            GoTo ProximoFicheiro
        End If
        If Not FileOutput_FileExists(fullPath) Then
            Call Debug_Registar(passo, promptId, "ERRO", "", "M10_CI_DOWNLOAD_NOFILE", _
                "Download reportado como OK, mas o ficheiro não existe no disco: " & fullPath, _
                "Possível falha de permissões/OneDrive/antivírus/paths longos.")
            GoTo ProximoFicheiro
        End If
        Dim bytesLen As Double
        bytesLen = -1
        On Error Resume Next
        bytesLen = FileLen(fullPath)
        On Error GoTo Falha
        If bytesLen = 0 Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "M10_CI_ZERO_BYTES", _
                "Ficheiro gravado com 0 bytes: " & fullPath, _
                "Confirme se o ficheiro no container tem conteúdo.")
        End If
        savedCount = savedCount + 1
        ' Sidecar meta (best-effort; se path demasiado longo, regista alerta e segue)
        Dim metaPath As String, msgMeta As String
        metaPath = fullPath & ".meta.json"
        If FileOutput_PathLenOK(metaPath, maxPath, msgMeta) Then
            Call WriteMetaJson(fullPath, pipelineNome, promptId, resultado.responseId, "code_interpreter", overwriteMode, runFolder, runId, passo, container_id & ":" & file_id)
        Else
            Call Debug_Registar(passo, promptId, "ALERTA", "", "M10_META_PATH_TOO_LONG", msgMeta, _
                "Encurtar OUTPUT Folder / prefix / subfolder.")
        End If
        Dim notes As String
        notes = "process=code_interpreter"
        If usedFallback Then notes = notes & ";fallback=container_list"
        Call Try_Files_LogEventOutput(pipelineNome, promptId, runFolder, fullPath, "output(code_interpreter)", "DL", notes, resultado.responseId, runId, passo, i - 1, "OUTPUT")
        Call AppendList(outFilesUsed, "OUT:" & Replace(fullPath, runFolder & "\", ""))
        Call AppendList(outFilesOps, "DL:" & Replace(fullPath, runFolder & "\", ""))
ProximoFicheiro:
    Next i
    If usedFallback Then
        Process_CodeInterpreter = "[FILE OUTPUT/CI] (fallback list container) " & CStr(savedCount) & " ficheiro(s) descarregado(s) para " & runFolder
    Else
        Process_CodeInterpreter = "[FILE OUTPUT/CI] " & CStr(savedCount) & " ficheiro(s) descarregado(s) para " & runFolder
    End If
    Exit Function
Falha:
    Process_CodeInterpreter = "[ERRO] Process_CodeInterpreter: " & Err.Description
End Function

' ============================================================
' Structured Outputs - text.format fragments
' ============================================================
Private Function FileOutput_TextFormat_JsonSchema() As String
    Dim schema As String
    schema = FileOutput_ManifestJsonSchema()

    FileOutput_TextFormat_JsonSchema = _
        """text"":{""format"":{""type"":""json_schema"",""name"":""file_manifest"",""schema"":" & schema & ",""strict"":true}}"
End Function

Private Function FileOutput_TextFormat_JsonObject() As String
    FileOutput_TextFormat_JsonObject = """text"":{""format"":{""type"":""json_object""}}"
End Function

Private Function FileOutput_ManifestJsonSchema() As String
    FileOutput_ManifestJsonSchema = _
        "{""type"":""object"",""additionalProperties"":false,""properties"":{" & _
        """output_kind"":{""type"":""string"",""enum"":[""file""]}," & _
        """files"":{""type"":""array"",""items"":{""type"":""object"",""additionalProperties"":false,""properties"":{" & _
            """file_name"":{""type"":""string""}," & _
            """file_type"":{""type"":""string""}," & _
            """subfolder"":{""type"":""string""}," & _
            """payload_kind"":{""type"":""string"",""enum"":[""text"",""markdown"",""structure"",""base64""]}," & _
            """payload"":{""type"":""string""}" & _
        "},""required"":[""file_name"",""file_type"",""payload_kind"",""payload""]}}}" & _
        "},""required"":[""output_kind"",""files""]}"
End Function

Private Sub ExtraFragment_Append(ByRef extraFragment As String, ByVal fragmentSemChavesExternas As String)
    Dim f As String
    f = Trim$(fragmentSemChavesExternas)
    If f = "" Then Exit Sub

    ' O M05 faz: json = json & "," & extraFragment
    ' Logo: aqui NÃO queremos que extraFragment comece por vírgula.
    Dim e As String
    e = Trim$(extraFragment)

    If Left$(e, 1) = "," Then e = Mid$(e, 2)
    If Right$(e, 1) = "," Then e = Left$(e, Len(e) - 1)

    If Left$(f, 1) = "," Then f = Mid$(f, 2)
    If Right$(f, 1) = "," Then f = Left$(f, Len(f) - 1)

    If Trim$(e) = "" Then
        extraFragment = f
    Else
        extraFragment = e & "," & f
    End If
End Sub

' ============================================================

' FLOW_TEMPLATE leitura (por Prompt ID)

' ============================================================

Private Function FlowTemplate_GetPromptRow(ByVal promptId As String) As Object

    On Error GoTo Falha



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("FLOW_TEMPLATE")



    Dim mapa As Object

    Set mapa = HeaderMap(ws, 1)



    ' Suporta 2 layouts:
    '   A) Tabela simples (coluna "Prompt ID" + colunas dedicadas output_kind/process_mode/...)
    '   B) Layout catálogo/blocos (coluna "ID" + coluna "Config extra" com linhas chave: valor)

    Dim colId As Long

    Dim modoTabelaSimples As Boolean

    modoTabelaSimples = False



    If mapa.exists("Prompt ID") Then

        colId = CLng(mapa("Prompt ID"))

        modoTabelaSimples = True

    ElseIf mapa.exists("ID") Then

        colId = CLng(mapa("ID"))

        modoTabelaSimples = False

    Else

        Set FlowTemplate_GetPromptRow = Nothing

        Exit Function

    End If



    Dim lr As Long

    lr = ws.Cells(ws.rowS.Count, colId).End(xlUp).Row



    Dim r As Long

    For r = 2 To lr

        If Trim$(CStr(ws.Cells(r, colId).value)) = Trim$(promptId) Then

            Dim d As Object

            Set d = CreateObject("Scripting.Dictionary")



            If modoTabelaSimples Then

                d("output_kind") = Flow_GetCell(ws, mapa, r, "output_kind")

                d("process_mode") = Flow_GetCell(ws, mapa, r, "process_mode")

                d("auto_save") = Flow_GetCell(ws, mapa, r, "auto_save")

                d("overwrite_mode") = Flow_GetCell(ws, mapa, r, "overwrite_mode")

                d("file_name_prefix_template") = Flow_GetCell(ws, mapa, r, "file_name_prefix_template")

                d("subfolder_template") = Flow_GetCell(ws, mapa, r, "subfolder_template")

                d("pptx_mode") = Flow_GetCell(ws, mapa, r, "pptx_mode")

                d("xlsx_mode") = Flow_GetCell(ws, mapa, r, "xlsx_mode")

                d("pdf_mode") = Flow_GetCell(ws, mapa, r, "pdf_mode")

                d("image_mode") = Flow_GetCell(ws, mapa, r, "image_mode")

                d("structured_outputs_mode") = Flow_GetCell(ws, mapa, r, "structured_outputs_mode")

            Else

                ' Layout catálogo/blocos: ler da coluna "Config extra" e fazer parse só às chaves de File Output
                Dim colCfg As Long

                colCfg = 0

                If mapa.exists("Config extra") Then

                    colCfg = CLng(mapa("Config extra"))

                ElseIf mapa.exists("Config extra (amigável)") Then

                    colCfg = CLng(mapa("Config extra (amigável)"))

                End If



                Dim cfg As String

                cfg = ""

                If colCfg > 0 Then cfg = CStr(ws.Cells(r, colCfg).value)



                Dim ov As Object

                Set ov = FileOutput_ParseFileOutputKeysFromConfigExtra(cfg)



                d("output_kind") = ""

                d("process_mode") = ""

                d("auto_save") = ""

                d("overwrite_mode") = ""

                d("file_name_prefix_template") = ""

                d("subfolder_template") = ""

                d("pptx_mode") = ""

                d("xlsx_mode") = ""

                d("pdf_mode") = ""

                d("image_mode") = ""

                d("structured_outputs_mode") = ""



                If Not ov Is Nothing Then

                    If ov.exists("output_kind") Then d("output_kind") = CStr(ov("output_kind"))

                    If ov.exists("process_mode") Then d("process_mode") = CStr(ov("process_mode"))

                    If ov.exists("auto_save") Then d("auto_save") = CStr(ov("auto_save"))

                    If ov.exists("overwrite_mode") Then d("overwrite_mode") = CStr(ov("overwrite_mode"))

                    If ov.exists("file_name_prefix_template") Then d("file_name_prefix_template") = CStr(ov("file_name_prefix_template"))

                    If ov.exists("subfolder_template") Then d("subfolder_template") = CStr(ov("subfolder_template"))

                    If ov.exists("pptx_mode") Then d("pptx_mode") = CStr(ov("pptx_mode"))

                    If ov.exists("xlsx_mode") Then d("xlsx_mode") = CStr(ov("xlsx_mode"))

                    If ov.exists("pdf_mode") Then d("pdf_mode") = CStr(ov("pdf_mode"))

                    If ov.exists("image_mode") Then d("image_mode") = CStr(ov("image_mode"))

                    If ov.exists("structured_outputs_mode") Then d("structured_outputs_mode") = CStr(ov("structured_outputs_mode"))

                End If

            End If



            Set FlowTemplate_GetPromptRow = d

            Exit Function

        End If

    Next r



    Set FlowTemplate_GetPromptRow = Nothing

    Exit Function



Falha:

    Set FlowTemplate_GetPromptRow = Nothing

End Function


Private Function Flow_GetCell(ByVal ws As Worksheet, ByVal mapa As Object, ByVal r As Long, ByVal header As String) As String
    On Error GoTo Falha
    If Not mapa.exists(header) Then
        Flow_GetCell = ""
        Exit Function
    End If
    Flow_GetCell = Trim$(CStr(ws.Cells(r, CLng(mapa(header))).value))
    Exit Function
Falha:
    Flow_GetCell = ""
End Function

' ============================================================

' Parser mínimo (File Output): extrair chaves internas a partir de "Config extra"

'   - Aceita linhas "chave: valor" (uma por linha)

'   - Ignora comentários (# ou // no início da linha)

'   - Devolve apenas as chaves relevantes para File Output

' ============================================================

Private Function FileOutput_ParseFileOutputKeysFromConfigExtra(ByVal configExtraTexto As String) As Object

    On Error GoTo Falha



    Dim t As String

    t = CStr(configExtraTexto)



    t = Replace(t, vbCrLf, vbLf)

    t = Replace(t, vbCr, vbLf)



    ' Se vier colado com "\n" literal (sem quebras reais), converter para vbLf
    If InStr(1, t, vbLf) = 0 Then

        If InStr(1, t, "\n") > 0 Then t = Replace(t, "\n", vbLf)

    End If



    If Trim$(t) = "" Then

        Set FileOutput_ParseFileOutputKeysFromConfigExtra = Nothing

        Exit Function

    End If



    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")

    d.CompareMode = vbTextCompare



    Dim linhas() As String

    linhas = Split(t, vbLf)



    Dim i As Long

    For i = 0 To UBound(linhas)

        Dim linha As String

        linha = Trim$(CStr(linhas(i)))



        If linha = "" Then GoTo ProxLinha

        If Left$(linha, 1) = "#" Then GoTo ProxLinha

        If Left$(linha, 2) = "//" Then GoTo ProxLinha



        Dim k As String, v As String

        If Not FO_SplitPrimeiro(linha, ":", k, v) Then GoTo ProxLinha



        k = LCase$(Trim$(k))

        v = Trim$(v)



        If FO_IsFileOutputKey(k) Then d(k) = v



ProxLinha:

    Next i



    If d.Count = 0 Then

        Set FileOutput_ParseFileOutputKeysFromConfigExtra = Nothing

    Else

        Set FileOutput_ParseFileOutputKeysFromConfigExtra = d

    End If



    Exit Function



Falha:

    Set FileOutput_ParseFileOutputKeysFromConfigExtra = Nothing

End Function



Private Function FO_IsFileOutputKey(ByVal k As String) As Boolean

    Dim kk As String

    kk = LCase$(Trim$(k))



    FO_IsFileOutputKey = (kk = "output_kind" Or kk = "process_mode" Or kk = "auto_save" Or kk = "overwrite_mode" Or kk = "file_name_prefix_template" Or kk = "pptx_mode" Or kk = "xlsx_mode" Or kk = "pdf_mode" Or kk = "image_mode" Or kk = "structured_outputs_mode")


End Function



Private Function FO_SplitPrimeiro(ByVal texto As String, ByVal separador As String, ByRef outK As String, ByRef outV As String) As Boolean

    Dim p As Long

    p = InStr(1, texto, separador)



    If p = 0 Then

        FO_SplitPrimeiro = False

        Exit Function

    End If



    outK = Trim$(Left$(texto, p - 1))

    outV = Trim$(Mid$(texto, p + Len(separador)))



    FO_SplitPrimeiro = (outK <> "")

End Function


' ============================================================
' Config helpers
' ============================================================
Private Function Config_Get(ByVal key As String, ByVal defaultValue As String) As String
    On Error GoTo Falha
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")

    Dim lr As Long
    lr = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 1 To lr
        If Trim$(CStr(ws.Cells(r, 1).value)) = Trim$(key) Then
            Config_Get = Trim$(CStr(ws.Cells(r, 2).value))
            If Config_Get = "" Then Config_Get = defaultValue
            Exit Function
        End If
    Next r

Falha:
    Config_Get = defaultValue
End Function
Private Function FileOutput_MaxPathSafe() As Long
    ' [POR CONFIRMAR] O limite real pode variar (política de Long Paths no Windows).
    ' Usamos um default conservador para reduzir falhas de I/O em VBA/Office.
    ' Pode ser parametrizado em Config: FILE_MAX_PATH_SAFE (número).
    On Error GoTo Falha
    Dim s As String
    s = Config_Get("FILE_MAX_PATH_SAFE", CStr(MAX_PATH_SAFE_DEFAULT))

    If IsNumeric(s) Then
        FileOutput_MaxPathSafe = CLng(s)
    Else
        FileOutput_MaxPathSafe = MAX_PATH_SAFE_DEFAULT
    End If
    Exit Function

Falha:
    FileOutput_MaxPathSafe = MAX_PATH_SAFE_DEFAULT
End Function

Private Function FileOutput_PathLenOK(ByVal fullPath As String, ByVal maxLen As Long, ByRef outMsg As String) As Boolean
    Dim L As Long
    L = Len(fullPath)

    If L <= maxLen Then
        FileOutput_PathLenOK = True
    Else
        outMsg = "Caminho demasiado longo (" & CStr(L) & " chars), limite=" & CStr(maxLen) & ": " & fullPath
        FileOutput_PathLenOK = False
    End If
End Function

Private Function FileOutput_FileExists(ByVal fullPath As String) As Boolean
    On Error GoTo Falha
    FileOutput_FileExists = (Dir(fullPath) <> "")
    Exit Function
Falha:
    FileOutput_FileExists = False
End Function


' ============================================================
' Naming, folders, collisions, placeholders
' ============================================================
Private Function FileOutput_GetRunId(ByVal pipelineIndex As Long) As String
    If pipelineIndex < 0 Or pipelineIndex > 10 Then pipelineIndex = 0

    If Trim$(gRunId(pipelineIndex)) <> "" Then
        FileOutput_GetRunId = gRunId(pipelineIndex)
        Exit Function
    End If

    Randomize
    gRunId(pipelineIndex) = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(1000 + Rnd() * 8999), "0000")
    FileOutput_GetRunId = gRunId(pipelineIndex)
End Function

Private Function FileOutput_BuildRunFolder(ByVal outputFolderBase As String, ByVal pipelineNome As String, ByVal runId As String) As String
    Dim strat As String
    strat = LCase$(Config_Get("FILE_FOLDER_STRATEGY_DEFAULT", "by_pipeline_run"))

    Dim baseOut As String
    baseOut = outputFolderBase
    If Trim$(baseOut) = "" Then baseOut = ThisWorkbook.path

    baseOut = Replace(baseOut, "/", "\")
    If Right$(baseOut, 1) = "\" Then baseOut = Left$(baseOut, Len(baseOut) - 1)

    pipelineNome = FileOutput_SafeFolderName(pipelineNome)

    If strat = "by_pipeline_run" Then
        FileOutput_BuildRunFolder = baseOut & "\" & pipelineNome & "\" & Format$(Date, "yyyy-mm-dd") & "\Run_" & runId
    ElseIf strat = "root" Or strat = "flat_root" Then
        FileOutput_BuildRunFolder = baseOut
    Else
        FileOutput_BuildRunFolder = baseOut & "\" & pipelineNome
    End If

End Function

Private Function FileOutput_ResolvePrefix(ByVal pipelineNome As String, ByVal promptId As String, ByVal passo As Long, ByVal runId As String, ByVal tpl As String) As String
    Dim t As String
    t = Trim$(CStr(tpl))
    If t = "" Or LCase$(t) = "inherit" Then
        FileOutput_ResolvePrefix = ""
        Exit Function
    End If
    FileOutput_ResolvePrefix = FileOutput_ApplyPlaceholders(t, pipelineNome, promptId, passo, runId)
End Function

Private Function FileOutput_ResolveSubfolder(ByVal runFolder As String, ByVal pipelineNome As String, ByVal promptId As String, ByVal passo As Long, ByVal runId As String, ByVal subFolder As String, ByVal tpl As String) As String
    Dim rel As String
    rel = ""

    If Trim$(subFolder) <> "" Then
        rel = FileOutput_ApplyPlaceholders(subFolder, pipelineNome, promptId, passo, runId)
    ElseIf Trim$(tpl) <> "" And LCase$(Trim$(tpl)) <> "inherit" Then
        rel = FileOutput_ApplyPlaceholders(tpl, pipelineNome, promptId, passo, runId)
    End If

    rel = Replace(rel, "/", "\")
    rel = Trim$(rel)

    ' segurança: bloquear path traversal e absolutos
    If rel <> "" Then
        If InStr(1, rel, "..", vbTextCompare) > 0 Then rel = ""
        If InStr(1, rel, ":\", vbTextCompare) > 0 Then rel = ""
        If Left$(rel, 1) = "\" Then rel = ""
    End If

    Dim target As String
    target = runFolder
    If rel <> "" Then
        rel = FileOutput_SafeFolderPath(rel)
        target = runFolder & "\" & rel
    End If

    Call EnsureFolder(target)
    FileOutput_ResolveSubfolder = target
End Function

Private Function FileOutput_ApplyPlaceholders(ByVal tpl As String, ByVal pipelineNome As String, ByVal promptId As String, ByVal passo As Long, ByVal runId As String) As String
    Dim s As String
    s = tpl
    s = Replace(s, "{PIPELINE}", pipelineNome)
    s = Replace(s, "{PROMPT_ID}", Replace(promptId, "/", "_"))
    s = Replace(s, "{STEP}", Format$(passo, "00"))
    s = Replace(s, "{YYYYMMDD}", Format$(Date, "yyyymmdd"))
    s = Replace(s, "{HHMMSS}", Format$(Time, "hhnnss"))
    s = Replace(s, "{RUN_ID}", runId)
    s = Replace(s, "{USER}", Environ$("USERNAME"))
    FileOutput_ApplyPlaceholders = s
End Function

Private Function FileOutput_ResolveCollision(ByVal fullPath As String, ByVal overwriteMode As String) As String
    overwriteMode = LCase$(Trim$(overwriteMode))
    If overwriteMode = "" Then overwriteMode = "suffix"

    If Dir(fullPath) = "" Then
        FileOutput_ResolveCollision = fullPath
        Exit Function
    End If

    If overwriteMode = "overwrite" Then
        FileOutput_ResolveCollision = fullPath
        Exit Function
    End If

    If overwriteMode = "fail" Then
        Err.Raise vbObjectError + 513, "FileOutput_ResolveCollision", "Ficheiro já existe e overwrite_mode=fail: " & fullPath
    End If

    ' suffix
    Dim base As String, ext As String
    base = fullPath
    ext = ""
    If InStrRev(fullPath, ".") > InStrRev(fullPath, "\") Then
        ext = Mid$(fullPath, InStrRev(fullPath, "."))
        base = Left$(fullPath, Len(fullPath) - Len(ext))
    End If

    Dim i As Long
    For i = 1 To 999
        Dim cand As String
        cand = base & "_" & Format$(i, "000") & ext
        If Dir(cand) = "" Then
            FileOutput_ResolveCollision = cand
            Exit Function
        End If
    Next i

    FileOutput_ResolveCollision = fullPath
End Function

Private Function FileOutput_SafeFolderName(ByVal s As String) As String
    s = Trim$(CStr(s))
    If s = "" Then s = "PIPELINE"
    s = FileOutput_SafeCommon(s)
    FileOutput_SafeFolderName = s
End Function

Private Function FileOutput_SafeFolderPath(ByVal rel As String) As String
    Dim parts() As String
    parts = Split(rel, "\")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = FileOutput_SafeFolderName(parts(i))
    Next i
    FileOutput_SafeFolderPath = Join(parts, "\")
End Function

Private Function FileOutput_SafeFileName(ByVal s As String) As String
    s = Trim$(CStr(s))
    If s = "" Then s = "output"
    s = Replace(s, "/", "_")
    s = FileOutput_SafeCommon(s)

    ' não terminar em ponto/espaço
    Do While Right$(s, 1) = "." Or Right$(s, 1) = " "
        s = Left$(s, Len(s) - 1)
        If Len(s) = 0 Then s = "output": Exit Do
    Loop

    ' bloquear nomes reservados
    If FileOutput_IsReservedName(s) Then s = "_" & s

    FileOutput_SafeFileName = s
End Function

Private Function FileOutput_SafeCommon(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    FileOutput_SafeCommon = Trim$(s)
End Function

Private Function FileOutput_IsReservedName(ByVal s As String) As Boolean
    Dim base As String
    base = UCase$(s)
    If InStr(base, ".") > 0 Then base = Left$(base, InStr(base, ".") - 1)

    Dim reserved As Variant
    reserved = Array("CON", "PRN", "AUX", "NUL", _
                     "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
                     "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
    Dim i As Long
    For i = LBound(reserved) To UBound(reserved)
        If base = reserved(i) Then
            FileOutput_IsReservedName = True
            Exit Function
        End If
    Next i
    FileOutput_IsReservedName = False
End Function

' ============================================================
' IO helpers (UTF-8 + base64 + Office automation)
' ============================================================
Private Sub EnsureFolder(ByVal folderPath As String)
    On Error GoTo Falha
    Dim p As String
    p = Replace(folderPath, "/", "\")
    If p = "" Then Exit Sub

    Dim parts() As String
    parts = Split(p, "\")
    Dim cur As String
    Dim i As Long

    If InStr(p, ":\") > 0 Then
        cur = parts(0) & "\"
        i = 1
    Else
        cur = parts(0)
        i = 1
    End If

    For i = i To UBound(parts)
        If cur = "" Then
            cur = parts(i)
        Else
            If Right$(cur, 1) <> "\" Then cur = cur & "\"
            cur = cur & parts(i)
        End If
        If cur <> "" Then
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        End If
    Next i
    Exit Sub
Falha:
End Sub

Private Sub WriteTextUTF8(ByVal path As String, ByVal content As String)
    On Error GoTo Falha
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' text
    st.Charset = Config_Get("FILE_OUTPUT_ENCODING", "utf-8")
    st.Open
    st.WriteText content
    st.Position = 0
    st.SaveToFile path, 2 ' overwrite
    st.Close
    Exit Sub
Falha:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Sub

Private Sub WriteBinaryFromBase64(ByVal path As String, ByVal b64 As String)
    On Error GoTo Falha
    Dim xml As Object, node As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = b64

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1 ' binary
    st.Open
    st.Write node.nodeTypedValue
    st.Position = 0
    st.SaveToFile path, 2
    st.Close
    Exit Sub
Falha:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Sub

Private Function CreateDocx_FromText(ByVal fullPath As String, ByVal txt As String) As Boolean
    On Error GoTo Falha
    Dim wdApp As Object, doc As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set doc = wdApp.Documents.Add
    doc.content.text = txt
    doc.SaveAs2 fullPath, 16 ' wdFormatDocumentDefault
    doc.Close False
    wdApp.Quit
    CreateDocx_FromText = True
    Exit Function
Falha:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    CreateDocx_FromText = False
End Function

Private Function ExportPdf_FromText(ByVal pdfPath As String, ByVal txt As String) As Boolean
    On Error GoTo Falha
    Dim wdApp As Object, doc As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set doc = wdApp.Documents.Add
    doc.content.text = txt
    doc.ExportAsFixedFormat pdfPath, 17 ' wdExportFormatPDF
    doc.Close False
    wdApp.Quit
    ExportPdf_FromText = True
    Exit Function
Falha:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    ExportPdf_FromText = False
End Function

' ============================================================
' CODE_INTERPRETER download
' ============================================================
Private Function DownloadContainerFileEx(ByVal apiKey As String, ByVal containerId As String, ByVal fileId As String, ByVal savePath As String, ByRef outHttpStatus As Long, ByRef outErrText As String) As Boolean
    On Error GoTo Falha
    outHttpStatus = 0
    outErrText = ""

    Dim url As String
    url = "https://api.openai.com/v1/containers/" & containerId & "/files/" & fileId & "/content"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.Send

    outHttpStatus = http.status

    If outHttpStatus < 200 Or outHttpStatus >= 300 Then
        On Error Resume Next
        outErrText = http.ResponseText
        DownloadContainerFileEx = False
        Exit Function
    End If

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write http.ResponseBody
    st.Position = 0
    st.SaveToFile savePath, 2
    st.Close

    DownloadContainerFileEx = True
    Exit Function

Falha:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    If outHttpStatus = 0 Then outHttpStatus = -1
    If outErrText = "" Then outErrText = Err.Description
    DownloadContainerFileEx = False
End Function

' Wrapper (mantido por compatibilidade interna)
Private Function DownloadContainerFile(ByVal apiKey As String, ByVal containerId As String, ByVal fileId As String, ByVal savePath As String) As Boolean
    Dim st As Long, errT As String
    DownloadContainerFile = DownloadContainerFileEx(apiKey, containerId, fileId, savePath, st, errT)
End Function

Private Function CI_ExtractCitations(ByVal rawJson As String) As Collection
    Set CI_ExtractCitations = New Collection
    On Error GoTo Falha

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    Dim pos As Long
    pos = 1

    Do
        pos = InStr(pos, rawJson, """type"":""container_file_citation""", vbTextCompare)
        If pos = 0 Then Exit Do

        Dim window As String
        window = Mid$(rawJson, pos, 1200)

        Dim containerId As String, fileId As String, fileName As String
        containerId = Regex_FirstGroup(re, window, """container_id""\s*:\s*""([^""]+)""")
        fileId = Regex_FirstGroup(re, window, """file_id""\s*:\s*""([^""]+)""")
        fileName = Regex_FirstGroup(re, window, """filename""\s*:\s*""([^""]+)""")

        If containerId <> "" And fileId <> "" Then
            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            d("container_id") = containerId
            d("file_id") = fileId
            d("filename") = fileName
            CI_ExtractCitations.Add d
        End If

        pos = pos + 10
    Loop

    Exit Function
Falha:
End Function
Private Function CI_ExtractContainerIdFromCall(ByVal rawJson As String) As String
    On Error GoTo Falha

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    Dim pos As Long
    pos = InStr(1, rawJson, """type"":""code_interpreter_call""", vbTextCompare)

    ' fallback: procurar primeiro container_id disponível
    If pos = 0 Then
        pos = InStr(1, rawJson, """container_id""", vbTextCompare)
        If pos = 0 Then Exit Function
    End If

    Dim window As String
    window = Mid$(rawJson, pos, 2000)

    CI_ExtractContainerIdFromCall = Regex_FirstGroup(re, window, """container_id""\s*:\s*""([^""]+)""")
    Exit Function

Falha:
    CI_ExtractContainerIdFromCall = ""
End Function

Private Function CI_ListContainerFiles(ByVal apiKey As String, ByVal containerId As String, ByRef outHttpStatus As Long, ByRef outJson As String) As Collection
    Set CI_ListContainerFiles = New Collection
    outHttpStatus = 0
    outJson = ""
    On Error GoTo Falha

    Dim url As String
    url = "https://api.openai.com/v1/containers/" & containerId & "/files"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.Send

    outHttpStatus = http.status
    outJson = http.ResponseText

    If outHttpStatus < 200 Or outHttpStatus >= 300 Then
        Exit Function
    End If

    Dim txt As String
    txt = outJson
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    Dim ms As Object

    ' padrão principal (id + object=container_file + path)
    re.pattern = """id""\s*:\s*""([^""]+)""[\s\S]{0,200}?""object""\s*:\s*""container\.file""[\s\S]{0,800}?""path""\s*:\s*""([^""]+)"""
    Set ms = re.Execute(txt)

    Dim m As Object
    For Each m In ms
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("file_id") = CStr(m.SubMatches(0))
        d("path") = CStr(m.SubMatches(1))
        CI_ListContainerFiles.Add d
    Next m

    Exit Function

Falha:
End Function

Private Function CI_BuildCitationsFromContainerList(ByVal containerId As String, ByVal files As Collection, ByRef outEligible As Long) As Collection
    Set CI_BuildCitationsFromContainerList = New Collection
    outEligible = 0

    On Error GoTo Falha

    Dim i As Long
    For i = 1 To files.Count
        Dim it As Object
        Set it = files(i)

        Dim p As String
        p = ""
        On Error Resume Next
        p = CStr(it("path"))
        On Error GoTo Falha

        If CI_ShouldDownloadPath(p) Then
            outEligible = outEligible + 1

            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            d("container_id") = containerId
            d("file_id") = CStr(it("file_id"))
            d("filename") = CI_PathBaseName(p)
            CI_BuildCitationsFromContainerList.Add d
        End If
    Next i

    Exit Function

Falha:
End Function

Private Function CI_PathBaseName(ByVal p As String) As String
    p = Replace(CStr(p), "\", "/")
    If InStrRev(p, "/") > 0 Then
        CI_PathBaseName = Mid$(p, InStrRev(p, "/") + 1)
    Else
        CI_PathBaseName = p
    End If
End Function

Private Function CI_ShouldDownloadPath(ByVal p As String) As Boolean
    CI_ShouldDownloadPath = False

    p = Trim$(CStr(p))
    If p = "" Then Exit Function

    Dim base As String
    base = CI_PathBaseName(p)
    base = Trim$(base)
    If base = "" Then Exit Function
    If Left$(base, 1) = "." Then Exit Function

    Dim dot As Long
    dot = InStrRev(base, ".")
    If dot = 0 Then Exit Function

    Dim ext As String
    ext = LCase$(Mid$(base, dot + 1))

    Select Case ext
        Case "docx", "pptx", "xlsx", "pdf", "txt", "md", "csv", "png", "jpg", "jpeg", "webp", "gif", "zip", "json"
            CI_ShouldDownloadPath = True
    End Select
End Function


Private Function Regex_FirstGroup(ByVal re As Object, ByVal text As String, ByVal pattern As String) As String
    On Error GoTo Falha
    re.pattern = pattern
    Dim m As Object
    Set m = re.Execute(text)
    If m.Count = 0 Then
        Regex_FirstGroup = ""
    Else
        Regex_FirstGroup = CStr(m(0).SubMatches(0))
    End If
    Exit Function
Falha:
    Regex_FirstGroup = ""
End Function

' ============================================================
' Sidecar meta.json
' ============================================================
Private Sub WriteMetaJson( _
    ByVal fullPath As String, _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByVal responseId As String, _
    ByVal mode As String, _
    ByVal overwriteMode As String, _
    ByVal runFolder As String, _
    ByVal runId As String, _
    ByVal passo As Long, _
    ByVal extraRef As String _
)
    On Error GoTo Falha

    Dim metaPath As String
    metaPath = fullPath & ".meta.json"

    Dim sha As String
    sha = Try_SHA256_File(fullPath) ' best-effort

    Dim relPath As String
    relPath = Replace(fullPath, runFolder & "\", "")

    Dim j As String
    j = "{"
    j = j & """pipeline_name"":" & Json_Q(pipelineNome) & ","
    j = j & """prompt_id"":" & Json_Q(promptId) & ","
    j = j & """response_id"":" & Json_Q(responseId) & ","
    j = j & """timestamp"":" & Json_Q(Format$(Now, "yyyy-mm-ddThh:nn:ss")) & ","
    j = j & """mode"":" & Json_Q(mode) & ","
    j = j & """overwrite_mode"":" & Json_Q(overwriteMode) & ","
    j = j & """run_id"":" & Json_Q(runId) & ","
    j = j & """step"":" & CStr(passo) & ","
    j = j & """relative_path"":" & Json_Q(relPath) & ","
    j = j & """full_path"":" & Json_Q(fullPath) & ","
    j = j & """sha256"":" & Json_Q(sha)

    If Trim$(extraRef) <> "" Then
        j = j & ",""source_ref"":" & Json_Q(extraRef)
    End If

    j = j & "}"

    Call WriteTextUTF8(metaPath, j)
    Exit Sub
Falha:
End Sub

Private Function Try_SHA256_File(ByVal fullPath As String) As String
    ' Best-effort: tenta usar função existente no projecto (ex.: em M09).
    On Error Resume Next
    Dim v As Variant
    v = Application.Run("Files_SHA256_File", fullPath)
    If Err.Number <> 0 Then
        Try_SHA256_File = ""
        Err.Clear
    Else
        Try_SHA256_File = CStr(v)
    End If
    On Error GoTo 0
End Function

Private Sub Try_Files_LogEventOutput( _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByVal runFolder As String, _
    ByVal fullPath As String, _
    ByVal usageMode As String, _
    ByVal op As String, _
    ByVal notes As String, _
    ByVal responseId As String, _
    Optional ByVal runId As String = "", _
    Optional ByVal stepN As Long = 0, _
    Optional ByVal outputIndex As Long = -1, _
    Optional ByVal sourceType As String = "OUTPUT" _
)
    ' Best-effort: evita "Sub or Function not defined" se o wrapper ainda não existir/importado.
    On Error Resume Next
    Application.Run "Files_LogEventOutput", pipelineNome, promptId, runFolder, fullPath, usageMode, op, notes, responseId, runId, stepN, outputIndex, sourceType
    On Error GoTo 0
End Sub

Private Function Json_Q(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    Json_Q = """" & s & """"
End Function

' ============================================================
' JSON helpers (mínimos; suficiente para manifest determinístico)
' ============================================================
Private Function Json_GetString(ByVal json As String, ByVal key As String) As String
    On Error GoTo Falha

    Dim pos As Long
    pos = InStr(1, json, """" & key & """", vbTextCompare)
    If pos = 0 Then
        Json_GetString = ""
        Exit Function
    End If

    pos = InStr(pos, json, ":", vbTextCompare)
    If pos = 0 Then
        Json_GetString = ""
        Exit Function
    End If
    pos = pos + 1

    Do While pos <= Len(json)
        Dim ch As String
        ch = Mid$(json, pos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        pos = pos + 1
    Loop

    If pos > Len(json) Then
        Json_GetString = ""
        Exit Function
    End If

    If Mid$(json, pos, 1) <> """" Then
        Json_GetString = ""
        Exit Function
    End If

    Dim raw As String
    raw = Json_ReadQuoted(json, pos)

    Json_GetString = Json_Unescape(raw)
    Exit Function

Falha:
    Json_GetString = ""
End Function

Private Function Json_GetArrayRaw(ByVal json As String, ByVal key As String) As String
    On Error GoTo Falha

    Dim pos As Long
    pos = InStr(1, json, """" & key & """", vbTextCompare)
    If pos = 0 Then Json_GetArrayRaw = "": Exit Function

    pos = InStr(pos, json, ":", vbTextCompare)
    If pos = 0 Then Json_GetArrayRaw = "": Exit Function
    pos = pos + 1

    Do While pos <= Len(json)
        Dim ch As String
        ch = Mid$(json, pos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        pos = pos + 1
    Loop

    If pos > Len(json) Then Json_GetArrayRaw = "": Exit Function
    If Mid$(json, pos, 1) <> "[" Then Json_GetArrayRaw = "": Exit Function

    Dim endPos As Long
    endPos = Json_FindMatching(json, pos, "[", "]")
    If endPos = 0 Then Json_GetArrayRaw = "": Exit Function

    Json_GetArrayRaw = Mid$(json, pos, endPos - pos + 1)
    Exit Function

Falha:
    Json_GetArrayRaw = ""
End Function

Private Function Json_SplitArrayObjects(ByVal arrJson As String) As Collection
    Set Json_SplitArrayObjects = New Collection
    On Error GoTo Falha

    Dim pos As Long
    pos = 1

    Do While pos <= Len(arrJson)
        Dim ch As String
        ch = Mid$(arrJson, pos, 1)

        If ch = "{" Then
            Dim endPos As Long
            endPos = Json_FindMatching(arrJson, pos, "{", "}")
            If endPos = 0 Then Exit Do
            Json_SplitArrayObjects.Add Mid$(arrJson, pos, endPos - pos + 1)
            pos = endPos + 1
        Else
            pos = pos + 1
        End If
    Loop

    Exit Function
Falha:
End Function

Private Function Json_FindMatching(ByVal s As String, ByVal startPos As Long, ByVal openCh As String, ByVal closeCh As String) As Long
    Dim depth As Long
    depth = 0

    ' IMPORTANTE: não usar nome "inStr" (colide semanticamente com a função InStr)
    Dim inString As Boolean
    inString = False

    Dim i As Long
    i = startPos

    Do While i <= Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch = """" Then
            If Not Json_IsEscaped(s, i) Then
                inString = Not inString
            End If
        End If

        If Not inString Then
            If ch = openCh Then
                depth = depth + 1
            ElseIf ch = closeCh Then
                depth = depth - 1
                If depth = 0 Then
                    Json_FindMatching = i
                    Exit Function
                End If
            End If
        End If

        i = i + 1
    Loop

    Json_FindMatching = 0
End Function

Private Function Json_IsEscaped(ByVal s As String, ByVal pos As Long) As Boolean
    Dim cnt As Long
    cnt = 0

    Dim i As Long
    i = pos - 1
    Do While i >= 1 And Mid$(s, i, 1) = "\"
        cnt = cnt + 1
        i = i - 1
    Loop

    Json_IsEscaped = (cnt Mod 2 = 1)
End Function

Private Function Json_ReadQuoted(ByVal s As String, ByVal startQuotePos As Long) As String
    Dim i As Long
    i = startQuotePos + 1

    Dim out As String
    out = ""

    Do While i <= Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch = """" And Not Json_IsEscaped(s, i) Then
            Exit Do
        End If

        out = out & ch
        i = i + 1
    Loop

    Json_ReadQuoted = out
End Function

Private Function Json_Unescape(ByVal s As String) As String
    ' Unescape robusto (evita erros de Replace em sequências como \\n)
    Dim i As Long
    i = 1

    Dim res As String
    res = ""

    Do While i <= Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch <> "\" Then
            res = res & ch
            i = i + 1
        Else
            If i = Len(s) Then Exit Do

            Dim nxt As String
            nxt = Mid$(s, i + 1, 1)

            Select Case nxt
                Case """": res = res & """": i = i + 2
                Case "\": res = res & "\": i = i + 2
                Case "/": res = res & "/": i = i + 2
                Case "n": res = res & vbLf: i = i + 2
                Case "r": res = res & vbCr: i = i + 2
                Case "t": res = res & vbTab: i = i + 2
                Case "b": res = res & Chr$(8): i = i + 2
                Case "f": res = res & Chr$(12): i = i + 2

                Case "u"
                    ' \uXXXX
                    If i + 5 <= Len(s) Then
                        Dim hex4 As String
                        hex4 = Mid$(s, i + 2, 4)
                        If IsHex4(hex4) Then
                            res = res & ChrW$(CLng("&H" & hex4))
                            i = i + 6
                        Else
                            ' fallback: mantém literal
                            res = res & "\u"
                            i = i + 2
                        End If
                    Else
                        res = res & "\u"
                        i = i + 2
                    End If

                Case Else
                    res = res & nxt
                    i = i + 2
            End Select
        End If
    Loop

    Json_Unescape = res
End Function

Private Function IsHex4(ByVal texto As String) As Boolean
    Dim i As Long
    If Len(texto) <> 4 Then IsHex4 = False: Exit Function

    For i = 1 To 4
        Dim c As String
        c = Mid$(texto, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", c, vbBinaryCompare) = 0 Then
            IsHex4 = False
            Exit Function
        End If
    Next i

    IsHex4 = True
End Function

' ============================================================
' Util
' ============================================================
Private Function Nz(ByVal v As Variant) As String
    If IsError(v) Then
        Nz = ""
    ElseIf IsMissing(v) Then
        Nz = ""
    ElseIf IsNull(v) Then
        Nz = ""
    Else
        Nz = CStr(v)
    End If
End Function

Private Sub AppendList(ByRef cur As String, ByVal item As String)
    If Trim$(item) = "" Then Exit Sub
    If Trim$(cur) = "" Then
        cur = item
    Else
        cur = cur & "; " & item
    End If
End Sub

Private Function HeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim lc As Long
    lc = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lc
        Dim h As String
        h = Trim$(CStr(ws.Cells(headerRow, c).value))
        If h <> "" Then d(h) = c
    Next c

    Set HeaderMap = d
End Function

' ============================================================
' TESTES (macro)
' ============================================================
Public Sub Test_FileOutput()
    On Error GoTo Falha

    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets("PAINEL")

    Dim outputFolderBase As String
    outputFolderBase = Trim$(CStr(wsP.Range("B3").value))
    If outputFolderBase = "" Then
        MsgBox "PAINEL!B3 (OUTPUT Folder) está vazio. Preenche e volta a executar.", vbExclamation
        Exit Sub
    End If

    Dim apiKey As String
    Dim apiKeySource As String
    Dim apiKeyAlert As String
    Dim apiKeyError As String

    If Not Config_ResolveOpenAIApiKey(apiKey, apiKeySource, apiKeyAlert, apiKeyError) Then
        MsgBox "OPENAI_API_KEY ausente para Test_FileOutput: " & apiKeyError, vbExclamation
        Exit Sub
    End If

    Dim pipelineNome As String: pipelineNome = "TEST"
    Dim passo As Long: passo = 1
    Dim promptId As String: promptId = "TEST/FileOutput"

    Dim res As ApiResultado
    res.httpStatus = 200
    res.responseId = "resp_TEST"
    res.rawResponseJson = "{""id"":""resp_TEST"",""output"":[]}"
    res.Erro = ""

    Dim manifest As String
    manifest = "{""output_kind"":""file"",""files"":[" & _
        "{""file_name"":""exemplo.txt"",""file_type"":""txt"",""payload_kind"":""text"",""payload"":""Olá\nMundo""}," & _
        "{""file_name"":""exemplo.docx"",""file_type"":""docx"",""payload_kind"":""text"",""payload"":""Documento Word de teste""}" & _
        "]}"

    res.outputText = manifest

    Dim filesUsed As String, filesOps As String
    filesUsed = "": filesOps = ""

    Dim logTxt As String
    logTxt = FileOutput_ProcessAfterResponse(apiKey, outputFolderBase, pipelineNome, 0, passo, promptId, res, _
        "file", "metadata", "Sim", "suffix", "{PIPELINE}_{PROMPT_ID}_{STEP}_{YYYYMMDD}_{HHMMSS}", "docs", _
        "structure", "structure", "export", "base64", filesUsed, filesOps)

    ' overwrite_mode=suffix: correr 2x para forçar _001
    Dim logTxt2 As String
    logTxt2 = FileOutput_ProcessAfterResponse(apiKey, outputFolderBase, pipelineNome, 0, passo, promptId, res, _
        "file", "metadata", "Sim", "suffix", "{PIPELINE}_{PROMPT_ID}_{STEP}_{YYYYMMDD}_{HHMMSS}", "docs", _
        "structure", "structure", "export", "base64", filesUsed, filesOps)

    ' Teste Seguimento: output longo -> múltiplas linhas sem truncagem (M02)
    Dim big As String
    big = String$(SAFE_LIMIT + 5000, "A") & String$(SAFE_LIMIT + 5000, "B")

    Dim p As PromptDefinicao
    p.Id = "TEST/LongOutput"
    p.textoPrompt = "(teste)"
    Call Seguimento_Registar(1, p, "TEST_MODEL", "", 200, "resp_TEST_LONG", big, "TEST", "", "", "", "")

    MsgBox "Test_FileOutput concluído." & vbCrLf & _
           "- Ver pasta OUTPUT do PAINEL (subpasta TEST\...\Run_*) " & vbCrLf & _
           "- Ver FILES_MANAGEMENT (novas entradas no topo)" & vbCrLf & _
           "- Ver Seguimento (linhas de continuação).", vbInformation
    Exit Sub

Falha:
    MsgBox "Test_FileOutput falhou: " & Err.Description, vbCritical
End Sub





