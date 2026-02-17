Attribute VB_Name = "M13_ContextKV"
Option Explicit

' =============================================================================
' Módulo: M13_ContextKV
' Propósito:
' - Capturar variáveis chave-valor do output e injetá-las em prompts seguintes.
' - Resolver placeholders/directivas de contexto com observabilidade e fallback seguro.
'
' Atualizações:
' - 2026-02-17 | Codex | Correcao de escape/unescape JSON no ContextKV
'   - Corrige Replace com padroes invertidos que removiam barras e neutralizavam aspas no modulo.
'   - Alinha ContextKV_JsonEscape/ContextKV_JsonUnescape com convencao usada em outros modulos.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - ContextKV_EnsureLayout (Sub): rotina pública do módulo.
' - ContextKV_InjectForStep (Function): rotina pública do módulo.
' - ContextKV_WriteInjectedVars (Sub): rotina pública do módulo.
' - ContextKV_CaptureRow (Sub): rotina pública do módulo.
' - SelfTest_ContextKV_Parse_RESULTS_JSON (Sub): rotina pública do módulo.
' - SelfTest_ContextKV_Placeholder_Replace (Sub): rotina pública do módulo.
' - SelfTest_ContextKV_FileFallback (Sub): rotina pública do módulo.
' - SelfTest_ContextKV_OutputRef (Sub): rotina pública do módulo.
' - SelfTest_RunAll_ContextKV (Sub): rotina pública do módulo.
' =============================================================================
' =============================================================================
' M11_ContextKV
' Captura + Injeção de Variáveis (Key-Value) usando a folha "Seguimento"
'
' Objectivos:
' 1) CAPTURA: Após cada passo, extrair blocos (Key-Value) do Output (texto)
' registado no Seguimento e guardar em:
' - captured_vars
' - captured_vars_meta
' 2) INJEÇÃO: Antes do passo seguinte, se a prompt pedir explicitamente:
' - {{VAR:NOME}}
' - VARS: NOME1, NOME2
' - {@OUTPUT: "Prompt Anterior" | "Todas as prompts" | "<Prompt ID>"}
' então resolver valores a partir do Seguimento e injectar no input final.
' 3) Observabilidade: logging na folha DEBUG (sem dumps grandes nem segredos).
'
' Compatibilidade:
' - Não altera comportamento de pipelines antigos, porque só injecta quando há
' placeholders/directivas no texto da prompt (ou VARS no INPUTS).
' - CAPTURA escreve apenas colunas novas no Seguimento (não mexe no Output).
'
' NOTA: Este módulo não depende de bibliotecas externas (late-binding para RegExp
' e ADODB.Stream).
' =============================================================================
Private Const CONTEXTKV_DEFAULT_MAX_CELL_CHARS As Long = 32000
Private Const CONTEXTKV_DEFAULT_SUBFOLDER As String = "vars"
' Marcador usado pelo PIPELINER para outputs longos (ver Seguimento_EscreverOutputSemTruncagem)
Private Const MARKER_FULL_OUTPUT_SAVED As String = "[[FULL_OUTPUT_SAVED:"
Private Function ContextKV_Fence() As String
' Devolve uma "fence" de 3 backticks (sem os escrever literalmente no código)
ContextKV_Fence = String$(3, ChrW(96))
End Function
' -----------------------------
' API pública (chamada por M07)
' -----------------------------
Public Sub ContextKV_EnsureLayout()
On Error Resume Next
ContextKV_EnsureSeguimentoColumns
ContextKV_EnsureDebugColumns
ContextKV_EnsureHistoricoColumns
On Error GoTo 0
End Sub
Public Function ContextKV_InjectForStep( _
ByVal pipelineName As String, _
ByVal stepN As Long, _
ByVal promptId As String, _
ByVal outputFolderBase As String, _
ByVal runToken As String, _
ByRef promptText As String, _
ByRef outInjectedVarsJson As String, _
ByRef outErro As String _
) As Boolean
Dim strictMode As Boolean
strictMode = ContextKV_GetBoolConfig("CONTEXT_KV_STRICT", False)

outErro = ""
outInjectedVarsJson = ""
ContextKV_InjectForStep = True

If Not ContextKV_GetBoolConfig("CONTEXT_KV_ENABLED", True) Then Exit Function

Call ContextKV_EnsureLayout

Dim wsS As Worksheet
Set wsS = ContextKV_GetSheet("Seguimento")
If wsS Is Nothing Then Exit Function

Dim dictNeeded As Object
Set dictNeeded = CreateObject("Scripting.Dictionary") ' key -> True

Dim dictInjected As Object
Set dictInjected = CreateObject("Scripting.Dictionary") ' key -> json fragment

Dim dictValuesCache As Object
Set dictValuesCache = CreateObject("Scripting.Dictionary") ' key -> value (string)

' 1) VARS: ... (no texto da prompt) - extrai e remove a directiva do texto
Dim varsFromPrompt As Object
Set varsFromPrompt = CreateObject("Scripting.Dictionary")
ContextKV_ExtractVarsDirectiveKeys promptText, varsFromPrompt, True ' remove do texto

' 2) VARS: ... (no INPUTS: da prompt no catálogo)
Dim inputsText As String
inputsText = ContextKV_TryReadInputsTextByPromptId(promptId)
Dim varsFromInputs As Object
Set varsFromInputs = CreateObject("Scripting.Dictionary")
If Trim$(inputsText) <> "" Then
    ContextKV_ExtractVarsDirectiveKeys inputsText, varsFromInputs, False
End If

Dim k As Variant
For Each k In varsFromPrompt.keys
    If Not dictNeeded.exists(k) Then dictNeeded.Add k, True
Next k
For Each k In varsFromInputs.keys
    If Not dictNeeded.exists(k) Then dictNeeded.Add k, True
Next k

' 3) Placeholders {{VAR:KEY}}
Dim placeholders As Collection
Set placeholders = ContextKV_FindVarPlaceholders(promptText)

If placeholders.Count > 0 Then
    ContextKV_LogEvent pipelineName, stepN, promptId, "PLACEHOLDER_FOUND", "", "INFO", "Encontrados " & CStr(placeholders.Count) & " placeholders {{VAR:...}}."
End If

Dim i As Long
For i = 1 To placeholders.Count
    Dim ph As Variant
    ph = placeholders(i) ' Array(key, firstIndex0based, length)
    Dim keyName As String
    keyName = CStr(ph(0))
    If Not dictNeeded.exists(keyName) Then dictNeeded.Add keyName, True
Next i

' 4) Referências {@OUTPUT: ...}
Dim outputRefs As Collection
Set outputRefs = ContextKV_FindOutputRefs(promptText)

If outputRefs.Count > 0 Then
    ContextKV_LogEvent pipelineName, stepN, promptId, "OUTPUT_REF_FOUND", "", "INFO", "Encontradas " & CStr(outputRefs.Count) & " referências {@OUTPUT: ...}."
End If

' Se não há nada para injectar, sai (compatibilidade)
If dictNeeded.Count = 0 And outputRefs.Count = 0 Then Exit Function

ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_START", "", "INFO", "Início da injeção (ContextKV)."

' 5) Resolver variáveis pedidas (dictNeeded) e fazer substituições:
'    5.1) substituir placeholders no texto
'    5.2) se vierem de VARS: sem placeholder, anexar no fim como blocos determinísticos
'
' Primeiro, resolver valores (cache) para todas as keys pedidas
For Each k In dictNeeded.keys
    Dim val As String, srcStep As Long, srcPrompt As String, srcHow As String, sha As String, errMsg As String
    val = ContextKV_ResolveVarValue(pipelineName, stepN, CStr(k), outputFolderBase, srcStep, srcPrompt, srcHow, sha, errMsg)
    
    If val = "" And errMsg <> "" Then
        ' Variável pedida mas não encontrada
        ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_MISS", CStr(k), IIf(strictMode, "ERRO", "ALERTA"), errMsg
        Call ContextKV_DictSet(dictInjected, CStr(k), ContextKV_JsonObj_KVError(errMsg))
        
        If strictMode Then
            outErro = "Variável exigida em falta: " & CStr(k) & " | " & errMsg
            ContextKV_InjectForStep = False
            Exit Function
        End If
        
        ' Cache com vazio (para substituição)
        If Not dictValuesCache.exists(CStr(k)) Then dictValuesCache.Add CStr(k), ""
        
    Else
        If Not dictValuesCache.exists(CStr(k)) Then dictValuesCache.Add CStr(k), val
        
        Dim srcInfo As String
        srcInfo = "from_step=" & CStr(srcStep) & "; from_prompt=" & srcPrompt & "; source=" & srcHow
        
        ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_OK", CStr(k), "INFO", srcInfo
        
        Call ContextKV_DictSet(dictInjected, CStr(k), ContextKV_JsonObj_Source(srcStep, srcPrompt, sha, srcHow))
    End If
Next k

' 5.1) Substituir placeholders em reverse (para não estragar índices)
If placeholders.Count > 0 Then
    ' Ordenar por posição desc (já vêm por ordem de varrimento; fazemos reverse)
    For i = placeholders.Count To 1 Step -1
        ph = placeholders(i)
        keyName = CStr(ph(0))
        Dim idx0 As Long: idx0 = CLng(ph(1))
        Dim ln As Long: ln = CLng(ph(2))
        
        Dim rep As String
        If dictValuesCache.exists(keyName) Then
            rep = CStr(dictValuesCache(keyName))
        Else
            rep = ""
        End If
        
        promptText = ContextKV_ReplaceSpan(promptText, idx0, ln, rep)
        ContextKV_LogEvent pipelineName, stepN, promptId, "PLACEHOLDER_REPLACED", keyName, "INFO", "Placeholder substituído (" & CStr(Len(rep)) & " chars)."
    Next i
End If

' 5.2) Injectar VARS (sem placeholder) no fim, como blocos determinísticos
Dim appended As String
appended = ""

For Each k In varsFromPrompt.keys
    keyName = CStr(k)
    If ContextKV_PromptHasPlaceholderForKey(placeholders, keyName) Then
        ' já foi injectado no lugar
    Else
        appended = appended & ContextKV_FormatVarBlock(keyName, CStr(dictValuesCache(keyName))) & vbCrLf & vbCrLf
    End If
Next k

For Each k In varsFromInputs.keys
    keyName = CStr(k)
    If ContextKV_PromptHasPlaceholderForKey(placeholders, keyName) Then
    Else
        If appended = "" Or InStr(1, appended, "{VAR:" & keyName & "}", vbTextCompare) = 0 Then
            appended = appended & ContextKV_FormatVarBlock(keyName, CStr(dictValuesCache(keyName))) & vbCrLf & vbCrLf
        End If
    End If
Next k

If Trim$(appended) <> "" Then
    promptText = promptText & vbCrLf & vbCrLf & appended
End If

' 6) Processar {@OUTPUT: ...} - substituir no texto por blocos de output
If outputRefs.Count > 0 Then
    Dim outIdx As Long
    outIdx = 0
    
    For i = outputRefs.Count To 1 Step -1
        Dim ref As Variant
        ref = outputRefs(i) ' Array(arg, firstIndex0based, length)
        
        Dim arg As String
        arg = Trim$(CStr(ref(0)))
        idx0 = CLng(ref(1))
        ln = CLng(ref(2))
        
        outIdx = outIdx + 1
        Dim outKey As String
        outKey = "@OUTPUT_" & Format$(outIdx, "00")
        
        Dim outText As String, outSrcStep As Long, outSrcPrompt As String, outSha As String, outHow As String, outErr As String
        outText = ContextKV_ResolveOutputRef(pipelineName, stepN, promptId, arg, outSrcStep, outSrcPrompt, outHow, outSha, outErr)
        
        If outText = "" And outErr <> "" Then
            ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_MISS", outKey, IIf(strictMode, "ERRO", "ALERTA"), outErr
            Call ContextKV_DictSet(dictInjected, outKey, ContextKV_JsonObj_KVError(outErr))
            
            If strictMode Then
                outErro = "Falha a resolver {@OUTPUT: " & arg & "}: " & outErr
                ContextKV_InjectForStep = False
                Exit Function
            End If
            
            promptText = ContextKV_ReplaceSpan(promptText, idx0, ln, "")
        Else
            Call ContextKV_DictSet(dictInjected, outKey, ContextKV_JsonObj_Source(outSrcStep, outSrcPrompt, outSha, outHow))
            ContextKV_LogEvent pipelineName, stepN, promptId, "OUTPUT_REF_INJECTED", outKey, "INFO", outHow & " (" & CStr(Len(outText)) & " chars)"
            promptText = ContextKV_ReplaceSpan(promptText, idx0, ln, outText)
        End If
    Next i
End If

' 7) Produzir injected_vars JSON (para ser registado na linha do passo actual)
outInjectedVarsJson = ContextKV_JsonObjectFromFragments(dictInjected)

ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_END", "", "INFO", "Injeção concluída. injected_vars chars=" & CStr(Len(outInjectedVarsJson)) & "."
End Function
Public Sub ContextKV_WriteInjectedVars( _
ByVal pipelineName As String, _
ByVal stepN As Long, _
ByVal promptId As String, _
ByVal injectedVarsJson As String, _
ByVal outputFolderBase As String, _
ByVal runToken As String _
)
On Error GoTo EH
If Not ContextKV_GetBoolConfig("CONTEXT_KV_ENABLED", True) Then Exit Sub
If Trim$(injectedVarsJson) = "" Then Exit Sub

Call ContextKV_EnsureSeguimentoColumns

Dim wsS As Worksheet
Set wsS = ContextKV_GetSheet("Seguimento")
If wsS Is Nothing Then Exit Sub

Dim rowS As Long
rowS = ContextKV_FindSeguimentoRow(wsS, pipelineName, stepN, promptId)
If rowS = 0 Then
    ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_ERROR", "", "ERRO", "Não foi possível localizar a linha no Seguimento para escrever injected_vars."
    Exit Sub
End If

Dim colInjected As Long
colInjected = ContextKV_FindColumnByHeader(wsS, "injected_vars")
If colInjected = 0 Then Exit Sub

Dim toWrite As String
toWrite = injectedVarsJson

Dim maxChars As Long
maxChars = ContextKV_GetLongConfig("CONTEXT_KV_MAX_CELL_CHARS", CONTEXTKV_DEFAULT_MAX_CELL_CHARS)

If Len(toWrite) > maxChars Then
    Dim pointerJson As String, savedPath As String
    pointerJson = ContextKV_SaveTextAsFileAndReturnPointer(outputFolderBase, runToken, pipelineName, stepN, promptId, "injected_vars", toWrite, "json", savedPath)
    toWrite = pointerJson
    ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_FILE_FALLBACK", "", "ALERTA", "injected_vars guardado em ficheiro: " & savedPath
End If

wsS.Cells(rowS, colInjected).value = toWrite
Exit Sub
EH:
ContextKV_LogEvent pipelineName, stepN, promptId, "INJECT_ERROR", "", "ERRO", "Erro ao escrever injected_vars: " & Err.Description
End Sub
Public Sub ContextKV_CaptureRow( _
ByVal pipelineName As String, _
ByVal stepN As Long, _
ByVal promptId As String, _
ByVal outputFolderBase As String, _
ByVal runToken As String _
)
On Error GoTo EH
If Not ContextKV_GetBoolConfig("CONTEXT_KV_ENABLED", True) Then Exit Sub

Call ContextKV_EnsureSeguimentoColumns

Dim wsS As Worksheet
Set wsS = ContextKV_GetSheet("Seguimento")
If wsS Is Nothing Then Exit Sub

Dim rowS As Long
rowS = ContextKV_FindSeguimentoRow(wsS, pipelineName, stepN, promptId)
If rowS = 0 Then
    ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_ERROR", "", "ERRO", "Não foi possível localizar a linha no Seguimento para captura."
    Exit Sub
End If

Dim colOut As Long
colOut = ContextKV_FindColumnByHeader(wsS, "Output (texto)")
If colOut = 0 Then
    ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_ERROR", "", "ERRO", "Coluna 'Output (texto)' não encontrada no Seguimento."
    Exit Sub
End If

Dim rawOut As String
rawOut = CStr(wsS.Cells(rowS, colOut).value)

Dim fullOut As String
fullOut = ContextKV_ResolveFullOutput(rawOut)

ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_START", "", "INFO", "Início da captura (ContextKV). Output chars=" & CStr(Len(fullOut)) & "."

Dim reserved As Variant
reserved = ContextKV_ReservedKeys()

Dim dictCap As Object, dictMeta As Object
Set dictCap = CreateObject("Scripting.Dictionary")  ' key -> value
Set dictMeta = CreateObject("Scripting.Dictionary") ' key -> meta fragment json

Dim key As Variant
For Each key In reserved
    Dim v As String, method As String, fmt As String, errMsg As String, labelFound As Boolean
    v = "": method = "": fmt = "": errMsg = "": labelFound = False
    
    If ContextKV_ParseOutputForKey(fullOut, CStr(key), v, method, fmt, errMsg, labelFound) Then
        Call ContextKV_DictSet(dictCap, CStr(key), v)
        
        Dim sha As String
        sha = ContextKV_TrySHA256(v)
        
        Call ContextKV_DictSet(dictMeta, CStr(key), ContextKV_JsonObj_MetaOK(fmt, Len(v), sha, method))
    ElseIf labelFound Then
        ' O rótulo existe mas não foi possível extrair
        Call ContextKV_DictSet(dictMeta, CStr(key), ContextKV_JsonObj_MetaError("parse_failed", method, errMsg))
    End If
Next key

If dictCap.Count = 0 And dictMeta.Count = 0 Then
    ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_MISS", "", "INFO", "Sem variáveis capturadas (nenhum rótulo encontrado)."
    Exit Sub
End If

Dim capturedJson As String
capturedJson = ContextKV_BuildCapturedVarsJson(dictCap, outputFolderBase, runToken, pipelineName, stepN, promptId)

Dim metaJson As String
metaJson = ContextKV_JsonObjectFromFragments(dictMeta)

Dim colCap As Long, colMeta As Long
colCap = ContextKV_FindColumnByHeader(wsS, "captured_vars")
colMeta = ContextKV_FindColumnByHeader(wsS, "captured_vars_meta")

If colCap = 0 Or colMeta = 0 Then Exit Sub

Dim maxChars As Long
maxChars = ContextKV_GetLongConfig("CONTEXT_KV_MAX_CELL_CHARS", CONTEXTKV_DEFAULT_MAX_CELL_CHARS)

If Len(capturedJson) > maxChars Then
    Dim pointerJson As String, savedPath As String
    pointerJson = ContextKV_SaveTextAsFileAndReturnPointer(outputFolderBase, runToken, pipelineName, stepN, promptId, "captured_vars", capturedJson, "json", savedPath)
    capturedJson = pointerJson
    ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_FILE_FALLBACK", "captured_vars", "ALERTA", "captured_vars guardado em ficheiro: " & savedPath
End If

If Len(metaJson) > maxChars Then
    Dim pointerMeta As String, savedPath2 As String
    pointerMeta = ContextKV_SaveTextAsFileAndReturnPointer(outputFolderBase, runToken, pipelineName, stepN, promptId, "captured_vars_meta", metaJson, "json", savedPath2)
    metaJson = pointerMeta
    ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_FILE_FALLBACK", "captured_vars_meta", "ALERTA", "captured_vars_meta guardado em ficheiro: " & savedPath2
End If

wsS.Cells(rowS, colCap).value = capturedJson
wsS.Cells(rowS, colMeta).value = metaJson

ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_OK", "", "INFO", "Captura concluída. Keys=" & CStr(dictCap.Count) & " | captured_vars chars=" & CStr(Len(capturedJson)) & "."
Exit Sub
EH:
ContextKV_LogEvent pipelineName, stepN, promptId, "CAPTURE_ERROR", "", "ERRO", "Erro na captura: " & Err.Description
End Sub
' -----------------------------
' SelfTests (idempotentes)
' -----------------------------
Public Sub SelfTest_ContextKV_Parse_RESULTS_JSON()
On Error GoTo EH
Dim sample As String
sample = "A) TABELA" & vbCrLf & _
         "..." & vbCrLf & _
         "B) RESULTS_JSON" & vbCrLf & _
         ContextKV_Fence() & "json" & vbCrLf & _
         "[{""a"":1},{""b"":2}]" & vbCrLf & _
         ContextKV_Fence() & "" & vbCrLf & _
         "C) REGISTO_PESQUISAS" & vbCrLf & _
         "1. query" & vbCrLf

Dim v As String, method As String, fmt As String, errMsg As String, labelFound As Boolean
v = "": method = "": fmt = "": errMsg = "": labelFound = False

Dim ok As Boolean
ok = ContextKV_ParseOutputForKey(sample, "RESULTS_JSON", v, method, fmt, errMsg, labelFound)

If ok And InStr(1, v, """a"":1", vbTextCompare) > 0 Then
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_PASS", "Parse_RESULTS_JSON", "INFO", "PASS | method=" & method & " | fmt=" & fmt
Else
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "Parse_RESULTS_JSON", "ERRO", "FAIL | ok=" & CStr(ok) & " | err=" & errMsg
End If
Exit Sub
EH:
ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "Parse_RESULTS_JSON", "ERRO", "Erro: " & Err.Description
End Sub
Public Sub SelfTest_ContextKV_Placeholder_Replace()
On Error GoTo EH
Dim prompt As String
prompt = "Resumo:" & vbCrLf & "{{VAR:RESULTS_JSON}}"

Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
dict.Add "RESULTS_JSON", "[1,2]"

Dim ph As Collection
Set ph = ContextKV_FindVarPlaceholders(prompt)

Dim i As Long
For i = ph.Count To 1 Step -1
    Dim item As Variant
    item = ph(i)
    Dim key As String: key = CStr(item(0))
    Dim idx0 As Long: idx0 = CLng(item(1))
    Dim ln As Long: ln = CLng(item(2))
    prompt = ContextKV_ReplaceSpan(prompt, idx0, ln, CStr(dict(key)))
Next i

If InStr(1, prompt, "[1,2]", vbTextCompare) > 0 Then
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_PASS", "Placeholder_Replace", "INFO", "PASS"
Else
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "Placeholder_Replace", "ERRO", "FAIL | Resultado=" & prompt
End If
Exit Sub
EH:
ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "Placeholder_Replace", "ERRO", "Erro: " & Err.Description
End Sub
Public Sub SelfTest_ContextKV_FileFallback()
On Error GoTo EH
Dim big As String
big = String$(CONTEXTKV_DEFAULT_MAX_CELL_CHARS + 500, "A")

Dim baseFolder As String
baseFolder = ThisWorkbook.path

Dim pointer As String, savedPath As String
pointer = ContextKV_SaveTextAsFileAndReturnPointer(baseFolder, "SELFTEST", "SELFTEST", 0, "SELFTEST", "BIGVALUE", big, "txt", savedPath)

If InStr(1, pointer, """@file""", vbTextCompare) > 0 And ContextKV_FileExists(savedPath) Then
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_PASS", "FileFallback", "INFO", "PASS | " & savedPath
Else
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "FileFallback", "ERRO", "FAIL | pointer=" & pointer
End If
Exit Sub
EH:
ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "FileFallback", "ERRO", "Erro: " & Err.Description
End Sub
Public Sub SelfTest_ContextKV_OutputRef()
On Error GoTo EH
Dim blk As String
blk = ContextKV_FormatOutputBlock("X/01", "texto")

If InStr(1, blk, "OUTPUT_PROMPT_ID:", vbTextCompare) > 0 And InStr(1, blk, "OUTPUT_TEXT_BEGIN", vbTextCompare) > 0 Then
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_PASS", "OutputRefFormat", "INFO", "PASS"
Else
    ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "OutputRefFormat", "ERRO", "FAIL"
End If
Exit Sub
EH:
ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_FAIL", "OutputRefFormat", "ERRO", "Erro: " & Err.Description
End Sub
Public Sub SelfTest_RunAll_ContextKV()
ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_START", "", "INFO", "Início SelfTest_RunAll_ContextKV"
SelfTest_ContextKV_Parse_RESULTS_JSON
SelfTest_ContextKV_Placeholder_Replace
SelfTest_ContextKV_FileFallback
SelfTest_ContextKV_OutputRef
ContextKV_LogEvent "SELFTEST", 0, "SELFTEST", "SELFTEST_END", "", "INFO", "Fim SelfTest_RunAll_ContextKV"
End Sub
' =============================================================================
' Implementação interna (helpers)
' =============================================================================
Private Function ContextKV_ReservedKeys() As Variant
ContextKV_ReservedKeys = Array( _
"RESULTS_JSON", _
"REGISTO_PESQUISAS", _
"NEXT_PROMPT_ID", _
"DECISION", _
"MEMORY_SHORT", _
"MEMORY_OVERALL", _
"LESSONS", _
"VARIABLES", _
"CRITERIA", _
"RATIONALE" _
)
End Function
Private Sub ContextKV_EnsureSeguimentoColumns()
Dim ws As Worksheet
Set ws = ContextKV_GetSheet("Seguimento")
If ws Is Nothing Then Exit Sub
ContextKV_EnsureHeader ws, "captured_vars"
ContextKV_EnsureHeader ws, "captured_vars_meta"
ContextKV_EnsureHeader ws, "injected_vars"
End Sub
Private Sub ContextKV_EnsureDebugColumns()
Dim ws As Worksheet
Set ws = ContextKV_GetSheet("DEBUG")
If ws Is Nothing Then Exit Sub
ContextKV_EnsureHeader ws, "pipeline_name"
ContextKV_EnsureHeader ws, "action"
ContextKV_EnsureHeader ws, "var"
ContextKV_EnsureHeader ws, "result"
ContextKV_EnsureHeader ws, "details"
End Sub
Private Sub ContextKV_EnsureHistoricoColumns()
Dim ws As Worksheet
Set ws = ContextKV_GetSheet("HISTÓRICO")
If ws Is Nothing Then Exit Sub
ContextKV_EnsureHeader ws, "captured_vars"
ContextKV_EnsureHeader ws, "captured_vars_meta"
ContextKV_EnsureHeader ws, "injected_vars"
End Sub
Private Sub ContextKV_EnsureHeader(ByVal ws As Worksheet, ByVal headerName As String)
On Error GoTo EH
Dim map As Object
Set map = ContextKV_HeaderMap(ws)

Dim key As String
key = ContextKV_NormalizeHeader(headerName)

If map.exists(key) Then Exit Sub

Dim lastCol As Long
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
If lastCol < 1 Then lastCol = 1

ws.Cells(1, lastCol + 1).value = headerName
Exit Sub
EH:
' silencioso (layout), mas evita crash
End Sub
Private Function ContextKV_GetSheet(ByVal sheetName As String) As Worksheet
On Error Resume Next
Set ContextKV_GetSheet = ThisWorkbook.Worksheets(sheetName)
On Error GoTo 0
End Function
Private Function ContextKV_HeaderMap(ByVal ws As Worksheet) As Object
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")
Dim lastCol As Long
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
If lastCol < 1 Then
    Set ContextKV_HeaderMap = d
    Exit Function
End If

Dim c As Long
For c = 1 To lastCol
    Dim h As String
    h = Trim$(CStr(ws.Cells(1, c).value))
    If h <> "" Then
        Dim k As String
        k = ContextKV_NormalizeHeader(h)
        If Not d.exists(k) Then d.Add k, c
    End If
Next c

Set ContextKV_HeaderMap = d
End Function
Private Function ContextKV_FindColumnByHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long
Dim map As Object
Set map = ContextKV_HeaderMap(ws)
Dim key As String
key = ContextKV_NormalizeHeader(headerName)

If map.exists(key) Then
    ContextKV_FindColumnByHeader = CLng(map(key))
Else
    ContextKV_FindColumnByHeader = 0
End If
End Function


Private Function ContextKV_NormalizeHeader(ByVal s As String) As String
    s = Trim$(s)
    s = LCase$(s)
    s = ContextKV_RemoveDiacritics(s)
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    
    Do While InStr(s, "  ") > 0
       s = Replace(s, vbCrLf, " ")
        s = Replace(s, vbCr, " ")
        s = Replace(s, vbLf, " ")
    Loop
    
    ContextKV_NormalizeHeader = s
End Function



Private Function ContextKV_RemoveDiacritics(ByVal s As String) As String
' Remoção simples (PT): suficiente para headers e comparações
Dim a As Variant, b As Variant, i As Long
a = Array("á", "à", "ã", "â", "ä", "é", "è", "ê", "ë", "í", "ì", "î", "ï", "ó", "ò", "õ", "ô", "ö", "ú", "ù", "û", "ü", "ç", _
"Á", "À", "Ã", "Â", "Ä", "É", "È", "Ê", "Ë", "Í", "Ì", "Î", "Ï", "Ó", "Ò", "Õ", "Ô", "Ö", "Ú", "Ù", "Û", "Ü", "Ç")
b = Array("a", "a", "a", "a", "a", "e", "e", "e", "e", "i", "i", "i", "i", "o", "o", "o", "o", "o", "u", "u", "u", "u", "c", _
"a", "a", "a", "a", "a", "e", "e", "e", "e", "i", "i", "i", "i", "o", "o", "o", "o", "o", "u", "u", "u", "u", "c")
For i = LBound(a) To UBound(a)
s = Replace(s, CStr(a(i)), CStr(b(i)))
Next i
ContextKV_RemoveDiacritics = s
End Function
' -----------------------------
' Parsing placeholders e directivas
' -----------------------------
Private Function ContextKV_FindVarPlaceholders(ByVal text As String) As Collection
Dim col As New Collection
On Error GoTo EH
Dim re As Object, matches As Object, m As Object
Set re = CreateObject("VBScript.RegExp")
re.Global = True
re.IgnoreCase = True
re.pattern = "\{\{\s*VAR\s*:\s*([A-Za-z0-9_]+)\s*\}\}"

Set matches = re.Execute(text)

Dim i As Long
For i = 0 To matches.Count - 1
    Set m = matches(i)
    Dim key As String
    key = UCase$(CStr(m.SubMatches(0)))
    Dim item As Variant
    item = Array(key, CLng(m.FirstIndex), CLng(m.Length))
    col.Add item
Next i

Set ContextKV_FindVarPlaceholders = col
Exit Function
EH:
Set ContextKV_FindVarPlaceholders = col
End Function
Private Function ContextKV_FindOutputRefs(ByVal text As String) As Collection
Dim col As New Collection
On Error GoTo EH
Dim re As Object, matches As Object, m As Object
Set re = CreateObject("VBScript.RegExp")
re.Global = True
re.IgnoreCase = True
re.pattern = "\{\@OUTPUT\s*:\s*([^\}]*)\}"

Set matches = re.Execute(text)

Dim i As Long
For i = 0 To matches.Count - 1
    Set m = matches(i)
    Dim arg As String
    arg = CStr(m.SubMatches(0))
    col.Add Array(arg, CLng(m.FirstIndex), CLng(m.Length))
Next i

Set ContextKV_FindOutputRefs = col
Exit Function
EH:
Set ContextKV_FindOutputRefs = col
End Function
Private Sub ContextKV_ExtractVarsDirectiveKeys(ByRef text As String, ByVal outDict As Object, ByVal removeFromText As Boolean)
On Error GoTo EH
Dim re As Object, matches As Object, m As Object
Set re = CreateObject("VBScript.RegExp")
re.Global = True
re.IgnoreCase = True
re.MultiLine = True
re.pattern = "^\s*VARS\s*:\s*([A-Za-z0-9_,\s]+)\s*$"

Set matches = re.Execute(text)

Dim i As Long
For i = 0 To matches.Count - 1
    Set m = matches(i)
    Dim listText As String
    listText = CStr(m.SubMatches(0))
    
    Dim parts() As String
    parts = Split(listText, ",")
    
    Dim j As Long
    For j = LBound(parts) To UBound(parts)
        Dim key As String
        key = UCase$(Trim$(parts(j)))
        If key <> "" Then
            If Not outDict.exists(key) Then outDict.Add key, True
        End If
    Next j
Next i

If removeFromText And matches.Count > 0 Then
    ' Remove as linhas VARS: do texto (para não ir para o modelo)
    text = re.Replace(text, "")
    text = Trim$(text)
End If

Exit Sub
EH:
' silencioso
End Sub
Private Function ContextKV_PromptHasPlaceholderForKey(ByVal placeholders As Collection, ByVal keyName As String) As Boolean
Dim i As Long
For i = 1 To placeholders.Count
Dim ph As Variant
ph = placeholders(i)
If UCase$(CStr(ph(0))) = UCase$(keyName) Then
ContextKV_PromptHasPlaceholderForKey = True
Exit Function
End If
Next i
ContextKV_PromptHasPlaceholderForKey = False
End Function
Private Function ContextKV_ReplaceSpan(ByVal s As String, ByVal startIndex0 As Long, ByVal lengthToReplace As Long, ByVal replacement As String) As String
' startIndex0 é 0-based (VBScript.RegExp), VBA strings são 1-based
Dim leftPart As String, rightPart As String
If startIndex0 < 0 Then startIndex0 = 0
If lengthToReplace < 0 Then lengthToReplace = 0

leftPart = Left$(s, startIndex0)
rightPart = Mid$(s, startIndex0 + lengthToReplace + 1)

ContextKV_ReplaceSpan = leftPart & replacement & rightPart
End Function
' -----------------------------
' Resolver INPUTS: da prompt (catálogo)
' -----------------------------
Private Function ContextKV_TryReadInputsTextByPromptId(ByVal promptId As String) As String
On Error GoTo EH
Dim wsName As String
wsName = ContextKV_ParseSheetNameFromPromptId(promptId)
If wsName = "" Then Exit Function

Dim ws As Worksheet
Set ws = ContextKV_GetSheet(wsName)
If ws Is Nothing Then Exit Function

Dim cel As Range
Set cel = ws.Columns(1).Find(What:=promptId, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
If cel Is Nothing Then Exit Function

' Layout típico do catálogo:
' Linha da prompt: Col A=ID
' Linha +2: Col C="INPUTS:" ; Col D="<lista>"
Dim lbl As String, val As String
lbl = CStr(ws.Cells(cel.Row + 2, 3).value)
val = CStr(ws.Cells(cel.Row + 2, 4).value)

If InStr(1, UCase$(ContextKV_RemoveDiacritics(lbl)), "INPUTS", vbTextCompare) > 0 Then
    ContextKV_TryReadInputsTextByPromptId = val
Else
    ' fallback: devolve D na mesma (é melhor do que falhar silenciosamente)
    ContextKV_TryReadInputsTextByPromptId = val
End If

Exit Function
EH:
ContextKV_TryReadInputsTextByPromptId = ""
End Function
Private Function ContextKV_ParseSheetNameFromPromptId(ByVal promptId As String) As String
Dim p As Long
p = InStr(1, promptId, "/", vbTextCompare)
If p <= 1 Then
ContextKV_ParseSheetNameFromPromptId = ""
Else
ContextKV_ParseSheetNameFromPromptId = Left$(promptId, p - 1)
End If
End Function
' -----------------------------
' Resolver variáveis a partir do Seguimento
' -----------------------------
Private Function ContextKV_ResolveVarValue( _
ByVal pipelineName As String, _
ByVal stepN As Long, _
ByVal keyName As String, _
ByVal outputFolderBase As String, _
ByRef outSourceStep As Long, _
ByRef outSourcePrompt As String, _
ByRef outHow As String, _
ByRef outSha As String, _
ByRef outErr As String _
) As String
outSourceStep = 0
outSourcePrompt = ""
outHow = ""
outSha = ""
outErr = ""

Dim wsS As Worksheet
Set wsS = ContextKV_GetSheet("Seguimento")
If wsS Is Nothing Then
    outErr = "Folha Seguimento não encontrada."
    Exit Function
End If

Dim srcRow As Long
srcRow = ContextKV_FindBestSourceRow(wsS, pipelineName, stepN, keyName)
If srcRow = 0 Then
    outErr = "Não foi encontrada linha fonte (passo anterior / histórico)."
    Exit Function
End If

Dim colPasso As Long, colPrompt As Long
colPasso = ContextKV_FindColumnByHeader(wsS, "Passo")
colPrompt = ContextKV_FindColumnByHeader(wsS, "Prompt ID")

If colPasso > 0 Then outSourceStep = CLng(val(CStr(wsS.Cells(srcRow, colPasso).value)))
If colPrompt > 0 Then outSourcePrompt = CStr(wsS.Cells(srcRow, colPrompt).value)

' 1) tentar via captured_vars
Dim colCap As Long
colCap = ContextKV_FindColumnByHeader(wsS, "captured_vars")

If colCap > 0 Then
    Dim capJson As String
    capJson = CStr(wsS.Cells(srcRow, colCap).value)
    
    If Trim$(capJson) <> "" Then
        capJson = ContextKV_ResolveJsonPointerIfNeeded(capJson, outputFolderBase)
        
        Dim found As Boolean, value As String, fromFile As Boolean
        found = ContextKV_JsonGetCapturedValue(capJson, keyName, outputFolderBase, value, fromFile)
        
        If found Then
            outHow = "captured_vars" & IIf(fromFile, "_file", "")
            outSha = ContextKV_TrySHA256(value)
            ContextKV_ResolveVarValue = value
            Exit Function
        End If
    End If
End If

' 2) lazy parse do Output (texto) do srcRow e, se encontrar, persistir em captured_vars
Dim colOut As Long
colOut = ContextKV_FindColumnByHeader(wsS, "Output (texto)")
If colOut = 0 Then
    outErr = "Coluna Output (texto) não encontrada."
    Exit Function
End If

Dim rawOut As String, fullOut As String
rawOut = CStr(wsS.Cells(srcRow, colOut).value)
fullOut = ContextKV_ResolveFullOutput(rawOut)

Dim v As String, method As String, fmt As String, errMsg As String, labelFound As Boolean
v = "": method = "": fmt = "": errMsg = "": labelFound = False

If ContextKV_ParseOutputForKey(fullOut, keyName, v, method, fmt, errMsg, labelFound) Then
    outHow = "output_parse_lazy:" & method
    outSha = ContextKV_TrySHA256(v)
    ContextKV_ResolveVarValue = v
    
    ' Persistir de forma lazy no captured_vars do srcRow (para futuras injecções)
    On Error Resume Next
    ContextKV_UpsertCapturedVar wsS, srcRow, keyName, v, fmt, method, outputFolderBase, "LAZY"
    On Error GoTo 0
    
    Exit Function
End If

outErr = "Variável '" & keyName & "' não encontrada no passo fonte (nem em captured_vars nem no output)."
End Function
Private Function ContextKV_FindBestSourceRow(ByVal wsS As Worksheet, ByVal pipelineName As String, ByVal stepN As Long, ByVal keyName As String) As Long
' Preferência:
' 1) passo anterior (Passo = stepN-1) no mesmo pipeline
' 2) último passo anterior com captured_vars contendo a key
' 3) último passo anterior do pipeline (fallback)
On Error GoTo EH
Dim colPipe As Long, colPasso As Long, colCap As Long
colPipe = ContextKV_FindColumnByHeader(wsS, "pipeline_name")
colPasso = ContextKV_FindColumnByHeader(wsS, "Passo")
colCap = ContextKV_FindColumnByHeader(wsS, "captured_vars")

If colPipe = 0 Or colPasso = 0 Then Exit Function

Dim lastRow As Long
lastRow = wsS.Cells(wsS.rowS.Count, colPasso).End(xlUp).Row
If lastRow < 2 Then Exit Function

Dim r As Long
Dim bestFallback As Long: bestFallback = 0

' 1) passo anterior directo
For r = lastRow To 2 Step -1
    If Trim$(CStr(wsS.Cells(r, colPipe).value)) = pipelineName Then
        Dim sVal As String
        sVal = Trim$(CStr(wsS.Cells(r, colPasso).value))
        If sVal <> "" And IsNumeric(sVal) Then
            If CLng(sVal) = stepN - 1 Then
                ContextKV_FindBestSourceRow = r
                Exit Function
            End If
        End If
    End If
Next r

' Se não há keyName (ex.: {@OUTPUT:"Prompt Anterior"}), devolve apenas o último passo anterior
If Trim$(keyName) = "" Then
    For r = lastRow To 2 Step -1
        If Trim$(CStr(wsS.Cells(r, colPipe).value)) = pipelineName Then
            Dim sVal0 As String
            sVal0 = Trim$(CStr(wsS.Cells(r, colPasso).value))
            If sVal0 <> "" And IsNumeric(sVal0) Then
                If CLng(sVal0) < stepN Then
                    ContextKV_FindBestSourceRow = r
                    Exit Function
                End If
            End If
        End If
    Next r
    Exit Function
End If

' 2) procurar último passo anterior com captured_vars contendo a key
If colCap > 0 Then
    For r = lastRow To 2 Step -1
        If Trim$(CStr(wsS.Cells(r, colPipe).value)) = pipelineName Then
            sVal = Trim$(CStr(wsS.Cells(r, colPasso).value))
            If sVal <> "" And IsNumeric(sVal) Then
                If CLng(sVal) < stepN Then
                    Dim cap As String
                    cap = CStr(wsS.Cells(r, colCap).value)
                    If InStr(1, cap, """" & keyName & """", vbTextCompare) > 0 Then
                        ContextKV_FindBestSourceRow = r
                        Exit Function
                    End If
                    If bestFallback = 0 Then bestFallback = r
                End If
            End If
        End If
    Next r
End If

' 3) fallback: último passo anterior do pipeline
If bestFallback <> 0 Then
    ContextKV_FindBestSourceRow = bestFallback
End If

Exit Function
EH:
ContextKV_FindBestSourceRow = 0
End Function
Private Function ContextKV_FindSeguimentoRow(ByVal wsS As Worksheet, ByVal pipelineName As String, ByVal stepN As Long, ByVal promptId As String) As Long
' Procura do fim para o início (mais robusto quando há múltiplas execuções)
On Error GoTo EH
Dim colPipe As Long, colPasso As Long, colPrompt As Long
colPipe = ContextKV_FindColumnByHeader(wsS, "pipeline_name")
colPasso = ContextKV_FindColumnByHeader(wsS, "Passo")
colPrompt = ContextKV_FindColumnByHeader(wsS, "Prompt ID")

If colPipe = 0 Or colPasso = 0 Or colPrompt = 0 Then Exit Function

Dim lastRow As Long
lastRow = wsS.Cells(wsS.rowS.Count, colPasso).End(xlUp).Row
If lastRow < 2 Then Exit Function

Dim r As Long
For r = lastRow To 2 Step -1
    If Trim$(CStr(wsS.Cells(r, colPipe).value)) = pipelineName Then
        Dim sStep As String
        sStep = Trim$(CStr(wsS.Cells(r, colPasso).value))
        If sStep <> "" And IsNumeric(sStep) Then
            If CLng(sStep) = stepN Then
                If Trim$(CStr(wsS.Cells(r, colPrompt).value)) = promptId Then
                    ContextKV_FindSeguimentoRow = r
                    Exit Function
                End If
            End If
        End If
    End If
Next r

Exit Function
EH:
ContextKV_FindSeguimentoRow = 0
End Function
' -----------------------------
' Output refs {@OUTPUT: ...}
' -----------------------------
Private Function ContextKV_ResolveOutputRef( _
ByVal pipelineName As String, _
ByVal stepN As Long, _
ByVal promptIdCurrent As String, _
ByVal argRaw As String, _
ByRef outSourceStep As Long, _
ByRef outSourcePrompt As String, _
ByRef outHow As String, _
ByRef outSha As String, _
ByRef outErr As String _
) As String
outSourceStep = 0
outSourcePrompt = ""
outHow = ""
outSha = ""
outErr = ""

Dim arg As String
arg = ContextKV_Unquote(Trim$(argRaw))
arg = ContextKV_RemoveDiacritics(arg)

Dim wsS As Worksheet
Set wsS = ContextKV_GetSheet("Seguimento")
If wsS Is Nothing Then
    outErr = "Folha Seguimento não encontrada."
    Exit Function
End If

Dim colPipe As Long, colPasso As Long, colPrompt As Long, colOut As Long
colPipe = ContextKV_FindColumnByHeader(wsS, "pipeline_name")
colPasso = ContextKV_FindColumnByHeader(wsS, "Passo")
colPrompt = ContextKV_FindColumnByHeader(wsS, "Prompt ID")
colOut = ContextKV_FindColumnByHeader(wsS, "Output (texto)")

If colPipe = 0 Or colPasso = 0 Or colPrompt = 0 Or colOut = 0 Then
    outErr = "Cabeçalhos necessários em Seguimento não encontrados."
    Exit Function
End If

Dim lastRow As Long
lastRow = wsS.Cells(wsS.rowS.Count, colPasso).End(xlUp).Row

If arg = "" Or UCase$(arg) = "PROMPT ANTERIOR" Then
    Dim rPrev As Long
    rPrev = ContextKV_FindBestSourceRow(wsS, pipelineName, stepN, "") ' "" => passo anterior
    If rPrev = 0 Then
        outErr = "Não foi possível localizar o passo anterior."
        Exit Function
    End If
    
    outSourceStep = CLng(val(CStr(wsS.Cells(rPrev, colPasso).value)))
    outSourcePrompt = CStr(wsS.Cells(rPrev, colPrompt).value)
    Dim t As String
    t = ContextKV_ResolveFullOutput(CStr(wsS.Cells(rPrev, colOut).value))
    
    Dim block As String
    block = ContextKV_FormatOutputBlock(outSourcePrompt, t)
    
    outHow = "previous_step"
    outSha = ContextKV_TrySHA256(block)
    ContextKV_ResolveOutputRef = block
    Exit Function
End If

If UCase$(arg) = "TODAS AS PROMPTS" Then
    Dim blocks As String
    blocks = ""
    
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(wsS.Cells(r, colPipe).value)) = pipelineName Then
            Dim sStep As String
            sStep = Trim$(CStr(wsS.Cells(r, colPasso).value))
            If sStep <> "" And IsNumeric(sStep) Then
                If CLng(sStep) < stepN Then
                    Dim pid As String
                    pid = CStr(wsS.Cells(r, colPrompt).value)
                    Dim outTxt As String
                    outTxt = ContextKV_ResolveFullOutput(CStr(wsS.Cells(r, colOut).value))
                    blocks = blocks & ContextKV_FormatOutputBlock(pid, outTxt) & vbCrLf & vbCrLf
                End If
            End If
        End If
    Next r
    
    outHow = "all_previous_prompts"
    outSha = ContextKV_TrySHA256(blocks)
    ContextKV_ResolveOutputRef = blocks
    Exit Function
End If

' Caso contrário: assume Prompt ID específico
Dim targetPromptId As String
targetPromptId = Trim$(argRaw)
targetPromptId = ContextKV_Unquote(targetPromptId)

Dim foundRow As Long
foundRow = 0
Dim r2 As Long
For r2 = lastRow To 2 Step -1
    If Trim$(CStr(wsS.Cells(r2, colPipe).value)) = pipelineName Then
        If Trim$(CStr(wsS.Cells(r2, colPrompt).value)) = targetPromptId Then
            Dim sStep2 As String
            sStep2 = Trim$(CStr(wsS.Cells(r2, colPasso).value))
            If sStep2 <> "" And IsNumeric(sStep2) Then
                If CLng(sStep2) < stepN Then
                    foundRow = r2
                    Exit For
                End If
            End If
        End If
    End If
Next r2

If foundRow = 0 Then
    outErr = "Não foi encontrada referência {@OUTPUT} para Prompt ID: " & targetPromptId
    Exit Function
End If

outSourceStep = CLng(val(CStr(wsS.Cells(foundRow, colPasso).value)))
outSourcePrompt = CStr(wsS.Cells(foundRow, colPrompt).value)
t = ContextKV_ResolveFullOutput(CStr(wsS.Cells(foundRow, colOut).value))
block = ContextKV_FormatOutputBlock(outSourcePrompt, t)

outHow = "last_by_prompt_id"
outSha = ContextKV_TrySHA256(block)
ContextKV_ResolveOutputRef = block
End Function
Private Function ContextKV_FormatOutputBlock(ByVal sourcePromptId As String, ByVal outText As String) As String
Dim s As String
s = "{OUTPUT_PROMPT_ID:" & sourcePromptId & "}" & vbCrLf & _
"{OUTPUT_TEXT_BEGIN}" & vbCrLf & _
outText & vbCrLf & _
"{OUTPUT_TEXT_END}"
ContextKV_FormatOutputBlock = s
End Function
Private Function ContextKV_FormatVarBlock(ByVal keyName As String, ByVal value As String) As String
Dim s As String
s = "{VAR:" & keyName & "}" & vbCrLf & _
"{VAR_VALUE_BEGIN}" & vbCrLf & _
value & vbCrLf & _
"{VAR_VALUE_END}"
ContextKV_FormatVarBlock = s
End Function
Private Function ContextKV_Unquote(ByVal s As String) As String
s = Trim$(s)
If Len(s) >= 2 Then
If Left$(s, 1) = """" And Right$(s, 1) = """" Then
ContextKV_Unquote = Mid$(s, 2, Len(s) - 2)
Exit Function
End If
End If
ContextKV_Unquote = s
End Function
' -----------------------------
' Parse de output (captura)
' -----------------------------
Private Function ContextKV_ParseOutputForKey( _
ByVal outputText As String, _
ByVal keyName As String, _
ByRef outValue As String, _
ByRef outMethod As String, _
ByRef outFmt As String, _
ByRef outErr As String, _
ByRef outLabelFound As Boolean _
) As Boolean
ContextKV_ParseOutputForKey = False
outValue = ""
outMethod = ""
outFmt = ""
outErr = ""
outLabelFound = False

Dim t As String
t = outputText
t = Replace(t, vbCrLf, vbLf)
t = Replace(t, vbCr, vbLf)

Dim lines() As String
lines = Split(t, vbLf)

Dim i As Long
Dim lblIndex As Long: lblIndex = -1

Dim keyU As String
keyU = UCase$(keyName)

For i = LBound(lines) To UBound(lines)
    Dim ln As String
    ln = Trim$(lines(i))
    If ln <> "" Then
        Dim lnClean As String
        lnClean = ContextKV_CleanLineForLabel(ln)
        If ContextKV_LineIsLabelForKey(lnClean, keyU) Then
            lblIndex = i
            outLabelFound = True
            Exit For
        End If
    End If
Next i

If lblIndex = -1 Then Exit Function

' P3) Linha "KEY: valor"
Dim lblLine As String
lblLine = Trim$(lines(lblIndex))
If InStr(1, UCase$(lblLine), keyU & ":", vbTextCompare) > 0 Then
    Dim p As Long
    p = InStr(1, lblLine, ":", vbTextCompare)
    If p > 0 Then
        Dim vLine As String
        vLine = Trim$(Mid$(lblLine, p + 1))
        If vLine <> "" Then
            outValue = vLine
            outMethod = "key_colon_line"
            outFmt = "text"
            ContextKV_ParseOutputForKey = True
            Exit Function
        End If
    End If
End If

' P1) Bloco fenced (fence de 3 backticks)
Dim j As Long
For j = lblIndex + 1 To UBound(lines)
    ln = Trim$(lines(j))
    If ln = "" Then
        ' skip
    ElseIf Left$(ln, 3) = ContextKV_Fence() Then
        Dim k As Long
        Dim startContent As Long
        startContent = j + 1
        For k = startContent To UBound(lines)
            If Left$(Trim$(lines(k)), 3) = ContextKV_Fence() Then
                ' capturar startContent..k-1
                outValue = ContextKV_JoinLines(lines, startContent, k - 1)
                outValue = Trim$(outValue)
                outMethod = "fenced_block_after_label"
                outFmt = ContextKV_InferFmt(outValue)
                ContextKV_ParseOutputForKey = (outValue <> "")
                Exit Function
            End If
        Next k
        
        outErr = "Fence iniciado mas não foi encontrado fecho (fence de 3 backticks)."
        outMethod = "fenced_block_after_label"
        Exit Function
    Else
        Exit For ' primeiro conteúdo não-fence
    End If
Next j

' P2) Secção rotulada até ao próximo cabeçalho (A) B) C) ... ou fim)
Dim startIdx As Long
startIdx = lblIndex + 1

' avançar linhas vazias
Do While startIdx <= UBound(lines)
    If Trim$(lines(startIdx)) <> "" Then Exit Do
    startIdx = startIdx + 1
Loop

If startIdx > UBound(lines) Then
    outErr = "Rótulo encontrado mas sem conteúdo subsequente."
    outMethod = "section_until_next_header"
    Exit Function
End If

Dim endIdx As Long
endIdx = UBound(lines)
For j = startIdx To UBound(lines)
    ln = Trim$(lines(j))
    If ln <> "" Then
        Dim lnClean2 As String
        lnClean2 = ContextKV_CleanLineForLabel(ln)
        If ContextKV_LineLooksLikeSectionHeader(lnClean2) Then
            endIdx = j - 1
            Exit For
        End If
    End If
Next j

outValue = ContextKV_JoinLines(lines, startIdx, endIdx)
outValue = Trim$(outValue)
outMethod = "section_until_next_header"
outFmt = ContextKV_InferFmt(outValue)

If outValue = "" Then
    outErr = "Conteúdo vazio após rótulo."
    ContextKV_ParseOutputForKey = False
Else
    ContextKV_ParseOutputForKey = True
End If
End Function
Private Function ContextKV_CleanLineForLabel(ByVal s As String) As String
' Remove adornos típicos (markdown) do início/fim para facilitar detecção de rótulos
Dim t As String
t = Trim$(s)
' remover **...**
Do While Left$(t, 2) = "**"
    t = Mid$(t, 3)
    t = Trim$(t)
Loop
Do While Right$(t, 2) = "**"
    t = Left$(t, Len(t) - 2)
    t = Trim$(t)
Loop

' remover bullets
If Left$(t, 1) = "-" Or Left$(t, 1) = "o" Then
    t = Trim$(Mid$(t, 2))
End If

ContextKV_CleanLineForLabel = t
End Function
Private Function ContextKV_LineIsLabelForKey(ByVal lineClean As String, ByVal keyU As String) As Boolean
Dim u As String
u = UCase$(Trim$(lineClean))
If u = keyU Then
    ContextKV_LineIsLabelForKey = True
    Exit Function
End If

If Left$(u, Len(keyU) + 1) = keyU & ":" Then
    ContextKV_LineIsLabelForKey = True
    Exit Function
End If

' Formatos: "B) KEY" ou "B) KEY:"
If Len(u) >= 3 Then
    If Mid$(u, 2, 1) = ")" Then
        Dim rest As String
        rest = Trim$(Mid$(u, 3))
        If rest = keyU Or Left$(rest, Len(keyU) + 1) = keyU & ":" Then
            ContextKV_LineIsLabelForKey = True
            Exit Function
        End If
    End If
End If

' fallback: linha contém a key como palavra inteira
If ContextKV_ContainsWholeWord(u, keyU) Then
    ContextKV_LineIsLabelForKey = True
    Exit Function
End If

ContextKV_LineIsLabelForKey = False
End Function
Private Function ContextKV_LineLooksLikeSectionHeader(ByVal lineClean As String) As Boolean
Dim t As String
t = Trim$(lineClean)
If Len(t) < 2 Then Exit Function
Dim c1 As String, c2 As String
c1 = Left$(t, 1)
c2 = Mid$(t, 2, 1)

If c2 = ")" Then
    If c1 Like "[A-Z]" Then
        ContextKV_LineLooksLikeSectionHeader = True
        Exit Function
    End If
End If

ContextKV_LineLooksLikeSectionHeader = False
End Function
Private Function ContextKV_ContainsWholeWord(ByVal textU As String, ByVal wordU As String) As Boolean
Dim p As Long
p = InStr(1, textU, wordU, vbTextCompare)
If p = 0 Then Exit Function
Dim beforeC As String, afterC As String
beforeC = ""
afterC = ""

If p > 1 Then beforeC = Mid$(textU, p - 1, 1)
If p + Len(wordU) <= Len(textU) Then afterC = Mid$(textU, p + Len(wordU), 1)

If beforeC <> "" Then
    If beforeC Like "[A-Z0-9_]" Then Exit Function
End If
If afterC <> "" Then
    If afterC Like "[A-Z0-9_]" Then Exit Function
End If

ContextKV_ContainsWholeWord = True
End Function
Private Function ContextKV_JoinLines(ByRef lines() As String, ByVal iStart As Long, ByVal iEnd As Long) As String
If iStart > iEnd Then
ContextKV_JoinLines = ""
Exit Function
End If
Dim i As Long, s As String
s = ""
For i = iStart To iEnd
    If s = "" Then
        s = CStr(lines(i))
    Else
        s = s & vbCrLf & CStr(lines(i))
    End If
Next i
ContextKV_JoinLines = s
End Function
Private Function ContextKV_InferFmt(ByVal s As String) As String
Dim t As String
t = Trim$(s)
If t = "" Then
ContextKV_InferFmt = "text"
ElseIf (Left$(t, 1) = "{" And Right$(t, 1) = "}") Or (Left$(t, 1) = "[" And Right$(t, 1) = "]") Then
ContextKV_InferFmt = "json_like"
Else
ContextKV_InferFmt = "text"
End If
End Function
' -----------------------------
' Serialização JSON mínima (objecto)
' -----------------------------
Private Function ContextKV_JsonObjectFromFragments(ByVal dictFragments As Object) As String
Dim k As Variant
Dim s As String
s = "{"
Dim first As Boolean
first = True
For Each k In dictFragments.keys
    If Not first Then s = s & ","
    s = s & """" & ContextKV_JsonEscape(CStr(k)) & """:" & CStr(dictFragments(k))
    first = False
Next k

s = s & "}"
ContextKV_JsonObjectFromFragments = s
End Function
Private Function ContextKV_BuildCapturedVarsJson(ByVal dictCap As Object, ByVal outputFolderBase As String, ByVal runToken As String, ByVal pipelineName As String, ByVal stepN As Long, ByVal promptId As String) As String
Dim maxChars As Long
maxChars = ContextKV_GetLongConfig("CONTEXT_KV_MAX_CELL_CHARS", CONTEXTKV_DEFAULT_MAX_CELL_CHARS)
Dim fragments As Object
Set fragments = CreateObject("Scripting.Dictionary") ' key -> jsonValue

Dim k As Variant
For Each k In dictCap.keys
    Dim v As String
    v = CStr(dictCap(k))
    
    Dim fmt As String
    fmt = ContextKV_InferFmt(v)
    
    Dim jsonValue As String
    jsonValue = """" & ContextKV_JsonEscape(v) & """"
    
    If Len(jsonValue) > maxChars Then
        Dim savedPath As String
        jsonValue = ContextKV_SaveTextAsFileAndReturnPointer(outputFolderBase, runToken, pipelineName, stepN, promptId, CStr(k), v, IIf(fmt = "json_like", "json", "txt"), savedPath)
    End If
    
    Call ContextKV_DictSet(fragments, CStr(k), jsonValue)
Next k

ContextKV_BuildCapturedVarsJson = ContextKV_JsonObjectFromFragments(fragments)
End Function
Private Function ContextKV_JsonObj_Source(ByVal fromStep As Long, ByVal fromPrompt As String, ByVal sha As String, ByVal sourceHow As String) As String
Dim s As String
s = "{"
s = s & """from_step"":" & CStr(fromStep) & ","
s = s & """from_prompt"":""" & ContextKV_JsonEscape(fromPrompt) & ""","
s = s & """sha256"":""" & ContextKV_JsonEscape(sha) & ""","
s = s & """source"":""" & ContextKV_JsonEscape(sourceHow) & """"
s = s & "}"
ContextKV_JsonObj_Source = s
End Function
Private Function ContextKV_JsonObj_KVError(ByVal errMsg As String) As String
ContextKV_JsonObj_KVError = "{""error"":""" & ContextKV_JsonEscape(errMsg) & """}"
End Function
Private Function ContextKV_JsonObj_MetaOK(ByVal fmt As String, ByVal chars As Long, ByVal sha As String, ByVal method As String) As String
Dim s As String
s = "{"
s = s & """fmt"":""" & ContextKV_JsonEscape(fmt) & ""","
s = s & """chars"":" & CStr(chars) & ","
s = s & """sha256"":""" & ContextKV_JsonEscape(sha) & ""","
s = s & """method"":""" & ContextKV_JsonEscape(method) & """"
s = s & "}"
ContextKV_JsonObj_MetaOK = s
End Function
Private Function ContextKV_JsonObj_MetaError(ByVal errCode As String, ByVal method As String, ByVal errMsg As String) As String
Dim s As String
s = "{"
s = s & """error"":""" & ContextKV_JsonEscape(errCode) & ""","
s = s & """method"":""" & ContextKV_JsonEscape(method) & ""","
s = s & """details"":""" & ContextKV_JsonEscape(errMsg) & """"
s = s & "}"
ContextKV_JsonObj_MetaError = s
End Function
Private Function ContextKV_JsonEscape(ByVal s As String) As String
Dim t As String
t = s
t = Replace(t, "\", "\\")
t = Replace(t, """", "\""")
t = Replace(t, vbCrLf, "\n")
t = Replace(t, vbCr, "\n")
t = Replace(t, vbLf, "\n")
t = Replace(t, vbTab, "\t")
ContextKV_JsonEscape = t
End Function
Private Function ContextKV_JsonUnescape(ByVal s As String) As String
' Unescape mínimo (para \n, \t, ", \)
Dim t As String
t = s
t = Replace(t, "\n", vbCrLf)
t = Replace(t, "\t", vbTab)
t = Replace(t, "\""", """")
t = Replace(t, "\\", "\")
ContextKV_JsonUnescape = t
End Function
' -----------------------------
' Leitura de captured_vars (JSON mínimo)
' -----------------------------
Private Function ContextKV_ResolveJsonPointerIfNeeded(ByVal jsonText As String, ByVal outputFolderBase As String) As String
Dim t As String
t = Trim$(jsonText)
If t = "" Then
ContextKV_ResolveJsonPointerIfNeeded = ""
Exit Function
End If
Dim rePtr As Object
Set rePtr = CreateObject("VBScript.RegExp")
rePtr.Global = False
rePtr.IgnoreCase = True
rePtr.pattern = "^\{\s*""@file""\s*:"
If rePtr.Test(t) Then
    Dim filePath As String
    filePath = ContextKV_JsonGetFilePointer(t)
    If filePath <> "" Then
        Dim fullPath As String
        fullPath = ContextKV_ResolveFilePath(outputFolderBase, filePath)
        If ContextKV_FileExists(fullPath) Then
            ContextKV_ResolveJsonPointerIfNeeded = ContextKV_ReadTextFileUTF8(fullPath)
            Exit Function
        End If
    End If
End If

ContextKV_ResolveJsonPointerIfNeeded = jsonText
End Function
Private Function ContextKV_JsonGetCapturedValue(ByVal capturedJson As String, ByVal keyName As String, ByVal outputFolderBase As String, ByRef outValue As String, ByRef outFromFile As Boolean) As Boolean
ContextKV_JsonGetCapturedValue = False
outValue = ""
outFromFile = False
Dim t As String
t = capturedJson
If Trim$(t) = "" Then Exit Function

Dim needle As String
needle = """" & keyName & """"

Dim p As Long
p = InStr(1, t, needle, vbTextCompare)
If p = 0 Then Exit Function

' encontrar ':' a seguir à key
Dim c As Long
c = InStr(p + Len(needle), t, ":", vbTextCompare)
If c = 0 Then Exit Function

' avançar espaços
Dim i As Long
i = c + 1
Do While i <= Len(t) And (Mid$(t, i, 1) = " " Or Mid$(t, i, 1) = vbTab)
    i = i + 1
Loop
If i > Len(t) Then Exit Function

Dim ch As String
ch = Mid$(t, i, 1)

If ch = """" Then
    ' string JSON
    Dim endPos As Long
    endPos = ContextKV_FindEndJsonString(t, i)
    If endPos = 0 Then Exit Function
    
    Dim raw As String
    raw = Mid$(t, i + 1, endPos - (i + 1))
    outValue = ContextKV_JsonUnescape(raw)
    ContextKV_JsonGetCapturedValue = True
    Exit Function
End If

If ch = "{" Then
    ' objecto - procurar @file
    Dim objEnd As Long
    objEnd = ContextKV_FindEndJsonObject(t, i)
    If objEnd = 0 Then Exit Function
    
    Dim objText As String
    objText = Mid$(t, i, objEnd - i + 1)
    
    Dim filePath As String
    filePath = ContextKV_JsonGetFilePointer(objText)
    If filePath <> "" Then
        Dim fullPath As String
        fullPath = ContextKV_ResolveFilePath(outputFolderBase, filePath)
        If ContextKV_FileExists(fullPath) Then
            outValue = ContextKV_ReadTextFileUTF8(fullPath)
            outFromFile = True
            ContextKV_JsonGetCapturedValue = True
            Exit Function
        End If
    End If
End If
End Function
Private Function ContextKV_FindEndJsonString(ByVal jsonText As String, ByVal startQuotePos As Long) As Long
' startQuotePos aponta para a aspa inicial
Dim i As Long
i = startQuotePos + 1
Do While i <= Len(jsonText)
    Dim ch As String
    ch = Mid$(jsonText, i, 1)
    If ch = "\" Then
        i = i + 2
    ElseIf ch = """" Then
        ContextKV_FindEndJsonString = i
        Exit Function
    Else
        i = i + 1
    End If
Loop

ContextKV_FindEndJsonString = 0
End Function
Private Function ContextKV_FindEndJsonObject(ByVal jsonText As String, ByVal startBracePos As Long) As Long
Dim depth As Long
depth = 0
Dim i As Long
For i = startBracePos To Len(jsonText)
    Dim ch As String
    ch = Mid$(jsonText, i, 1)
    
    If ch = """" Then
        ' saltar string
        Dim endS As Long
        endS = ContextKV_FindEndJsonString(jsonText, i)
        If endS = 0 Then Exit Function
        i = endS
    ElseIf ch = "{" Then
        depth = depth + 1
    ElseIf ch = "}" Then
        depth = depth - 1
        If depth = 0 Then
            ContextKV_FindEndJsonObject = i
            Exit Function
        End If
    End If
Next i

ContextKV_FindEndJsonObject = 0
End Function
Private Function ContextKV_JsonGetFilePointer(ByVal jsonObjText As String) As String
' Procura {"@file":"..."} e devolve o path (sem unescape complexo)
Dim t As String
t = jsonObjText
Dim p As Long
p = InStr(1, t, """@file""", vbTextCompare)
If p = 0 Then Exit Function

Dim c As Long
c = InStr(p, t, ":", vbTextCompare)
If c = 0 Then Exit Function

Dim i As Long
i = c + 1
Do While i <= Len(t) And (Mid$(t, i, 1) = " " Or Mid$(t, i, 1) = vbTab)
    i = i + 1
Loop

If i > Len(t) Then Exit Function
If Mid$(t, i, 1) <> """" Then Exit Function

Dim endPos As Long
endPos = ContextKV_FindEndJsonString(t, i)
If endPos = 0 Then Exit Function

Dim raw As String
raw = Mid$(t, i + 1, endPos - (i + 1))
ContextKV_JsonGetFilePointer = ContextKV_JsonUnescape(raw)
End Function
' -----------------------------
' Persistência @file
' -----------------------------
Private Function ContextKV_SaveTextAsFileAndReturnPointer( _
ByVal outputFolderBase As String, _
ByVal runToken As String, _
ByVal pipelineName As String, _
ByVal stepN As Long, _
ByVal promptId As String, _
ByVal varName As String, _
ByVal content As String, _
ByVal ext As String, _
ByRef outSavedPath As String _
) As String
On Error GoTo EH

Dim base As String
base = Trim$(outputFolderBase)
If base = "" Then base = ThisWorkbook.path

Dim subFolder As String
subFolder = ContextKV_GetTextConfig("CONTEXT_KV_OUTPUT_SUBFOLDER", CONTEXTKV_DEFAULT_SUBFOLDER)
If Trim$(subFolder) = "" Then subFolder = CONTEXTKV_DEFAULT_SUBFOLDER

Dim folderPath As String
folderPath = ContextKV_JoinPath(base, subFolder)
ContextKV_EnsureFolder folderPath

Dim stamp As String
stamp = Format$(Now, "yyyy-mm-dd_hhnn")

Dim safePipe As String, safeRun As String, safeVar As String
safePipe = ContextKV_SanitizeFilenamePart(pipelineName)
safeRun = ContextKV_SanitizeFilenamePart(runToken)
safeVar = ContextKV_SanitizeFilenamePart(varName)

Dim fname As String
fname = safePipe & "_" & safeRun & "_step" & Format$(stepN, "00") & "_" & safeVar & "_" & stamp & "." & ext

Dim fullPath As String
fullPath = ContextKV_JoinPath(folderPath, fname)

Call ContextKV_WriteTextFileUTF8(fullPath, content)
outSavedPath = fullPath

' Guardar ponteiro RELATIVO (subpasta + ficheiro) para portabilidade
Dim rel As String
rel = subFolder & "\" & fname

ContextKV_SaveTextAsFileAndReturnPointer = "{""@file"":""" & ContextKV_JsonEscape(rel) & """}"
Exit Function
EH:
outSavedPath = ""
' fallback: tenta guardar inline (ainda que possa falhar no caller)
ContextKV_SaveTextAsFileAndReturnPointer = "{""@file"":""" & ContextKV_JsonEscape("") & """}"
End Function
Private Sub ContextKV_EnsureFolder(ByVal folderPath As String)
On Error Resume Next
If Dir(folderPath, vbDirectory) = "" Then
MkDir folderPath
End If
On Error GoTo 0
End Sub
Private Function ContextKV_ResolveFilePath(ByVal outputFolderBase As String, ByVal maybeRelative As String) As String
Dim p As String
p = Trim$(maybeRelative)
If p = "" Then Exit Function
If InStr(1, p, ":\", vbTextCompare) > 0 Or Left$(p, 2) = "\\" Then
    ContextKV_ResolveFilePath = p
    Exit Function
End If

Dim base As String
base = Trim$(outputFolderBase)
If base = "" Then base = ThisWorkbook.path

ContextKV_ResolveFilePath = ContextKV_JoinPath(base, p)
End Function
Private Function ContextKV_JoinPath(ByVal a As String, ByVal b As String) As String
If a = "" Then
ContextKV_JoinPath = b
Exit Function
End If
If b = "" Then
ContextKV_JoinPath = a
Exit Function
End If
Dim aa As String, bb As String
aa = a: bb = b
If Right$(aa, 1) = "\" Or Right$(aa, 1) = "/" Then aa = Left$(aa, Len(aa) - 1)
If Left$(bb, 1) = "\" Or Left$(bb, 1) = "/" Then bb = Mid$(bb, 2)
ContextKV_JoinPath = aa & "\" & bb
End Function
Private Function ContextKV_SanitizeFilenamePart(ByVal s As String) As String
Dim t As String
t = ContextKV_RemoveDiacritics(s)
Dim i As Long, ch As String, out As String
out = ""
For i = 1 To Len(t)
    ch = Mid$(t, i, 1)
    If ch Like "[A-Za-z0-9]" Or ch = "_" Or ch = "-" Then
        out = out & ch
    ElseIf ch = " " Or ch = "|" Or ch = "/" Or ch = "\" Or ch = ":" Then
        out = out & "_"
    Else
        out = out & "_"
    End If
Next i

Do While InStr(out, "__") > 0
    out = Replace(out, "__", "_")
Loop

If Len(out) > 60 Then out = Left$(out, 60)
If out = "" Then out = "X"

ContextKV_SanitizeFilenamePart = out
End Function
Private Function ContextKV_FileExists(ByVal fullPath As String) As Boolean
On Error Resume Next
ContextKV_FileExists = (Dir(fullPath) <> "")
On Error GoTo 0
End Function
Private Sub ContextKV_WriteTextFileUTF8(ByVal fullPath As String, ByVal content As String)
Dim stm As Object
Set stm = CreateObject("ADODB.Stream")
stm.Type = 2 ' text
stm.Charset = "utf-8"
stm.Open
stm.WriteText content
stm.SaveToFile fullPath, 2 ' overwrite
stm.Close
End Sub
Private Function ContextKV_ReadTextFileUTF8(ByVal fullPath As String) As String
On Error GoTo EH
Dim stm As Object
Set stm = CreateObject("ADODB.Stream")
stm.Type = 2
stm.Charset = "utf-8"
stm.Open
stm.LoadFromFile fullPath
ContextKV_ReadTextFileUTF8 = stm.ReadText
stm.Close
Exit Function
EH:
ContextKV_ReadTextFileUTF8 = ""
End Function
' -----------------------------
' Output longo: ler "FULL_OUTPUT_SAVED"
' -----------------------------
Private Function ContextKV_ResolveFullOutput(ByVal rawOut As String) As String
Dim t As String
t = CStr(rawOut)
Dim p As Long
p = InStr(1, t, MARKER_FULL_OUTPUT_SAVED, vbTextCompare)
If p > 0 Then
    Dim pEnd As Long
    pEnd = InStr(p, t, "]]", vbTextCompare)
    If pEnd > p Then
        Dim inside As String
        inside = Mid$(t, p + Len(MARKER_FULL_OUTPUT_SAVED), pEnd - (p + Len(MARKER_FULL_OUTPUT_SAVED)))
        inside = Trim$(inside)
        ' inside pode começar com espaço
        If Left$(inside, 1) = " " Then inside = Trim$(inside)
        
        If ContextKV_FileExists(inside) Then
            Dim content As String
            content = ContextKV_ReadTextFileUTF8(inside)
            If content <> "" Then
                ContextKV_ResolveFullOutput = content
                Exit Function
            End If
        End If
    End If
End If

ContextKV_ResolveFullOutput = rawOut
End Function
' -----------------------------
' Lazy upsert em captured_vars (srcRow)
' -----------------------------
Private Sub ContextKV_UpsertCapturedVar( _
ByVal wsS As Worksheet, _
ByVal srcRow As Long, _
ByVal keyName As String, _
ByVal value As String, _
ByVal fmt As String, _
ByVal method As String, _
ByVal outputFolderBase As String, _
ByVal howTag As String _
)
On Error GoTo EH
Dim colCap As Long, colMeta As Long
colCap = ContextKV_FindColumnByHeader(wsS, "captured_vars")
colMeta = ContextKV_FindColumnByHeader(wsS, "captured_vars_meta")
If colCap = 0 Or colMeta = 0 Then Exit Sub

Dim capJson As String
capJson = CStr(wsS.Cells(srcRow, colCap).value)
capJson = ContextKV_ResolveJsonPointerIfNeeded(capJson, outputFolderBase)

Dim metaJson As String
metaJson = CStr(wsS.Cells(srcRow, colMeta).value)
metaJson = ContextKV_ResolveJsonPointerIfNeeded(metaJson, outputFolderBase)

' Estratégia simples: se vazio, criar novo objecto; caso contrário, se já tem key, não mexe
If InStr(1, capJson, """" & keyName & """", vbTextCompare) > 0 Then Exit Sub

Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
' Não tentamos parse completo (mínimo): só anexar no fim antes de "}"

Dim jsonVal As String
jsonVal = """" & ContextKV_JsonEscape(value) & """"

Dim maxChars As Long
maxChars = ContextKV_GetLongConfig("CONTEXT_KV_MAX_CELL_CHARS", CONTEXTKV_DEFAULT_MAX_CELL_CHARS)

If Len(jsonVal) > maxChars Then
    Dim savedPath As String
    jsonVal = ContextKV_SaveTextAsFileAndReturnPointer(outputFolderBase, "LAZY", "LAZY", 0, "", keyName, value, IIf(fmt = "json_like", "json", "txt"), savedPath)
End If

If Trim$(capJson) = "" Then
    capJson = "{""" & keyName & """:" & jsonVal & "}"
ElseIf Right$(Trim$(capJson), 1) = "}" Then
    capJson = Left$(Trim$(capJson), Len(Trim$(capJson)) - 1)
    If Right$(capJson, 1) <> "{" Then capJson = capJson & ","
    capJson = capJson & """" & keyName & """:" & jsonVal & "}"
Else
    ' não é JSON esperado; aborta
    Exit Sub
End If

Dim sha As String
sha = ContextKV_TrySHA256(value)
Dim metaFrag As String
metaFrag = ContextKV_JsonObj_MetaOK(fmt, Len(value), sha, method & "|lazy=" & howTag)

If Trim$(metaJson) = "" Then
    metaJson = "{""" & keyName & """:" & metaFrag & "}"
ElseIf Right$(Trim$(metaJson), 1) = "}" Then
    metaJson = Left$(Trim$(metaJson), Len(Trim$(metaJson)) - 1)
    If Right$(metaJson, 1) <> "{" Then metaJson = metaJson & ","
    metaJson = metaJson & """" & keyName & """:" & metaFrag & "}"
Else
    ' ignora meta
End If

If Len(capJson) <= maxChars Then wsS.Cells(srcRow, colCap).value = capJson
If Len(metaJson) <= maxChars Then wsS.Cells(srcRow, colMeta).value = metaJson

Exit Sub
EH:
' silencioso
End Sub
' -----------------------------
' Config helpers
' -----------------------------
Private Function ContextKV_GetTextConfig(ByVal keyName As String, ByVal defaultValue As String) As String
On Error GoTo EH
Dim ws As Worksheet
Set ws = ContextKV_GetSheet("Config")
If ws Is Nothing Then
    ContextKV_GetTextConfig = defaultValue
    Exit Function
End If

Dim lastRow As Long
lastRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row

Dim r As Long
For r = 1 To lastRow
    If UCase$(Trim$(CStr(ws.Cells(r, 1).value))) = UCase$(keyName) Then
        ContextKV_GetTextConfig = CStr(ws.Cells(r, 2).value)
        Exit Function
    End If
Next r

ContextKV_GetTextConfig = defaultValue
Exit Function
EH:
ContextKV_GetTextConfig = defaultValue
End Function
Private Function ContextKV_GetBoolConfig(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
Dim t As String
t = UCase$(Trim$(ContextKV_GetTextConfig(keyName, IIf(defaultValue, "TRUE", "FALSE"))))
If t = "TRUE" Or t = "1" Or t = "YES" Or t = "SIM" Then
    ContextKV_GetBoolConfig = True
ElseIf t = "FALSE" Or t = "0" Or t = "NO" Or t = "NAO" Or t = "NÃO" Then
    ContextKV_GetBoolConfig = False
Else
    ContextKV_GetBoolConfig = defaultValue
End If
End Function
Private Function ContextKV_GetLongConfig(ByVal keyName As String, ByVal defaultValue As Long) As Long
Dim t As String
t = Trim$(ContextKV_GetTextConfig(keyName, CStr(defaultValue)))
If IsNumeric(t) Then
ContextKV_GetLongConfig = CLng(t)
Else
ContextKV_GetLongConfig = defaultValue
End If
End Function
' -----------------------------
' Hash (opcional; usa função existente em M09 se existir)
' -----------------------------
Private Function ContextKV_TrySHA256(ByVal text As String) As String
On Error GoTo EH
' Preferência: função já existente no PIPELINER (M09_FilesManagement)
' Se não existir, devolve vazio (não falha).
ContextKV_TrySHA256 = Files_SHA256_Text(text)
Exit Function
EH:
ContextKV_TrySHA256 = ""
End Function
' -----------------------------
' Logging (DEBUG)
' -----------------------------
Private Sub ContextKV_LogEvent(ByVal pipelineName As String, ByVal stepN As Long, ByVal promptId As String, ByVal action As String, ByVal varName As String, ByVal result As String, ByVal details As String)
On Error Resume Next
Dim ws As Worksheet
Set ws = ContextKV_GetSheet("DEBUG")
If ws Is Nothing Then Exit Sub

' garantir colunas
ContextKV_EnsureDebugColumns

Dim map As Object
Set map = ContextKV_HeaderMap(ws)

Dim newRow As Long
newRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1
If newRow < 2 Then newRow = 2

' preencher colunas antigas (compatibilidade)
If map.exists(ContextKV_NormalizeHeader("Timestamp")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("Timestamp"))).value = Now
If map.exists(ContextKV_NormalizeHeader("Passo")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("Passo"))).value = stepN
If map.exists(ContextKV_NormalizeHeader("Prompt ID")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("Prompt ID"))).value = promptId
If map.exists(ContextKV_NormalizeHeader("Severidade")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("Severidade"))).value = result

' usar campos existentes para manter auditabilidade em templates antigos
If map.exists(ContextKV_NormalizeHeader("Parâmetro")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("Parâmetro"))).value = action & IIf(varName <> "", "|" & varName, "")
If map.exists(ContextKV_NormalizeHeader("Problema")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("Problema"))).value = details
If map.exists(ContextKV_NormalizeHeader("Sugestão")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("Sugestão"))).value = ""

' novas colunas estruturadas (se existirem)
If map.exists(ContextKV_NormalizeHeader("pipeline_name")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("pipeline_name"))).value = pipelineName
If map.exists(ContextKV_NormalizeHeader("action")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("action"))).value = action
If map.exists(ContextKV_NormalizeHeader("var")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("var"))).value = varName
If map.exists(ContextKV_NormalizeHeader("result")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("result"))).value = result
If map.exists(ContextKV_NormalizeHeader("details")) Then ws.Cells(newRow, map(ContextKV_NormalizeHeader("details"))).value = details
End Sub
Private Sub ContextKV_DictSet(ByVal d As Object, ByVal k As String, ByVal v As Variant)
On Error Resume Next
If d.exists(k) Then
d(k) = v
Else
d.Add k, v
End If
On Error GoTo 0
End Sub


