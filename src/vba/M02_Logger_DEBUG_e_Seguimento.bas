Attribute VB_Name = "M02_Logger_DEBUG_e_Seguimento"
Option Explicit

Private Const SHEET_DEBUG As String = "DEBUG"
Private Const SHEET_SEGUIMENTO As String = "Seguimento"
Private Const SHEET_HISTORICO As String = "HISTÓRICO"

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
    ByVal sugestao As String _
)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEBUG)

    Dim mapa As Object
    Set mapa = Debug_MapaCabecalhos(ws)

    Dim novaLinha As Long
    novaLinha = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1

    Debug_SetValue ws, mapa, novaLinha, "Timestamp", Now
    Debug_SetValue ws, mapa, novaLinha, "Passo", passo
    Debug_SetValue ws, mapa, novaLinha, "Prompt ID", promptId
    Debug_SetValue ws, mapa, novaLinha, "Severidade", severidade
    Debug_SetValue ws, mapa, novaLinha, "Linha (Config extra)", linhaConfigExtra
    Debug_SetValue ws, mapa, novaLinha, "Parametro", parametro          ' aceita "Parâmetro" no Excel
    Debug_SetValue ws, mapa, novaLinha, "Problema", problema
    Debug_SetValue ws, mapa, novaLinha, "Sugestao", sugestao            ' aceita "Sugestão" no Excel
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
