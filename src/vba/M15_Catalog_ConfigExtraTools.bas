Attribute VB_Name = "M15_Catalog_ConfigExtraTools"
Option Explicit

' =============================================================================
' Módulo: M15_Catalog_ConfigExtraTools
' Propósito:
' - Criar folha de catálogo compatível com o layout do PIPELINER (blocos de 5 linhas).
' - Executar diagnóstico sequencial de Config Extra para troubleshooting de payload JSON.
' - Gerar evidência auditável numa folha dedicada sem alterar layout estrutural de folhas core.
'
' Atualizações:
' - 2026-02-17 | Codex | Criação de toolkit de catálogo + diagnóstico de Config Extra
'   - Adiciona macro para criar folha catálogo modelo (headers + bloco base + instruções Next PROMPT).
'   - Adiciona macro de testes sequenciais de Config Extra com relatório em CONFIG_EXTRA_TESTS.
'   - Inclui pré-validação estrutural de JSON para detetar fecho sem abertura e outras quebras.
'
' Funções e procedimentos:
' - TOOL_CreateCatalogTemplateSheet()
'   - Cria uma folha de catálogo pronta para uso no formato esperado pelo PIPELINER.
' - TOOL_RunConfigExtraSequentialDiagnostics()
'   - Corre bateria de casos de Config Extra e regista resultados na folha CONFIG_EXTRA_TESTS + DEBUG.
' =============================================================================

Private Const DEFAULT_TEMPLATE_SHEET As String = "CATALOGO_MODELO"
Private Const DIAG_SHEET As String = "CONFIG_EXTRA_TESTS"

Public Sub TOOL_CreateCatalogTemplateSheet()
    On Error GoTo EH

    Dim ws As Worksheet
    Dim nome As String

    nome = Trim$(InputBox("Nome da nova folha de catálogo:", "PIPELINER - Criar Catálogo", DEFAULT_TEMPLATE_SHEET))
    If nome = "" Then Exit Sub

    If WorksheetExists(nome) Then
        MsgBox "Já existe uma folha com o nome '" & nome & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nome

    Call Catalog_WriteHeaders(ws)
    Call Catalog_WriteBlockSkeleton(ws, nome)

    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 42
    ws.Columns("D").ColumnWidth = 120
    ws.Columns("E:H").ColumnWidth = 20
    ws.Columns("I:K").ColumnWidth = 28
    ws.Rows("1:6").EntireRow.AutoFit

    MsgBox "Folha de catálogo criada: " & nome, vbInformation
    Exit Sub

EH:
    MsgBox "TOOL_CreateCatalogTemplateSheet falhou: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub TOOL_RunConfigExtraSequentialDiagnostics()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = EnsureDiagSheet()

    Call Diag_WriteHeaders(ws)

    Dim casos As Collection
    Set casos = BuildConfigExtraCases()

    Dim i As Long
    Dim linha As Long
    linha = 2

    For i = 1 To casos.Count
        Dim item As Object
        Set item = casos(i)

        Dim configExtra As String
        configExtra = CStr(item("config"))

        Dim auditJson As String
        Dim inputJson As String
        Dim extraFragment As String

        Call ConfigExtra_Converter(configExtra, "PROMPT_FALLBACK", i, "DIAG/ConfigExtra/" & CStr(i) & "/A", auditJson, inputJson, extraFragment)

        Dim modos As String
        modos = "Web search"

        Dim fragComFO As String
        fragComFO = extraFragment
        Call FileOutput_PrepareRequest("file", "metadata", "json_schema", modos, fragComFO)

        Dim payload As String
        payload = BuildPayloadProbe(fragComFO)

        Dim detail As String
        Dim okJson As Boolean
        okJson = JsonStructuralQuickCheck(payload, detail)

        ws.Cells(linha, 1).Value = i
        ws.Cells(linha, 2).Value = CStr(item("nome"))
        ws.Cells(linha, 3).Value = configExtra
        ws.Cells(linha, 4).Value = auditJson
        ws.Cells(linha, 5).Value = inputJson
        ws.Cells(linha, 6).Value = fragComFO
        ws.Cells(linha, 7).Value = IIf(okJson, "OK", "ERRO")
        ws.Cells(linha, 8).Value = detail
        ws.Cells(linha, 9).Value = Left$(payload, 700)

        If okJson Then
            Call Debug_Registar(0, "M15_CONFIG_EXTRA_DIAG", "INFO", "", "CONFIG_EXTRA_CASE_OK", _
                "Caso " & CStr(i) & " válido: " & CStr(item("nome")), "Folha " & DIAG_SHEET & " contém detalhes.")
        Else
            Call Debug_Registar(0, "M15_CONFIG_EXTRA_DIAG", "ERRO", "", "CONFIG_EXTRA_CASE_FAIL", _
                "Caso " & CStr(i) & " inválido: " & CStr(item("nome")) & " | " & detail, _
                "Rever Config extra e fragment merge (Config extra + File Output).")
        End If

        linha = linha + 1
    Next i

    ws.Columns("A:I").EntireColumn.AutoFit
    ws.Activate
    MsgBox "Diagnóstico concluído. Ver folha '" & DIAG_SHEET & "' e DEBUG.", vbInformation
    Exit Sub

EH:
    MsgBox "TOOL_RunConfigExtraSequentialDiagnostics falhou: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub Catalog_WriteHeaders(ByVal ws As Worksheet)
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "Nome curto"
    ws.Range("C1").Value = "Nome descritivo"
    ws.Range("D1").Value = "Texto prompt"
    ws.Range("E1").Value = "Modelo"
    ws.Range("F1").Value = "Modos"
    ws.Range("G1").Value = "Storage"
    ws.Range("H1").Value = "Config extra"
    ws.Range("I1").Value = "Comentários"
    ws.Range("J1").Value = "Notas para desenvolvimento"
    ws.Range("K1").Value = "Histórico de versões"

    ws.Rows(1).Font.Bold = True
End Sub

Private Sub Catalog_WriteBlockSkeleton(ByVal ws As Worksheet, ByVal sheetName As String)
    Dim idBase As String
    idBase = sheetName & "/01/NomeCurto/A"

    ws.Range("A2").Value = idBase
    ws.Range("B2").Value = "NomeCurto"
    ws.Range("C2").Value = "Descrição do prompt"
    ws.Range("D2").Value = "ROLE" & vbLf & "Descreva aqui o prompt principal."
    ws.Range("E2").Value = "gpt-5.2"
    ws.Range("F2").Value = "Web search"
    ws.Range("G2").Value = "TRUE"
    ws.Range("H2").Value = "output_kind: file" & vbLf & "process_mode: metadata" & vbLf & "structured_outputs_mode: json_schema"
    ws.Range("I2").Value = "Exemplo base"
    ws.Range("K2").Value = "A — versão inicial"

    ws.Range("B3").Value = "Next PROMPT: STOP"
    ws.Range("B4").Value = "Next PROMPT default: STOP"
    ws.Range("B5").Value = "Next PROMPT allowed: STOP"

    ws.Range("C3").Value = "Descrição textual:"
    ws.Range("D3").Value = "Resumo do objetivo do prompt."
    ws.Range("C4").Value = "INPUTS:"
    ws.Range("D4").Value = "URLS_ENTRADA: <https://exemplo.pt/noticia>" & vbLf & "FILES: GUIA_DE_ESTILO.pdf (latest) (as pdf)"
    ws.Range("C5").Value = "OUTPUTS:"
    ws.Range("D5").Value = "1 ficheiro TXT UTF-8 (manifest metadata)."

    ws.Range("A2:K5").WrapText = True
End Sub

Private Function BuildConfigExtraCases() As Collection
    Dim c As New Collection

    c.Add MakeCase("Escalar simples", "truncation: auto")
    c.Add MakeCase("Lista include válida", "include: [web_search_call.action.sources]")
    c.Add MakeCase("Nesting com pontos", "text.format.type: json_schema")
    c.Add MakeCase("Objeto simples válido", "metadata: {projeto: CPSA, versao: 1}")
    c.Add MakeCase("Bloco input válido", "input:" & vbLf & "  role: user" & vbLf & "  content: Mensagem de teste")
    c.Add MakeCase("Linha sem separador", "linha_sem_dois_pontos")
    c.Add MakeCase("Conflito conversation + previous_response_id", "conversation: conv_123" & vbLf & "previous_response_id: resp_123")
    c.Add MakeCase("Chave proibida tools", "tools: [{type:web_search}]")
    c.Add MakeCase("Objeto mal formado", "text.format: {type: json_schema")
    c.Add MakeCase("Caso semelhante ao incidente", "output_kind: file" & vbLf & "process_mode: metadata" & vbLf & "structured_outputs_mode: json_schema" & vbLf & "truncation: auto" & vbLf & "include: [web_search_call.action.sources]" & vbLf & "auto_save: TRUE" & vbLf & "overwrite_mode: suffix")

    Set BuildConfigExtraCases = c
End Function

Private Function MakeCase(ByVal nome As String, ByVal config As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("nome") = nome
    d("config") = config
    Set MakeCase = d
End Function

Private Function BuildPayloadProbe(ByVal extraFragment As String) As String
    Dim p As String
    p = "{""model"":""gpt-5.2"",""input"":[{""role"":""user"",""content"":""probe""}]"

    If Trim$(extraFragment) <> "" Then
        p = p & "," & extraFragment
    End If

    p = p & "}"
    BuildPayloadProbe = p
End Function

Private Function JsonStructuralQuickCheck(ByVal jsonText As String, ByRef outDetail As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim stack As String

    inString = False
    escaped = False
    stack = ""
    outDetail = "ok"

    For i = 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)

        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            End If
        Else
            Select Case ch
                Case """"
                    inString = True
                Case "{"
                    stack = stack & "}"
                Case "["
                    stack = stack & "]"
                Case "}", "]"
                    If Len(stack) = 0 Then
                        outDetail = "fecho_sem_abertura @pos=" & CStr(i) & " char=" & ch
                        JsonStructuralQuickCheck = False
                        Exit Function
                    End If
                    If Right$(stack, 1) <> ch Then
                        outDetail = "fecho_incompativel @pos=" & CStr(i) & " esperado=" & Right$(stack, 1) & " recebido=" & ch
                        JsonStructuralQuickCheck = False
                        Exit Function
                    End If
                    stack = Left$(stack, Len(stack) - 1)
            End Select
        End If
    Next i

    If inString Then
        outDetail = "string_nao_fechada"
        JsonStructuralQuickCheck = False
        Exit Function
    End If

    If Len(stack) > 0 Then
        outDetail = "estrutura_nao_fechada esperado=" & Right$(stack, 1)
        JsonStructuralQuickCheck = False
        Exit Function
    End If

    JsonStructuralQuickCheck = True
End Function

Private Function EnsureDiagSheet() As Worksheet
    Dim ws As Worksheet

    If WorksheetExists(DIAG_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(DIAG_SHEET)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DIAG_SHEET
    End If

    Set EnsureDiagSheet = ws
End Function

Private Sub Diag_WriteHeaders(ByVal ws As Worksheet)
    ws.Range("A1").Value = "#"
    ws.Range("B1").Value = "Caso"
    ws.Range("C1").Value = "Config extra (input)"
    ws.Range("D1").Value = "Config extra (audit JSON)"
    ws.Range("E1").Value = "input_json"
    ws.Range("F1").Value = "extraFragment + FileOutput"
    ws.Range("G1").Value = "Preflight estrutural"
    ws.Range("H1").Value = "Detalhe"
    ws.Range("I1").Value = "Payload preview"

    ws.Rows(1).Font.Bold = True
    ws.Columns("C:F").ColumnWidth = 60
    ws.Columns("I").ColumnWidth = 80
    ws.Columns("C:I").WrapText = True
End Sub

Private Function WorksheetExists(ByVal nome As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Worksheets(nome) Is Nothing
    On Error GoTo 0
End Function
