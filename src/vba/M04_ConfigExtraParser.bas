Attribute VB_Name = "M04_ConfigExtraParser"
Option Explicit

' =============================================================================
' Módulo: M04_ConfigExtraParser
' Propósito:
' - Converter o campo amigável "Config extra" em JSON válido para a API.
' - Validar sintaxe, chaves proibidas e coerência de parâmetros com logging em DEBUG.
'
' Atualizações:
' - 2026-02-16 | Codex | Correção de serialização de dicionários aninhados e helper de dump
'   - Evita erro 450 ao ler itens Object de Scripting.Dictionary sem Set.
'   - Adiciona helper público para dump recursivo de dicionários na janela Immediate.
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - ConfigExtra_Converter (Sub): rotina pública do módulo.
' - ConfigExtra_DebugDumpDictionary (Sub): imprime recursivamente keys/tipos/valores no Immediate.
' =============================================================================


' =============================================================================
' PARSER DIDÁCTICO: "Config extra" amigável -> JSON
'
' Convenção:
'   - Cada linha: chave: valor
'   - Separador de linhas: ALT+ENTER (Excel)
'   - input como bloco:
'       input:
'         role: user
'         content: ...
'
' Comportamento:
'   - Se encontrar erros de sintaxe: regista alerta no DEBUG e ignora apenas essa parte
'   - Se houver conflitos (ex.: conversation + previous_response_id): ignora previous_response_id e alerta
'   - Proíbe chaves dedicadas (model/temperature/max_output_tokens/store/tools): ignora e alerta
' =============================================================================

Private Function NormalizarQuebrasDeLinha(ByVal texto As String) As String

    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)

    ' Se vier colado com "\n" literal (sem quebras reais), converter para vbLf
    If InStr(1, texto, vbLf) = 0 Then
        If InStr(1, texto, "\n") > 0 Then texto = Replace(texto, "\n", vbLf)
    End If

    NormalizarQuebrasDeLinha = texto

End Function



Private Function IsIndentado(ByVal linhaOriginal As String) As Boolean
    If Len(linhaOriginal) = 0 Then
        IsIndentado = False
    Else
        Dim c As String
        c = Left$(linhaOriginal, 1)
        IsIndentado = (c = " " Or c = vbTab)
    End If
End Function


Private Function SplitPrimeiro(ByVal s As String, ByVal separador As String, ByRef esquerda As String, ByRef direita As String) As Boolean
    Dim p As Long
    p = InStr(1, s, separador)
    If p = 0 Then
        SplitPrimeiro = False
        Exit Function
    End If
    esquerda = Trim$(Left$(s, p - 1))
    direita = Trim$(Mid$(s, p + Len(separador)))
    SplitPrimeiro = True
End Function


Private Function JsonEscapar(ByVal s As String) As String
    JsonEscapar = Json_EscapeString(CStr(s))
End Function


' ---- util: detectar chaves proibidas (top-level e dot-notation)
Private Function ChaveProibida(ByVal chave As String) As Boolean
    Dim k As String
    k = LCase$(Trim$(chave))
    ChaveProibida = (k = "model" Or k = "temperature" Or k = "max_output_tokens" Or k = "store" Or k = "tools")
End Function


' ---- cria/obtém um dicionário (late binding) para permitir nesting via keys com "."
Private Function NovoDict() As Object
    Set NovoDict = CreateObject("Scripting.Dictionary")
End Function


Private Sub Dict_SetPathValue(ByVal dictRaiz As Object, ByVal path As String, ByVal valorTipo As String, ByVal valorJsonLiteral As String)
    Dim partes() As String
    partes = Split(path, ".")

    Dim d As Object
    Set d = dictRaiz

    Dim i As Long
    For i = 0 To UBound(partes) - 1
        Dim p As String
        p = Trim$(partes(i))
        If Not d.exists(p) Then
            Dim nd As Object
            Set nd = NovoDict()
            d.Add p, nd
        End If
        Set d = d(p)
    Next i

    Dim folha As String
    folha = Trim$(partes(UBound(partes)))

    Dim pack(1) As Variant
    pack(0) = valorTipo         ' "raw" ou "str"
    pack(1) = valorJsonLiteral  ' já pronto para JSON (sem chave)
    If d.exists(folha) Then d.Remove folha
    d.Add folha, pack
End Sub


Private Function Dict_ExistsTop(ByVal d As Object, ByVal key As String) As Boolean
    Dict_ExistsTop = d.exists(key)
End Function


Private Sub Dict_RemoveTop(ByVal d As Object, ByVal key As String)
    If d.exists(key) Then d.Remove key
End Sub


Private Function Dict_ToJsonObject(ByVal d As Object) As String
    Dim partes As Collection
    Set partes = New Collection

    Dim k As Variant
    For Each k In d.keys
        Dim v As Variant
        If IsObject(d.Item(k)) Then
            Set v = d.Item(k)
        Else
            v = d.Item(k)
        End If

        Dim jsonValor As String
        If IsObject(v) Then
            jsonValor = Dict_ToJsonObject(v)
        ElseIf IsArray(v) Then
            If v(0) = "raw" Then
                jsonValor = CStr(v(1))
            Else
                jsonValor = """" & JsonEscapar(CStr(v(1))) & """"
            End If
        Else
            ' fallback
            jsonValor = """" & JsonEscapar(CStr(v)) & """"
        End If

        partes.Add """" & JsonEscapar(CStr(k)) & """:" & jsonValor
    Next k

    Dim s As String
    s = "{"
    Dim i As Long
    For i = 1 To partes.Count
        If i > 1 Then s = s & ","
        s = s & partes(i)
    Next i
    s = s & "}"
    Dict_ToJsonObject = s
End Function




Public Sub ConfigExtra_DebugDumpDictionary(ByVal d As Object, Optional ByVal indent As Long = 0)
    If d Is Nothing Then
        Debug.Print Space$(indent) & "<Nothing>"
        Exit Sub
    End If

    Dim k As Variant
    For Each k In d.keys
        Dim v As Variant
        If IsObject(d.Item(k)) Then
            Set v = d.Item(k)
        Else
            v = d.Item(k)
        End If

        If IsObject(v) Then
            Debug.Print Space$(indent) & CStr(k) & " -> " & TypeName(v)
            ConfigExtra_DebugDumpDictionary v, indent + 2
        ElseIf IsArray(v) Then
            Dim payloadType As String
            Dim payloadValue As String
            payloadType = CStr(v(0))
            payloadValue = CStr(v(1))
            Debug.Print Space$(indent) & CStr(k) & " -> Variant() [" & payloadType & "] " & payloadValue
        Else
            Debug.Print Space$(indent) & CStr(k) & " -> " & TypeName(v) & " = " & CStr(v)
        End If
    Next k
End Sub

Private Function JsonArrayDeStrings(ByVal valores As Collection) As String
    Dim s As String
    s = "["
    Dim i As Long
    For i = 1 To valores.Count
        If i > 1 Then s = s & ","
        s = s & """" & JsonEscapar(CStr(valores(i))) & """"
    Next i
    s = s & "]"
    JsonArrayDeStrings = s
End Function


Private Function ParseListaSimples(ByVal valor As String, ByVal passo As Long, ByVal promptId As String, ByVal linhaN As Long, ByVal parametro As String, ByRef ok As Boolean) As String
    ' Espera: [a, b, c] (strings; com ou sem aspas)
    ok = False
    valor = Trim$(valor)

    If Left$(valor, 1) <> "[" Or Right$(valor, 1) <> "]" Then
        Debug_Registar passo, promptId, "ALERTA", linhaN, parametro, _
            "Lista mal formada: deve começar por '[' e terminar em ']'.", _
            "Ex.: include: [web_search_call.action.sources]"
        Exit Function
    End If

    Dim inner As String
    inner = Trim$(Mid$(valor, 2, Len(valor) - 2))

    Dim itens As New Collection
    If inner <> "" Then
        Dim partes() As String
        partes = Split(inner, ",")
        Dim i As Long
        For i = 0 To UBound(partes)
            Dim item As String
            item = Trim$(partes(i))
            If item <> "" Then
                ' remove aspas se existirem
                If Left$(item, 1) = """" And Right$(item, 1) = """" And Len(item) >= 2 Then
                    item = Mid$(item, 2, Len(item) - 2)
                End If
                itens.Add item
            End If
        Next i
    End If

    ok = True
    ParseListaSimples = JsonArrayDeStrings(itens)
End Function


Private Function ParseObjectoSimples(ByVal valor As String, ByVal passo As Long, ByVal promptId As String, ByVal linhaN As Long, ByVal parametro As String, ByRef ok As Boolean) As String
    ' Espera: {k: v, k2: v2} (valores simples; sem nesting profundo)
    ok = False
    valor = Trim$(valor)

    If Left$(valor, 1) <> "{" Or Right$(valor, 1) <> "}" Then
        Debug_Registar passo, promptId, "ALERTA", linhaN, parametro, _
            "Objecto mal formado: deve começar por '{' e terminar em '}'.", _
            "Ex.: metadata: {projeto: AvalCap, versao: A}"
        Exit Function
    End If

    Dim inner As String
    inner = Trim$(Mid$(valor, 2, Len(valor) - 2))

    Dim d As Object
    Set d = NovoDict()

    If inner <> "" Then
        Dim pares() As String
        pares = Split(inner, ",")
        Dim i As Long
        For i = 0 To UBound(pares)
            Dim par As String
            par = Trim$(pares(i))
            If par = "" Then GoTo ProximoPar

            Dim k As String, v As String
            If Not SplitPrimeiro(par, ":", k, v) Then
                Debug_Registar passo, promptId, "ALERTA", linhaN, parametro, _
                    "Par inválido em objecto (falta ':'): " & par, _
                    "Use o formato {chave: valor, ...}"
                GoTo ProximoPar
            End If

            ' valores: boolean/número/strings
            Dim vJsonTipo As String, vJsonLiteral As String
            Call ConverterValorParaJson(v, passo, promptId, linhaN, parametro, vJsonTipo, vJsonLiteral)

            ' guardar como folha simples (sem dot nesting dentro do objecto)
            Dim pack(1) As Variant
            pack(0) = vJsonTipo
            If vJsonTipo = "raw" Then
                pack(1) = vJsonLiteral
            Else
                pack(1) = v ' string original (será escapada no Dict_ToJsonObject)
            End If

            If d.exists(k) Then d.Remove k
            d.Add k, pack

ProximoPar:
        Next i
    End If

    ok = True
    ParseObjectoSimples = Dict_ToJsonObject(d)
End Function


Private Sub ConverterValorParaJson( _
    ByVal valorTexto As String, _
    ByVal passo As Long, ByVal promptId As String, ByVal linhaN As Long, ByVal parametro As String, _
    ByRef outTipo As String, ByRef outJsonLiteralOrText As String _
)
    ' outTipo:
    '   - "raw": outJsonLiteralOrText contém literal JSON (true/123/{"a":1}/["x"])
    '   - "str": outJsonLiteralOrText contém texto (será escapado e entre aspas)

    Dim v As String
    v = Trim$(valorTexto)

    If v = "" Then
        outTipo = "str"
        outJsonLiteralOrText = ""
        Exit Sub
    End If

    Dim vl As String
    vl = LCase$(v)

    If vl = "true" Or vl = "false" Then
        outTipo = "raw": outJsonLiteralOrText = vl
        Exit Sub
    End If

    If IsNumeric(Replace(v, ",", ".")) Then
        outTipo = "raw": outJsonLiteralOrText = Replace(v, ",", ".")
        Exit Sub
    End If

    ' listas/objectos amigáveis
    If Left$(v, 1) = "[" Then
        Dim okList As Boolean
        Dim jsonList As String
        jsonList = ParseListaSimples(v, passo, promptId, linhaN, parametro, okList)
        If okList Then
            outTipo = "raw": outJsonLiteralOrText = jsonList
        Else
            outTipo = "str": outJsonLiteralOrText = "" ' ignora
        End If
        Exit Sub
    End If

    If Left$(v, 1) = "{" Then
        Dim okObj As Boolean
        Dim jsonObj As String
        jsonObj = ParseObjectoSimples(v, passo, promptId, linhaN, parametro, okObj)
        If okObj Then
            outTipo = "raw": outJsonLiteralOrText = jsonObj
        Else
            outTipo = "str": outJsonLiteralOrText = "" ' ignora
        End If
        Exit Sub
    End If

    ' aspas opcionais
    If Left$(v, 1) = """" And Right$(v, 1) = """" And Len(v) >= 2 Then
        v = Mid$(v, 2, Len(v) - 2)
    End If

    outTipo = "str": outJsonLiteralOrText = v
End Sub

Private Function IsChaveInternaFileOutput(ByVal chave As String) As Boolean

    Dim k As String
    k = LCase$(Trim$(chave))

    ' File Output (interno PIPELINER) — não deve ir para o request /v1/responses
    If k = "output_kind" Or k = "process_mode" Or k = "auto_save" Or k = "overwrite_mode" Then
        IsChaveInternaFileOutput = True
        Exit Function
    End If

    If k = "file_name_prefix_template" Or k = "subfolder_template" Then
        IsChaveInternaFileOutput = True
        Exit Function
    End If

    If k = "pptx_mode" Or k = "xlsx_mode" Or k = "pdf_mode" Or k = "image_mode" Or k = "structured_outputs_mode" Then
        IsChaveInternaFileOutput = True
        Exit Function
    End If

    ' Qualquer chave "file_*" é tratada como interna ao PIPELINER (ex.: file_output_encoding)
    If Left$(k, 5) = "file_" Then
        IsChaveInternaFileOutput = True
        Exit Function
    End If

    IsChaveInternaFileOutput = False

End Function


Public Sub ConfigExtra_Converter( _
    ByVal configExtraTexto As String, _
    ByVal textoPromptFallback As String, _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByRef outAuditJson As String, _
    ByRef outInputJsonLiteral As String, _
    ByRef outExtraFragmentSemInput As String _
)
    Dim texto As String
    texto = NormalizarQuebrasDeLinha(CStr(configExtraTexto))

    Dim linhas() As String
    If Trim$(texto) = "" Then
        outAuditJson = "{}"
        outInputJsonLiteral = ""
        outExtraFragmentSemInput = ""
        Exit Sub
    End If
    linhas = Split(texto, vbLf)

    Dim d As Object
    Set d = NovoDict()

    Dim emInputBloco As Boolean
    emInputBloco = False

    Dim inputRole As String, inputContent As String
    inputRole = "": inputContent = ""

    Dim i As Long
    For i = 0 To UBound(linhas)
        Dim linhaOriginal As String
        linhaOriginal = CStr(linhas(i))

        Dim linhaTrim As String
        linhaTrim = Trim$(linhaOriginal)

        Dim linhaN As Long
        linhaN = i + 1 ' humano

        If linhaTrim = "" Then
            ' linha vazia termina bloco input
            If emInputBloco Then emInputBloco = False
            GoTo ProximaLinha
        End If

        If Left$(linhaTrim, 1) = "#" Or Left$(linhaTrim, 2) = "//" Then
            GoTo ProximaLinha
        End If

        ' Detectar início de bloco input:
        If LCase$(linhaTrim) = "input:" Then
            emInputBloco = True
            GoTo ProximaLinha
        End If

        If emInputBloco Then
            If Not IsIndentado(linhaOriginal) Then
                ' terminou bloco input
                emInputBloco = False
                ' continua a processar esta linha como normal
            Else
                Dim kIn As String, vIn As String
                If Not SplitPrimeiro(linhaTrim, ":", kIn, vIn) Then
                    Debug_Registar passo, promptId, "ALERTA", linhaN, "input", _
                        "Linha inválida no bloco input (falta ':'): " & linhaTrim, _
                        "Use, por exemplo: 'role: user' e 'content: ...' (indentados)."
                    GoTo ProximaLinha
                End If

                If LCase$(kIn) = "role" Then
                    inputRole = vIn
                ElseIf LCase$(kIn) = "content" Then
                    inputContent = vIn
                Else
                    Debug_Registar passo, promptId, "ALERTA", linhaN, "input." & kIn, _
                        "Chave não reconhecida no bloco input: " & kIn, _
                        "Use apenas 'role' e 'content' (ou deixe vazio)."
                End If

                GoTo ProximaLinha
            End If
        End If

        ' Linha normal: chave: valor
        Dim chave As String, valor As String
        If Not SplitPrimeiro(linhaTrim, ":", chave, valor) Then
            Debug_Registar passo, promptId, "ALERTA", linhaN, "(linha)", _
                "Linha ignorada (não segue 'chave: valor'): " & linhaTrim, _
                "Ex.: truncation: auto"
            GoTo ProximaLinha
        End If

        If ChaveProibida(chave) Then
            Debug_Registar passo, promptId, "ALERTA", linhaN, chave, _
                "Parâmetro proibido em Config extra (já existe em colunas dedicadas).", _
                "Remova '" & chave & "' daqui e use a coluna apropriada (Modelo/Storage/Modos/etc.)."
            GoTo ProximaLinha
        End If

        Dim tipo As String, valJsonOrText As String
        Call ConverterValorParaJson(valor, passo, promptId, linhaN, chave, tipo, valJsonOrText)

        ' Se houve erro de parsing em listas/objectos, ConverterValorParaJson pode devolver "" para string
        If tipo = "str" Then
            Dict_SetPathValue d, chave, "str", valJsonOrText
        Else
            Dict_SetPathValue d, chave, "raw", valJsonOrText
        End If

ProximaLinha:
    Next i

    ' ---- Aplicar regras pós-parse
    ' Conflito conversation + previous_response_id
    If Dict_ExistsTop(d, "conversation") And Dict_ExistsTop(d, "previous_response_id") Then
        Dict_RemoveTop d, "previous_response_id"
        Debug_Registar passo, promptId, "ALERTA", "", "previous_response_id", _
            "Conflito: não usar 'conversation' e 'previous_response_id' em simultâneo. Ignorado previous_response_id.", _
            "Remova um dos dois. Se quer conversa persistente use 'conversation'; caso contrário use 'previous_response_id'."
    End If

    ' ---- Input override (se foi definido em bloco input)
    outInputJsonLiteral = ""
    If Trim$(inputRole) <> "" Or Trim$(inputContent) <> "" Then
        If Trim$(inputRole) = "" Then inputRole = "user"
        If Trim$(inputContent) = "" Then inputContent = textoPromptFallback

        ' input como array de mensagens
        outInputJsonLiteral = "[{""role"":""" & JsonEscapar(inputRole) & """,""content"":""" & JsonEscapar(inputContent) & """}]"

        ' Para auditoria, guardamos também no JSON convertido:
        Dict_SetPathValue d, "input", "raw", outInputJsonLiteral
    End If

    ' ---- Gerar JSON auditável (config extra convertido)
    outAuditJson = Dict_ToJsonObject(d)

    ' ---- Gerar fragmento sem "input" para merge no request
    Dim d2 As Object
    Set d2 = NovoDict()

    Dim k As Variant
    Dim kName As String, kLc As String

    For Each k In d.keys

        kName = CStr(k)
        kLc = LCase$(Trim$(kName))

        If kLc <> "input" Then
            If Not IsChaveInternaFileOutput(kLc) Then
                d2.Add k, d(k)
            End If
        End If

    Next k


    Dim objSemInput As String
    objSemInput = Dict_ToJsonObject(d2)

    If objSemInput = "{}" Then
        outExtraFragmentSemInput = ""
    Else
        outExtraFragmentSemInput = Mid$(objSemInput, 2, Len(objSemInput) - 2) ' remove { }
    End If
End Sub



