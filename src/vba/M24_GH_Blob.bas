Attribute VB_Name = "M24_GH_Blob"
Option Explicit

' =============================================================================
' MÃ³dulo: M24_GH_Blob
' PropÃ³sito:
' - Criar blobs no endpoint GitHub Git Database API (/git/blobs) a partir de ficheiros locais.
' - Aplicar validaÃ§Ãµes de tamanho e encoding (utf-8/base64) antes do envio.
' - Atualizar metadados do fileItem e registar eventos canÃ³nicos no DEBUG.
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | CriaÃ§Ã£o do mÃ³dulo de blobs GitHub
'   - Adiciona CreateBlob(cfg, fileItem) com leitura de bytes e classificaÃ§Ã£o texto/binÃ¡rio.
'   - Implementa limite GH_MAX_FILE_MB e logging GH_BLOB_OK/GH_BLOB_TOO_LARGE.
'   - Atualiza fileItem com BytesLen/BlobSha apÃ³s criaÃ§Ã£o bem-sucedida.
'
' FunÃ§Ãµes e procedimentos:
' - CreateBlob(cfg As Object, fileItem As Object) As String
'   - Cria blob no GitHub, devolve blob_sha e atualiza metadados do fileItem.
' =============================================================================

Private Const DEFAULT_GH_MAX_FILE_MB As Double = 20#

Public Function CreateBlob(ByVal cfg As Object, ByVal fileItem As Object) As String
    On Error GoTo EH

    Dim resolvedPath As String
    resolvedPath = GH_GetText(fileItem, "ResolvedPath", "")
    If Len(resolvedPath) = 0 Then resolvedPath = GH_GetText(fileItem, "FullPath", "")
    If Len(resolvedPath) = 0 Then resolvedPath = GH_GetText(fileItem, "Path", "")

    If Len(resolvedPath) = 0 Then
        Call GH_LogEvent(fileItem, "ERRO", "GH_BLOB_PATH_MISSING", "fileItem sem path resolvido para upload de blob.", "Preencha ResolvedPath/FullPath/Path antes de CreateBlob.")
        CreateBlob = ""
        Exit Function
    End If

    Dim fileBytes() As Byte
    fileBytes = GH_ReadAllBytes(resolvedPath)

    Dim bytesLen As Long
    bytesLen = GH_ByteLen(fileBytes)
    Call GH_SetValue(fileItem, "BytesLen", bytesLen)

    Dim maxMb As Double
    maxMb = GH_GetNumber(cfg, "GH_MAX_FILE_MB", DEFAULT_GH_MAX_FILE_MB)

    If maxMb > 0# Then
        Dim maxBytes As Double
        maxBytes = maxMb * 1024# * 1024#
        If CDbl(bytesLen) > maxBytes Then
            Call GH_LogEvent(fileItem, "ALERTA", "GH_BLOB_TOO_LARGE", _
                "Ficheiro excede GH_MAX_FILE_MB e nÃ£o serÃ¡ enviado.", _
                "path=" & resolvedPath & " | bytes=" & CStr(bytesLen) & " | limit_mb=" & Replace$(CStr(maxMb), ",", "."))
            CreateBlob = ""
            Exit Function
        End If
    End If

    Dim isText As Boolean
    isText = GH_DetermineTextMode(cfg, fileItem, resolvedPath)

    Dim payload As String
    payload = GH_BuildBlobPayload(fileBytes, isText)

    Dim apiUrl As String
    apiUrl = GH_GetText(cfg, "GH_BLOBS_URL", "")
    If Len(apiUrl) = 0 Then apiUrl = GH_GetText(cfg, "BlobsUrl", "")
    If Len(apiUrl) = 0 Then
        Call GH_LogEvent(fileItem, "ERRO", "GH_BLOB_URL_MISSING", "Config sem endpoint de blobs.", "Defina GH_BLOBS_URL em cfg.")
        CreateBlob = ""
        Exit Function
    End If

    Dim token As String
    token = GH_GetText(cfg, "GH_TOKEN", "")
    If Len(token) = 0 Then token = GH_GetText(cfg, "Token", "")

    Dim responseText As String
    Dim statusCode As Long
    statusCode = GH_PostBlob(apiUrl, token, payload, responseText)

    If statusCode < 200 Or statusCode >= 300 Then
        Call GH_LogEvent(fileItem, "ERRO", "GH_BLOB_HTTP_FAIL", "Falha ao criar blob no GitHub.", "status=" & CStr(statusCode) & " | body=" & Left$(responseText, 400))
        CreateBlob = ""
        Exit Function
    End If

    Dim blobSha As String
    blobSha = GH_ExtractJsonString(responseText, "sha")
    If Len(blobSha) = 0 Then
        Call GH_LogEvent(fileItem, "ERRO", "GH_BLOB_SHA_MISSING", "Resposta sem campo sha.", "body=" & Left$(responseText, 400))
        CreateBlob = ""
        Exit Function
    End If

    Call GH_SetValue(fileItem, "BlobSha", blobSha)
    CreateBlob = blobSha

    Call GH_LogEvent(fileItem, "INFO", "GH_BLOB_OK", "Blob criado com sucesso no GitHub.", "sha=" & blobSha & " | bytes=" & CStr(bytesLen) & " | text_mode=" & CStr(isText))
    Exit Function

EH:
    Call GH_LogEvent(fileItem, "ERRO", "GH_BLOB_EXCEPTION", "Erro inesperado em CreateBlob.", "Err " & CStr(Err.Number) & ": " & Err.Description)
    CreateBlob = ""
End Function

Private Function GH_ReadAllBytes(ByVal fullPath As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile fullPath
    GH_ReadAllBytes = st.Read
    st.Close
End Function

Private Function GH_ByteLen(ByRef b() As Byte) As Long
    On Error GoTo EH
    GH_ByteLen = (UBound(b) - LBound(b) + 1)
    Exit Function
EH:
    GH_ByteLen = 0
End Function

Private Function GH_DetermineTextMode(ByVal cfg As Object, ByVal fileItem As Object, ByVal fullPath As String) As Boolean
    Dim forceText As String
    forceText = UCase$(Trim$(GH_GetText(fileItem, "BlobForceText", GH_GetText(fileItem, "IsText", ""))))

    If forceText = "TRUE" Or forceText = "1" Or forceText = "TEXT" Or forceText = "UTF-8" Or forceText = "UTF8" Then
        GH_DetermineTextMode = True
        Exit Function
    End If

    If forceText = "FALSE" Or forceText = "0" Or forceText = "BINARY" Or forceText = "BIN" Then
        GH_DetermineTextMode = False
        Exit Function
    End If

    Dim ext As String
    ext = LCase$(GH_FileExt(fullPath))

    Select Case ext
        Case "txt", "md", "json", "csv", "tsv", "xml", "yaml", "yml", "ini", "cfg", "log", "vba", "bas", "cls", "frm", "html", "htm", "css", "js", "ts", "py", "java", "c", "cpp", "h", "hpp", "sql"
            GH_DetermineTextMode = True
        Case Else
            GH_DetermineTextMode = False
    End Select
End Function

Private Function GH_FileExt(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, ".")
    If p <= 0 Then
        GH_FileExt = ""
    Else
        GH_FileExt = Mid$(fullPath, p + 1)
    End If
End Function

Private Function GH_BuildBlobPayload(ByRef b() As Byte, ByVal isText As Boolean) As String
    Dim content As String
    Dim enc As String

    If isText Then
        content = GH_BytesToUtf8(b)
        enc = "utf-8"
    Else
        content = GH_Base64Encode(b)
        enc = "base64"
    End If

    GH_BuildBlobPayload = "{""content"":""" & GH_JsonEscape(content) & """,""encoding"":""" & enc & """}"
End Function

Private Function GH_BytesToUtf8(ByRef b() As Byte) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write b
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    GH_BytesToUtf8 = st.ReadText(-1)
    st.Close
End Function

Private Function GH_Base64Encode(ByRef b() As Byte) As String
    Dim dom As Object, node As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    If dom Is Nothing Then Set dom = CreateObject("MSXML2.DOMDocument.3.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = b
    GH_Base64Encode = Replace$(node.Text, vbLf, "")
    GH_Base64Encode = Replace$(GH_Base64Encode, vbCr, "")
End Function

Private Function GH_PostBlob(ByVal apiUrl As String, ByVal token As String, ByVal payload As String, ByRef outResponse As String) As Long
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", apiUrl, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/vnd.github+json"
    http.SetRequestHeader "X-GitHub-Api-Version", "2022-11-28"

    If Len(Trim$(token)) > 0 Then
        http.SetRequestHeader "Authorization", "Bearer " & token
    End If

    http.Send payload
    GH_PostBlob = CLng(http.Status)
    outResponse = CStr(http.ResponseText)
End Function

Private Function GH_ExtractJsonString(ByVal json As String, ByVal key As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """" & GH_RegexEscape(key) & """" & "\s*:\s*""([^""]*)"""

    If re.Test(json) Then
        Set m = re.Execute(json)(0)
        GH_ExtractJsonString = CStr(m.SubMatches(0))
    Else
        GH_ExtractJsonString = ""
    End If
End Function

Private Function GH_RegexEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace$(t, "\", "\\")
    t = Replace$(t, ".", "\.")
    t = Replace$(t, "+", "\+")
    t = Replace$(t, "*", "\*")
    t = Replace$(t, "?", "\?")
    t = Replace$(t, "^", "\^")
    t = Replace$(t, "$", "\$")
    t = Replace$(t, "(", "\(")
    t = Replace$(t, ")", "\)")
    t = Replace$(t, "[", "\[")
    t = Replace$(t, "]", "\]")
    t = Replace$(t, "{", "\{")
    t = Replace$(t, "}", "\}")
    t = Replace$(t, "|", "\|")
    GH_RegexEscape = t
End Function

Private Function GH_JsonEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace$(t, "\", "\\")
    t = Replace$(t, Chr$(34), "\" & Chr$(34))
    t = Replace$(t, vbCrLf, "\n")
    t = Replace$(t, vbCr, "\n")
    t = Replace$(t, vbLf, "\n")
    t = Replace$(t, vbTab, "\t")
    GH_JsonEscape = t
End Function

Private Sub GH_LogEvent(ByVal fileItem As Object, ByVal severidade As String, ByVal parametro As String, ByVal problema As String, ByVal sugestao As String)
    On Error GoTo EH
    Call Debug_Registar(0, GH_GetText(fileItem, "PromptId", ""), severidade, "", parametro, problema, sugestao)
    Exit Sub
EH:
End Sub

Private Function GH_GetText(ByVal obj As Object, ByVal key As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback

    If obj Is Nothing Then
        GH_GetText = defaultValue
        Exit Function
    End If

    If TypeName(obj) = "Dictionary" Or TypeName(obj) = "Scripting.Dictionary" Then
        If obj.Exists(key) Then
            GH_GetText = CStr(obj(key))
        Else
            GH_GetText = defaultValue
        End If
        Exit Function
    End If

    Dim v As Variant
    v = CallByName(obj, key, VbGet)
    If IsNull(v) Or IsEmpty(v) Then
        GH_GetText = defaultValue
    Else
        GH_GetText = CStr(v)
    End If
    Exit Function

Fallback:
    GH_GetText = defaultValue
End Function

Private Function GH_GetNumber(ByVal obj As Object, ByVal key As String, ByVal defaultValue As Double) As Double
    Dim txt As String
    txt = Trim$(GH_GetText(obj, key, ""))

    If Len(txt) = 0 Then
        GH_GetNumber = defaultValue
        Exit Function
    End If

    txt = Replace$(txt, ",", ".")
    If IsNumeric(txt) Then
        GH_GetNumber = CDbl(txt)
    Else
        GH_GetNumber = defaultValue
    End If
End Function

Private Sub GH_SetValue(ByVal obj As Object, ByVal key As String, ByVal v As Variant)
    On Error GoTo EH

    If obj Is Nothing Then Exit Sub

    If TypeName(obj) = "Dictionary" Or TypeName(obj) = "Scripting.Dictionary" Then
        obj(key) = v
        Exit Sub
    End If

    CallByName obj, key, VbLet, v
    Exit Sub
EH:
End Sub
