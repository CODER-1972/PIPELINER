Attribute VB_Name = "M24_GH_Blob"
Option Explicit

' =============================================================================
' Modulo: M24_GH_Blob
' Proposito:
' - Encapsular criacao de blobs GitHub e regras de encoding/tamanho.
' - Aplicar limite GH_MAX_FILE_MB por artefacto antes de enviar a API.
' - Fornecer helpers de serializacao JSON e Base64 reutilizaveis.
'
' Atualizacoes:
' - 2026-03-04 | Codex | Refactor de blobs para modulo dedicado
'   - Move POST /git/blobs para funcao unica GH_Blob_Create.
'   - Implementa escolha de encoding por tipo de conteudo (utf-8/base64).
'   - Adiciona validacao de tamanho com logging canonico GH_BLOB_TOO_LARGE.
'
' Funcoes e procedimentos:
' - GH_Blob_Create(cfg, filePath, content, pipelineNome, blobSha, errReason) As Boolean
'   - Cria blob no repo alvo e devolve SHA; aplica limite de tamanho e encoding.
' - GH_Blob_JsonEscape(value As String) As String
'   - Escapa string para uso seguro em JSON.
' - GH_Blob_Base64FromText(text As String) As String
'   - Converte texto em bytes UTF-8 e codifica em Base64.
' =============================================================================

Public Function GH_Blob_Create( _
    ByVal cfg As Object, _
    ByVal filePath As String, _
    ByVal content As String, _
    ByVal pipelineNome As String, _
    ByRef blobSha As String, _
    ByRef errReason As String) As Boolean

    blobSha = ""
    errReason = ""

    Dim maxFileMb As Long
    maxFileMb = GH_Config_GetLong(cfg, "max_file_mb", 50)

    Dim bytesCount As Long
    bytesCount = GH_Blob_Utf8ByteLength(content)

    If bytesCount > maxFileMb * 1024 * 1024 Then
        errReason = "Ficheiro excede GH_MAX_FILE_MB"
        Call GH_LogWarn(0, pipelineNome, GH_EVT_BLOB_TOO_LARGE, filePath, "bytes=" & CStr(bytesCount) & "; max_mb=" & CStr(maxFileMb))
        Exit Function
    End If

    Dim isBinary As Boolean
    isBinary = GH_Blob_IsLikelyBinary(filePath, content)

    Dim payload As String
    payload = GH_Blob_BuildPayload(cfg, content, isBinary)

    Dim statusCode As Long
    Dim responseText As String
    Dim httpErr As String
    Dim url As String

    url = GH_TreeCommit_GitBlobUrl(cfg)
    If Not GH_HTTP_SendJson("POST", url, cfg, payload, statusCode, responseText, httpErr, pipelineNome) Then
        errReason = "Erro em /git/blobs (status=" & CStr(statusCode) & ")"
        If httpErr <> "" Then errReason = errReason & " " & httpErr
        Exit Function
    End If

    blobSha = GH_TreeCommit_JsonPick(responseText, "sha")
    If blobSha = "" Then
        errReason = "Resposta /git/blobs sem sha"
        Exit Function
    End If

    If GH_Config_GetBoolean(cfg, "log_blob_sha", True) Then
        Call GH_LogInfo(0, pipelineNome, GH_EVT_BLOB_OK, filePath, "sha=" & Left$(blobSha, 10))
    End If

    GH_Blob_Create = True
End Function

Private Function GH_Blob_BuildPayload(ByVal cfg As Object, ByVal content As String, ByVal isBinary As Boolean) As String
    Dim encodingName As String
    Dim payloadContent As String

    If isBinary Then
        encodingName = GH_Config_GetString(cfg, "binary_mode", "base64")
        payloadContent = GH_Blob_Base64FromText(content)
    Else
        encodingName = GH_Config_GetString(cfg, "encoding_text", "utf-8")
        payloadContent = content
    End If

    GH_Blob_BuildPayload = "{""content"":""" & GH_Blob_JsonEscape(payloadContent) & """,""encoding"":""" & GH_Blob_JsonEscape(encodingName) & """}"
End Function

Public Function GH_Blob_Base64FromText(ByVal text As String) As String
    On Error GoTo EH

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.Position = 0
    stm.Type = 1

    Dim bytes() As Byte
    bytes = stm.Read
    stm.Close

    GH_Blob_Base64FromText = GH_Blob_Base64FromBytes(bytes)
    Exit Function
EH:
    GH_Blob_Base64FromText = ""
End Function

Public Function GH_Blob_JsonEscape(ByVal value As String) As String
    Dim s As String
    s = value
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", Chr$(92) & """")
    s = Replace$(s, vbCrLf, "\n")
    s = Replace$(s, vbCr, "\n")
    s = Replace$(s, vbLf, "\n")
    s = Replace$(s, vbTab, "\t")
    GH_Blob_JsonEscape = s
End Function

Private Function GH_Blob_Base64FromBytes(ByRef bytes() As Byte) As String
    On Error GoTo EH

    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument")

    Dim node As Object
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes

    GH_Blob_Base64FromBytes = Replace$(node.Text, vbLf, "")
    GH_Blob_Base64FromBytes = Replace$(GH_Blob_Base64FromBytes, vbCr, "")
    Exit Function
EH:
    GH_Blob_Base64FromBytes = ""
End Function

Private Function GH_Blob_Utf8ByteLength(ByVal text As String) As Long
    On Error GoTo Fallback
    Dim b64 As String
    b64 = GH_Blob_Base64FromText(text)
    If b64 = "" Then
        GH_Blob_Utf8ByteLength = LenB(text)
    Else
        GH_Blob_Utf8ByteLength = ((Len(b64) * 3) \ 4)
    End If
    Exit Function
Fallback:
    GH_Blob_Utf8ByteLength = LenB(text)
End Function

Private Function GH_Blob_IsLikelyBinary(ByVal filePath As String, ByVal content As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "png", "jpg", "jpeg", "gif", "webp", "pdf", "zip", "doc", "docx", "xls", "xlsx", "xlsm", "ppt", "pptx"
            GH_Blob_IsLikelyBinary = True
            Exit Function
    End Select

    GH_Blob_IsLikelyBinary = (InStr(1, content, Chr$(0), vbBinaryCompare) > 0)
End Function
