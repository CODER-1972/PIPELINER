Attribute VB_Name = "M24_GH_Blob"
Option Explicit

' =============================================================================
' MÃ³dulo: M24_GH_Blob
' PropÃ³sito:
' - Fornecer utilitÃ¡rios de conteÃºdo (blob) para exportaÃ§Ã£o GitHub.
' - Codificar texto UTF-8 para Base64, compatÃ­vel com endpoint /contents.
' - Escapar strings JSON de forma consistente entre mÃ³dulos.
'
' AtualizaÃ§Ãµes:
' - 2026-03-05 | Codex | Hardening do escaping JSON
'   - Substitui escape de aspas por construÃ§Ã£o explÃ­cita com Chr$(34) para reduzir ambiguidades no VBE.
' - 2026-03-04 | Codex | CriaÃ§Ã£o do mÃ³dulo de blobs GitHub
'   - Adiciona encoding UTF-8 + Base64 via ADODB.Stream/MSXML bin.base64.
'   - Adiciona helper de escaping JSON para payloads robustos.
'
' FunÃ§Ãµes e procedimentos:
' - GH_Blob_Base64FromText(text) As String
'   - Converte texto Unicode para UTF-8 e codifica em Base64.
' - GH_Blob_JsonEscape(value) As String
'   - Escapa aspas, barras e control chars para uso seguro em JSON.
' =============================================================================

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
    s = Replace$(s, Chr$(34), "\" & Chr$(34))
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
