Attribute VB_Name = "M25_GH_TreeCommit"
Option Explicit

' =============================================================================
' MÃ³dulo: M25_GH_TreeCommit
' PropÃ³sito:
' - Isolar composiÃ§Ã£o de endpoints e payloads para commit de export GitHub.
' - Fornecer naming determinÃ­stico para ficheiros de exportaÃ§Ã£o DEBUG.
' - Concentrar regras de metadados de commit/mensagem em helpers reutilizÃ¡veis.
'
' AtualizaÃ§Ãµes:
' - 2026-03-04 | Codex | CriaÃ§Ã£o do mÃ³dulo de Ã¡rvore/commit GitHub
'   - Adiciona builder de URL para endpoint /repos/{owner}/{repo}/contents/{path}.
'   - Adiciona builder de payload JSON de update/create de ficheiro.
'   - Adiciona helper de mensagem padrÃ£o para commits de exportaÃ§Ã£o.
'
' FunÃ§Ãµes e procedimentos:
' - GH_TreeCommit_ContentsUrl(baseUrl, owner, repo, repoPath) As String
'   - Monta URL final do endpoint contents com encoding de path.
' - GH_TreeCommit_BuildContentsPayload(message, branch, base64Content, existingSha) As String
'   - Monta payload JSON para criar/atualizar ficheiro no GitHub.
' - GH_TreeCommit_DefaultMessage(pipelineIndex) As String
'   - Gera mensagem padrÃ£o com timestamp e pipeline.
' =============================================================================

Public Function GH_TreeCommit_ContentsUrl( _
    ByVal baseUrl As String, _
    ByVal owner As String, _
    ByVal repo As String, _
    ByVal repoPath As String) As String

    Dim normalizedBase As String
    normalizedBase = Trim$(baseUrl)
    If Right$(normalizedBase, 1) = "/" Then
        normalizedBase = Left$(normalizedBase, Len(normalizedBase) - 1)
    End If

    GH_TreeCommit_ContentsUrl = normalizedBase & "/repos/" & _
                                GH_TreeCommit_UrlEncode(owner) & "/" & _
                                GH_TreeCommit_UrlEncode(repo) & "/contents/" & _
                                GH_TreeCommit_EncodePath(repoPath)
End Function

Public Function GH_TreeCommit_BuildContentsPayload( _
    ByVal message As String, _
    ByVal branch As String, _
    ByVal base64Content As String, _
    Optional ByVal existingSha As String = "") As String

    Dim json As String
    json = "{" & _
           """message"":""" & GH_Blob_JsonEscape(message) & """," & _
           """branch"":""" & GH_Blob_JsonEscape(branch) & """," & _
           """content"":""" & GH_Blob_JsonEscape(base64Content) & """"

    If Len(Trim$(existingSha)) > 0 Then
        json = json & ",""sha"":""" & GH_Blob_JsonEscape(existingSha) & """"
    End If

    json = json & "}"
    GH_TreeCommit_BuildContentsPayload = json
End Function

Public Function GH_TreeCommit_DefaultMessage(ByVal pipelineIndex As Long) As String
    GH_TreeCommit_DefaultMessage = "PIPELINER debug export | pipeline=" & CStr(pipelineIndex) & _
                                   " | " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Private Function GH_TreeCommit_EncodePath(ByVal repoPath As String) As String
    Dim parts() As String
    parts = Split(repoPath, "/")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = GH_TreeCommit_UrlEncode(parts(i))
    Next i

    GH_TreeCommit_EncodePath = Join(parts, "/")
End Function

Private Function GH_TreeCommit_UrlEncode(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                out = out & ch
            Case Else
                out = out & "%" & Right$("0" & Hex$(code And &HFF), 2)
        End Select
    Next i

    GH_TreeCommit_UrlEncode = out
End Function
