Attribute VB_Name = "M16_ErrorMessageFormatter"
Option Explicit

' =============================================================================
' Modulo: M16_ErrorMessageFormatter
' Proposito:
' - Padronizar mensagens de erro/alerta para DEBUG com contexto minimo acionavel.
' - Fornecer helpers para gerar mensagens curtas, consistentes e rastreaveis.
' - Nao faz I/O de rede; apenas formata texto para logging.
'
' Atualizacoes:
' - 2026-02-23 | Codex | Criacao de utilitarios de formatacao de diagnosticos
'   - Adiciona `Diag_Format` para montar mensagens com campos estaveis.
'   - Adiciona `Diag_WithRetryHint` para mensagens com proxima acao sugerida.
'   - Adiciona `Diag_ErrorFingerprint` para identificar classe de erro sem expor dados sensiveis.
'
' Funcoes e procedimentos:
' - Diag_Format(scopeTag, problem, impact, nextAction, Optional details) As String
'   - Produz mensagem padrao com labels fixos (PROBLEMA/IMPACTO/ACAO/DETALHE).
' - Diag_WithRetryHint(scopeTag, problem, retryAction, Optional details) As String
'   - Atalho para alertas recuperaveis com orientacao de retry.
' - Diag_ErrorFingerprint(errNumber, errDescription, Optional httpStatus) As String
'   - Produz fingerprint estavel para agrupar incidentes no DEBUG.
' =============================================================================

Public Function Diag_Format( _
    ByVal scopeTag As String, _
    ByVal problem As String, _
    ByVal impact As String, _
    ByVal nextAction As String, _
    Optional ByVal details As String = "" _
) As String
    Dim msg As String

    msg = "[" & SanitizeOneLine(scopeTag) & "] " & _
          "PROBLEMA=" & SanitizeOneLine(problem) & " | " & _
          "IMPACTO=" & SanitizeOneLine(impact) & " | " & _
          "ACAO=" & SanitizeOneLine(nextAction)

    If Len(Trim$(details)) > 0 Then
        msg = msg & " | DETALHE=" & SanitizeOneLine(details)
    End If

    Diag_Format = msg
End Function

Public Function Diag_WithRetryHint( _
    ByVal scopeTag As String, _
    ByVal problem As String, _
    ByVal retryAction As String, _
    Optional ByVal details As String = "" _
) As String
    Diag_WithRetryHint = Diag_Format( _
        scopeTag:=scopeTag, _
        problem:=problem, _
        impact:="Execucao interrompida temporariamente", _
        nextAction:="Repetir teste apos: " & retryAction, _
        details:=details _
    )
End Function

Public Function Diag_ErrorFingerprint( _
    ByVal errNumber As Long, _
    ByVal errDescription As String, _
    Optional ByVal httpStatus As Long = 0 _
) As String
    Dim normalizedDesc As String

    normalizedDesc = UCase$(Trim$(errDescription))
    normalizedDesc = Replace$(normalizedDesc, vbCr, " ")
    normalizedDesc = Replace$(normalizedDesc, vbLf, " ")
    normalizedDesc = Replace$(normalizedDesc, "  ", " ")

    If Len(normalizedDesc) > 64 Then
        normalizedDesc = Left$(normalizedDesc, 64)
    End If

    Diag_ErrorFingerprint = "ERR#" & CStr(errNumber) & _
                            "|HTTP=" & CStr(httpStatus) & _
                            "|SIG=" & Replace$(normalizedDesc, "|", "/")
End Function

Private Function SanitizeOneLine(ByVal s As String) As String
    Dim t As String

    t = Trim$(s)
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")
    t = Replace$(t, vbTab, " ")
    t = Replace$(t, "|", "/")

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace$(t, "  ", " ")
    Loop

    SanitizeOneLine = t
End Function
