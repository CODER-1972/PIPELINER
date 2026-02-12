Attribute VB_Name = "M00_JsonUtil"
Option Explicit

' =============================================================================
' Módulo: M00_JsonUtil
' Propósito:
' - Fornecer utilitários de escape/normalização de strings JSON usados em payloads e logs.
' - Centralizar regras de escaping para reduzir erros de serialização em módulos de API/parser.
'
' Atualizações:
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - Json_EscapeString (Function): rotina pública do módulo.
' =============================================================================

' ============================================================
' JSON Utils — Escape estrito para strings em JSON
' - Escapa \ e "
' - TAB -> \t
' - CR/LF -> \n (CRLF tratado como 1 quebra)
' - Outros control chars (<32) -> \u00XX
' ============================================================

Public Function Json_EscapeString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    out = ""
    i = 1

    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)

        Select Case code
            Case 34 ' "
                out = out & "\"""          ' -> \"
            Case 92 ' \
                out = out & "\\"           ' -> \\
            Case 9  ' TAB
                out = out & "\t"
            Case 10 ' LF
                out = out & "\n"
            Case 13 ' CR  (se CRLF, consome o LF)
                If i < Len(s) Then
                    If AscW(Mid$(s, i + 1, 1)) = 10 Then
                        out = out & "\n"
                        i = i + 1
                    Else
                        out = out & "\n"
                    End If
                Else
                    out = out & "\n"
                End If
            Case 0 To 31
                ' Outros control chars ilegais em JSON string literal
                out = out & "\u" & Right$("000" & Hex$(code), 4)
            Case Else
                out = out & ch
        End Select

        i = i + 1
    Loop

    Json_EscapeString = out
End Function

