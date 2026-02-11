Attribute VB_Name = "M03_Catalogo_ler_prompt_porID"
Option Explicit


Public Function Catalogo_ObterPromptPorID(ByVal promptId As String) As PromptDefinicao
    Dim p As PromptDefinicao
    Dim lookupId As String
    lookupId = Trim$(promptId)

    p.nomeFolha = ExtrairNomeFolhaDoID(lookupId)

    If Trim$(p.nomeFolha) = "" Then
        Catalogo_ObterPromptPorID = p
        Exit Function
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(p.nomeFolha)
    On Error GoTo 0

    If ws Is Nothing Then
        Catalogo_ObterPromptPorID = p
        Exit Function
    End If

    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultimaLinha < 2 Then
        Catalogo_ObterPromptPorID = p
        Exit Function
    End If

    Dim linha As Long
    For linha = 2 To ultimaLinha
        If Trim$(CStr(ws.Cells(linha, 1).value)) = lookupId Then
            p.Id = lookupId
            p.NomeCurto = CStr(ws.Cells(linha, 2).value)
            p.NomeDescritivo = CStr(ws.Cells(linha, 3).value)
            p.textoPrompt = CStr(ws.Cells(linha, 4).value)

            p.modelo = CStr(ws.Cells(linha, 5).value)
            p.modos = Trim$(CStr(ws.Cells(linha, 6).value))
            If p.modos = "" Then p.modos = "Nenhum"

            Dim storageTxt As String
            storageTxt = Trim$(CStr(ws.Cells(linha, 7).value))
            If storageTxt = "" Then
                p.storage = True
            Else
                p.storage = TextoParaBooleano(storageTxt)
            End If

            p.ConfigExtra = CStr(ws.Cells(linha, 8).value)

            p.Comentarios = CStr(ws.Cells(linha, 9).value)
            p.NotasDev = CStr(ws.Cells(linha, 10).value)
            p.HistoricoVersoes = CStr(ws.Cells(linha, 11).value)
            Exit For
        End If
    Next linha

    Catalogo_ObterPromptPorID = p
End Function


Private Function ExtrairNomeFolhaDoID(ByVal promptId As String) As String
    Dim p As Long
    p = InStr(1, promptId, "/")
    If p = 0 Then
        ExtrairNomeFolhaDoID = promptId
    Else
        ExtrairNomeFolhaDoID = Left$(promptId, p - 1)
    End If
End Function


Private Function TextoParaBooleano(ByVal valor As String) As Boolean
    Dim v As String
    v = UCase$(Trim$(valor))
    TextoParaBooleano = (v = "TRUE" Or v = "VERDADEIRO" Or v = "1" Or v = "SIM")
End Function
