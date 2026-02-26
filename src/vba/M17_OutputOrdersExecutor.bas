Attribute VB_Name = "M17_OutputOrdersExecutor"
Option Explicit

' =============================================================================
' MÃ³dulo: M17_OutputOrdersExecutor
' PropÃ³sito:
' - Interpretar ordens EXECUTE emitidas no output do modelo apÃ³s sucesso HTTP.
' - Implementar whitelist inicial de comandos (LOAD_CSV) com validaÃ§Ãµes de seguranÃ§a.
' - Importar CSV para nova worksheet com logging completo em DEBUG e resumo para Seguimento.
'
' AtualizaÃ§Ãµes:
' - 2026-02-26 | Codex | Corrige dependÃªncias internas do SelfTest
'   - Adiciona helpers locais EnsureFolder e WriteTextUTF8 usados pela bateria T1..T9.
'   - Remove acoplamento implÃ­cito a helpers Private de outros mÃ³dulos.
' - 2026-02-23 | Codex | ImplementaÃ§Ã£o inicial de Output Orders (v1.3)
'   - Adiciona parser de linhas EXECUTE com tolerÃ¢ncia a variantes seguras.
'   - Implementa LOAD_CSV com resoluÃ§Ã£o automÃ¡tica de ficheiro, prÃ©-check e importaÃ§Ã£o.
'   - Inclui SelfTest_OutputOrders_RunAll com testes idempotentes T1..T9.
'
' FunÃ§Ãµes e procedimentos:
' - OutputOrders_TryExecute(...): executa ordens reconhecidas e devolve append para files_ops_log.
' - SelfTest_OutputOrders_RunAll(): bateria idempotente de testes do executor.
' =============================================================================

Private Const OUTPUT_ORDERS_MAX As Long = 3
Private Const OUTPUT_EXECUTE_TRACKER As String = "OUTPUT_EXECUTE"

Public Function OutputOrders_TryExecute( _
    ByVal passo As Long, _
    ByVal promptId As String, _
    ByVal responseId As String, _
    ByVal outputText As String, _
    ByVal outputFolder As String, _
    ByVal downloadedFiles As Variant _
) As String
    On Error GoTo EH

    OutputOrders_TryExecute = ""
    If Trim$(outputText) = "" Then Exit Function

    Dim directives As Collection
    Set directives = ParseExecuteDirectives(outputText)

    Call Debug_Registar(passo, promptId, "INFO", "", "OUTPUT_EXECUTE_FOUND", _
        "response_id=" & Trim$(responseId) & " | directives=" & CStr(directives.Count), "")

    If directives.Count = 0 Then Exit Function

    Dim limitN As Long
    limitN = directives.Count
    If limitN > OUTPUT_ORDERS_MAX Then limitN = OUTPUT_ORDERS_MAX

    Dim i As Long
    For i = 1 To limitN
        Dim item As Object
        Set item = directives(i)

        Dim cmd As String, fileName As String
        cmd = UCase$(Trim$(CStr(item("cmd"))))
        fileName = Trim$(CStr(item("argFileName")))

        Call Debug_Registar(passo, promptId, "INFO", "", "OUTPUT_EXECUTE_PARSED", _
            "index=" & CStr(i) & " | cmd=" & cmd & " | file=" & fileName, "")

        If cmd <> "LOAD_CSV" Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "OUTPUT_EXECUTE_UNKNOWN_CMD", _
                "Comando nÃ£o permitido: " & cmd, "Whitelist atual: LOAD_CSV.")
            GoTo NextDirective
        End If

        If Not ValidateCsvFileName(fileName) Then
            Call Debug_Registar(passo, promptId, "ERRO", "", "OUTPUT_EXECUTE_INVALID_FILENAME", _
                "Filename invÃ¡lido: " & fileName, "Use apenas basename.csv sem paths/sÃ­mbolos perigosos.")
            GoTo NextDirective
        End If

        Dim csvPath As String
        csvPath = ResolveCsvSource(fileName, outputFolder, downloadedFiles)
        If Trim$(csvPath) = "" Then
            Call Debug_Registar(passo, promptId, "ERRO", "", "OUTPUT_EXECUTE_FILE_NOT_FOUND", _
                "NÃ£o foi possÃ­vel resolver ficheiro CSV: " & fileName & " | outputFolder=" & outputFolder, _
                "Confirme download/geraÃ§Ã£o do ficheiro e nome exato no EXECUTE.")
            GoTo NextDirective
        End If

        Dim bomPass As Boolean, crlfPass As Boolean, colsHint As Long
        Call PrecheckCsv_BomAndCrLf(csvPath, bomPass, crlfPass, colsHint)

        Call Debug_Registar(passo, promptId, "INFO", "", "OUTPUT_EXECUTE_CSV_PRECHECK", _
            "file=" & fileName & " | BOM_" & IIf(bomPass, "PASS", "FAIL") & _
            " | CRLF_" & IIf(crlfPass, "PASS", "FAIL") & _
            " | colsHint=" & CStr(colsHint), "")

        If Not bomPass Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "OUTPUT_EXECUTE_CSV_BOM_FAIL", _
                "CSV sem BOM UTF-8 (EF BB BF): " & csvPath, "Reexporte com utf-8-sig para melhor compatibilidade com Excel.")
        End If
        If Not crlfPass Then
            Call Debug_Registar(passo, promptId, "ALERTA", "", "OUTPUT_EXECUTE_CSV_CRLF_IN_FIELDS", _
                "Detetadas quebras de linha reais em campo quoted: " & csvPath, "Substituir CR/LF por literal \\n no exportador.")
        End If

        Dim sheetName As String
        sheetName = DeriveSheetNameFromCsv(csvPath)

        Dim ws As Worksheet
        Set ws = CreateOutputSheetAfterPainel(sheetName)
        Call Debug_Registar(passo, promptId, "INFO", "", "OUTPUT_EXECUTE_SHEET_CREATED", _
            "sheetName=" & ws.Name, "")

        Dim rowsImported As Long, colsImported As Long, importErr As String
        If Not LoadCsvIntoSheet_QueryTable(csvPath, ws, rowsImported, colsImported, importErr) Then
            If Not LoadCsvIntoSheet_OpenTextFallback(csvPath, ws, rowsImported, colsImported, importErr) Then
                Call Debug_Registar(passo, promptId, "ERRO", "", "OUTPUT_EXECUTE_IMPORT_FAIL", _
                    "Falha ao importar CSV: " & csvPath & " | " & importErr, "Validar delimitador ';' e encoding UTF-8.")
                GoTo NextDirective
            End If
        End If

        Call Debug_Registar(passo, promptId, "INFO", "", "OUTPUT_EXECUTE_CSV_IMPORTED", _
            "sheetName=" & ws.Name & " | rows=" & CStr(rowsImported) & " | cols=" & CStr(colsImported), "")

        Dim verifyOk As Boolean, verifyEvidence As String
        verifyOk = VerifyImportedSheet(ws, colsHint, verifyEvidence, rowsImported, colsImported)
        Call Debug_Registar(passo, promptId, "INFO", "", "OUTPUT_EXECUTE_VERIFIED", _
            "result=" & IIf(verifyOk, "PASS", "FAIL") & " | " & verifyEvidence, "")

        If verifyOk Then
            Dim msg As String
            msg = "CREATED AND LOADED Excel Sheet " & ws.Name & " importing " & fileName & ", and verified."
            OutputOrders_TryExecute = OutputOrders_AppendLog(OutputOrders_TryExecute, msg)
        End If

NextDirective:
    Next i

    Exit Function
EH:
    Call Debug_Registar(passo, promptId, "ERRO", "", OUTPUT_EXECUTE_TRACKER & "_UNHANDLED", _
        "Err " & CStr(Err.Number) & ": " & Err.Description, "Reveja parser/importaÃ§Ã£o de Output Orders.")
End Function

Public Function ParseExecuteDirectives(ByVal outputText As String) As Collection
    Dim out As New Collection
    Dim lines() As String
    lines = Split(Replace(outputText, vbCrLf, vbLf), vbLf)

    Dim inCode As Boolean
    inCode = False

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim raw As String
        raw = CStr(lines(i))

        If Left$(Trim$(raw), 3) = "```" Then
            inCode = Not inCode
            GoTo NextLine
        End If
        If inCode Then GoTo NextLine

        Dim parsed As Object
        Set parsed = ParseExecuteLine(raw)
        If Not parsed Is Nothing Then out.Add parsed

NextLine:
    Next i

    Set ParseExecuteDirectives = out
End Function

Private Function ParseExecuteLine(ByVal rawLine As String) As Object
    Dim t As String
    t = Trim$(rawLine)
    If t = "" Then Exit Function

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "^<?\s*EXECUTE:\s*([A-Z_]+)\s*\((.+)\)\s*>?$"

    If Not re.Test(t) Then Exit Function

    Dim m As Object
    Set m = re.Execute(t)(0)

    Dim cmd As String
    Dim rawArg As String
    cmd = UCase$(Trim$(m.SubMatches(0)))
    rawArg = Trim$(CStr(m.SubMatches(1)))

    Dim fileName As String
    fileName = ParseExecuteFileArg(rawArg)

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cmd", cmd
    d.Add "argFileName", fileName
    d.Add "rawLine", rawLine
    d.Add "index", 0

    Set ParseExecuteLine = d
End Function

Private Function ParseExecuteFileArg(ByVal rawArg As String) As String
    Dim t As String
    t = Trim$(rawArg)

    If Left$(t, 1) = "[" And Right$(t, 1) = "]" Then
        ParseExecuteFileArg = Trim$(Mid$(t, 2, Len(t) - 2))
        Exit Function
    End If

    If Left$(t, 1) = """" And Right$(t, 1) = """" Then
        ParseExecuteFileArg = Mid$(t, 2, Len(t) - 2)
        Exit Function
    End If

    ParseExecuteFileArg = t
End Function

Public Function ValidateCsvFileName(ByVal fileName As String) As Boolean
    Dim n As String
    n = Trim$(fileName)

    If n = "" Then Exit Function
    If LCase$(Right$(n, 4)) <> ".csv" Then Exit Function

    If InStr(1, n, "..", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, n, ":", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, n, "\", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, n, "/", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, n, "%", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, n, "~", vbBinaryCompare) > 0 Then Exit Function

    ValidateCsvFileName = True
End Function

Public Function ResolveCsvSource(ByVal fileName As String, ByVal outputFolder As String, ByVal downloadedFiles As Variant) As String
    On Error GoTo EH

    Dim absFromDownloads As String
    absFromDownloads = ResolveCsvFromDownloaded(fileName, outputFolder, downloadedFiles)
    If absFromDownloads <> "" Then
        ResolveCsvSource = absFromDownloads
        Exit Function
    End If

    Dim p1 As String
    p1 = BuildPath(outputFolder, fileName)
    If FileExistsFast(p1) Then
        ResolveCsvSource = p1
        Exit Function
    End If

    ResolveCsvSource = FindFileRecursiveByBasename(outputFolder, fileName)
    Exit Function
EH:
    ResolveCsvSource = ""
End Function

Private Function ResolveCsvFromDownloaded(ByVal fileName As String, ByVal outputFolder As String, ByVal downloadedFiles As Variant) As String
    On Error GoTo EH

    Dim needle As String
    needle = LCase$(Trim$(fileName))

    If IsObject(downloadedFiles) Then
        Dim it As Variant
        For Each it In downloadedFiles
            Dim p As String
            p = CStr(it)
            If LCase$(GetBaseName(p)) = needle Then
                If FileExistsFast(p) Then
                    ResolveCsvFromDownloaded = p
                    Exit Function
                End If
                If FileExistsFast(BuildPath(outputFolder, p)) Then
                    ResolveCsvFromDownloaded = BuildPath(outputFolder, p)
                    Exit Function
                End If
            End If
        Next it
    Else
        Dim token As Variant
        Dim norm As String
        norm = Replace(Replace(CStr(downloadedFiles), ";", "|"), vbCrLf, "|")
        For Each token In Split(norm, "|")
            Dim t As String
            t = Trim$(CStr(token))
            If LCase$(Left$(t, 3)) = "dl:" Or LCase$(Left$(t, 4)) = "out:" Or LCase$(Left$(t, 3)) = "sv:" Then
                t = Trim$(Mid$(t, InStr(1, t, ":", vbBinaryCompare) + 1))
            End If
            If LCase$(GetBaseName(t)) = needle Then
                If FileExistsFast(t) Then
                    ResolveCsvFromDownloaded = t
                    Exit Function
                End If
                If FileExistsFast(BuildPath(outputFolder, t)) Then
                    ResolveCsvFromDownloaded = BuildPath(outputFolder, t)
                    Exit Function
                End If
            End If
        Next token
    End If

    Exit Function
EH:
    ResolveCsvFromDownloaded = ""
End Function

Public Sub PrecheckCsv_BomAndCrLf(ByVal csvPath As String, ByRef bomPass As Boolean, ByRef crlfPass As Boolean, ByRef colsHint As Long)
    bomPass = CsvHasUtf8Bom(csvPath)
    crlfPass = Not CsvHasQuotedNewLine(csvPath)
    colsHint = CsvHeaderColsHint(csvPath)
End Sub

Private Function CsvHasUtf8Bom(ByVal filePath As String) As Boolean
    On Error GoTo EH
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Binary Access Read As #ff

    If LOF(ff) >= 3 Then
        Dim b1 As Byte, b2 As Byte, b3 As Byte
        Get #ff, 1, b1
        Get #ff, 2, b2
        Get #ff, 3, b3
        CsvHasUtf8Bom = (b1 = &HEF And b2 = &HBB And b3 = &HBF)
    End If

    Close #ff
    Exit Function
EH:
    On Error Resume Next
    Close #ff
    CsvHasUtf8Bom = False
End Function

Private Function CsvHasQuotedNewLine(ByVal filePath As String) As Boolean
    On Error GoTo EH

    Dim txt As String
    txt = ReadTextFileUtf8(filePath)
    If txt = "" Then Exit Function

    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    inQuotes = False

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch = """" Then
            If inQuotes Then
                If i < Len(txt) And Mid$(txt, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                inQuotes = True
            End If
        ElseIf inQuotes And (ch = vbCr Or ch = vbLf) Then
            CsvHasQuotedNewLine = True
            Exit Function
        End If
    Next i

    Exit Function
EH:
    CsvHasQuotedNewLine = False
End Function

Private Function CsvHeaderColsHint(ByVal filePath As String) As Long
    On Error GoTo EH
    Dim txt As String
    txt = ReadTextFileUtf8(filePath)
    If txt = "" Then Exit Function

    Dim firstLine As String
    Dim p As Long
    p = InStr(1, txt, vbLf, vbBinaryCompare)
    If p > 0 Then
        firstLine = Left$(txt, p - 1)
    Else
        firstLine = txt
    End If
    firstLine = Replace(firstLine, vbCr, "")

    CsvHeaderColsHint = UBound(Split(firstLine, ";")) + 1
    Exit Function
EH:
    CsvHeaderColsHint = 0
End Function

Public Function DeriveSheetNameFromCsv(ByVal csvPath As String) As String
    Dim baseName As String
    baseName = "CSV_" & FileNameNoExt(GetBaseName(csvPath))

    Dim f As Integer
    Dim i As Long
    Dim lineText As String

    On Error GoTo Fallback
    f = FreeFile
    Open csvPath For Input As #f

    For i = 1 To 50
        If EOF(f) Then Exit For
        Line Input #f, lineText
        Dim c1 As String
        c1 = CsvFirstField(lineText)
        If InStr(1, c1, "/", vbBinaryCompare) > 1 Then
            Dim prefix As String
            prefix = Split(c1, "/")(0)
            If Trim$(prefix) <> "" Then
                baseName = prefix
                Exit For
            End If
        End If
    Next i

    Close #f
    DeriveSheetNameFromCsv = CreateUniqueWorksheetName(SanitizeWorksheetName(baseName))
    Exit Function

Fallback:
    On Error Resume Next
    Close #f
    DeriveSheetNameFromCsv = CreateUniqueWorksheetName(SanitizeWorksheetName(baseName))
End Function

Public Function CreateUniqueWorksheetName(ByVal baseName As String) As String
    Dim candidate As String
    candidate = baseName
    If Len(candidate) = 0 Then candidate = "CSV_IMPORT"
    If Len(candidate) > 31 Then candidate = Left$(candidate, 31)

    If Not SheetExists(candidate) Then
        CreateUniqueWorksheetName = candidate
        Exit Function
    End If

    Dim n As Long
    n = 1
    Do
        Dim suffix As String
        suffix = "_" & Format$(n, "00")
        candidate = Left$(baseName, 31 - Len(suffix)) & suffix
        If Not SheetExists(candidate) Then
            CreateUniqueWorksheetName = candidate
            Exit Function
        End If
        n = n + 1
    Loop
End Function

Private Function CreateOutputSheetAfterPainel(ByVal requestedName As String) As Worksheet
    Dim ws As Worksheet
    If SheetExists("PAINEL") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("PAINEL"))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
    ws.Name = requestedName
    Set CreateOutputSheetAfterPainel = ws
End Function

Public Function LoadCsvIntoSheet_QueryTable(ByVal csvPath As String, ByVal ws As Worksheet, ByRef outRows As Long, ByRef outCols As Long, ByRef outErr As String) As Boolean
    On Error GoTo EH

    ws.Cells.Clear

    Dim qt As QueryTable
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFilePlatform = 65001
        .AdjustColumnWidth = False
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    outRows = ws.UsedRange.Rows.Count
    outCols = ws.UsedRange.Columns.Count
    LoadCsvIntoSheet_QueryTable = True
    Exit Function
EH:
    outErr = "QueryTable: " & Err.Description
    On Error Resume Next
    If Not qt Is Nothing Then qt.Delete
    LoadCsvIntoSheet_QueryTable = False
End Function

Public Function LoadCsvIntoSheet_OpenTextFallback(ByVal csvPath As String, ByVal ws As Worksheet, ByRef outRows As Long, ByRef outCols As Long, ByRef outErr As String) As Boolean
    On Error GoTo EH

    Dim wbTmp As Workbook
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=csvPath, Origin:=65001, DataType:=xlDelimited, Semicolon:=True, Comma:=False, TextQualifier:=xlTextQualifierDoubleQuote
    Set wbTmp = ActiveWorkbook

    ws.Cells.Clear
    wbTmp.Worksheets(1).UsedRange.Copy Destination:=ws.Range("A1")

    outRows = ws.UsedRange.Rows.Count
    outCols = ws.UsedRange.Columns.Count

    wbTmp.Close SaveChanges:=False
    Application.DisplayAlerts = True

    LoadCsvIntoSheet_OpenTextFallback = True
    Exit Function
EH:
    outErr = "OpenText: " & Err.Description
    On Error Resume Next
    If Not wbTmp Is Nothing Then wbTmp.Close SaveChanges:=False
    Application.DisplayAlerts = True
    LoadCsvIntoSheet_OpenTextFallback = False
End Function

Public Function VerifyImportedSheet(ByVal ws As Worksheet, ByVal expectedCols As Long, ByRef evidence As String, ByRef rowsImported As Long, ByRef colsImported As Long) As Boolean
    rowsImported = ws.UsedRange.Rows.Count
    colsImported = ws.UsedRange.Columns.Count

    If rowsImported = 1 And colsImported = 1 And Trim$(CStr(ws.Range("A1").Value)) = "" Then
        evidence = "Sheet vazia apÃ³s importaÃ§Ã£o."
        Exit Function
    End If

    If colsImported < 2 Then
        evidence = "colsImported<2"
        Exit Function
    End If

    Dim filled As Long, c As Long
    filled = 0
    For c = 1 To colsImported
        If Trim$(CStr(ws.Cells(1, c).Value)) <> "" Then filled = filled + 1
    Next c
    If filled <= 1 Then
        evidence = "header com <=1 coluna preenchida"
        Exit Function
    End If

    If expectedCols > 1 And colsImported <> expectedCols Then
        evidence = "colsImported=" & CStr(colsImported) & " esperado=" & CStr(expectedCols)
        Exit Function
    End If

    evidence = "rows=" & CStr(rowsImported) & " cols=" & CStr(colsImported)
    VerifyImportedSheet = True
End Function

Public Sub SelfTest_OutputOrders_RunAll()
    On Error GoTo EH

    Dim base As String
    base = Environ$("TEMP") & "\PIPELINER_OutputOrders_SelfTest"
    EnsureFolder base

    Dim tPass As Long, tFail As Long
    tPass = 0: tFail = 0

    SelfTest_Assert "T1 Parser sem EXECUTE", ParseExecuteDirectives("texto normal").Count = 0, tPass, tFail
    SelfTest_Assert "T2 Parser com EXECUTE", ParseExecuteDirectives("<EXECUTE: LOAD_CSV([file.csv])>").Count = 1, tPass, tFail

    Dim r3 As String
    r3 = OutputOrders_TryExecute(1, "SELFTEST/T3", "resp", "EXECUTE: DELETE_FILE([x.csv])", base, "")
    SelfTest_Assert "T3 Cmd desconhecido ignora", Trim$(r3) = "", tPass, tFail

    Dim r4 As String
    r4 = OutputOrders_TryExecute(1, "SELFTEST/T4", "resp", "EXECUTE: LOAD_CSV([..\\evil.csv])", base, "")
    SelfTest_Assert "T4 Path traversal bloqueado", Trim$(r4) = "", tPass, tFail

    Dim noBomCsv As String
    noBomCsv = BuildPath(base, "nobom.csv")
    WriteAnsiText noBomCsv, "A;B" & vbCrLf & "1;2"
    Dim bPass As Boolean, cPass As Boolean, hint As Long
    PrecheckCsv_BomAndCrLf noBomCsv, bPass, cPass, hint
    SelfTest_Assert "T5 BOM fail detetado", (bPass = False), tPass, tFail

    Dim crlfCsv As String
    crlfCsv = BuildPath(base, "crlf.csv")
    WriteTextUTF8 crlfCsv, ChrW(&HFEFF) & "A;B" & vbCrLf & """" & "a" & vbCrLf & "b""" & ";2"
    PrecheckCsv_BomAndCrLf crlfCsv, bPass, cPass, hint
    SelfTest_Assert "T6 CRLF em campo quoted", (cPass = False), tPass, tFail

    Dim okCsv As String
    okCsv = BuildPath(base, "catalogo_ok.csv")
    WriteTextUTF8 okCsv, ChrW(&HFEFF) & "ID;Nome;Prompt" & vbCrLf & "AvalCap/01/Teste/A;N;P" & vbCrLf & ";;;;;;;;;;"

    Dim r7 As String
    r7 = OutputOrders_TryExecute(1, "SELFTEST/T7", "resp", "EXECUTE: LOAD_CSV([catalogo_ok.csv])", base, "DL:catalogo_ok.csv")
    SelfTest_Assert "T7 Import OK", InStr(1, r7, "CREATED AND LOADED Excel Sheet", vbTextCompare) > 0, tPass, tFail

    Dim r8 As String
    r8 = OutputOrders_TryExecute(1, "SELFTEST/T8", "resp", "EXECUTE: LOAD_CSV([catalogo_ok.csv])", base, "DL:catalogo_ok.csv")
    SelfTest_Assert "T8 IdempotÃªncia cria sufixo", InStr(1, r8, "_01", vbTextCompare) > 0 Or InStr(1, r8, "_02", vbTextCompare) > 0, tPass, tFail

    SelfTest_Assert "T9 Frase Seguimento exacta", InStr(1, r7, "CREATED AND LOADED Excel Sheet ", vbBinaryCompare) = 1 And Right$(r7, Len(", and verified.")) = ", and verified.", tPass, tFail

    Call Debug_Registar(0, "SELFTEST_OUTPUT_ORDERS", "INFO", "", "SELFTEST_SUMMARY", _
        "PASS=" & CStr(tPass) & " | FAIL=" & CStr(tFail), "")

    MsgBox "SelfTest_OutputOrders_RunAll concluÃ­do. PASS=" & tPass & " FAIL=" & tFail, IIf(tFail = 0, vbInformation, vbExclamation)
    Exit Sub
EH:
    Call Debug_Registar(0, "SELFTEST_OUTPUT_ORDERS", "ERRO", "", "SELFTEST_CRASH", "Err " & Err.Number & ": " & Err.Description, "")
    MsgBox "Falha no SelfTest_OutputOrders_RunAll: " & Err.Description, vbCritical
End Sub

Private Sub SelfTest_Assert(ByVal testName As String, ByVal cond As Boolean, ByRef passN As Long, ByRef failN As Long)
    If cond Then
        passN = passN + 1
        Call Debug_Registar(0, "SELFTEST_OUTPUT_ORDERS", "INFO", "", testName, "PASS", "")
    Else
        failN = failN + 1
        Call Debug_Registar(0, "SELFTEST_OUTPUT_ORDERS", "ERRO", "", testName, "FAIL", "Rever implementaÃ§Ã£o do Output Orders.")
    End If
End Sub

Private Function OutputOrders_AppendLog(ByVal currentText As String, ByVal extraText As String) As String
    If Trim$(currentText) = "" Then
        OutputOrders_AppendLog = extraText
    Else
        OutputOrders_AppendLog = currentText & " | " & extraText
    End If
End Function

Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    On Error GoTo EH
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    ReadTextFileUtf8 = st.ReadText(-1)
    st.Close
    Exit Function
EH:
    ReadTextFileUtf8 = ""
End Function

Private Function FindFileRecursiveByBasename(ByVal startFolder As String, ByVal baseName As String) As String
    On Error GoTo EH
    If Trim$(startFolder) = "" Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(startFolder) Then Exit Function

    Dim folder As Object
    Set folder = fso.GetFolder(startFolder)

    Dim f As Object
    For Each f In folder.Files
        If LCase$(CStr(f.Name)) = LCase$(baseName) Then
            FindFileRecursiveByBasename = CStr(f.Path)
            Exit Function
        End If
    Next f

    Dim subf As Object
    For Each subf In folder.SubFolders
        Dim hit As String
        hit = FindFileRecursiveByBasename(CStr(subf.Path), baseName)
        If hit <> "" Then
            FindFileRecursiveByBasename = hit
            Exit Function
        End If
    Next subf

    Exit Function
EH:
    FindFileRecursiveByBasename = ""
End Function

Private Function CsvFirstField(ByVal lineText As String) As String
    Dim p As Long
    p = InStr(1, lineText, ";", vbBinaryCompare)
    If p > 0 Then
        CsvFirstField = Trim$(Left$(lineText, p - 1))
    Else
        CsvFirstField = Trim$(lineText)
    End If

    If Left$(CsvFirstField, 1) = """" And Right$(CsvFirstField, 1) = """" And Len(CsvFirstField) >= 2 Then
        CsvFirstField = Mid$(CsvFirstField, 2, Len(CsvFirstField) - 2)
    End If
End Function

Private Function SanitizeWorksheetName(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If t = "" Then t = "CSV_IMPORT"

    t = Replace(t, ":", "")
    t = Replace(t, "\", "")
    t = Replace(t, "/", "")
    t = Replace(t, "?", "")
    t = Replace(t, "*", "")
    t = Replace(t, "[", "")
    t = Replace(t, "]", "")

    If Len(t) > 31 Then t = Left$(t, 31)
    If Len(t) = 0 Then t = "CSV_IMPORT"

    SanitizeWorksheetName = t
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function FileExistsFast(ByVal p As String) As Boolean
    On Error Resume Next
    FileExistsFast = (Len(Dir$(p)) > 0)
    On Error GoTo 0
End Function

Private Function BuildPath(ByVal p1 As String, ByVal p2 As String) As String
    If Right$(p1, 1) = "\" Then
        BuildPath = p1 & p2
    Else
        BuildPath = p1 & "\" & p2
    End If
End Function

Private Function GetBaseName(ByVal p As String) As String
    Dim t As String
    t = Replace(Replace(Trim$(p), "/", "\"), "\\", "\")
    If InStrRev(t, "\") > 0 Then
        GetBaseName = Mid$(t, InStrRev(t, "\") + 1)
    Else
        GetBaseName = t
    End If
End Function

Private Function FileNameNoExt(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 1 Then
        FileNameNoExt = Left$(fileName, p - 1)
    Else
        FileNameNoExt = fileName
    End If
End Function

Private Sub WriteAnsiText(ByVal filePath As String, ByVal txt As String)
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, txt
    Close #ff
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    On Error Resume Next
    If Len(Trim$(folderPath)) = 0 Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Sub WriteTextUTF8(ByVal filePath As String, ByVal txt As String)
    On Error GoTo EH
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile filePath, 2
    st.Close
    Exit Sub
EH:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Sub
