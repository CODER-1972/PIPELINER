Attribute VB_Name = "M14_ConfigApiKey"
Option Explicit

' =============================================================================
' Módulo: M14_ConfigApiKey
' Propósito:
' - Resolver OPENAI_API_KEY com precedência segura (variável de ambiente antes de Config!B1).
' - Fornecer diagnóstico sem exposição de segredos para consumo de DEBUG e self-tests.
'
' Atualizações:
' - 2026-02-16 | Codex | Resolver de API key com prioridade para ambiente
'   - Adiciona Config_ResolveOpenAIApiKey para uso transversal no motor.
'   - Mantém fallback retrocompatível para Config!B1 sem obrigar alterações estruturais no Excel.
'   - Expõe helper de self-test sem ler ambiente real.
'
' Funções e procedimentos:
' - Config_ResolveOpenAIApiKey(ByRef outApiKey, ByRef outSource, ByRef outAlert, ByRef outError) As Boolean
'   - Resolve key efetiva; não escreve logs diretamente.
' - Config_SelfTest_ResolveOpenAIApiKey(ByVal envValue, ByVal configB1Value, ByRef outApiKey, ByRef outSource, ByRef outAlert, ByRef outError) As Boolean
'   - Variante determinística para SelfTests.
' =============================================================================

Private Const CFG_SHEET As String = "Config"
Private Const CFG_CELL_API_KEY As String = "B1"
Private Const ENV_OPENAI_API_KEY As String = "OPENAI_API_KEY"

Public Function Config_ResolveOpenAIApiKey( _
    ByRef outApiKey As String, _
    ByRef outSource As String, _
    ByRef outAlert As String, _
    ByRef outError As String _
) As Boolean
    Dim cfgValue As String
    Dim envValue As String

    cfgValue = Config_ReadApiKeyCell()
    envValue = Trim$(Environ$(ENV_OPENAI_API_KEY))

    Config_ResolveOpenAIApiKey = ResolveOpenAIApiKeyCore(envValue, cfgValue, outApiKey, outSource, outAlert, outError)
End Function

Public Function Config_SelfTest_ResolveOpenAIApiKey( _
    ByVal envValue As String, _
    ByVal configB1Value As String, _
    ByRef outApiKey As String, _
    ByRef outSource As String, _
    ByRef outAlert As String, _
    ByRef outError As String _
) As Boolean
    Config_SelfTest_ResolveOpenAIApiKey = ResolveOpenAIApiKeyCore(envValue, configB1Value, outApiKey, outSource, outAlert, outError)
End Function

Private Function ResolveOpenAIApiKeyCore( _
    ByVal envValue As String, _
    ByVal configB1Value As String, _
    ByRef outApiKey As String, _
    ByRef outSource As String, _
    ByRef outAlert As String, _
    ByRef outError As String _
) As Boolean
    Dim envKey As String
    Dim cfgRaw As String

    envKey = Trim$(CStr(envValue))
    cfgRaw = Trim$(CStr(configB1Value))

    outApiKey = ""
    outSource = ""
    outAlert = ""
    outError = ""

    If envKey <> "" Then
        outApiKey = envKey
        outSource = "ENV"

        If Config_IsUsableLiteralKey(cfgRaw) Then
            outAlert = "Config!B1 contém API key literal, mas OPENAI_API_KEY do ambiente foi priorizada. Recomenda-se remover a key literal da folha Config."
        End If

        ResolveOpenAIApiKeyCore = True
        Exit Function
    End If

    If Config_IsEnvDirective(cfgRaw) Then
        outError = "Config!B1 está configurada para usar Environ(\'OPENAI_API_KEY\'), mas a variável de ambiente OPENAI_API_KEY está vazia/ausente."
        ResolveOpenAIApiKeyCore = False
        Exit Function
    End If

    If Config_IsUsableLiteralKey(cfgRaw) Then
        outApiKey = cfgRaw
        outSource = "CONFIG_B1"
        outAlert = "OPENAI_API_KEY não encontrada no ambiente; foi usado fallback em Config!B1. Recomenda-se migrar para variável de ambiente."
        ResolveOpenAIApiKeyCore = True
        Exit Function
    End If

    outError = "OPENAI_API_KEY ausente: variável de ambiente OPENAI_API_KEY vazia e Config!B1 sem valor válido."
    ResolveOpenAIApiKeyCore = False
End Function

Private Function Config_ReadApiKeyCell() As String
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)

    Config_ReadApiKeyCell = Trim$(CStr(ws.Range(CFG_CELL_API_KEY).value))
    Exit Function
EH:
    Config_ReadApiKeyCell = ""
End Function

Private Function Config_IsEnvDirective(ByVal cfgValue As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(CStr(cfgValue)))

    If s = "" Then Exit Function

    s = Replace$(s, " ", "")
    s = Replace$(s, "'", "")

    If InStr(1, s, "environ(\"" & LCase$(ENV_OPENAI_API_KEY) & " \ ")", vbTextCompare) > 0 Then
        Config_IsEnvDirective = True
        Exit Function
    End If

    If s = "env:" & LCase$(ENV_OPENAI_API_KEY) Or s = "${" & LCase$(ENV_OPENAI_API_KEY) & "}" Then
        Config_IsEnvDirective = True
        Exit Function
    End If
End Function

Private Function Config_IsUsableLiteralKey(ByVal cfgValue As String) As Boolean
    Dim s As String
    s = Trim$(CStr(cfgValue))

    If s = "" Then Exit Function
    If Config_IsEnvDirective(s) Then Exit Function

    Select Case LCase$(s)
        Case "openai_api_key", "your_openai_api_key", "<openai_api_key>", "(environ(\'openai_api_key\'))"
            Exit Function
    End Select

    If InStr(1, s, "insira", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "placeholder", vbTextCompare) > 0 Then Exit Function

    Config_IsUsableLiteralKey = True
End Function

Public Sub Diagnostico_Encoding_BOM_M14()
    Dim p As String
    p = ThisWorkbook.path & "\..\src\vba\M14_ConfigApiKey.bas"
    p = " C:\Users\pnico\OneDrive - Universidade de Lisboa\Investiga  o - T CNICA IA\Git_PIPELINER\src\vba\M14_ConfigApiKey.bas"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object, bytes() As Byte
    Dim i As Long

    ' Ler 3 primeiros bytes
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 'bin rio
    stm.Open
    stm.LoadFromFile fso.GetAbsolutePathName(p)
    bytes = stm.Read(3)
    stm.Close

    Debug.Print "Ficheiro:", p
    Debug.Print "3 bytes iniciais:", Hex$(bytes(0)), Hex$(bytes(1)), Hex$(bytes(2))
    Debug.Print "BOM UTF-8? (EF BB BF) =", (bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF)
End Sub



