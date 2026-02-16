Attribute VB_Name = "M14_ConfigApiKey"
Option Explicit

' =============================================================================
' MÃ³dulo: M14_ConfigApiKey
' PropÃ³sito:
' - Resolver OPENAI_API_KEY com precedÃªncia segura (variÃ¡vel de ambiente antes de Config!B1).
' - Fornecer diagnÃ³stico sem exposiÃ§Ã£o de segredos para consumo de DEBUG e self-tests.
'
' AtualizaÃ§Ãµes:
' - 2026-02-16 | Codex | Resolver de API key com prioridade para ambiente
'   - Adiciona Config_ResolveOpenAIApiKey para uso transversal no motor.
'   - MantÃ©m fallback retrocompatÃ­vel para Config!B1 sem obrigar alteraÃ§Ãµes estruturais no Excel.
'   - ExpÃµe helper de self-test sem ler ambiente real.
'
' FunÃ§Ãµes e procedimentos:
' - Config_ResolveOpenAIApiKey(ByRef outApiKey, ByRef outSource, ByRef outAlert, ByRef outError) As Boolean
'   - Resolve key efetiva; nÃ£o escreve logs diretamente.
' - Config_SelfTest_ResolveOpenAIApiKey(ByVal envValue, ByVal configB1Value, ByRef outApiKey, ByRef outSource, ByRef outAlert, ByRef outError) As Boolean
'   - Variante determinÃ­stica para SelfTests.
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
            outAlert = "Config!B1 contÃ©m API key literal, mas OPENAI_API_KEY do ambiente foi priorizada. Recomenda-se remover a key literal da folha Config."
        End If

        ResolveOpenAIApiKeyCore = True
        Exit Function
    End If

    If Config_IsEnvDirective(cfgRaw) Then
        outError = "Config!B1 estÃ¡ configurada para usar Environ(""OPENAI_API_KEY""), mas a variÃ¡vel de ambiente OPENAI_API_KEY estÃ¡ vazia/ausente."
        ResolveOpenAIApiKeyCore = False
        Exit Function
    End If

    If Config_IsUsableLiteralKey(cfgRaw) Then
        outApiKey = cfgRaw
        outSource = "CONFIG_B1"
        outAlert = "OPENAI_API_KEY nÃ£o encontrada no ambiente; foi usado fallback em Config!B1. Recomenda-se migrar para variÃ¡vel de ambiente."
        ResolveOpenAIApiKeyCore = True
        Exit Function
    End If

    outError = "OPENAI_API_KEY ausente: variÃ¡vel de ambiente OPENAI_API_KEY vazia e Config!B1 sem valor vÃ¡lido."
    ResolveOpenAIApiKeyCore = False
End Function

Private Function Config_ReadApiKeyCell() As String
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)

    Config_ReadApiKeyCell = Trim$(CStr(ws.Range(CFG_CELL_API_KEY).Value))
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

    If InStr(1, s, "environ(""" & LCase$(ENV_OPENAI_API_KEY) & """)", vbTextCompare) > 0 Then
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
        Case "openai_api_key", "your_openai_api_key", "<openai_api_key>", "(environ(""openai_api_key""))"
            Exit Function
    End Select

    If InStr(1, s, "insira", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "placeholder", vbTextCompare) > 0 Then Exit Function

    Config_IsUsableLiteralKey = True
End Function
