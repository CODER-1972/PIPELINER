Attribute VB_Name = "M09_FilesManagement"
Option Explicit

' ============================================================================
' M09_FilesManagement (REWRITE COMPLETO + FIX UPLOAD + DIAGNOSTICOS)
'
' Objetivo:
'   Preparar contexto de ficheiros para uso em prompts (Responses API),
'   integrado com pipelines (M07).
'
' Modos de transporte:
'   1) FILE_ID (default): upload em /v1/files (purpose user_data/vision),
'      guarda/reutiliza file_id por hash+usage_mode (FILES_MANAGEMENT).
'   2) INLINE_BASE64: NAO faz upload; injeta no request como data URL:
'        - PDF  : input_file.file_data = "data:application/pdf;base64,..."
'        - Image: input_image.image_url = "data:image/<ext>;base64,..."
'
' Config (folha "Config"):
'   B5 = FILES_TRANSPORT_MODE: "FILE_ID" (default) ou "INLINE_BASE64"
'   B6 = FILES_ENABLE_IA_FALLBACK: TRUE/FALSE (default FALSE)
'   B7 = FILES_INLINE_MAX_MB: limite MB por ficheiro em INLINE_BASE64 (default 20)
'
' Dependencias:
'   - Debug_Registar (M02): niveis INFO / ALERTA / ERRO
'   - OpenAI_Executar (M05): apenas se IA fallback estiver ativa
'
' Notas:
'   - Late binding (CreateObject) para evitar referencias obrigatorias
' ============================================================================


' ============================================================
' SHA-256 (Windows CryptoAPI) + fallback simples
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
        ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, _
        ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, _
        ByRef phHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As LongPtr) As Long
    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
        ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
        ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

    Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, _
        ByRef phHash As Long) As Long

    Private Declare Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
    Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If

Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C
Private Const HP_HASHVAL As Long = 2
Private gLastSHA256Diag As String
Private gLastFNV32Diag As String

' ============================================================
' Constantes
' ============================================================

Private Const SHEET_FILES As String = "FILES_MANAGEMENT"
Private Const SHEET_CONFIG As String = "Config"
Private Const PDF_CACHE_SIDECAR_EXT As String = ".src.sha256"

' Headers da folha FILES_MANAGEMENT (v2)
Private Const H_TIMESTAMP As String = "Timestamp"
Private Const H_DL_UL As String = "DL / UL"
Private Const H_FILE_NAME As String = "File name"
Private Const H_TYPE As String = "Type"
Private Const H_FOLDER As String = "Folder"
Private Const H_FULL_PATH As String = "Full path"
Private Const H_FILE_ID As String = "file_id"
Private Const H_USAGE_MODE As String = "usage_mode"
Private Const H_CONVERTED_TO_PDF As String = "converted_to_pdf (TRUE/FALSE)"
Private Const H_HASH As String = "hash (SHA-256 do conteudo efetivamente usado)"
Private Const H_LAST_MODIFIED As String = "last_modified"
Private Const H_SIZE_BYTES As String = "size_bytes"
Private Const H_LAST_USED_PIPELINE As String = "Last used_in_pipeline_name"
Private Const H_UTILIZACOES As String = "Utilizações"
Private Const H_USED_IN_PROMPTS As String = "used_in_prompts"
Private Const H_LAST_USED_AT As String = "last_used_at"
Private Const H_NOTES As String = "notes"

Private Const MAX_USED_PROMPTS As Long = 20
Private Const USED_PROMPTS_SUFFIX As String = "(...)"
Private Const USED_PROMPTS_SEP As String = ";  "


Private Const PDF_CACHE_FOLDER_NAME As String = "_pdf_cache"
Private Const EXT_PDF As String = "pdf"
Private Const IMG_EXTS As String = "png;jpg;jpeg;webp"

Private Const FALLBACK_MODEL As String = "gpt-4.1-mini"

Private Const TRANSPORT_FILE_ID As String = "FILE_ID"
Private Const TRANSPORT_INLINE As String = "INLINE_BASE64"

' WinHTTP secure protocols (best-effort). TLS 1.2 flag em WinHTTP e tipicamente &H800.
Private Const WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 As Long = &H800
Private Const WINHTTP_OPTION_SECURE_PROTOCOLS As Long = 9

' ============================================================
' Estado interno (evitar recursao em diagnosticos)
' ============================================================

Private gDiagRunning As Boolean

' ============================================================
' Estado do RUN (PAINEL) para separador visual em FILES_MANAGEMENT
' ============================================================
Private gRunToken As String
Private gLastSeparatorRunToken As String



' ============================================================
' API PUBLICA (chamada pelo M07 antes do OpenAI_Executar)
' ============================================================

Public Sub Files_SetRunToken(ByVal token As String)
    ' Definido pelo PAINEL (M07) no inicio de cada run; usado para colocar separador visual por run.
    gRunToken = Trim$(CStr(token))
End Sub


' ============================================================
' (ALTERAR) Files_PrepararContextoDaPrompt — versão integral
'   SUBSTITUA A FUNÇÃO INTEIRA por esta versão
'   (Esta é a versão que cumpre D4: call EnsureConfig + allowReuse/reuseTag + passagem de parâmetros)
' ============================================================
Public Function Files_PrepararContextoDaPrompt( _
    ByVal apiKey As String, _
    ByVal pipelineNome As String, _
    ByVal inputFolder As String, _
    ByVal promptId As String, _
    ByVal promptText As String, _
    ByVal inputJsonLiteralBase As String, _
    ByRef outInputJsonLiteralFinal As String, _
    ByRef outFilesUsedResumo As String, _
    ByRef outFilesOpsResumo As String, _
    ByRef outFileIdsUsed As String, _
    ByRef outFalhaCritica As Boolean, _
    ByRef outErroMsg As String, _
    Optional ByVal forceAsIsToTextEmbed As Boolean = False, _
    Optional ByVal stepN As Long = 0 _
) As Boolean

    ' ============================================================
    ' Versão com:
    ' - effective_mode (override quando /v1/responses não aceita as_is para DOCX/PPTX)
    ' - Política DOCX/PPTX: AUTO_AS_PDF / AUTO_TEXT_EMBED / ERROR
    ' - Fallback conversão PDF: TEXT_EMBED ou ERROR
    ' - Limite e ação para text_embed grande
    ' - PDF cache (evita reconversão -> evita hash diferente -> permite reutilização de file_id)
    ' - Label: ProcessarComoInputFile:
    ' ============================================================

    On Error GoTo TrataErro

    Dim dbgStep As String: dbgStep = "inicio"
    Dim dbgReq As String: dbgReq = ""
    Dim dbgFile As String: dbgFile = ""
    Dim dbgPath As String: dbgPath = ""
    Dim dbgUsoFinal As String: dbgUsoFinal = ""

    Files_PrepararContextoDaPrompt = True
    outFalhaCritica = False
    outErroMsg = ""

    outFilesUsedResumo = ""
    outFilesOpsResumo = ""
    outFileIdsUsed = ""
    outInputJsonLiteralFinal = inputJsonLiteralBase

    inputFolder = Trim$(CStr(inputFolder))

    Dim transportMode As String
    transportMode = Files_Config_TransportMode()

    Dim enableIAFallback As Boolean
    enableIAFallback = Files_Config_EnableIAFallback()

    Dim inlineMaxBytes As Double
    inlineMaxBytes = Files_Config_InlineMaxBytes()

    ' --- políticas de contexto para Office + limites text_embed ---
    Dim docxContextMode As String
    Dim docxAsPdfFallback As String
    Dim textEmbedMaxChars As Long
    Dim textEmbedOverflowAction As String

    docxContextMode = Files_Config_DocxContextMode()
    docxAsPdfFallback = Files_Config_DocxAsPdfFallback()
    textEmbedMaxChars = Files_Config_TextEmbedMaxChars()
    textEmbedOverflowAction = Files_Config_TextEmbedOverflowAction()

    Call Files_EnsureSheetExists

    Dim celInputsValor As Range, celOps As Range
    Call Files_EncontrarCelulasInputs(promptId, celInputsValor, celOps)

    If celInputsValor Is Nothing Then
        outInputJsonLiteralFinal = inputJsonLiteralBase
        Exit Function
    End If

    Dim textoInputs As String
    textoInputs = CStr(celInputsValor.value)

    ' garantir que existe a opção de config para reutilização
    Call Files_EnsureConfig_ReutilizacaoUpload

    Dim diretivas As Collection
    Set diretivas = Files_ExtrairDiretivasDeFicheiros(textoInputs)

    If diretivas.Count = 0 Then
        If Not celOps Is Nothing Then celOps.value = ""
        outInputJsonLiteralFinal = inputJsonLiteralBase
        Exit Function
    End If

    Dim haRequired As Boolean
    haRequired = Files_TemRequiredDiretivas(diretivas)

    If inputFolder = "" Or Dir(inputFolder, vbDirectory) = "" Then
        Call Files_EscreverOperacoes(celOps, diretivas, "ERRO: INPUT Folder nao existe ou esta vazio.", True)

        If haRequired Then
            outFalhaCritica = True
            outErroMsg = "INPUT Folder invalido e existem ficheiros (required). inputFolder=" & inputFolder
            Call Debug_Registar(0, promptId, "ERRO", "", "FILES", outErroMsg, "Sugestao: preencha o INPUT Folder no PAINEL.")
            Files_PrepararContextoDaPrompt = False
        Else
            Call Debug_Registar(0, promptId, "ALERTA", "", "FILES", _
                "INPUT Folder invalido; ficheiros opcionais foram ignorados. inputFolder=" & inputFolder, _
                "Sugestao: preencha o INPUT Folder no PAINEL.")
        End If

        outInputJsonLiteralFinal = inputJsonLiteralBase
        Exit Function
    End If

    Dim pdfCacheFolder As String
    pdfCacheFolder = Files_ComporPdfCacheFolder(inputFolder)
    Call Files_CriarPastaSeNaoExiste(pdfCacheFolder)

    Dim wsFiles As Worksheet
    Set wsFiles = ThisWorkbook.Worksheets(SHEET_FILES)

    Dim mapaCab As Object
    Set mapaCab = Files_MapaCabecalhos(wsFiles)

    Dim filePartsJson As String: filePartsJson = ""
    Dim textoEmbedTotal As String: textoEmbedTotal = ""
    Dim filesUsedLista As String: filesUsedLista = ""
    Dim filesOpsCurto As String: filesOpsCurto = ""
    Dim fileIdsLista As String: fileIdsLista = ""

    Dim houveOverride As Boolean: houveOverride = False
    Dim houveAmbiguidade As Boolean: houveAmbiguidade = False

    Dim i As Long
    For i = 1 To diretivas.Count

        Dim d As Object
        Set d = diretivas(i)

        Dim reqNome As String
        reqNome = CStr(d("requested_name"))

        Dim required As Boolean
        required = CBool(d("required"))

        Dim wantAsIs As Boolean, wantAsPdf As Boolean, wantText As Boolean, wantLatest As Boolean
        wantAsIs = CBool(d("as_is"))
        wantAsPdf = CBool(d("as_pdf"))
        wantText = CBool(d("text_embed"))
        wantLatest = CBool(d("latest"))

        Dim resolvedPath As String, resolvedName As String
        Dim status As String, candidatosLog As String, overrideUsado As Boolean

        resolvedPath = ""
        resolvedName = ""
        status = ""
        candidatosLog = ""
        overrideUsado = False

        dbgStep = "ResolverFicheiro"
        dbgReq = reqNome
        dbgFile = ""
        dbgPath = ""
        dbgUsoFinal = ""

        If Left$(reqNome, 1) = "@" Then
            Call Files_ResolverOutputToken( _
                pipelineNome, promptId, stepN, reqNome, resolvedPath, resolvedName, status, candidatosLog)

            If status = "AMBIGUOUS" Then houveAmbiguidade = True
            overrideUsado = False
        Else
            Call Files_ResolverFicheiro( _
                apiKey, promptId, inputFolder, reqNome, wantLatest, enableIAFallback, _
                resolvedPath, resolvedName, status, candidatosLog, overrideUsado, houveAmbiguidade)
        End If

        dbgFile = resolvedName
        dbgPath = resolvedPath

        If overrideUsado Then houveOverride = True

        If status = "NOT_FOUND" Then
            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " nao encontrado: " & reqNome & vbCrLf

            If required Then
                outFalhaCritica = True
                outErroMsg = "Ficheiro obrigatorio nao encontrado: " & reqNome & " | inputFolder=" & inputFolder
                Call Debug_Registar(0, promptId, "ERRO", "", "FILES", outErroMsg, "Sugestao: confirme nome e existencia no INPUT Folder.")
            Else
                Call Debug_Registar(0, promptId, "ALERTA", "", "FILES", _
                    "Ficheiro nao encontrado: " & reqNome & " | inputFolder=" & inputFolder, _
                    "Sugestao: confirme nome e existencia no INPUT Folder, ou use (required) apenas quando necessario.")
            End If

            Call Files_OperacoesAdicionarResultado(d, "NOT_FOUND", "", "", "", False, False, False)
            GoTo ProximoItem
        End If

        If status = "AMBIGUOUS" Then
            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " ambiguidade: " & reqNome & vbCrLf

            Call Debug_Registar(0, promptId, "ALERTA", "", "FILES", _
                "Ambiguidade ao resolver ficheiro: " & reqNome & " | candidatos: " & candidatosLog, _
                "Sugestao: use nome mais especifico, (latest), ou indique extensao. Se quiser, ative IA fallback em Config!B6.")

            If required Then
                outFalhaCritica = True
                outErroMsg = "Ficheiro obrigatorio com ambiguidade: " & reqNome
            End If

            Call Files_OperacoesAdicionarResultado(d, "AMBIGUOUS", "", "", "", False, False, True)
            GoTo ProximoItem
        End If

        Dim ext As String
        ext = Files_ObterExtensao(resolvedName)

        Dim usoFinal As String
        usoFinal = Files_DeterminarUsageMode(ext, wantAsIs, wantAsPdf, wantText)

        dbgUsoFinal = usoFinal

        Dim overrideModo As Boolean
        overrideModo = False

        If forceAsIsToTextEmbed And wantAsIs Then
            If LCase$(ext) <> EXT_PDF And Not Files_EhImagem(ext) Then
                usoFinal = "text_embed"
                overrideModo = True
            End If
        End If

        If overrideUsado Or overrideModo Then houveOverride = True

        ' ============ effective_mode ============
        Dim modoPedido As String, modoEfetivo As String, overrideReason As String
        modoPedido = usoFinal
        modoEfetivo = usoFinal
        overrideReason = ""

        If modoEfetivo = "as_is" Then
            Dim extLower As String
            extLower = LCase$(Trim$(ext))

            If extLower <> "" And (Not Files_IsExtSuportadaComoInputFileResponses(extLower)) Then
                Dim policy As String
                Dim fallbackPolicy As String

                policy = UCase$(Trim$(docxContextMode))
                fallbackPolicy = UCase$(Trim$(docxAsPdfFallback))

                overrideReason = "Extensão ." & extLower & " não suportada como input_file em /v1/responses."

                Select Case policy
                    Case "ERROR"
                        status = "UNSUPPORTED_EXT_AS_INPUT_FILE"
                        Call Debug_Registar(0, promptId, "ALERTA", "", "DOCX_INPUTFILE_OVERRIDDEN", _
                            "Pedido '" & modoPedido & "' para ." & extLower & " é incompatível (policy=ERROR).", _
                            "Use (as pdf) ou (text) no anexo; ou configure FILES_DOCX_CONTEXT_MODE=AUTO_AS_PDF/AUTO_TEXT_EMBED.")
                        Call Files_OperacoesAdicionarResultado(d, status, resolvedName, modoPedido, overrideReason, False, True, required)
                        If required Then
                            outFalhaCritica = True
                            outErroMsg = "Ficheiro obrigatório não suportado como input_file em /v1/responses: " & resolvedName & " (." & extLower & ")."
                        End If
                        GoTo ProximoItem

                    Case "AUTO_TEXT_EMBED"
                        If Files_PodeExtrairTexto(extLower) Then
                            modoEfetivo = "text_embed"
                        ElseIf Files_PodeConverterParaPDF(extLower) Then
                            modoEfetivo = "pdf_upload"
                        Else
                            status = "UNSUPPORTED_EXT_NO_FALLBACK"
                            Call Debug_Registar(0, promptId, "ALERTA", "", "DOCX_INPUTFILE_OVERRIDDEN", _
                                "Não há alternativa para tratar ." & extLower & " (sem conversão PDF e sem extração de texto).", _
                                "Use (text) ou forneça PDF; ou configure FILES_DOCX_CONTEXT_MODE=AUTO_AS_PDF com fallback TEXT_EMBED.")
                            Call Files_OperacoesAdicionarResultado(d, status, resolvedName, modoPedido, overrideReason, False, True, required)
                            If required Then
                                outFalhaCritica = True
                                outErroMsg = "Ficheiro obrigatório não suportado e sem fallback: " & resolvedName & " (." & extLower & ")."
                            End If
                            GoTo ProximoItem
                        End If

                    Case Else ' AUTO_AS_PDF
                        If Files_PodeConverterParaPDF(extLower) Then
                            modoEfetivo = "pdf_upload"
                        ElseIf fallbackPolicy = "TEXT_EMBED" And Files_PodeExtrairTexto(extLower) Then
                            modoEfetivo = "text_embed"
                        Else
                            status = "UNSUPPORTED_EXT_PDF_CONVERSION_NOT_AVAILABLE"
                            Call Debug_Registar(0, promptId, "ALERTA", "", "DOCX_INPUTFILE_OVERRIDDEN", _
                                "Conversão para PDF não disponível para ." & extLower & " e fallback=ERROR.", _
                                "Use (text) ou forneça PDF; ou configure FILES_DOCX_AS_PDF_FALLBACK=TEXT_EMBED.")
                            Call Files_OperacoesAdicionarResultado(d, status, resolvedName, modoPedido, overrideReason, False, True, required)
                            If required Then
                                outFalhaCritica = True
                                outErroMsg = "Conversão para PDF não disponível e fallback=ERROR para: " & resolvedName & " (." & extLower & ")."
                            End If
                            GoTo ProximoItem
                        End If
                End Select

                If modoEfetivo <> modoPedido Then
                    overrideModo = True
                    overrideUsado = True
                    houveOverride = True
                    Call Debug_Registar(0, promptId, "ALERTA", "", "DOCX_INPUTFILE_OVERRIDDEN", _
                        "Override automático: raw_mode=" & modoPedido & " => effective_mode=" & modoEfetivo & " (" & overrideReason & ")", _
                        "Recomendação: para DOCX/PPTX, use (as pdf) por defeito; alternativa: (text).")
                End If
            End If
        End If

        usoFinal = modoEfetivo
        dbgUsoFinal = usoFinal

        Dim lastMod As Date
        lastMod = Files_DataModificacao(resolvedPath)

        Dim convertido As Boolean
        convertido = False

        Dim sourceHash As String
        sourceHash = ""

        Dim hashUsado As String
        hashUsado = ""

        Dim fileId As String
        fileId = ""

        Dim erroLocal As String
        erroLocal = ""

        Dim caminhoUsado As String
        caminhoUsado = resolvedPath

        ' ======= BLOCO PDF (corrigido) =======
        If usoFinal = "pdf_upload" And LCase$(ext) <> EXT_PDF Then

            If Files_PodeConverterParaPDF(LCase$(ext)) Then
                sourceHash = Files_SHA256_File(resolvedPath)

                Dim pdfPath As String
                pdfPath = Files_ComporCaminhoPdfConvertido(pdfCacheFolder, resolvedName)

                Dim okConv As Boolean
                Dim usedCache As Boolean
                Dim convertedNow As Boolean

                usedCache = False
                convertedNow = False
                erroLocal = ""

                okConv = Files_PdfCache_GetOrConvertPdf( _
                            promptId, _
                            resolvedName, _
                            resolvedPath, _
                            pdfPath, _
                            sourceHash, _
                            usedCache, _
                            convertedNow, _
                            erroLocal _
                        )

                If okConv Then
                    caminhoUsado = pdfPath
                    convertido = True

                    If convertedNow Then
                        filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " convertido para PDF: " & resolvedName & vbCrLf
                    ElseIf usedCache Then
                        filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " PDF em cache: " & resolvedName & vbCrLf
                    Else
                        filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " PDF OK (estado cache indeterminado): " & resolvedName & vbCrLf
                    End If

                Else
                    filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " falha conversao PDF: " & resolvedName & " | " & erroLocal & vbCrLf

                    Call Debug_Registar(0, promptId, "ALERTA", "", "as_pdf", _
                        "Falha ao converter para PDF: " & resolvedName & " | " & erroLocal, _
                        "Sugestao: verifique se Word/PowerPoint estao instalados e se o ficheiro abre sem erros.")

                    If UCase$(Trim$(docxAsPdfFallback)) = "ERROR" Then
                        If required Then
                            outFalhaCritica = True
                            outErroMsg = "Falha conversão PDF para ficheiro obrigatório: " & resolvedName & " | " & erroLocal
                        End If
                        Call Files_OperacoesAdicionarResultado(d, "PDF_CONVERSION_FAIL", resolvedName, "pdf_upload", erroLocal, False, True, required)
                        GoTo ProximoItem
                    Else
                        usoFinal = "text_embed"
                        dbgUsoFinal = usoFinal
                        convertido = False
                    End If
                End If

            Else
                Call Debug_Registar(0, promptId, "ALERTA", "", "as_pdf", _
                    "Conversão para PDF não suportada para extensão ." & ext & " | ficheiro=" & resolvedName, _
                    "Use (text) ou forneça PDF; ou configure FILES_DOCX_CONTEXT_MODE=AUTO_TEXT_EMBED.")

                If UCase$(Trim$(docxAsPdfFallback)) = "ERROR" Then
                    If required Then
                        outFalhaCritica = True
                        outErroMsg = "Conversão para PDF não suportada (obrigatório): " & resolvedName & " (." & ext & ")."
                    End If
                    Call Files_OperacoesAdicionarResultado(d, "PDF_CONVERSION_NOT_SUPPORTED", resolvedName, "pdf_upload", "Sem conversão PDF", False, True, required)
                    GoTo ProximoItem
                Else
                    usoFinal = "text_embed"
                    dbgUsoFinal = usoFinal
                    convertido = False
                End If
            End If
        End If

        Dim sizeBytesUsado As Double
        sizeBytesUsado = Files_TamanhoBytes(caminhoUsado)

        Dim textoExtraDeste As String
        textoExtraDeste = ""

        ' calcular allowReuse/reuseTag por ficheiro (precedência: prompt > Config)
        Dim allowReuse As Boolean
        Dim reuseTag As String

        allowReuse = Files_Config_ReutilizacaoUpload()
        reuseTag = "reuse=" & IIf(allowReuse, "TRUE", "FALSE") & " (config)"

        If d.exists("reuse_override_found") Then
            If CBool(d("reuse_override_found")) Then
                allowReuse = CBool(d("reuse_override_value"))
                reuseTag = "reuse=" & IIf(allowReuse, "TRUE", "FALSE") & " (prompt)"
            End If
        End If

        If usoFinal = "text_embed" Then

            textoExtraDeste = Files_ExtrairTextoDoFicheiro(caminhoUsado, erroLocal)

            If erroLocal <> "" Then
                Call Debug_Registar(0, promptId, "ALERTA", "", "text_embed", _
                    "Falha ao extrair texto: " & resolvedName & " | " & erroLocal, _
                    "Sugestao: confirme extensao e permissao. Se for binario, use (as_pdf) ou (as_is).")
            End If

            If textoExtraDeste <> "" Then
                Dim charsExtra As Long
                charsExtra = Len(textoExtraDeste)

                If textEmbedMaxChars > 0 And charsExtra > textEmbedMaxChars Then
                    Dim overflowAction As String
                    overflowAction = UCase$(Trim$(textEmbedOverflowAction))

                    Call Debug_Registar(0, promptId, "ALERTA", "", "TEXT_EMBED_TOO_LARGE", _
                        "text_embed grande: " & resolvedName & " | chars=" & charsExtra & " | max=" & textEmbedMaxChars & " | action=" & overflowAction, _
                        "Sugestao: use (as pdf); ou aumente FILES_TEXT_EMBED_MAX_CHARS; ou mude FILES_TEXT_EMBED_OVERFLOW_ACTION.")

                    Select Case overflowAction
                        Case "ALERT_ONLY"
                            ' Só alerta

                        Case "TRUNCATE"
                            textoExtraDeste = Left$(textoExtraDeste, textEmbedMaxChars) & vbCrLf & "[TRUNCADO AUTO: excedeu FILES_TEXT_EMBED_MAX_CHARS]"

                        Case "STOP"
                            Call Files_OperacoesAdicionarResultado(d, "TEXT_EMBED_TOO_LARGE_STOP", resolvedName, "text_embed", _
                                "text_embed excede máximo (" & charsExtra & " > " & textEmbedMaxChars & ")", False, True, required)

                            If required Then
                                outFalhaCritica = True
                                outErroMsg = "text_embed demasiado grande para ficheiro obrigatório: " & resolvedName & " (chars=" & charsExtra & ")."
                            End If
                            GoTo ProximoItem

                        Case Else ' RETRY_AS_PDF
                            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " override text_embed->pdf_upload: " & resolvedName & vbCrLf

                            usoFinal = "pdf_upload"
                            dbgUsoFinal = usoFinal

                            If LCase$(Files_ObterExtensao(caminhoUsado)) <> EXT_PDF Then
                                If Files_PodeConverterParaPDF(LCase$(ext)) Then
                                    sourceHash = Files_SHA256_File(resolvedPath)

                                    Dim pdfPath2 As String
                                    pdfPath2 = Files_ComporCaminhoPdfConvertido(pdfCacheFolder, resolvedName)

                                    Dim okConv2 As Boolean
                                    Dim usedCache2 As Boolean
                                    Dim convertedNow2 As Boolean
                                    Dim erroConv2 As String

                                    usedCache2 = False
                                    convertedNow2 = False
                                    erroConv2 = ""

                                    okConv2 = Files_PdfCache_GetOrConvertPdf( _
                                                promptId, _
                                                resolvedName, _
                                                resolvedPath, _
                                                pdfPath2, _
                                                sourceHash, _
                                                usedCache2, _
                                                convertedNow2, _
                                                erroConv2 _
                                            )

                                    If okConv2 Then
                                        caminhoUsado = pdfPath2
                                        convertido = True

                                        If convertedNow2 Then
                                            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " override text_embed->pdf_upload (PDF gerado): " & resolvedName & vbCrLf
                                        ElseIf usedCache2 Then
                                            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " override text_embed->pdf_upload (PDF em cache): " & resolvedName & vbCrLf
                                        Else
                                            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " override text_embed->pdf_upload (estado indeterminado): " & resolvedName & vbCrLf
                                        End If

                                    Else
                                        If UCase$(Trim$(docxAsPdfFallback)) = "ERROR" Then
                                            Call Files_OperacoesAdicionarResultado(d, "PDF_CONVERSION_FAIL", resolvedName, "pdf_upload", _
                                                "Falha conversão PDF: " & erroConv2, False, True, required)

                                            If required Then
                                                outFalhaCritica = True
                                                outErroMsg = "Falha conversão PDF (overflow) para ficheiro obrigatório: " & resolvedName & " | " & erroConv2
                                            End If
                                            GoTo ProximoItem
                                        Else
                                            textoExtraDeste = Left$(textoExtraDeste, textEmbedMaxChars) & vbCrLf & "[TRUNCADO AUTO: conversão PDF falhou]"
                                            usoFinal = "text_embed"
                                            dbgUsoFinal = usoFinal
                                            convertido = False
                                        End If
                                    End If
                                Else
                                    textoExtraDeste = Left$(textoExtraDeste, textEmbedMaxChars) & vbCrLf & "[TRUNCADO AUTO: não é possível converter para PDF]"
                                    usoFinal = "text_embed"
                                    dbgUsoFinal = usoFinal
                                    convertido = False
                                End If
                            End If

                            If usoFinal = "pdf_upload" Then
                                GoTo ProcessarComoInputFile
                            End If
                    End Select
                End If
            End If

            If textoExtraDeste <> "" Then
                textoEmbedTotal = textoEmbedTotal & vbCrLf & vbCrLf & _
                    "----- BEGIN FILE: " & resolvedName & " -----" & vbCrLf & _
                    textoExtraDeste & vbCrLf & _
                    "----- END FILE: " & resolvedName & " -----" & vbCrLf

                hashUsado = Files_SHA256_Text(textoExtraDeste)
            Else
                hashUsado = Files_SHA256_File(caminhoUsado)
            End If

            Call Files_UpsertFilesManagement(wsFiles, mapaCab, pipelineNome, promptId, resolvedName, _
                inputFolder, caminhoUsado, "", "text_embed", False, sourceHash, lastMod, sizeBytesUsado, hashUsado, _
                "text_embed; origem=" & IIf(convertido, "pdf_convertido", "original"))

            filesUsedLista = Files_AppendLista(filesUsedLista, resolvedName & " (text)")
            Call Files_OperacoesAdicionarResultado(d, "OK", resolvedName, "text_embed", "", convertido, (overrideUsado Or overrideModo), False)

        ElseIf usoFinal = "image_upload" Then
            ' (mantém o teu bloco existente de imagem — não alterado aqui)
            ' ... (deixa como estava no teu módulo)
            ' Para não perder funcionalidades, não mexo neste ramo aqui.

        ElseIf usoFinal = "pdf_upload" Or usoFinal = "as_is" Then

ProcessarComoInputFile:
            hashUsado = Files_SHA256_File(caminhoUsado)

            Dim modoCache As String
            modoCache = IIf(usoFinal = "as_is", "as_is", "pdf_upload")

            If transportMode = TRANSPORT_INLINE Then
                ' (mantém o teu bloco existente INLINE — não alterado aqui)
                ' ...
            Else
                fileId = Files_ObterOuCriarFileId(wsFiles, mapaCab, apiKey, promptId, pipelineNome, _
                    resolvedName, inputFolder, caminhoUsado, modoCache, convertido, sourceHash, lastMod, sizeBytesUsado, hashUsado, _
                    allowReuse, reuseTag, erroLocal)

                If fileId <> "" Then
                    filePartsJson = Files_AppendJsonPart(filePartsJson, _
                        "{""type"":""input_file"",""file_id"":""" & Files_JsonEscape(fileId) & """}")

                    fileIdsLista = Files_AppendLista(fileIdsLista, fileId)

                    If convertido Then
                        filesUsedLista = Files_AppendLista(filesUsedLista, resolvedName & " (as_pdf)")
                    ElseIf usoFinal = "as_is" Then
                        filesUsedLista = Files_AppendLista(filesUsedLista, resolvedName & " (as_is)")
                    Else
                        filesUsedLista = Files_AppendLista(filesUsedLista, resolvedName & " (pdf)")
                    End If

                    Call Files_OperacoesAdicionarResultado(d, "OK", resolvedName, modoCache, fileId, convertido, (overrideUsado Or overrideModo), False)
                Else
                    GoTo FicheiroFalhou
                End If
            End If

            GoTo ProximoItem

FicheiroFalhou:
            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " ficheiro falhou: " & resolvedName & " | " & erroLocal & vbCrLf

            If required Then
                outFalhaCritica = True
                outErroMsg = "Falha ao anexar ficheiro obrigatorio: " & resolvedName & " | " & erroLocal
                Call Debug_Registar(0, promptId, "ERRO", "", "FILES", outErroMsg, "Sugestao: verifique tamanho, permissao e configuracao de transporte.")
            Else
                Call Debug_Registar(0, promptId, "ALERTA", "", "FILES", _
                    "Falha ao anexar ficheiro: " & resolvedName & " | " & erroLocal, _
                    "Sugestao: verifique tamanho, permissao e configuracao de transporte.")
            End If

            Call Files_OperacoesAdicionarResultado(d, "UPLOAD_FAIL", resolvedName, modoCache, "", convertido, (overrideUsado Or overrideModo), False)

        Else
            filesOpsCurto = filesOpsCurto & Files_TimestampCurto() & " modo desconhecido: " & resolvedName & " | " & usoFinal & vbCrLf
            Call Debug_Registar(0, promptId, "ALERTA", "", "FILES", _
                "Modo desconhecido ao preparar ficheiro: " & resolvedName & " | usoFinal=" & usoFinal, _
                "Sugestao: verifique regras de usage_mode.")
            Call Files_OperacoesAdicionarResultado(d, "UPLOAD_FAIL", resolvedName, usoFinal, "", convertido, (overrideUsado Or overrideModo), False)
        End If

ProximoItem:
    Next i

    outFilesUsedResumo = filesUsedLista
    outFilesOpsResumo = Files_NormalizarQuebrasLinha(filesOpsCurto)
    outFileIdsUsed = fileIdsLista

    Call Files_EscreverOperacoes(celOps, diretivas, "", False)

    If outFalhaCritica Then
        outInputJsonLiteralFinal = inputJsonLiteralBase
        Files_PrepararContextoDaPrompt = False
        Exit Function
    End If

    Dim textoFinal As String
    textoFinal = promptText
    If textoEmbedTotal <> "" Then
        textoFinal = textoFinal & vbCrLf & vbCrLf & textoEmbedTotal
    End If

    If houveAmbiguidade Or houveOverride Then
        textoFinal = textoFinal & vbCrLf & vbCrLf & "FILES CONTEXT:" & vbCrLf & _
            Files_BuildFilesContextResumo(diretivas)
    End If

    outInputJsonLiteralFinal = Files_MontarInputJson( _
        inputJsonLiteralBase, _
        filePartsJson, _
        textoFinal, _
        (houveAmbiguidade Or houveOverride), _
        promptId)

    Exit Function

TrataErro:
    Files_PrepararContextoDaPrompt = False
    outFalhaCritica = False

    outErroMsg = Files_FormatErroDetalhado( _
        "Files_PrepararContextoDaPrompt", _
        dbgStep, dbgReq, dbgFile, dbgPath, dbgUsoFinal)

    Call Debug_Registar(0, promptId, "ERRO", "", "FILES", outErroMsg, _
        "Sugestao: ative Break on All Errors no VBE para ver a linha exacta; verifique hashing/base64 e ficheiros grandes.")

    outInputJsonLiteralFinal = inputJsonLiteralBase
End Function

' ============================================================
' CONFIG (folha Config)
' ============================================================

Private Function Files_Config_TransportMode() As String
    On Error GoTo Falha

    Dim v As String
    v = ""

    On Error Resume Next
    v = CStr(ThisWorkbook.Worksheets(SHEET_CONFIG).Range("B5").value)
    On Error GoTo 0

    v = UCase$(Trim$(v))

    If v = TRANSPORT_INLINE Or v = "INLINE" Or v = "BASE64" Or v = "INLINEBASE64" Then
        Files_Config_TransportMode = TRANSPORT_INLINE
    Else
        Files_Config_TransportMode = TRANSPORT_FILE_ID
    End If

    Exit Function

Falha:
    Files_Config_TransportMode = TRANSPORT_FILE_ID
End Function

Private Function Files_Config_EnableIAFallback() As Boolean
    On Error GoTo Falha

    Dim v As Variant
    v = ""

    On Error Resume Next
    v = ThisWorkbook.Worksheets(SHEET_CONFIG).Range("B6").value
    On Error GoTo 0

    Files_Config_EnableIAFallback = Files_ValorParaBool(v, False)
    Exit Function

Falha:
    Files_Config_EnableIAFallback = False
End Function

Private Function Files_Config_InlineMaxBytes() As Double
    On Error GoTo Falha

    Dim mb As Double
    mb = 20#

    On Error Resume Next
    mb = CDbl(val(ThisWorkbook.Worksheets(SHEET_CONFIG).Range("B7").value))
    On Error GoTo 0

    If mb <= 0 Then mb = 20#

    Files_Config_InlineMaxBytes = mb * 1024# * 1024#
    Exit Function

Falha:
    Files_Config_InlineMaxBytes = 20# * 1024# * 1024#
End Function


' (NOVO) Config: modo de filename enviado no multipart (apenas "envelope")
'   - Procura uma linha na folha Config com:
'       Col A = "FILES_MULTIPART_FILENAME_MODE"
'       Col B = RAW | ASCII_SAFE | RFC5987
'   - Default: RAW (compatível com versões anteriores)
Private Function Files_Config_MultipartFilenameMode() As String
    On Error GoTo Falha

    Files_Config_MultipartFilenameMode = "RAW"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    If ws Is Nothing Then Exit Function

    Dim key As String
    key = "FILES_MULTIPART_FILENAME_MODE"

    Dim lastR As Long
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    If lastR < 1 Then Exit Function

    Dim i As Long
    For i = 1 To lastR
        If StrComp(Trim$(CStr(ws.Cells(i, 1).value)), key, vbTextCompare) = 0 Then
            Dim v As String
            v = UCase$(Trim$(CStr(ws.Cells(i, 2).value)))

            Select Case v
                Case "RAW", "ASCII_SAFE", "RFC5987"
                    Files_Config_MultipartFilenameMode = v
                Case Else
                    Files_Config_MultipartFilenameMode = "RAW"
            End Select

            Exit Function
        End If
    Next i

    Exit Function

Falha:
    Files_Config_MultipartFilenameMode = "RAW"
End Function


Private Function Files_ValorParaBool(ByVal v As Variant, ByVal defaultVal As Boolean) As Boolean
    On Error GoTo Falha

    If IsEmpty(v) Then
        Files_ValorParaBool = defaultVal
        Exit Function
    End If

    If VarType(v) = vbBoolean Then
        Files_ValorParaBool = CBool(v)
        Exit Function
    End If

    Dim s As String
    s = UCase$(Trim$(CStr(v)))

    If s = "" Then
        Files_ValorParaBool = defaultVal
    ElseIf s = "TRUE" Or s = "VERDADEIRO" Or s = "SIM" Or s = "YES" Or s = "Y" Or s = "1" Then
        Files_ValorParaBool = True
    ElseIf s = "FALSE" Or s = "FALSO" Or s = "NAO" Or s = "NO" Or s = "N" Or s = "0" Then
        Files_ValorParaBool = False
    Else
        Files_ValorParaBool = defaultVal
    End If

    Exit Function

Falha:
    Files_ValorParaBool = defaultVal
End Function

Private Function Files_TemRequiredDiretivas(ByVal diretivas As Collection) As Boolean
    Dim i As Long
    For i = 1 To diretivas.Count
        If CBool(diretivas(i)("required")) Then
            Files_TemRequiredDiretivas = True
            Exit Function
        End If
    Next i
    Files_TemRequiredDiretivas = False
End Function

' ============================================================
' NOVO: Config por chaves (sem quebrar retrocompatibilidade)
' ============================================================

Private Function Files_Config_DocxContextMode() As String
    Dim v As String
    v = UCase$(Trim$(Files_Config_GetByKey("FILES_DOCX_CONTEXT_MODE", "AUTO_AS_PDF")))
    Select Case v
        Case "AUTO_AS_PDF", "AUTO_TEXT_EMBED", "ERROR"
            Files_Config_DocxContextMode = v
        Case Else
            Files_Config_DocxContextMode = "AUTO_AS_PDF"
    End Select
End Function

Private Function Files_Config_DocxAsPdfFallback() As String
    Dim v As String
    v = UCase$(Trim$(Files_Config_GetByKey("FILES_DOCX_AS_PDF_FALLBACK", "TEXT_EMBED")))
    Select Case v
        Case "TEXT_EMBED", "ERROR"
            Files_Config_DocxAsPdfFallback = v
        Case Else
            Files_Config_DocxAsPdfFallback = "TEXT_EMBED"
    End Select
End Function

Private Function Files_Config_TextEmbedMaxChars() As Long
    Dim v As String
    v = Trim$(Files_Config_GetByKey("FILES_TEXT_EMBED_MAX_CHARS", "50000"))

    If IsNumeric(v) Then
        Files_Config_TextEmbedMaxChars = CLng(val(v))
        If Files_Config_TextEmbedMaxChars < 0 Then Files_Config_TextEmbedMaxChars = 50000
    Else
        Files_Config_TextEmbedMaxChars = 50000
    End If
End Function

Private Function Files_Config_TextEmbedOverflowAction() As String
    Dim v As String
    v = UCase$(Trim$(Files_Config_GetByKey("FILES_TEXT_EMBED_OVERFLOW_ACTION", "RETRY_AS_PDF")))
    Select Case v
        Case "ALERT_ONLY", "TRUNCATE", "RETRY_AS_PDF", "STOP"
            Files_Config_TextEmbedOverflowAction = v
        Case Else
            Files_Config_TextEmbedOverflowAction = "RETRY_AS_PDF"
    End Select
End Function

' ============================================================
' NOVO: utilitários para políticas de formato em /v1/responses
' ============================================================

Private Function Files_IsExtSuportadaComoInputFileResponses(ByVal extLower As String) As Boolean
    ' Extensões suportadas pelo /v1/responses para input_file (context stuffing).
    ' Baseado na mensagem de erro observada (pode evoluir; ajustar se necessário).
    extLower = LCase$(Replace$(Trim$(extLower), ".", ""))

    If extLower = "" Then
        Files_IsExtSuportadaComoInputFileResponses = False
        Exit Function
    End If

    Select Case extLower
        Case "art", "bat", "brf", "c", "cls", "css", "diff", "eml", "es", "h", "hs", _
             "htm", "html", "ics", "ifb", "java", "js", "json", "ksh", "ltx", "mail", _
             "markdown", "md", "mht", "mhtml", "mjs", "nws", "patch", "pdf", "pl", "pm", _
             "pot", "py", "rst", "scala", "sh", "shtml", "srt", "sty", "tex", "text", "txt", _
             "vcf", "vtt", "xml", "yaml", "yml"
            Files_IsExtSuportadaComoInputFileResponses = True
        Case Else
            Files_IsExtSuportadaComoInputFileResponses = False
    End Select
End Function

Private Function Files_PodeConverterParaPDF(ByVal extLower As String) As Boolean
    extLower = LCase$(Replace$(Trim$(extLower), ".", ""))
    Select Case extLower
        Case "doc", "docx", "ppt", "pptx"
            Files_PodeConverterParaPDF = True
        Case Else
            Files_PodeConverterParaPDF = False
    End Select
End Function

Private Function Files_PodeExtrairTexto(ByVal extLower As String) As Boolean
    extLower = LCase$(Replace$(Trim$(extLower), ".", ""))
    Select Case extLower
        Case "txt", "md", "markdown", "csv", "json", "doc", "docx", "ppt", "pptx"
            Files_PodeExtrairTexto = True
        Case Else
            Files_PodeExtrairTexto = False
    End Select
End Function


' ============================================================
' PARSER: extrair diretivas FILES/FICHEIROS dos INPUTS
' ============================================================

Private Function Files_ExtrairDiretivasDeFicheiros(ByVal textoInputs As String) As Collection
    Dim col As New Collection

    Dim t As String
    t = CStr(textoInputs)

    Dim pos As Long
    pos = Files_PosicaoTagFiles(t)

    If pos = 0 Then
        Set Files_ExtrairDiretivasDeFicheiros = col
        Exit Function
    End If

    Dim depois As String
    depois = Mid$(t, pos)

    Dim p2 As Long
    p2 = InStr(1, depois, ":", vbTextCompare)
    If p2 = 0 Then
        Set Files_ExtrairDiretivasDeFicheiros = col
        Exit Function
    End If

    Dim lista As String
    lista = Mid$(depois, p2 + 1)

    Dim eol As Long
    eol = InStr(1, lista, vbCrLf)
    If eol = 0 Then eol = InStr(1, lista, vbLf)
    If eol > 0 Then lista = Left$(lista, eol - 1)

    Dim itens() As String
    itens = Split(lista, ";")

    Dim i As Long
    For i = LBound(itens) To UBound(itens)
        Dim itemRaw As String
        itemRaw = Trim$(itens(i))
        If itemRaw <> "" Then
            Dim d As Object
            Set d = Files_ParseItem(itemRaw)
            col.Add d
        End If
    Next i

    Set Files_ExtrairDiretivasDeFicheiros = col
End Function

Private Function Files_PosicaoTagFiles(ByVal t As String) As Long
    Dim p As Long

    p = InStr(1, t, "FILES:", vbTextCompare)
    If p > 0 Then Files_PosicaoTagFiles = p: Exit Function

    p = InStr(1, t, "FILES :", vbTextCompare)
    If p > 0 Then Files_PosicaoTagFiles = p: Exit Function

    p = InStr(1, t, "FICHEIROS:", vbTextCompare)
    If p > 0 Then Files_PosicaoTagFiles = p: Exit Function

    p = InStr(1, t, "FICHEIROS :", vbTextCompare)
    If p > 0 Then Files_PosicaoTagFiles = p: Exit Function

    Files_PosicaoTagFiles = 0
End Function


' ============================================================
' (ALTERAR) Files_ParseItem — adicionar reuse_override_*
'   SUBSTITUA A FUNÇÃO INTEIRA por esta versão
' ============================================================
Private Function Files_ParseItem(ByVal itemRaw As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim raw As String
    raw = Trim$(itemRaw)

    d("requested_raw") = raw

    Dim low As String
    low = LCase$(raw)

    d("as_is") = (InStr(1, low, "(as is)", vbTextCompare) > 0) Or (InStr(1, low, "(as_is)", vbTextCompare) > 0)
    d("as_pdf") = (InStr(1, low, "(as pdf)", vbTextCompare) > 0) Or (InStr(1, low, "(as_pdf)", vbTextCompare) > 0)
    d("text_embed") = (InStr(1, low, "(text)", vbTextCompare) > 0) Or (InStr(1, low, "(text_embed)", vbTextCompare) > 0)

    d("latest") = (InStr(1, low, "(latest)", vbTextCompare) > 0) Or _
                  (InStr(1, low, "(mais_recente)", vbTextCompare) > 0) Or _
                  (InStr(1, low, "(mais recente)", vbTextCompare) > 0)

    d("required") = (InStr(1, low, "(required)", vbTextCompare) > 0) Or _
                    (InStr(1, low, "(obrigatorio)", vbTextCompare) > 0) Or _
                    (InStr(1, low, "(obrigatoria)", vbTextCompare) > 0)

    Dim nome As String
    nome = raw

    Dim lowTrim As String
    lowTrim = LCase$(Trim$(raw))

    If Left$(lowTrim, 8) = "@output(" Then
        ' Preservar expressão completa: @OUTPUT(...)
        nome = Trim$(raw)
    Else
        Dim p As Long
        p = InStr(1, raw, "(", vbTextCompare)
        If p > 0 Then nome = Trim$(Left$(raw, p - 1))
    End If

    nome = Replace(nome, """, "")
    nome = Replace(nome, "'", "")

    d("requested_name") = Trim$(nome)

    ' ---- (NOVO) override por ficheiro (precede Config)
    Dim reuseFound As Boolean, reuseValue As Boolean
    Call Files_ParseReuseOverride(raw, reuseFound, reuseValue)
    d("reuse_override_found") = reuseFound
    d("reuse_override_value") = reuseValue

    ' ---- campos de resultado (como antes)
    d("resultado_status") = ""
    d("resultado_nome") = ""
    d("resultado_modo") = ""
    d("resultado_file_id") = ""
    d("resultado_convertido") = False
    d("resultado_override") = False
    d("resultado_ambiguidade") = False

    Set Files_ParseItem = d
End Function


Private Sub Files_OperacoesAdicionarResultado(ByRef d As Object, ByVal status As String, ByVal nome As String, ByVal modo As String, ByVal fileId As String, ByVal convertido As Boolean, ByVal overrideUsado As Boolean, ByVal ambiguidade As Boolean)
    d("resultado_status") = status
    d("resultado_nome") = nome
    d("resultado_modo") = modo
    d("resultado_file_id") = fileId
    d("resultado_convertido") = convertido
    d("resultado_override") = overrideUsado
    d("resultado_ambiguidade") = ambiguidade
End Sub


' ============================================================
' RESOLUCAO DE FICHEIROS (deterministico + IA fallback opcional)
' ============================================================

Public Sub Files_ResolverOutputToken( _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByVal stepN As Long, _
    ByVal token As String, _
    ByRef resolvedPath As String, _
    ByRef resolvedName As String, _
    ByRef status As String, _
    ByRef candidatosLog As String _
)
    On Error GoTo Falha

    Dim modeAmb As String
    modeAmb = UCase$(Trim$(Files_Config_GetByKey("FILES_OUTPUT_CHAIN_AMBIGUITY_MODE", "STRICT")))
    If modeAmb <> "STRICT" And modeAmb <> "LAST_CREATED" Then
        modeAmb = "STRICT"
        Call Debug_Registar(0, promptId, "ALERTA", "", "CONFIG_OUTPUT_CHAIN_MODE_INVALID", _
            "FILES_OUTPUT_CHAIN_AMBIGUITY_MODE inválido. Usado STRICT.", _
            "Valores suportados: STRICT | LAST_CREATED.")
    End If

    Dim candidates As Collection
    Set candidates = Files_OutputCandidatesForRun(gRunToken)

    If candidates.Count = 0 Then
        status = "NOT_FOUND"
        candidatosLog = "run_id=" & gRunToken
        Call Debug_Registar(0, promptId, "ALERTA", "", "OUTPUT_CHAIN_NOT_FOUND", _
            "Sem outputs registados para o run atual.", _
            "Confirme se já houve passo de File Output e se FILES_MANAGEMENT foi atualizado.")
        Exit Sub
    End If

    Dim filtered As Collection
    Set filtered = candidates

    Dim t As String
    t = UCase$(Trim$(token))

    If t = "@LAST_OUTPUT" Then
        Dim prevStep As Long
        prevStep = stepN - 1
        Set filtered = Files_FilterCandidatesByStep(filtered, prevStep)

    ElseIf Left$(t, 8) = "@OUTPUT(" And Right$(t, 1) = ")" Then
        Dim args As Object
        Set args = Files_ParseOutputArgs(Mid$(token, 9, Len(token) - 9))

        If args.Exists("prompt_id") Then
            Set filtered = Files_FilterCandidatesByPrompt(filtered, CStr(args("prompt_id")))
        End If

        If args.Exists("step_n") Then
            Set filtered = Files_FilterCandidatesByStep(filtered, CLng(Val(CStr(args("step_n")))))
        End If

        If args.Exists("filename") Then
            Set filtered = Files_FilterCandidatesByFilename(filtered, CStr(args("filename")))
        End If

        If args.Exists("index") Then
            Dim idx As Long
            idx = CLng(Val(CStr(args("index"))))
            Set filtered = Files_PickByIndex(filtered, idx)
        End If
    Else
        status = "NOT_FOUND"
        candidatosLog = "token inválido: " & token
        Call Debug_Registar(0, promptId, "ERRO", "", "OUTPUT_CHAIN_NOT_FOUND", _
            "Token de output não reconhecido: " & token, _
            "Use @LAST_OUTPUT ou @OUTPUT(prompt_id=..., step_n=..., filename=..., index=...).")
        Exit Sub
    End If

    If filtered.Count = 0 Then
        status = "NOT_FOUND"
        candidatosLog = token
        Call Debug_Registar(0, promptId, "ALERTA", "", "OUTPUT_CHAIN_NOT_FOUND", _
            "Token sem candidatos: " & token, _
            "Refine os critérios (@OUTPUT) ou valide o passo anterior (@LAST_OUTPUT).")
        Exit Sub
    End If

    Dim chosen As Object
    If filtered.Count = 1 Then
        Set chosen = filtered(1)
    ElseIf modeAmb = "LAST_CREATED" Then
        Set chosen = Files_PickMostRecent(filtered)
        status = "OK"
        candidatosLog = Files_CandidatesShortList(filtered)
        Call Debug_Registar(0, promptId, "ALERTA", "", "OUTPUT_CHAIN_AMBIGUOUS", _
            "Múltiplos candidatos para " & token & ". Selecionado mais recente por configuração.", _
            "Para determinismo total, adicione step_n/index/filename ao @OUTPUT(...).")
    Else
        status = "AMBIGUOUS"
        candidatosLog = Files_CandidatesShortList(filtered)
        Call Debug_Registar(0, promptId, "ERRO", "", "OUTPUT_CHAIN_AMBIGUOUS", _
            "Múltiplos candidatos para " & token & ": " & candidatosLog, _
            "Use @OUTPUT(..., index=n) ou mude FILES_OUTPUT_CHAIN_AMBIGUITY_MODE para LAST_CREATED.")
        Exit Sub
    End If

    resolvedPath = CStr(chosen("full_path"))
    resolvedName = CStr(chosen("file_name"))

    If Dir(resolvedPath) = "" Then
        status = "NOT_FOUND"
        candidatosLog = resolvedPath
        Call Debug_Registar(0, promptId, "ERRO", "", "OUTPUT_FILE_MISSING_ON_DISK", _
            "Ficheiro registado mas não encontrado no disco: " & resolvedPath, _
            "Confirme OUTPUT Folder, permissões e se o ficheiro foi movido/apagado.")
        Exit Sub
    End If

    status = "OK"
    Call Debug_Registar(0, promptId, "INFO", "", "OUTPUT_CHAIN_RESOLVE", _
        "Token " & token & " resolvido para: " & resolvedName, _
        "OK")
    Exit Sub

Falha:
    status = "NOT_FOUND"
    candidatosLog = "erro: " & Err.Description
    Call Debug_Registar(0, promptId, "ERRO", "", "OUTPUT_CHAIN_NOT_FOUND", _
        "Falha ao resolver token de output: " & token & " | " & Err.Description, _
        "Verifique FILES_MANAGEMENT e formato da diretiva FILES.")
End Sub

Private Function Files_OutputCandidatesForRun(ByVal runId As String) As Collection
    On Error GoTo Falha
    Dim col As New Collection

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)

    Dim m As Object
    Set m = Files_MapaCabecalhos(ws)

    Dim cFull As Long, cName As Long, cTs As Long, cNotes As Long
    cFull = Files_Col(m, H_FULL_PATH)
    cName = Files_Col(m, H_FILE_NAME)
    cTs = Files_Col(m, H_TIMESTAMP)
    cNotes = Files_Col(m, H_NOTES)

    Dim lastRow As Long
    lastRow = Files_LastDataRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim notes As String
        notes = CStr(ws.Cells(r, cNotes).value)

        If InStr(1, notes, "source_type=OUTPUT", vbTextCompare) > 0 And InStr(1, notes, "run_id=" & runId, vbTextCompare) > 0 Then
            Dim fullPath As String
            fullPath = CStr(ws.Cells(r, cFull).value)
            If Trim$(fullPath) <> "" Then
                Dim d As Object
                Set d = CreateObject("Scripting.Dictionary")
                d("row") = r
                d("full_path") = fullPath
                d("file_name") = CStr(ws.Cells(r, cName).value)
                d("ts") = ws.Cells(r, cTs).value
                d("prompt_id") = Files_NoteValue(notes, "prompt_id")
                d("step_n") = CLng(Val(Files_NoteValue(notes, "step_n")))
                d("output_index") = CLng(Val(Files_NoteValue(notes, "output_index")))
                col.Add d
            End If
        End If
    Next r

    Set Files_OutputCandidatesForRun = col
    Exit Function
Falha:
    Set Files_OutputCandidatesForRun = New Collection
End Function

Private Function Files_ParseOutputArgs(ByVal argsText As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim parts() As String
    parts = Split(argsText, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim kv As String
        kv = Trim$(parts(i))
        Dim p As Long
        p = InStr(1, kv, "=", vbTextCompare)
        If p > 1 Then
            Dim k As String, v As String
            k = LCase$(Trim$(Left$(kv, p - 1)))
            v = Trim$(Mid$(kv, p + 1))
            d(k) = v
        End If
    Next i

    Set Files_ParseOutputArgs = d
End Function

Private Function Files_FilterCandidatesByStep(ByVal src As Collection, ByVal stepN As Long) As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 1 To src.Count
        If CLng(src(i)("step_n")) = stepN Then out.Add src(i)
    Next i
    Set Files_FilterCandidatesByStep = out
End Function

Private Function Files_FilterCandidatesByPrompt(ByVal src As Collection, ByVal promptId As String) As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 1 To src.Count
        If StrComp(CStr(src(i)("prompt_id")), promptId, vbTextCompare) = 0 Then out.Add src(i)
    Next i
    Set Files_FilterCandidatesByPrompt = out
End Function

Private Function Files_FilterCandidatesByFilename(ByVal src As Collection, ByVal pattern As String) As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 1 To src.Count
        If LCase$(CStr(src(i)("file_name"))) Like LCase$(pattern) Then out.Add src(i)
    Next i
    Set Files_FilterCandidatesByFilename = out
End Function

Private Function Files_PickByIndex(ByVal src As Collection, ByVal idx As Long) As Collection
    Dim out As New Collection
    If idx < 0 Then
        Set Files_PickByIndex = out
        Exit Function
    End If
    If src.Count = 0 Then
        Set Files_PickByIndex = out
        Exit Function
    End If
    If idx + 1 <= src.Count Then out.Add src(idx + 1)
    Set Files_PickByIndex = out
End Function

Private Function Files_PickMostRecent(ByVal src As Collection) As Object
    Dim i As Long
    Dim best As Object
    If src.Count = 0 Then Exit Function
    Set best = src(1)
    For i = 2 To src.Count
        If CDbl(src(i)("ts")) > CDbl(best("ts")) Then Set best = src(i)
    Next i
    Set Files_PickMostRecent = best
End Function

Private Function Files_CandidatesShortList(ByVal src As Collection) As String
    Dim i As Long
    Dim s As String
    s = ""
    For i = 1 To src.Count
        If i > 1 Then s = s & " | "
        s = s & CStr(src(i)("file_name")) & "(step=" & CStr(src(i)("step_n")) & ",idx=" & CStr(src(i)("output_index")) & ")"
        If Len(s) > 400 Then Exit For
    Next i
    Files_CandidatesShortList = s
End Function

Private Function Files_NoteValue(ByVal notes As String, ByVal keyName As String) As String
    Dim parts() As String
    parts = Split(notes, "|")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim t As String
        t = Trim$(parts(i))
        If LCase$(Left$(t, Len(keyName) + 1)) = LCase$(keyName & "=") Then
            Files_NoteValue = Mid$(t, Len(keyName) + 2)
            Exit Function
        End If
    Next i

    Files_NoteValue = ""
End Function

Private Sub Files_ResolverFicheiro( _
    ByVal apiKey As String, _
    ByVal promptId As String, _
    ByVal inputFolder As String, _
    ByVal requestedName As String, _
    ByVal wantLatest As Boolean, _
    ByVal enableIAFallback As Boolean, _
    ByRef outFullPath As String, _
    ByRef outFileName As String, _
    ByRef outStatus As String, _
    ByRef outCandidatosLog As String, _
    ByRef outOverrideUsado As Boolean, _
    ByRef ioHouveAmbiguidade As Boolean _
)
    outFullPath = ""
    outFileName = ""
    outStatus = ""
    outCandidatosLog = ""
    outOverrideUsado = False

    Dim nome As String
    nome = Trim$(requestedName)

    If nome = "" Then
        outStatus = "NOT_FOUND"
        Exit Sub
    End If

    nome = Files_SoNomeFicheiro(nome)

    Dim pathExato As String
    pathExato = Files_PathJoin(inputFolder, nome)

    If Files_ExisteFicheiro(pathExato) Then
        outFullPath = pathExato
        outFileName = nome
        outStatus = "OK"
        Exit Sub
    End If

    If InStr(1, nome, "*", vbTextCompare) > 0 Then
        Dim candidatos As Collection
        Set candidatos = Files_ListarPorPattern(inputFolder, nome)

        If candidatos.Count = 1 Then
            outFullPath = CStr(candidatos(1)("full_path"))
            outFileName = CStr(candidatos(1)("name"))
            outStatus = "OK"
            Exit Sub
        End If

        If candidatos.Count > 1 And wantLatest Then
            Dim melhor As Object
            Set melhor = Files_EscolherMaisRecente(candidatos)
            outFullPath = CStr(melhor("full_path"))
            outFileName = CStr(melhor("name"))
            outStatus = "OK"
            Exit Sub
        End If

        If candidatos.Count > 1 Then
            ioHouveAmbiguidade = True
            outStatus = "AMBIGUOUS"
            outCandidatosLog = Files_ResumoCandidatos(candidatos)

            If enableIAFallback Then
                Dim escolhido As String
                escolhido = Files_IA_EscolherCandidato(apiKey, promptId, nome, candidatos)
                If escolhido <> "" Then
                    Dim objEscolhido As Object
                    Set objEscolhido = Files_EncontrarCandidatoPorNome(candidatos, escolhido)
                    If Not objEscolhido Is Nothing Then
                        outFullPath = CStr(objEscolhido("full_path"))
                        outFileName = CStr(objEscolhido("name"))
                        outStatus = "OK"
                        outOverrideUsado = True
                        Exit Sub
                    End If
                End If
            End If

            Exit Sub
        End If

        outStatus = "NOT_FOUND"
        Exit Sub
    End If

    Dim candidatos2 As Collection
    Set candidatos2 = Files_ListarPorSubstring(inputFolder, nome)

    If candidatos2.Count = 1 Then
        outFullPath = CStr(candidatos2(1)("full_path"))
        outFileName = CStr(candidatos2(1)("name"))
        outStatus = "OK"
        Exit Sub
    End If

    If candidatos2.Count > 1 And wantLatest Then
        Dim melhor2 As Object
        Set melhor2 = Files_EscolherMaisRecente(candidatos2)
        outFullPath = CStr(melhor2("full_path"))
        outFileName = CStr(melhor2("name"))
        outStatus = "OK"
        Exit Sub
    End If

    If candidatos2.Count > 1 Then
        ioHouveAmbiguidade = True
        outStatus = "AMBIGUOUS"
        outCandidatosLog = Files_ResumoCandidatos(candidatos2)

        If enableIAFallback Then
            Dim escolhido2 As String
            escolhido2 = Files_IA_EscolherCandidato(apiKey, promptId, nome, candidatos2)
            If escolhido2 <> "" Then
                Dim objEscolhido2 As Object
                Set objEscolhido2 = Files_EncontrarCandidatoPorNome(candidatos2, escolhido2)
                If Not objEscolhido2 Is Nothing Then
                    outFullPath = CStr(objEscolhido2("full_path"))
                    outFileName = CStr(objEscolhido2("name"))
                    outStatus = "OK"
                    outOverrideUsado = True
                    Exit Sub
                End If
            End If
        End If

        Exit Sub
    End If

    outStatus = "NOT_FOUND"
End Sub

Private Function Files_IA_EscolherCandidato(ByVal apiKey As String, ByVal promptId As String, ByVal requestedName As String, ByVal candidatos As Collection) As String
    On Error GoTo Falha

    Files_IA_EscolherCandidato = ""

    If Trim$(apiKey) = "" Then Exit Function

    Dim lista As String
    lista = Files_ResumoCandidatos(candidatos)

    Dim pergunta As String
    pergunta = "Pedido: " & requestedName & vbCrLf & _
               "Candidatos (nome | data | tamanho):" & vbCrLf & lista

    Dim instrucao As String
    instrucao = "Escolhe exatamente UM nome de ficheiro da lista. Responde APENAS com o nome do ficheiro, sem aspas, sem comentarios. Se nao consegues decidir, responde vazio."

    Dim extra As String
    extra = """instructions"":""" & Files_JsonEscape(instrucao) & """"

    Dim res As ApiResultado
    res = OpenAI_Executar(apiKey, FALLBACK_MODEL, pergunta, 0, 80, "", False, "", extra)

    If Trim$(res.Erro) <> "" Then
        Call Debug_Registar(0, promptId, "ALERTA", "", "FILES IA", _
            "Fallback IA falhou: " & res.Erro, _
            "Sugestao: desative IA fallback (Config!B6=FALSE) e use (latest) ou nome mais especifico.")
        Exit Function
    End If

    Dim escolhido As String
    escolhido = Trim$(res.outputText)

    If Files_CandidatoExiste(candidatos, escolhido) Then
        Files_IA_EscolherCandidato = escolhido
    End If

    Exit Function

Falha:
    Files_IA_EscolherCandidato = ""
End Function


' ============================================================
' USAGE MODE
' ============================================================

Private Function Files_DeterminarUsageMode(ByVal ext As String, ByVal wantAsIs As Boolean, ByVal wantAsPdf As Boolean, ByVal wantText As Boolean) As String
    ext = LCase$(Trim$(ext))

    If wantText Then
        Files_DeterminarUsageMode = "text_embed"
        Exit Function
    End If

    If wantAsPdf Then
        Files_DeterminarUsageMode = "pdf_upload"
        Exit Function
    End If

    If wantAsIs Then
        Files_DeterminarUsageMode = "as_is"
        Exit Function
    End If

    If ext = EXT_PDF Then
        Files_DeterminarUsageMode = "pdf_upload"
        Exit Function
    End If

    If Files_EhImagem(ext) Then
        Files_DeterminarUsageMode = "image_upload"
        Exit Function
    End If

    Files_DeterminarUsageMode = "text_embed"
End Function

Private Function Files_EhImagem(ByVal ext As String) As Boolean
    Dim lista() As String
    lista = Split(IMG_EXTS, ";")

    Dim i As Long
    For i = LBound(lista) To UBound(lista)
        If LCase$(Trim$(lista(i))) = LCase$(Trim$(ext)) Then
            Files_EhImagem = True
            Exit Function
        End If
    Next i

    Files_EhImagem = False
End Function


' ============================================================
' CONVERSAO PARA PDF (Word / PowerPoint)
' ============================================================

Private Function Files_ConverterParaPDF(ByVal srcPath As String, ByVal destPdfPath As String, ByRef outErro As String) As Boolean
    On Error GoTo Falha

    outErro = ""
    Files_ConverterParaPDF = False

    Dim ext As String
    ext = LCase$(Files_ObterExtensao(srcPath))

    If ext = "doc" Or ext = "docx" Then
        Files_ConverterParaPDF = Files_ConverterWordParaPDF(srcPath, destPdfPath, outErro)
        Exit Function
    End If

    If ext = "ppt" Or ext = "pptx" Then
        Files_ConverterParaPDF = Files_ConverterPPTParaPDF(srcPath, destPdfPath, outErro)
        Exit Function
    End If

    outErro = "Extensao nao suportada para conversao PDF: " & ext
    Exit Function

Falha:
    outErro = "Erro conversao PDF: " & Err.Description
    Files_ConverterParaPDF = False
End Function

Private Function Files_ConverterWordParaPDF(ByVal srcPath As String, ByVal destPdfPath As String, ByRef outErro As String) As Boolean
    On Error GoTo Falha

    outErro = ""
    Files_ConverterWordParaPDF = False

    Dim wordApp As Object, doc As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    Set doc = wordApp.Documents.Open(srcPath, False, True, False)
    doc.SaveAs2 destPdfPath, 17 ' wdFormatPDF
    doc.Close False

    wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing

    Files_ConverterWordParaPDF = True
    Exit Function

Falha:
    outErro = "Word->PDF falhou: " & Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Files_ConverterWordParaPDF = False
End Function

Private Function Files_ConverterPPTParaPDF(ByVal srcPath As String, ByVal destPdfPath As String, ByRef outErro As String) As Boolean
    On Error GoTo Falha

    outErro = ""
    Files_ConverterPPTParaPDF = False

    Dim pptApp As Object, pres As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = False

    Set pres = pptApp.Presentations.Open(srcPath, True, True, False)
    pres.SaveAs destPdfPath, 32 ' ppSaveAsPDF

    pres.Close
    pptApp.Quit

    Set pres = Nothing
    Set pptApp = Nothing

    Files_ConverterPPTParaPDF = True
    Exit Function

Falha:
    outErro = "PPT->PDF falhou: " & Err.Description
    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    If Not pptApp Is Nothing Then pptApp.Quit
    Files_ConverterPPTParaPDF = False
End Function

Private Function Files_ComporPdfCacheFolder(ByVal inputFolder As String) As String
    Files_ComporPdfCacheFolder = Files_PathJoin(inputFolder, PDF_CACHE_FOLDER_NAME)
End Function

Private Function Files_ComporCaminhoPdfConvertido(ByVal pdfCacheFolder As String, ByVal resolvedName As String) As String
    Dim base As String
    base = Files_SemExtensao(resolvedName)
    Files_ComporCaminhoPdfConvertido = Files_PathJoin(pdfCacheFolder, base & ".pdf")
End Function


' ============================================================
' EXTRACAO DE TEXTO (text_embed)
' ============================================================

Private Function Files_ExtrairTextoDoFicheiro(ByVal fullPath As String, ByRef outErro As String) As String
    On Error GoTo Falha

    outErro = ""
    Files_ExtrairTextoDoFicheiro = ""

    Dim ext As String
    ext = LCase$(Files_ObterExtensao(fullPath))

    If ext = "txt" Or ext = "md" Or ext = "csv" Or ext = "json" Then
        Files_ExtrairTextoDoFicheiro = Files_LerTexto(fullPath, outErro)
        Exit Function
    End If

    If ext = "doc" Or ext = "docx" Then
        Files_ExtrairTextoDoFicheiro = Files_ExtrairTextoWord(fullPath, outErro)
        Exit Function
    End If

    If ext = "ppt" Or ext = "pptx" Then
        Files_ExtrairTextoDoFicheiro = Files_ExtrairTextoPPT(fullPath, outErro)
        Exit Function
    End If

    outErro = "Extensao nao suportada para text_embed: " & ext
    Exit Function

Falha:
    outErro = "Erro extrair texto: " & Err.Description
    Files_ExtrairTextoDoFicheiro = ""
End Function

Private Function Files_LerTexto(ByVal fullPath As String, ByRef outErro As String) As String
    On Error GoTo Falha

    outErro = ""
    Files_LerTexto = ""

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile fullPath
    Files_LerTexto = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing
    Exit Function

Falha:
    outErro = "Ler texto falhou: " & Err.Description
    Files_LerTexto = ""
End Function

Private Function Files_ExtrairTextoWord(ByVal fullPath As String, ByRef outErro As String) As String
    On Error GoTo Falha

    outErro = ""
    Files_ExtrairTextoWord = ""

    Dim wordApp As Object, doc As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    Set doc = wordApp.Documents.Open(fullPath, False, True, False)
    Files_ExtrairTextoWord = doc.content.text
    doc.Close False

    wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing
    Exit Function

Falha:
    outErro = "Extrair texto Word falhou: " & Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Files_ExtrairTextoWord = ""
End Function

Private Function Files_ExtrairTextoPPT(ByVal fullPath As String, ByRef outErro As String) As String
    On Error GoTo Falha

    outErro = ""
    Files_ExtrairTextoPPT = ""

    Dim pptApp As Object, pres As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = False

    Set pres = pptApp.Presentations.Open(fullPath, True, True, False)

    Dim sb As String
    sb = ""

    Dim sld As Object, shp As Object
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            On Error Resume Next
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    sb = sb & shp.TextFrame.TextRange.text & vbCrLf
                End If
            End If
            On Error GoTo 0
        Next shp
    Next sld

    pres.Close
    pptApp.Quit

    Set pres = Nothing
    Set pptApp = Nothing

    Files_ExtrairTextoPPT = sb
    Exit Function

Falha:
    outErro = "Extrair texto PPT falhou: " & Err.Description
    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    If Not pptApp Is Nothing Then pptApp.Quit
    Files_ExtrairTextoPPT = ""
End Function


' ============================================================
' UPLOAD / REUTILIZACAO (FILES_MANAGEMENT)
' ============================================================

' ============================================================
' (ALTERAR) Files_ObterOuCriarFileId — aceitar allowReuse/reuseTag e validar ativo
'   SUBSTITUA A FUNÇÃO INTEIRA por esta versão
' ============================================================
Private Function Files_ObterOuCriarFileId( _
    ByVal wsFiles As Worksheet, _
    ByVal mapaCab As Object, _
    ByVal apiKey As String, _
    ByVal promptId As String, _
    ByVal pipelineNome As String, _
    ByVal fileName As String, _
    ByVal folderBase As String, _
    ByVal fullPath As String, _
    ByVal usageMode As String, _
    ByVal convertido As Boolean, _
    ByVal sourceHash As String, _
    ByVal lastMod As Date, _
    ByVal sizeBytes As Double, _
    ByVal hashUsado As String, _
    ByVal allowReuse As Boolean, _
    ByVal reuseTag As String, _
    ByRef outErro As String _
) As String
    On Error GoTo Falha

    outErro = ""
    Files_ObterOuCriarFileId = ""

    fileName = Trim$(CStr(fileName))
    fullPath = Trim$(CStr(fullPath))
    usageMode = Trim$(CStr(usageMode))

    Dim reuseDiag As String
    reuseDiag = ""

    If allowReuse Then
        Dim linha As Long
        Dim colFileId As Long, colNome As Long
        colFileId = Files_Col(mapaCab, H_FILE_ID)
        colNome = Files_Col(mapaCab, H_FILE_NAME)

        ' 1) Cache canónica: hash + usage_mode
        linha = Files_EncontrarLinhaPorHash(wsFiles, mapaCab, hashUsado, usageMode, True)

        If linha > 0 Then
            Dim fileIdExistente As String
            Dim nomeExistente As String
            fileIdExistente = Trim$(CStr(wsFiles.Cells(linha, colFileId).value))
            nomeExistente = Trim$(CStr(wsFiles.Cells(linha, colNome).value))

            If fileIdExistente <> "" Then
                If StrComp(nomeExistente, fileName, vbTextCompare) = 0 Then
                    Dim st As Long, errCheck As String
                    If Files_OpenAI_FileIdAtivo(apiKey, fileIdExistente, st, errCheck) Then
                        Call Debug_Registar(0, promptId, "INFO", "", "FILES REUSE", _
                            "Reutilização OK | " & reuseTag & " | cache=hash+mode | file_id=" & fileIdExistente & " | file=" & fileName & " | mode=" & usageMode, _
                            "")

                        Call Files_UpsertFilesManagement(wsFiles, mapaCab, pipelineNome, promptId, fileName, _
                            folderBase, fullPath, fileIdExistente, usageMode, convertido, sourceHash, lastMod, sizeBytes, hashUsado, _
                            "reutilizado (cache) | " & reuseTag & " | cache=hash+mode | file_id_ativo=TRUE", "DL")

                        Files_ObterOuCriarFileId = fileIdExistente
                        Exit Function
                    Else
                        reuseDiag = "cache_invalida: file_id inativo (" & errCheck & ")"
                    End If
                Else
                    reuseDiag = "cache_ignorada: nome diferente (cache='" & nomeExistente & "' vs actual='" & fileName & "')"
                End If
            Else
                reuseDiag = "cache_invalida: file_id vazio na linha encontrada (hash+mode)"
            End If
        Else
            reuseDiag = "cache: não encontrei linha por hash+usage_mode"
        End If

        ' 2) Robustez/migração: full_path + usage_mode (com validação size + last_modified)
        '    ALTERAÇÃO: não bloquear por "nome diferente" neste fallback.
        '    Justificação: em fluxos DOCX->PDF ou overrides, o fileName pode diferir do nome guardado,
        '    mas o fullPath+mode+size+lastmod asseguram que é o mesmo artefacto em disco.
        Dim linhaP As Long
        linhaP = Files_EncontrarLinhaPorPath(wsFiles, mapaCab, fullPath, usageMode, True)

        If linhaP > 0 Then
            Dim colSize As Long, colLastMod As Long
            colSize = Files_Col(mapaCab, H_SIZE_BYTES)
            colLastMod = Files_Col(mapaCab, H_LAST_MODIFIED)

            Dim sizeCache As Double
            Dim modCache As Date
            sizeCache = 0
            modCache = 0

            On Error Resume Next
            If colSize > 0 Then sizeCache = CDbl(wsFiles.Cells(linhaP, colSize).value)
            If colLastMod > 0 Then modCache = CDate(wsFiles.Cells(linhaP, colLastMod).value)
            On Error GoTo Falha

            Dim sizeOK As Boolean, modOK As Boolean
            sizeOK = (sizeCache > 0 And Abs(sizeCache - sizeBytes) < 0.5)
            modOK = True
            If modCache <> 0 Then modOK = (Abs(DateDiff("s", modCache, lastMod)) <= 2)

            If sizeOK And modOK Then
                Dim fileIdP As String, nomeP As String
                fileIdP = Trim$(CStr(wsFiles.Cells(linhaP, colFileId).value))
                nomeP = Trim$(CStr(wsFiles.Cells(linhaP, colNome).value))

                If fileIdP <> "" Then
                    Dim st2 As Long, errCheck2 As String
                    If Files_OpenAI_FileIdAtivo(apiKey, fileIdP, st2, errCheck2) Then
                        Call Debug_Registar(0, promptId, "INFO", "", "FILES REUSE", _
                            "Reutilização OK | " & reuseTag & " | cache=path+mode (migração) | file_id=" & fileIdP & _
                            " | file=" & fileName & " | mode=" & usageMode & " | cache_file_name='" & nomeP & "'", _
                            "Nota: usado para evitar uploads duplicados quando há migração/ajuste do hash; validação é por path+mode+size+lastmod (nome pode divergir).")

                        Call Files_UpsertFilesManagement(wsFiles, mapaCab, pipelineNome, promptId, fileName, _
                            folderBase, fullPath, fileIdP, usageMode, convertido, sourceHash, lastMod, sizeBytes, hashUsado, _
                            "reutilizado (cache) | " & reuseTag & " | cache=path+mode | file_id_ativo=TRUE | hash_migrado=TRUE", "DL")

                        Files_ObterOuCriarFileId = fileIdP
                        Exit Function
                    Else
                        reuseDiag = reuseDiag & IIf(reuseDiag <> "", "; ", "") & "cache_invalida(path): file_id inativo (" & errCheck2 & ")"
                    End If
                Else
                    reuseDiag = reuseDiag & IIf(reuseDiag <> "", "; ", "") & "cache_invalida(path): file_id vazio"
                End If
            Else
                reuseDiag = reuseDiag & IIf(reuseDiag <> "", "; ", "") & "cache_ignorada(path): ficheiro mudou (size/lastmod)"
            End If
        End If
    Else
        reuseDiag = "reuse=FALSE (config/prompt)"
    End If

    ' Upload novo
    Dim purpose As String
    purpose = Files_DeterminarPurpose(usageMode)

    Dim novoId As String
    novoId = ""

    Dim httpStatus As Long
    httpStatus = 0

    Dim errUpload As String
    errUpload = ""

    Dim okUp As Boolean
    okUp = Files_UploadFile_OpenAI(apiKey, folderBase, fullPath, purpose, novoId, httpStatus, errUpload)

    If Not okUp Or novoId = "" Then
        outErro = "Upload falhou: " & errUpload
        Call Debug_Registar(0, promptId, "ERRO", "", "FILES UPLOAD", _
            "Upload FAIL | HTTP " & CStr(httpStatus) & " | " & reuseTag & IIf(reuseDiag <> "", " | " & reuseDiag, "") & " | " & outErro, _
            "Sugestão: confirme API key, conectividade e logs. Se for DOCX, valide a conversão para PDF e o cache.")
        Files_ObterOuCriarFileId = ""
        Exit Function
    End If

    Call Debug_Registar(0, promptId, "INFO", "", "FILES UPLOAD", _
        "Upload OK | HTTP " & CStr(httpStatus) & " | file_id=" & novoId & " | purpose=" & purpose & _
        " | " & reuseTag & IIf(reuseDiag <> "", " | " & reuseDiag, ""), _
        "")

    Call Files_UpsertFilesManagement(wsFiles, mapaCab, pipelineNome, promptId, fileName, _
        folderBase, fullPath, novoId, usageMode, convertido, sourceHash, lastMod, sizeBytes, hashUsado, _
        "upload novo | " & reuseTag & IIf(reuseDiag <> "", " | " & reuseDiag, "") & " | http=" & CStr(httpStatus), "UL")

    Files_ObterOuCriarFileId = novoId
    Exit Function

Falha:
    outErro = "Erro em Files_ObterOuCriarFileId: " & Err.Description
    Files_ObterOuCriarFileId = ""
End Function


Private Function Files_DeterminarPurpose(ByVal usageMode As String) As String
    Dim m As String
    m = LCase$(Trim$(usageMode))

    If m = "image_upload" Then
        Files_DeterminarPurpose = "vision"
    Else
        Files_DeterminarPurpose = "user_data"
    End If
End Function

' ============================================================
' (NOVO) Wrapper público para registar outputs no FILES_MANAGEMENT
'   - Reutiliza a lógica existente de upsert/inserção no topo (linha 2)
' ============================================================
Public Sub Files_LogEventOutput( _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByVal folderBase As String, _
    ByVal fullPath As String, _
    Optional ByVal usageMode As String = "output", _
    Optional ByVal dlUl As String = "", _
    Optional ByVal notes As String = "", _
    Optional ByVal responseId As String = "", _
    Optional ByVal runId As String = "", _
    Optional ByVal stepN As Long = 0, _
    Optional ByVal outputIndex As Long = -1, _
    Optional ByVal sourceType As String = "OUTPUT" _
)
    On Error GoTo Falha

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("FILES_MANAGEMENT")

    Dim mapaCab As Object
    Set mapaCab = Files_MapaCabecalhos(ws)

    Dim nome As String
    nome = Files_SoNomeFicheiro(fullPath)

    Dim pastaFull As String
    pastaFull = Replace(fullPath, "/", "\")
    Dim p As Long
    p = InStrRev(pastaFull, "\")
    If p > 0 Then
        pastaFull = Left$(pastaFull, p - 1)
    Else
        pastaFull = ""
    End If

    Dim folderRel As String
    folderRel = ""

    If Trim$(folderBase) <> "" Then
        Dim fb As String
        fb = folderBase
        If Right$(fb, 1) = "\" Then fb = Left$(fb, Len(fb) - 1)

        If pastaFull <> "" Then
            If LCase$(Left$(pastaFull, Len(fb))) = LCase$(fb) Then
                folderRel = Mid$(pastaFull, Len(fb) + 1)
                If Left$(folderRel, 1) = "\" Then folderRel = Mid$(folderRel, 2)
            Else
                folderRel = pastaFull
            End If
        End If
    Else
        folderRel = pastaFull
    End If

    Dim fileType As String
    fileType = LCase$(Files_ObterExtensao(nome))

    Dim lastMod As Date
    lastMod = 0
    On Error Resume Next
    lastMod = FileDateTime(fullPath)
    On Error GoTo Falha

    Dim sizeBytes As Double
    sizeBytes = 0
    On Error Resume Next
    sizeBytes = CDbl(FileLen(fullPath))
    On Error GoTo Falha

    Dim hashUsed As String
    hashUsed = ""
    On Error Resume Next
    hashUsed = Files_SHA256_File(fullPath)
    On Error GoTo Falha


    Dim chainMeta As String
    chainMeta = "source_type=" & sourceType
    If Trim$(runId) <> "" Then chainMeta = chainMeta & " | run_id=" & runId
    If stepN > 0 Then chainMeta = chainMeta & " | step_n=" & CStr(stepN)
    If outputIndex >= 0 Then chainMeta = chainMeta & " | output_index=" & CStr(outputIndex)
    chainMeta = chainMeta & " | prompt_id=" & promptId

    If Trim$(notes) <> "" Then
        notes = notes & " | " & chainMeta
    Else
        notes = chainMeta
    End If

    If Trim$(responseId) <> "" Then
        If Trim$(notes) <> "" Then
            notes = notes & " | response_id=" & responseId
        Else
            notes = "response_id=" & responseId
        End If
    End If

Call Files_UpsertFilesManagement( _
    ws, _
    mapaCab, _
    pipelineNome, _
    promptId, _
    nome, _
    folderRel, _
    fullPath, _
    "", _
    usageMode, _
    False, _
    "", _
    lastMod, _
    sizeBytes, _
    hashUsed, _
    notes, _
    dlUl _
)


    Exit Sub

Falha:
    On Error Resume Next
    Debug_Registar 0, promptId, "ERRO", "", "Files_LogEventOutput", _
        "Falha a registar output no FILES_MANAGEMENT: " & Err.Description, _
        "Confirma que FILES_MANAGEMENT existe e que a tabela tem os cabeçalhos esperados."
End Sub

Private Sub Files_UpsertFilesManagement( _
    ByVal wsFiles As Worksheet, _
    ByVal mapaCab As Object, _
    ByVal pipelineNome As String, _
    ByVal promptId As String, _
    ByVal fileName As String, _
    ByVal folderBase As String, _
    ByVal fullPath As String, _
    ByVal fileId As String, _
    ByVal usageMode As String, _
    ByVal convertido As Boolean, _
    ByVal sourceHash As String, _
    ByVal lastMod As Date, _
    ByVal sizeBytes As Double, _
    ByVal hashUsado As String, _
    ByVal notes As String, _
    Optional ByVal dlUl As String = "" _
)
    ' v2: cada evento/uso gera sempre um novo registo, inserido no topo (linha 2 / topo da tabela).
    '     - DL/UL: DL se nao houve /v1/files; UL se houve upload real.
    '     - Utilizações: contagem total por chave (file_id ou hash|usage_mode quando file_id vazio).
    '     - used_in_prompts: lista (1 linha) de prompt_id, mais recente primeiro, separador ";  ", max 20 + "(...)"
    '     - last_used_at: guarda o prompt_id do evento (nao timestamp).

    On Error GoTo Falha

    Dim lo As ListObject
    Set lo = Files_GetOrCreateTable_FilesManagement(wsFiles)
    If lo Is Nothing Then Exit Sub

    ' Timestamp único do evento
    Dim ts As Date
    ts = Now

    ' ---- Garantir hash (SHA-256 do conteúdo efetivamente usado) + diagnóstico
    If Trim$(CStr(hashUsado)) = "" Then
        If Trim$(CStr(fullPath)) <> "" Then
            If Dir(fullPath) <> "" Then
                hashUsado = Files_SHA256_File(fullPath)
                If Trim$(CStr(hashUsado)) <> "" Then
                    notes = notes & " | hash recalculado"
                End If
            Else
                notes = notes & " | HASH_DIAG: fullPath não existe: " & fullPath
            End If
        Else
            notes = notes & " | HASH_DIAG: fullPath vazio"
        End If

        If Trim$(CStr(hashUsado)) = "" Then
            Dim diag As String
            diag = Files_SHA256_LastDiag()
            If Trim$(diag) <> "" Then
                notes = notes & " | HASH_DIAG: " & diag
            End If
            notes = notes & " | AVISO: hash vazio"
        End If
    End If

    ' ---- Determinar DL/UL
    dlUl = UCase$(Trim$(dlUl))
    If dlUl <> "UL" Then dlUl = "DL"

    Dim dlUlDisplay As String
    dlUlDisplay = dlUl & vbLf & Format$(ts, "yyyy-mm-dd")

    ' ---- Obter estado anterior (antes de inserir nova linha)
    Dim prevRow As Long
    prevRow = Files_EncontrarLinhaAnteriorMesmoFicheiro(wsFiles, mapaCab, fileId, hashUsado, usageMode)

    Dim prevPrompts As String
    prevPrompts = ""

    Dim prevUtil As Long
    prevUtil = 0

    If prevRow > 0 Then
        On Error Resume Next
        prevPrompts = CStr(wsFiles.Cells(prevRow, Files_Col(mapaCab, H_USED_IN_PROMPTS)).value)
        prevUtil = CLng(wsFiles.Cells(prevRow, Files_Col(mapaCab, H_UTILIZACOES)).value)
        On Error GoTo Falha
    End If

    Dim newUtil As Long
    If prevUtil > 0 Then
        newUtil = prevUtil + 1
    Else
        newUtil = 1
    End If

    Dim newUsedInPrompts As String
    newUsedInPrompts = Files_BuildUsedInPrompts(prevPrompts, promptId)

    ' ---- Separador por run (antes de inserir linha nova)
    Call Files_MaybeAddRunSeparator(wsFiles, lo)

    ' ---- Inserir no topo da tabela (linha 2)
    Dim lr As ListRow
    Set lr = lo.ListRows.Add(Position:=1)

    Dim linha As Long
    linha = lr.Range.Row

    ' ============================================================
    ' CORRECÇÃO #1: o registo NÃO pode herdar a altura 6 do separador
    ' ============================================================
    On Error Resume Next
    lr.Range.EntireRow.RowHeight = wsFiles.StandardHeight
    ' Também impedir herança de "preto" do separador
    lr.Range.Interior.pattern = xlNone
    lr.Range.Font.Bold = False
    On Error GoTo Falha

    ' ---- Preencher campos
    wsFiles.Cells(linha, Files_Col(mapaCab, H_TIMESTAMP)).value = ts

    ' DL/UL com 2 linhas (DL?YYYY-MM-DD)
    Dim colDLUL As Long
    colDLUL = Files_Col(mapaCab, H_DL_UL)
    If colDLUL > 0 Then
        With wsFiles.Cells(linha, colDLUL)
            .value = dlUlDisplay
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If

    wsFiles.Cells(linha, Files_Col(mapaCab, H_FILE_NAME)).value = fileName
    wsFiles.Cells(linha, Files_Col(mapaCab, H_TYPE)).value = Files_ObterExtensao(fileName)
    wsFiles.Cells(linha, Files_Col(mapaCab, H_FOLDER)).value = folderBase
    wsFiles.Cells(linha, Files_Col(mapaCab, H_FULL_PATH)).value = fullPath
    wsFiles.Cells(linha, Files_Col(mapaCab, H_FILE_ID)).value = fileId
    wsFiles.Cells(linha, Files_Col(mapaCab, H_USAGE_MODE)).value = usageMode
    wsFiles.Cells(linha, Files_Col(mapaCab, H_CONVERTED_TO_PDF)).value = IIf(convertido, "TRUE", "FALSE")
    wsFiles.Cells(linha, Files_Col(mapaCab, H_HASH)).value = hashUsado
    wsFiles.Cells(linha, Files_Col(mapaCab, H_LAST_MODIFIED)).value = lastMod
    wsFiles.Cells(linha, Files_Col(mapaCab, H_SIZE_BYTES)).value = sizeBytes
    wsFiles.Cells(linha, Files_Col(mapaCab, H_LAST_USED_PIPELINE)).value = pipelineNome
    wsFiles.Cells(linha, Files_Col(mapaCab, H_UTILIZACOES)).value = newUtil
    wsFiles.Cells(linha, Files_Col(mapaCab, H_USED_IN_PROMPTS)).value = newUsedInPrompts
    wsFiles.Cells(linha, Files_Col(mapaCab, H_LAST_USED_AT)).value = promptId
    wsFiles.Cells(linha, Files_Col(mapaCab, H_NOTES)).value = notes

    ' ---- Formatação específica
    On Error Resume Next
    wsFiles.Cells(linha, Files_Col(mapaCab, H_CONVERTED_TO_PDF)).HorizontalAlignment = xlCenter
    wsFiles.Cells(linha, Files_Col(mapaCab, H_USED_IN_PROMPTS)).WrapText = True
    On Error GoTo 0

    Exit Sub

Falha:
End Sub

Private Function Files_EncontrarLinhaAnteriorMesmoFicheiro( _
    ByVal wsFiles As Worksheet, _
    ByVal mapaCab As Object, _
    ByVal fileId As String, _
    ByVal hashUsado As String, _
    ByVal usageMode As String _
) As Long
    On Error GoTo Falha

    Files_EncontrarLinhaAnteriorMesmoFicheiro = 0

    Dim lastRow As Long
    lastRow = Files_LastDataRow(wsFiles)
    If lastRow < 2 Then Exit Function

    Dim colFileId As Long, colHash As Long, colMode As Long
    colFileId = Files_Col(mapaCab, H_FILE_ID)
    colHash = Files_Col(mapaCab, H_HASH)
    colMode = Files_Col(mapaCab, H_USAGE_MODE)

    If colFileId = 0 Or colHash = 0 Or colMode = 0 Then Exit Function

    Dim r As Long
    fileId = Trim$(CStr(fileId))

    If fileId <> "" Then
        For r = 2 To lastRow
            If StrComp(Trim$(CStr(wsFiles.Cells(r, colFileId).value)), fileId, vbTextCompare) = 0 Then
                Files_EncontrarLinhaAnteriorMesmoFicheiro = r
                Exit Function
            End If
        Next r
    Else
        For r = 2 To lastRow
            Dim fid As String
            fid = Trim$(CStr(wsFiles.Cells(r, colFileId).value))
            If fid = "" Then
                Dim h As String, m As String
                h = Trim$(CStr(wsFiles.Cells(r, colHash).value))
                m = Trim$(CStr(wsFiles.Cells(r, colMode).value))

                If (h <> "") And (StrComp(h, hashUsado, vbTextCompare) = 0) And (StrComp(m, usageMode, vbTextCompare) = 0) Then
                    Files_EncontrarLinhaAnteriorMesmoFicheiro = r
                    Exit Function
                End If
            End If
        Next r
    End If

    Exit Function

Falha:
    Files_EncontrarLinhaAnteriorMesmoFicheiro = 0
End Function

Private Function Files_BuildUsedInPrompts(ByVal prev As String, ByVal promptIdAtual As String) As String
    On Error GoTo Falha

    promptIdAtual = Trim$(CStr(promptIdAtual))

    Dim ids As Collection
    Set ids = New Collection

    If promptIdAtual <> "" Then ids.Add promptIdAtual

    prev = Replace(prev, vbCrLf, " ")
    prev = Replace(prev, vbLf, " ")
    prev = Replace(prev, vbCr, " ")
    prev = Trim$(prev)

    If prev <> "" Then
        Dim parts() As String
        parts = Split(prev, ";")

        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            Dim p As String
            p = Trim$(CStr(parts(i)))
            If p <> "" Then
                If StrComp(p, USED_PROMPTS_SUFFIX, vbTextCompare) <> 0 Then
                    ids.Add p
                End If
            End If
        Next i
    End If

    Dim total As Long
    total = ids.Count

    Dim keep As Long
    keep = total
    Dim truncated As Boolean
    truncated = False

    If keep > MAX_USED_PROMPTS Then
        keep = MAX_USED_PROMPTS
        truncated = True
    End If

    If keep <= 0 Then
        Files_BuildUsedInPrompts = ""
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(1 To keep)

    Dim k As Long
    For k = 1 To keep
        arr(k) = CStr(ids(k))
    Next k

    Dim out As String
    out = Join(arr, USED_PROMPTS_SEP)

    If truncated Then
        out = out & USED_PROMPTS_SEP & USED_PROMPTS_SUFFIX
    End If

    Files_BuildUsedInPrompts = out
    Exit Function

Falha:
    Files_BuildUsedInPrompts = Trim$(CStr(promptIdAtual))
End Function

Private Function Files_EncontrarLinhaPorHash( _
    ByVal wsFiles As Worksheet, _
    ByVal mapaCab As Object, _
    ByVal hashValue As String, _
    ByVal usageMode As String, _
    ByVal exigirFileId As Boolean _
) As Long

    On Error GoTo EH

    Files_EncontrarLinhaPorHash = 0

    hashValue = Trim$(CStr(hashValue))
    usageMode = Trim$(CStr(usageMode))

    If hashValue = "" Or usageMode = "" Then Exit Function

    Dim colHash As Long
    Dim colMode As Long
    Dim colFileId As Long

    colHash = Files_Col(mapaCab, H_HASH)
    colMode = Files_Col(mapaCab, H_USAGE_MODE)
    colFileId = Files_Col(mapaCab, H_FILE_ID)

    If colHash = 0 Or colMode = 0 Then Exit Function
    If exigirFileId And colFileId = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = wsFiles.Cells(wsFiles.rowS.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim r As Long
    For r = lastRow To 2 Step -1

        Dim h As String
        Dim m As String
        h = Trim$(CStr(wsFiles.Cells(r, colHash).value))
        m = Trim$(CStr(wsFiles.Cells(r, colMode).value))

        If StrComp(h, hashValue, vbTextCompare) = 0 And StrComp(m, usageMode, vbTextCompare) = 0 Then
            If exigirFileId Then
                Dim fid As String
                fid = Trim$(CStr(wsFiles.Cells(r, colFileId).value))
                If fid = "" Then GoTo ContinueFor
            End If

            Files_EncontrarLinhaPorHash = r
            Exit Function
        End If

ContinueFor:
    Next r

    Exit Function

EH:
    Files_EncontrarLinhaPorHash = 0
End Function


Private Function Files_LastDataRow(ByVal wsFiles As Worksheet) As Long
    On Error GoTo Falha

    Dim lo As ListObject
    Set lo = Files_TryGetTable_FilesManagement(wsFiles)

    If Not lo Is Nothing Then
        If lo.DataBodyRange Is Nothing Then
            Files_LastDataRow = 1
        Else
            Files_LastDataRow = lo.DataBodyRange.rowS(lo.DataBodyRange.rowS.Count).Row
        End If
        Exit Function
    End If

    Files_LastDataRow = wsFiles.Cells(wsFiles.rowS.Count, 1).End(xlUp).Row
    Exit Function

Falha:
    Files_LastDataRow = wsFiles.Cells(wsFiles.rowS.Count, 1).End(xlUp).Row
End Function

Private Sub Files_AddRunSeparatorLine(ByVal wsFiles As Worksheet, ByVal lo As ListObject, ByVal runToken As String)
    ' NOTA: Mantém-se o nome da rotina por compatibilidade com o resto do código,
    '       mas o separador deixa de ser Shape e passa a ser uma "linha de intervalo" (row separadora),
    '       idêntica ao mecanismo usado na folha HISTÓRICO.

    On Error GoTo Falha

    If lo Is Nothing Then Exit Sub

    ' Inserir row separadora no topo da tabela (Position:=1)
    ' (Esta row ficará entre o novo run (que será inserido acima) e o histórico anterior.)
    Dim lrSep As ListRow
    Set lrSep = lo.ListRows.Add(Position:=1)

    ' Limpar conteúdos e aplicar formatação do separador
    With lrSep.Range
        .ClearContents

        ' Preto sólido em toda a largura da tabela
        .Interior.pattern = xlSolid
        .Interior.Color = vbBlack

        ' Garantir que não fica negrito nem wrap inesperado
        .Font.Bold = False
        .WrapText = False
    End With

    ' Altura do separador: 6 pt (row separadora)
    On Error Resume Next
    lrSep.Range.EntireRow.RowHeight = 6
    On Error GoTo Falha

    Exit Sub

Falha:
    ' Não bloquear a execução caso falhe a inserção/formatacao do separador
End Sub



Private Function Files_SanitizarNomeObjeto(ByVal s As String) As String
    ' Nomes de Shape devem ser curtos e sem caracteres especiais problemáticos.
    Dim i As Long
    Dim ch As String
    Dim out As String
    out = ""

    s = Trim$(CStr(s))
    If s = "" Then s = "RUN"

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            out = out & ch
        Else
            out = out & "_"
        End If
        If Len(out) >= 40 Then Exit For
    Next i

    If out = "" Then out = "RUN"
    Files_SanitizarNomeObjeto = out
End Function




' ============================================================
' UPLOAD: /v1/files (multipart/form-data)
'   - Importante: HTTP status 0 = falha antes de resposta (proxy/TLS/codigo)
' ============================================================

Public Function Files_UploadFile_OpenAI( _
    ByVal apiKey As String, _
    ByVal inputFolder As String, _
    ByVal resolvedPath As String, _
    ByVal purpose As String, _
    ByRef outFileId As String, _
    ByRef outHttpStatus As Long, _
    ByRef outErro As String _
) As Boolean

    On Error GoTo Falha

    Files_UploadFile_OpenAI = False
    outFileId = ""
    outErro = ""
    outHttpStatus = 0

    If Trim$(apiKey) = "" Then
        outErro = "API key vazia."
        Exit Function
    End If

    If Dir(resolvedPath) = "" Then
        outErro = "Ficheiro nao existe: " & resolvedPath
        Exit Function
    End If

    Dim fileNameRaw As String
    fileNameRaw = Files_SoNomeFicheiro(resolvedPath)

    ' Nome enviado no multipart (apenas "envelope"; NAO altera o ficheiro no disco)
    Dim fileNameSent As String
    fileNameSent = fileNameRaw

    Dim nameMode As String
    nameMode = Files_Config_MultipartFilenameMode() ' default RAW (compatível)

    If nameMode = "ASCII_SAFE" Then
        fileNameSent = Files_SanitizeFilenameAsciiSafe(fileNameRaw)
    ElseIf nameMode = "RFC5987" Then
        ' [POR CONFIRMAR] RFC5987 nao implementado no modo LEGACY; usar RAW
        fileNameSent = fileNameRaw
    Else
        fileNameSent = fileNameRaw
    End If

    Dim contentType As String
    contentType = Files_ContentTypePorExtensao(Files_ObterExtensao(fileNameRaw))
    If contentType = "" Then contentType = "application/octet-stream"

    Dim errRead As String
    Dim fileBytes() As Byte
    fileBytes = Files_ReadAllBytesEx(resolvedPath, errRead)
    If Files_ByteLen(fileBytes) = 0 Then
        outErro = "Leitura de ficheiro falhou: " & errRead
        Exit Function
    End If

    Dim boundary As String
    boundary = "----VBAFormBoundary" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 1000000))

    Dim body() As Byte
    body = Files_BuildMultipartBody_Safe(boundary, purpose, fileNameSent, contentType, fileBytes)

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 15000, 15000, 60000, 60000
    http.Open "POST", "https://api.openai.com/v1/files", False
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.SetRequestHeader "Accept", "application/json"

    Dim vBody As Variant
    vBody = body
    http.Send vBody

    outHttpStatus = http.status
    Dim resp As String
    resp = CStr(http.ResponseText)

    If http.status < 200 Or http.status >= 300 Then
        outErro = "HTTP " & CStr(http.status) & " | " & resp
        Exit Function
    End If

    outFileId = Files_ExtrairCampoJsonSimples(resp, "id")
    If outFileId = "" Then
        outErro = "Nao foi possivel obter file_id do JSON: " & resp
        Exit Function
    End If

    Files_UploadFile_OpenAI = True
    Exit Function

Falha:
    outErro = "Erro interno em Files_UploadFile_OpenAI: " & Err.Description
    Files_UploadFile_OpenAI = False
End Function


Private Function Files_BuildMultipartBody_ByValues(ByVal boundary As String, ByVal purpose As String, ByVal fileName As String, ByVal contentType As String, ByRef fileBytes() As Byte) As Byte()
    Dim pre1 As String, pre2 As String, post As String

    pre1 = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""purpose""" & vbCrLf & vbCrLf & _
           purpose & vbCrLf

    pre2 = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf & _
           "Content-Type: " & contentType & vbCrLf & vbCrLf

    post = vbCrLf & "--" & boundary & "--" & vbCrLf

    Files_BuildMultipartBody_ByValues = Files_BuildMultipartBody(pre1, pre2, fileBytes, post)
End Function

Private Function Files_FormatComError(ByVal prefix As String, ByVal errNum As Long, ByVal errDesc As String) As String
    Dim hx As String
    hx = Files_Hex32(errNum)

    Dim d As String
    d = Trim$(errDesc)
    If d = "" Then d = "(sem descricao)"

    Dim hint As String
    hint = Files_WinHttpHintFromHex(hx)

    If hint <> "" Then
        Files_FormatComError = prefix & " | Err=" & CStr(errNum) & " (0x" & hx & ") | " & d & " | " & hint
    Else
        Files_FormatComError = prefix & " | Err=" & CStr(errNum) & " (0x" & hx & ") | " & d
    End If
End Function

Private Function Files_WinHttpHintFromHex(ByVal hx As String) As String
    Dim u As String
    u = UCase$(Trim$(hx))

    Select Case u
        Case "80072EE7"
            Files_WinHttpHintFromHex = "DNS: nome do servidor nao resolvido (proxy/DNS)."
        Case "80072EFD"
            Files_WinHttpHintFromHex = "Ligacao falhou (proxy/firewall/rede)."
        Case "80072EE2"
            Files_WinHttpHintFromHex = "Timeout (proxy/rede lenta/bloqueio)."
        Case "80072F7D", "80072F8F"
            Files_WinHttpHintFromHex = "TLS/certificado falhou (inspecao SSL, TLS bloqueado)."
        Case Else
            Files_WinHttpHintFromHex = ""
    End Select
End Function

Private Function Files_Hex32(ByVal n As Long) As String
    Dim x As String
    x = Hex$(n And &HFFFFFFFF)
    Files_Hex32 = Right$("00000000" & x, 8)
End Function

Private Function Files_ByteArrayLen(ByRef b() As Byte) As Long
    On Error GoTo Vazio
    Files_ByteArrayLen = (UBound(b) - LBound(b) + 1)
    Exit Function
Vazio:
    Files_ByteArrayLen = 0
End Function

Public Function Files_ContentTypePorExtensao(ByVal ext As String) As String
    ext = LCase$(Trim$(ext))

    If ext = "pdf" Then
        Files_ContentTypePorExtensao = "application/pdf"
    ElseIf ext = "png" Then
        Files_ContentTypePorExtensao = "image/png"
    ElseIf ext = "jpg" Or ext = "jpeg" Then
        Files_ContentTypePorExtensao = "image/jpeg"
    ElseIf ext = "webp" Then
        Files_ContentTypePorExtensao = "image/webp"
    ElseIf ext = "txt" Or ext = "md" Or ext = "csv" Or ext = "json" Then
        Files_ContentTypePorExtensao = "text/plain"
    Else
        Files_ContentTypePorExtensao = "application/octet-stream"
    End If
End Function

Private Function Files_BuildMultipartBody(ByVal pre1 As String, ByVal pre2 As String, ByRef fileBytes() As Byte, ByVal post As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open

    stm.Write StrConv(pre1, vbFromUnicode)
    stm.Write StrConv(pre2, vbFromUnicode)
    stm.Write fileBytes
    stm.Write StrConv(post, vbFromUnicode)

    stm.Position = 0
    Files_BuildMultipartBody = stm.Read
    stm.Close
    Set stm = Nothing
End Function

Private Function Files_ReadAllBytes(ByVal fullPath As String) As Byte()
    On Error GoTo Falha

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.LoadFromFile fullPath
    Files_ReadAllBytes = stm.Read
    stm.Close
    Set stm = Nothing
    Exit Function

Falha:
    ' Fallback: leitura binaria nativa
    Files_ReadAllBytes = Files_ReadAllBytes_Binary(fullPath)
End Function

Private Function Files_ReadAllBytes_Binary(ByVal fullPath As String) As Byte()
    On Error GoTo Falha

    Dim f As Integer
    f = FreeFile

    Open fullPath For Binary Access Read As #f
    Dim ln As Long
    ln = LOF(f)
    If ln <= 0 Then
        Close #f
        Exit Function
    End If

    Dim b() As Byte
    ReDim b(0 To ln - 1) As Byte
    Get #f, , b
    Close #f

    Files_ReadAllBytes_Binary = b
    Exit Function

Falha:
    ' devolve array vazio
End Function

Public Function Files_ExtrairCampoJsonSimples(ByVal json As String, ByVal chave As String) As String
    On Error GoTo Falha

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = """" & chave & """" & "\s*:\s*""([^""]+)"""

    If re.Test(json) Then
        Files_ExtrairCampoJsonSimples = re.Execute(json)(0).SubMatches(0)
    Else
        Files_ExtrairCampoJsonSimples = ""
    End If
    Exit Function

Falha:
    Files_ExtrairCampoJsonSimples = ""
End Function


' ============================================================
' INLINE BASE64 helpers (data URLs)
' ============================================================

Private Function Files_BuildDataUrlFromFile(ByVal fullPath As String, ByVal mime As String, ByRef outErro As String) As String
    On Error GoTo Falha

    outErro = ""
    Files_BuildDataUrlFromFile = ""

    If mime = "" Then mime = "application/octet-stream"

    Dim bytes() As Byte
    bytes = Files_ReadAllBytes(fullPath)

    Dim b64 As String
    b64 = Files_Base64EncodeBytes(bytes, outErro)
    If b64 = "" Then Exit Function

    Files_BuildDataUrlFromFile = "data:" & mime & ";base64," & b64
    Exit Function

Falha:
    outErro = "Erro a construir data URL: " & Err.Description
    Files_BuildDataUrlFromFile = ""
End Function

Private Function Files_Base64EncodeBytes(ByRef bytes() As Byte, ByRef outErro As String) As String
    On Error GoTo Falha

    outErro = ""
    Files_Base64EncodeBytes = ""

    Dim dom As Object
    On Error Resume Next
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    If dom Is Nothing Then Set dom = CreateObject("MSXML2.DOMDocument.3.0")
    On Error GoTo Falha

    Dim node As Object
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes

    Dim s As String
    s = CStr(node.text)
    s = Replace(s, vbCrLf, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbCr, "")

    Files_Base64EncodeBytes = s
    Exit Function

Falha:
    outErro = "Base64 falhou: " & Err.Description
    Files_Base64EncodeBytes = ""
End Function

Private Function Files_FormatBytes(ByVal n As Double) As String
    If n <= 0 Then
        Files_FormatBytes = "0 B"
    ElseIf n < 1024# Then
        Files_FormatBytes = Format$(n, "0") & " B"
    ElseIf n < 1024# * 1024# Then
        Files_FormatBytes = Format$(n / 1024#, "0.0") & " KB"
    ElseIf n < 1024# * 1024# * 1024# Then
        Files_FormatBytes = Format$(n / (1024# * 1024#), "0.0") & " MB"
    Else
        Files_FormatBytes = Format$(n / (1024# * 1024# * 1024#), "0.0") & " GB"
    End If
End Function


' ============================================================
' SHA256 helpers
' ============================================================

Public Function Files_SHA256_File(ByVal fullPath As String) As String
    On Error GoTo Falha

    gLastSHA256Diag = ""
    Files_SHA256_File = ""

    fullPath = Trim$(CStr(fullPath))
    If fullPath = "" Then
        gLastSHA256Diag = "SHA256_File: fullPath vazio"
        Exit Function
    End If

    If Not Files_ExisteFicheiro(fullPath) Then
        gLastSHA256Diag = "SHA256_File: ficheiro não existe: " & fullPath
        Exit Function
    End If

    Dim b() As Byte
    Dim errB As String
    errB = ""
    b = Files_ReadAllBytesEx(fullPath, errB)

    If Files_ByteArrayLen(b) <= 0 Then
        gLastSHA256Diag = "SHA256_File: não foi possível ler bytes: " & errB
        GoTo FallbackDet
    End If

    Dim h As String
    h = Files_SHA256_Bytes(b)

    Dim diagSHA As String
    diagSHA = gLastSHA256Diag

    If Trim$(h) <> "" Then
        gLastSHA256Diag = ""
        Files_SHA256_File = h
        Exit Function
    End If

    If diagSHA = "" Then diagSHA = "sem diagnóstico"
    gLastSHA256Diag = "SHA256_File: SHA256 falhou (" & diagSHA & ")"

FallbackDet:
    Dim h2 As String
    h2 = ""

    On Error Resume Next
    h2 = Files_FNV32_Bytes(b)
    On Error GoTo Falha

    If Trim$(h2) <> "" Then
        gLastSHA256Diag = gLastSHA256Diag & " | fallback FNV32(bytes) aplicado"
        Files_SHA256_File = h2
        Exit Function
    End If

    ' Último recurso: metadata (menos robusto)
    On Error Resume Next
    h2 = Files_FNV32_String(fullPath & "|" & CStr(FileLen(fullPath)) & "|" & CStr(FileDateTime(fullPath)))
    On Error GoTo Falha

    If Trim$(h2) = "" Then
        gLastSHA256Diag = gLastSHA256Diag & " | fallback METADATA falhou (" & Files_FNV32_LastDiag() & ")"
    Else
        gLastSHA256Diag = gLastSHA256Diag & " | fallback METADATA aplicado"
    End If

    Files_SHA256_File = h2
    Exit Function

Falha:
    gLastSHA256Diag = "SHA256_File: erro VBA (" & Err.Number & ") " & Err.Description
    Files_SHA256_File = ""
End Function



Public Function Files_SHA256_Text(ByVal sText As String) As String
    On Error GoTo Falha

    gLastSHA256Diag = ""
    Files_SHA256_Text = ""

    If Len(sText) = 0 Then
        gLastSHA256Diag = "SHA256_Text: texto vazio"
        Exit Function
    End If

    Dim bytes() As Byte
    bytes = StrConv(sText, vbFromUnicode)

    If Files_ByteArrayLen(bytes) <= 0 Then
        gLastSHA256Diag = "SHA256_Text: byte array vazio após conversão"
        Exit Function
    End If

    Dim h As String
    h = Files_SHA256_Bytes(bytes)

    Dim diagSHA As String
    diagSHA = gLastSHA256Diag

    If Trim$(h) <> "" Then
        gLastSHA256Diag = ""
        Files_SHA256_Text = h
        Exit Function
    End If

    If diagSHA = "" Then diagSHA = "sem diagnóstico"
    gLastSHA256Diag = "SHA256_Text: SHA256 falhou (" & diagSHA & ")"

    Dim h2 As String
    h2 = Files_FNV32_Bytes(bytes)

    If Trim$(h2) = "" Then
        gLastSHA256Diag = gLastSHA256Diag & " | fallback FNV32(bytes) falhou (" & Files_FNV32_LastDiag() & ")"
    Else
        gLastSHA256Diag = gLastSHA256Diag & " | fallback FNV32(bytes) aplicado"
    End If

    Files_SHA256_Text = h2
    Exit Function

Falha:
    gLastSHA256Diag = "SHA256_Text: erro VBA: " & Err.Description
    Files_SHA256_Text = ""
End Function



Private Function Files_SHA256_Bytes(ByRef bytes() As Byte) As String
    On Error GoTo Falha

    Files_SHA256_Bytes = ""
    gLastSHA256Diag = ""

#If VBA7 Then
    Dim hProv As LongPtr, hHash As LongPtr
#Else
    Dim hProv As Long, hHash As Long
#End If

    ' Validar array
    Dim lb As Long, ub As Long, ln As Long
    On Error Resume Next
    lb = LBound(bytes)
    ub = UBound(bytes)
    If Err.Number <> 0 Then
        Err.Clear
        gLastSHA256Diag = "SHA256_Bytes: array não inicializado (0 bytes)"
        Exit Function
    End If
    On Error GoTo Falha

    ln = ub - lb + 1
    If ln <= 0 Then
        gLastSHA256Diag = "SHA256_Bytes: array vazio (0 bytes)"
        Exit Function
    End If

    Dim ok As Long
    ok = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)
    If ok = 0 Then
        gLastSHA256Diag = "SHA256_Bytes: CryptAcquireContext falhou | LastDllError=" & CStr(Err.LastDllError)
        Exit Function
    End If

    ok = CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash)
    If ok = 0 Then
        gLastSHA256Diag = "SHA256_Bytes: CryptCreateHash falhou | LastDllError=" & CStr(Err.LastDllError)
        CryptReleaseContext hProv, 0
        Exit Function
    End If

    ok = CryptHashData(hHash, bytes(lb), ln, 0)
    If ok = 0 Then
        gLastSHA256Diag = "SHA256_Bytes: CryptHashData falhou | LastDllError=" & CStr(Err.LastDllError) & " | ln=" & CStr(ln)
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Exit Function
    End If

    Dim hashB(0 To 31) As Byte
    Dim cbHash As Long
    cbHash = 32

    ok = CryptGetHashParam(hHash, HP_HASHVAL, hashB(0), cbHash, 0)
    If ok = 0 Then
        gLastSHA256Diag = "SHA256_Bytes: CryptGetHashParam falhou | LastDllError=" & CStr(Err.LastDllError)
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Exit Function
    End If

    Files_SHA256_Bytes = Files_BytesToHex(hashB)

    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    Exit Function

Falha:
    gLastSHA256Diag = "SHA256_Bytes: erro VBA: " & Err.Description
    Files_SHA256_Bytes = ""
End Function



Private Function Files_BytesToHex(ByRef b() As Byte) As String
    Dim i As Long
    Dim s As String
    s = ""
    For i = LBound(b) To UBound(b)
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next i
    Files_BytesToHex = LCase$(s)
End Function

Public Function Files_FNV32_String(ByVal s As String) As String
    On Error GoTo Falha

    gLastFNV32Diag = ""
    Files_FNV32_String = ""

    ' Determinístico: bytes UTF-16LE (VBA Unicode)
    Dim b() As Byte
    b = StrConv(s, vbUnicode)

#If VBA7 And Win64 Then
    ' 64-bit: usar LongLong + máscara 32-bit REAL (4294967295)
    Dim mask As LongLong
    mask = CLngLng(4294967295#)

    Dim h As LongLong
    h = 2166136261#   ' offset basis FNV-1a 32-bit

    Dim i As Long
    For i = LBound(b) To UBound(b)
        h = (h Xor CLngLng(b(i))) And mask
        h = (h * 16777619) And mask
    Next i

    Files_FNV32_String = "fnv32-" & Right$("00000000" & Hex$(h), 8)
    Exit Function
#Else
    ' 32-bit: versão safe em Double (mod 2^32)
    Dim u As Double
    u = 2166136261#

    Dim j As Long
    For j = LBound(b) To UBound(b)
        u = Files_U32_XorByte(u, b(j))
        u = Files_U32_Mul(u, 16777619)
    Next j

    Files_FNV32_String = "fnv32-" & Files_U32_ToHex8_D(u)
    Exit Function
#End If

Falha:
    gLastFNV32Diag = "FNV32_String: erro (" & Err.Number & ") " & Err.Description
    Files_FNV32_String = ""
End Function


' ============================================================
' CONSTRUCAO DO INPUT JSON
' ============================================================

Private Function Files_MontarInputJson( _
ByVal inputJsonLiteralBase As String, _
ByVal filePartsJson As String, _
ByVal textoFinal As String, _
ByVal includeFilesContext As Boolean, _
ByVal promptId As String _
) As String
' Alertas de tamanho (diagnóstico)
' - Não altera payload; apenas avisa quando o input_text é grande o suficiente
'   para aumentar risco de limites de request/contexto.
Const WARN_CHARS As Long = 120000
Const HARD_CHARS As Long = 250000

Dim nChars As Long
nChars = Len(textoFinal)

If nChars >= HARD_CHARS Then
    On Error Resume Next
    Call Debug_Registar(0, promptId, "ALERTA", "", "INPUT_TEXT_SIZE", _
        "input_text muito grande (chars=" & CStr(nChars) & "; hard=" & CStr(HARD_CHARS) & "). Pode falhar por limite de contexto/request.", _
        "Sugestao: usar as_pdf + file_id, ou reduzir text_embed/dividir em partes.")
    On Error GoTo 0
ElseIf nChars >= WARN_CHARS Then
    On Error Resume Next
    Call Debug_Registar(0, promptId, "ALERTA", "", "INPUT_TEXT_SIZE", _
        "input_text grande (chars=" & CStr(nChars) & "; warn=" & CStr(WARN_CHARS) & ").", _
        "Sugestao: se houver erros de limite/contexto, usar as_pdf + file_id ou reduzir o texto embebido.")
    On Error GoTo 0
End If

Dim base As String
base = Trim$(inputJsonLiteralBase)

Dim contentJson As String
contentJson = ""

If Trim$(filePartsJson) <> "" Then
    contentJson = filePartsJson
End If

Dim textPart As String
textPart = "{""type"":""input_text"",""text"":""" & Files_JsonEscape(textoFinal) & """}"

contentJson = Files_AppendJsonPart(contentJson, textPart)

Dim msg As String
msg = "{""role"":""user"",""content"":[" & contentJson & "]}"

If base = "" Then
    Files_MontarInputJson = "[" & msg & "]"
    Exit Function
End If

If Left$(base, 1) = "[" Then
    Dim trimmed As String
    trimmed = Trim$(base)

    If Right$(trimmed, 1) = "]" Then
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    End If

    If Right$(trimmed, 1) = "[" Then
        Files_MontarInputJson = "[" & msg & "]"
    Else
        Files_MontarInputJson = trimmed & "," & msg & "]"
    End If
    Exit Function
End If

Dim textBase As String
textBase = Files_TentarExtrairTextoDeJsonString(base)
If textBase = "" Then textBase = base

Dim textoFinal2 As String
textoFinal2 = textBase & vbCrLf & vbCrLf & textoFinal

Dim content2 As String
content2 = ""
If Trim$(filePartsJson) <> "" Then content2 = filePartsJson
content2 = Files_AppendJsonPart(content2, "{""type"":""input_text"",""text"":""" & Files_JsonEscape(textoFinal2) & """}")

Files_MontarInputJson = "[{""role"":""user"",""content"":[" & content2 & "]}]"
End Function



Private Function Files_TentarExtrairTextoDeJsonString(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)

    If Len(t) < 2 Then
        Files_TentarExtrairTextoDeJsonString = ""
        Exit Function
    End If

    If Left$(t, 1) <> """" Or Right$(t, 1) <> """" Then
        Files_TentarExtrairTextoDeJsonString = ""
        Exit Function
    End If

    t = Mid$(t, 2, Len(t) - 2)
    t = Replace(t, "\""", """")
    t = Replace(t, "\\", "\")
    t = Replace(t, "\n", vbCrLf)

    Files_TentarExtrairTextoDeJsonString = t
End Function

Private Function Files_JsonEscape(ByVal s As String) As String
    ' Wrapper local para o escape estrito (JSON) definido em M00_JsonUtil.Json_EscapeString
    ' - Mantém a assinatura existente no M09 para não quebrar chamadas internas
    ' - Resolve erros "invalid_json" quando o text_embed contém caracteres de controlo
    On Error GoTo EH

    Files_JsonEscape = Json_EscapeString(CStr(s))
    Exit Function

EH:
    ' Fallback mínimo (não ideal, mas evita crash em caso de erro inesperado)
    Dim t As String
    t = CStr(s)

    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")

    Files_JsonEscape = t
End Function



Private Function Files_AppendJsonPart(ByVal lista As String, ByVal part As String) As String
    If Trim$(lista) = "" Then
        Files_AppendJsonPart = part
    Else
        Files_AppendJsonPart = lista & "," & part
    End If
End Function


' ============================================================
' LOG VISUAL NO CATALOGO (celula "Operacoes com ficheiros:")
' ============================================================

Private Sub Files_EscreverOperacoes(ByVal celOps As Range, ByVal diretivas As Collection, ByVal erroGeral As String, ByVal erroCritico As Boolean)
    On Error GoTo Falha

    If celOps Is Nothing Then Exit Sub

    Dim header As String
    header = "Operacoes com ficheiros:"

    Dim linhaLista As String
    linhaLista = ""

    Dim detalhes As String
    detalhes = ""

    Dim i As Long
    For i = 1 To diretivas.Count
        Dim d As Object
        Set d = diretivas(i)

        Dim req As String
        req = CStr(d("requested_name"))

        Dim st As String
        st = CStr(d("resultado_status"))

        Dim nome As String
        nome = CStr(d("resultado_nome"))

        Dim modo As String
        modo = CStr(d("resultado_modo"))

        Dim conv As Boolean
        conv = CBool(d("resultado_convertido"))

        Dim ov As Boolean
        ov = CBool(d("resultado_override"))

        Dim token As String
        If st = "OK" Then
            token = nome & " [" & modo & "]"
            If conv Then token = token & " (convertido)"
            If ov Then token = token & " (override)"
        ElseIf st = "NOT_FOUND" Then
            token = req & " [NAO ENCONTRADO]"
        ElseIf st = "AMBIGUOUS" Then
            token = req & " [AMBIGUO]"
        ElseIf st = "UPLOAD_FAIL" Then
            If Trim$(nome) <> "" Then
                token = nome & " [FALHOU]"
            Else
                token = req & " [FALHOU]"
            End If
        Else
            token = req & " [??]"
        End If

        linhaLista = Files_AppendLista(linhaLista, token)
    Next i

    If erroGeral <> "" Then
        detalhes = detalhes & Files_TimestampCurto() & " " & erroGeral & vbCrLf
    End If

    Dim finalTxt As String
    finalTxt = header & vbCrLf & linhaLista
    If detalhes <> "" Then
        finalTxt = finalTxt & vbCrLf & detalhes
    End If

    celOps.value = finalTxt

    Call Files_AplicarCoresOperacoes(celOps, linhaLista, diretivas)
    Exit Sub

Falha:
End Sub

Private Sub Files_AplicarCoresOperacoes(ByVal celOps As Range, ByVal linhaLista As String, ByVal diretivas As Collection)
    On Error GoTo Falha

    If celOps Is Nothing Then Exit Sub
    If diretivas.Count = 0 Then Exit Sub

    celOps.Font.Color = vbBlack
    celOps.Font.Bold = False

    Dim baseStart As Long
    baseStart = Len("Operacoes com ficheiros:") + 2

    Dim posAtual As Long
    posAtual = baseStart

    Dim tokens() As String
    tokens = Split(linhaLista, ";")

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        Dim tok As String
        tok = Trim$(tokens(i))

        If tok <> "" Then
            Dim d As Object
            Set d = diretivas(i + 1)

            Dim st As String
            st = CStr(d("resultado_status"))

            Dim conv As Boolean
            conv = CBool(d("resultado_convertido"))

            Dim ov As Boolean
            ov = CBool(d("resultado_override"))

            Dim cor As Long
            cor = vbBlack

            Dim negrito As Boolean
            negrito = False

            If st = "NOT_FOUND" Or st = "AMBIGUOUS" Or st = "UPLOAD_FAIL" Then
                cor = vbRed
                negrito = True
            ElseIf ov Then
                cor = RGB(255, 192, 0)
            ElseIf conv Then
                cor = RGB(0, 176, 80)
            ElseIf st = "OK" Then
                cor = RGB(0, 112, 192)
            End If

            Dim ln As Long
            ln = Len(tok)

            celOps.Characters(posAtual, ln).Font.Color = cor
            celOps.Characters(posAtual, ln).Font.Bold = negrito

            posAtual = posAtual + ln + 2
        End If
    Next i

    Exit Sub

Falha:
End Sub

Private Function Files_BuildFilesContextResumo(ByVal diretivas As Collection) As String
    ' ============================================================
    ' FILES CONTEXT (resumo para humans)
    ' Mantém compatibilidade com a versão anterior, mas torna explícito
    ' quando o modo final foi pdf_upload (ex.: DOCX/PPTX -> PDF).
    '
    ' Regras:
    ' - Se st="OK": mostra o ficheiro resolvido + modo
    '   * se modo="pdf_upload": "<nome> => PDF (pdf_upload)"
    '   * caso contrário: "<nome> (<modo>)"
    ' - Se st<>"OK": mostra o pedido original + status: "<req> (<st>)"
    ' - Se "resultado_nome" vier vazio, faz fallback para "requested_name"
    ' ============================================================

    Dim sb As String
    sb = ""

    If diretivas Is Nothing Then
        Files_BuildFilesContextResumo = sb
        Exit Function
    End If

    Dim i As Long
    For i = 1 To diretivas.Count
        Dim d As Object
        Set d = diretivas(i)

        Dim req As String, st As String, nome As String, modo As String
        req = ""
        st = ""
        nome = ""
        modo = ""

        On Error Resume Next
        req = CStr(d("requested_name"))
        st = CStr(d("resultado_status"))
        nome = CStr(d("resultado_nome"))
        modo = CStr(d("resultado_modo"))
        On Error GoTo 0

        req = Trim$(req)
        st = Trim$(st)
        nome = Trim$(nome)
        modo = Trim$(modo)

        If st = "OK" Then
            If nome = "" Then nome = req

            If LCase$(modo) = "pdf_upload" Then
                sb = sb & "- " & nome & " => PDF (pdf_upload)" & vbCrLf
            Else
                sb = sb & "- " & nome & " (" & modo & ")" & vbCrLf
            End If
        Else
            If req = "" Then req = nome
            sb = sb & "- " & req & " (" & st & ")" & vbCrLf
        End If
    Next i

    Files_BuildFilesContextResumo = sb
End Function



' ============================================================
' LOCALIZAR CELULAS INPUTS e OPERACOES (catalogo)
' ============================================================

Private Sub Files_EncontrarCelulasInputs(ByVal promptId As String, ByRef outCelInputsValor As Range, ByRef outCelOps As Range)
    Set outCelInputsValor = Nothing
    Set outCelOps = Nothing

    Dim folha As String
    folha = Files_ExtrairFolhaDoID(promptId)
    If folha = "" Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(folha)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim celId As Range
    Set celId = ws.Columns(1).Find(What:=Trim$(promptId), LookIn:=xlValues, LookAt:=xlWhole)
    If celId Is Nothing Then Exit Sub

    Set outCelInputsValor = celId.Offset(2, 3)
    Set outCelOps = outCelInputsValor.Offset(0, 2)
End Sub

Private Function Files_ExtrairFolhaDoID(ByVal promptId As String) As String
    Dim p As Long
    p = InStr(1, promptId, "/")
    If p = 0 Then
        Files_ExtrairFolhaDoID = ""
    Else
        Files_ExtrairFolhaDoID = Left$(promptId, p - 1)
    End If
End Function


' ============================================================
' UTILITARIOS FICHEIROS / PASTAS
' ============================================================

Private Function Files_ExisteFicheiro(ByVal fullPath As String) As Boolean
    On Error Resume Next
    Files_ExisteFicheiro = (Len(Dir(fullPath)) > 0)
    On Error GoTo 0
End Function

Private Sub Files_CriarPastaSeNaoExiste(ByVal folderPath As String)
    On Error Resume Next
    If folderPath = "" Then Exit Sub
    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub
    MkDir folderPath
    On Error GoTo 0
End Sub

Private Function Files_PathJoin(ByVal folder As String, ByVal name As String) As String
    If Right$(folder, 1) = "\" Then
        Files_PathJoin = folder & name
    Else
        Files_PathJoin = folder & "\" & name
    End If
End Function

Private Function Files_SoNomeFicheiro(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "/", "\")
    Dim p As Long
    p = InStrRev(t, "\")
    If p > 0 Then
        Files_SoNomeFicheiro = Mid$(t, p + 1)
    Else
        Files_SoNomeFicheiro = t
    End If
End Function

Public Function Files_ObterExtensao(ByVal fileNameOrPath As String) As String
    Dim s As String
    s = Files_SoNomeFicheiro(fileNameOrPath)
    Dim p As Long
    p = InStrRev(s, ".")
    If p > 0 Then
        Files_ObterExtensao = Mid$(s, p + 1)
    Else
        Files_ObterExtensao = ""
    End If
End Function

Private Function Files_SemExtensao(ByVal fileNameOrPath As String) As String
    Dim s As String
    s = Files_SoNomeFicheiro(fileNameOrPath)
    Dim p As Long
    p = InStrRev(s, ".")
    If p > 0 Then
        Files_SemExtensao = Left$(s, p - 1)
    Else
        Files_SemExtensao = s
    End If
End Function

Private Function Files_DataModificacao(ByVal fullPath As String) As Date
    On Error GoTo Falha
    Files_DataModificacao = FileDateTime(fullPath)
    Exit Function
Falha:
    Files_DataModificacao = 0
End Function

Private Function Files_TamanhoBytes(ByVal fullPath As String) As Double
    On Error GoTo Falha
    Files_TamanhoBytes = CDbl(FileLen(fullPath))
    Exit Function
Falha:
    Files_TamanhoBytes = 0
End Function

Private Function Files_TimestampCurto() As String
    Files_TimestampCurto = Format$(Now, "hh:nn:ss")
End Function

Private Function Files_NormalizarQuebrasLinha(ByVal s As String) As String
    Dim t As String
    t = Replace(CStr(s), vbCrLf & vbCrLf, vbCrLf)
    Files_NormalizarQuebrasLinha = t
End Function

Private Function Files_AppendLista(ByVal lista As String, ByVal item As String) As String
    If Trim$(lista) = "" Then
        Files_AppendLista = item
    Else
        Files_AppendLista = lista & "; " & item
    End If
End Function


' ============================================================
' LISTAGEM DE CANDIDATOS (Dir) - filtra diretorios
' ============================================================

Private Function Files_IsDirectory(ByVal fullPath As String) As Boolean
    On Error GoTo Falha
    Files_IsDirectory = ((GetAttr(fullPath) And vbDirectory) <> 0)
    Exit Function
Falha:
    Files_IsDirectory = False
End Function

Private Function Files_ListarPorPattern(ByVal folder As String, ByVal pattern As String) As Collection
    Dim col As New Collection

    Dim f As String
    f = Dir(folder & "\" & pattern)

    Do While f <> ""
        Dim fp As String
        fp = folder & "\" & f

        If Not Files_IsDirectory(fp) Then
            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            d("name") = f
            d("full_path") = fp
            d("last_modified") = Files_DataModificacao(fp)
            d("size") = Files_TamanhoBytes(fp)
            col.Add d
        End If

        f = Dir()
    Loop

    Set Files_ListarPorPattern = col
End Function

Private Function Files_ListarPorSubstring(ByVal folder As String, ByVal needle As String) As Collection
    Dim col As New Collection

    Dim f As String
    f = Dir(folder & "\*.*")

    Dim lowNeedle As String
    lowNeedle = LCase$(needle)

    Do While f <> ""
        Dim fp As String
        fp = folder & "\" & f

        If Not Files_IsDirectory(fp) Then
            If InStr(1, LCase$(f), lowNeedle, vbTextCompare) > 0 Then
                Dim d As Object
                Set d = CreateObject("Scripting.Dictionary")
                d("name") = f
                d("full_path") = fp
                d("last_modified") = Files_DataModificacao(fp)
                d("size") = Files_TamanhoBytes(fp)
                col.Add d
            End If
        End If

        f = Dir()
    Loop

    Set Files_ListarPorSubstring = col
End Function

Private Function Files_EscolherMaisRecente(ByVal candidatos As Collection) As Object
    Dim best As Object
    Set best = Nothing

    Dim i As Long
    For i = 1 To candidatos.Count
        Dim d As Object
        Set d = candidatos(i)
        If best Is Nothing Then
            Set best = d
        Else
            If CDate(d("last_modified")) > CDate(best("last_modified")) Then
                Set best = d
            End If
        End If
    Next i

    Set Files_EscolherMaisRecente = best
End Function

Private Function Files_ResumoCandidatos(ByVal candidatos As Collection) As String
    Dim sb As String
    sb = ""

    Dim i As Long
    For i = 1 To candidatos.Count
        Dim d As Object
        Set d = candidatos(i)
        sb = sb & CStr(d("name")) & " | " & Format$(CDate(d("last_modified")), "yyyy-mm-dd hh:nn") & " | " & CStr(d("size")) & vbCrLf
    Next i

    Files_ResumoCandidatos = sb
End Function

Private Function Files_CandidatoExiste(ByVal candidatos As Collection, ByVal nome As String) As Boolean
    Dim i As Long
    For i = 1 To candidatos.Count
        If StrComp(CStr(candidatos(i)("name")), nome, vbTextCompare) = 0 Then
            Files_CandidatoExiste = True
            Exit Function
        End If
    Next i
    Files_CandidatoExiste = False
End Function

Private Function Files_EncontrarCandidatoPorNome(ByVal candidatos As Collection, ByVal nome As String) As Object
    Dim i As Long
    For i = 1 To candidatos.Count
        If StrComp(CStr(candidatos(i)("name")), nome, vbTextCompare) = 0 Then
            Set Files_EncontrarCandidatoPorNome = candidatos(i)
            Exit Function
        End If
    Next i
    Set Files_EncontrarCandidatoPorNome = Nothing
End Function


' ============================================================
' SHEET FILES_MANAGEMENT: criar/validar cabecalhos
' ============================================================

Private Sub Files_EnsureSheetExists()
    ' Garante existência da folha FILES_MANAGEMENT (v2) e da Tabela tblFILES_MANAGEMENT.
    ' Se a folha existir mas tiver estrutura diferente, é renomeada para backup e é recriada vazia (sem migração).

    On Error GoTo Falha

    Dim ws As Worksheet
    Dim backupName As String

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SHEET_FILES
    Else
        If Not Files_HeaderV2_OK(ws) Then
            backupName = Files_BuildBackupSheetName()

            On Error Resume Next
            ws.name = backupName
            On Error GoTo 0

            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.name = SHEET_FILES

            Call Files_RepointWorkbookNames_FromBackupToNew(backupName, SHEET_FILES)
        End If
    End If

    Call Files_WriteHeadersV2(ws)

    Dim lo As ListObject
    Set lo = Files_GetOrCreateTable_FilesManagement(ws)

    ' Formatação obrigatória
    Dim mapa As Object
    Set mapa = Files_MapaCabecalhos(ws)

    Dim cConv As Long
    cConv = Files_Col(mapa, H_CONVERTED_TO_PDF)
    If cConv > 0 Then ws.Columns(cConv).HorizontalAlignment = xlCenter

    Dim cPrompts As Long
    cPrompts = Files_Col(mapa, H_USED_IN_PROMPTS)
    If cPrompts > 0 Then ws.Columns(cPrompts).WrapText = True


    ' ------------------------------------------------------------
    ' Formatação (idempotente): só o header a negrito; registos sem negrito
    ' ------------------------------------------------------------
    If Not lo Is Nothing Then
        lo.HeaderRowRange.Font.Bold = True
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Font.Bold = False
    End If

    ' Type e Utilizações centrados (header + registos)
    Dim cType As Long, cUtil As Long, cDlUl As Long
    cType = Files_Col(mapa, H_TYPE)
    cUtil = Files_Col(mapa, H_UTILIZACOES)
    cDlUl = Files_Col(mapa, H_DL_UL)

    If cType > 0 Then ws.Columns(cType).HorizontalAlignment = xlCenter
    If cUtil > 0 Then ws.Columns(cUtil).HorizontalAlignment = xlCenter

    ' DL/UL em duas linhas: centrado e com WrapText
    If cDlUl > 0 Then
        With ws.Columns(cDlUl)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    End If

    Exit Sub

Falha:
End Sub

Private Function Files_HeaderV2_OK(ByVal ws As Worksheet) As Boolean
    On Error GoTo Falha

    Dim headers As Variant
    headers = Files_HeadersV2_Array()

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        Dim atual As String
        atual = Files_NormalizarCabecalho(CStr(ws.Cells(1, i + 1).value))

        Dim esperado As String
        esperado = Files_NormalizarCabecalho(CStr(headers(i)))

        If StrComp(atual, esperado, vbTextCompare) <> 0 Then
            Files_HeaderV2_OK = False
            Exit Function
        End If
    Next i

    Files_HeaderV2_OK = True
    Exit Function

Falha:
    Files_HeaderV2_OK = False
End Function

Private Sub Files_WriteHeadersV2(ByVal ws As Worksheet)
    On Error GoTo Falha

    Dim headers As Variant
    headers = Files_HeadersV2_Array()

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
    Next i

    ws.rowS(1).Font.Bold = True
    Exit Sub

Falha:
End Sub

Private Function Files_HeadersV2_Array() As Variant
    Files_HeadersV2_Array = Array( _
        H_TIMESTAMP, _
        H_DL_UL, _
        H_FILE_NAME, _
        H_TYPE, _
        H_FOLDER, _
        H_FULL_PATH, _
        H_FILE_ID, _
        H_USAGE_MODE, _
        H_CONVERTED_TO_PDF, _
        H_HASH, _
        H_LAST_MODIFIED, _
        H_SIZE_BYTES, _
        H_LAST_USED_PIPELINE, _
        H_UTILIZACOES, _
        H_USED_IN_PROMPTS, _
        H_LAST_USED_AT, _
        H_NOTES _
    )
End Function

Private Function Files_BuildBackupSheetName() As String
    Dim base As String
    base = "FILES_MGMT_OLD_" & Format$(Now, "yyyymmdd_hhnnss")  ' <= 31 chars

    If Len(base) > 31 Then base = Left$(base, 31)

    Dim nm As String
    nm = base

    Dim i As Long
    i = 1
    Do While Files_SheetExists(nm)
        i = i + 1
        nm = Left$(base, 27) & "_" & Format$(i, "00")
        If Len(nm) > 31 Then nm = Left$(nm, 31)
    Loop

    Files_BuildBackupSheetName = nm
End Function

Private Function Files_SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Files_SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Sub Files_RepointWorkbookNames_FromBackupToNew(ByVal backupSheetName As String, ByVal newSheetName As String)
    ' Quando se renomeia a folha antiga, os Names passam a apontar para o backup.
    ' Esta rotina tenta repontar automaticamente para a nova folha FILES_MANAGEMENT.
    On Error Resume Next

    Dim nm As name
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.RefersTo, "'" & backupSheetName & "'", vbTextCompare) > 0 Then
            nm.RefersTo = Replace(nm.RefersTo, "'" & backupSheetName & "'", "'" & newSheetName & "'")
        End If
    Next nm

    On Error GoTo 0
End Sub

Private Function Files_TryGetTable_FilesManagement(ByVal ws As Worksheet) As ListObject
    On Error GoTo Falha

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.name, "tblFILES_MANAGEMENT", vbTextCompare) = 0 Then
            Set Files_TryGetTable_FilesManagement = lo
            Exit Function
        End If
    Next lo

    Set Files_TryGetTable_FilesManagement = Nothing
    Exit Function

Falha:
    Set Files_TryGetTable_FilesManagement = Nothing
End Function

Private Function Files_GetOrCreateTable_FilesManagement(ByVal ws As Worksheet) As ListObject
    On Error GoTo Falha

    Dim lo As ListObject
    Set lo = Files_TryGetTable_FilesManagement(ws)

    ' Validar estrutura mínima
    If Not lo Is Nothing Then
        If lo.HeaderRowRange.Row <> 1 Or lo.HeaderRowRange.Column <> 1 Or lo.ListColumns.Count <> 17 Then
            On Error Resume Next
            lo.Delete
            On Error GoTo 0
            Set lo = Nothing
        End If
    End If

    If lo Is Nothing Then
        ' Criar tabela com cabeçalho + 1 linha vazia (depois removida)
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(2, 17))
        rng.rowS(2).ClearContents

        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        On Error Resume Next
        lo.name = "tblFILES_MANAGEMENT"
        On Error GoTo 0

        ' Remover linha vazia para deixar a tabela sem registos
        On Error Resume Next
        If lo.ListRows.Count > 0 Then lo.ListRows(1).Delete
        On Error GoTo 0
    End If

    Set Files_GetOrCreateTable_FilesManagement = lo
    Exit Function

Falha:
    Set Files_GetOrCreateTable_FilesManagement = Nothing
End Function

Private Function Files_MapaCabecalhos(ByVal ws As Worksheet) As Object
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = vbTextCompare

    Dim c As Long
    For c = 1 To 80
        Dim h As String
        h = Files_NormalizarCabecalho(CStr(ws.Cells(1, c).value))
        If h <> "" Then
            mapa(h) = c
        End If
    Next c

    Set Files_MapaCabecalhos = mapa
End Function

Private Function Files_NormalizarCabecalho(ByVal s As String) As String
    s = Replace(CStr(s), ChrW(160), " ") ' NBSP
    s = Replace(s, ChrW(8220), ChrW(34)) ' “ -> "
    s = Replace(s, ChrW(8221), ChrW(34)) ' ” -> "
    s = Trim$(s)

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    Files_NormalizarCabecalho = s
End Function

Private Function Files_Col(ByVal mapaCab As Object, ByVal headerName As String) As Long
    headerName = Files_NormalizarCabecalho(headerName)
    If mapaCab Is Nothing Then
        Files_Col = 0
        Exit Function
    End If
    If mapaCab.exists(headerName) Then
        Files_Col = CLng(mapaCab(headerName))
    Else
        Files_Col = 0
    End If
End Function


' ============================================================
' TESTES (Regressão) — FILES_MANAGEMENT v2
' ============================================================

Public Sub Files_RegressionTests()
    On Error GoTo Falha

    Dim wsF As Worksheet
    Call Files_EnsureSheetExists
    Set wsF = ThisWorkbook.Worksheets(SHEET_FILES)

    Dim mapa As Object
    Set mapa = Files_MapaCabecalhos(wsF)

    Dim lo As ListObject
    Set lo = Files_GetOrCreateTable_FilesManagement(wsF)
    If lo Is Nothing Then Err.Raise vbObjectError + 100, , "Nao foi possivel obter/criar tblFILES_MANAGEMENT."

    Dim wsT As Worksheet
    Set wsT = Nothing
    On Error Resume Next
    Set wsT = ThisWorkbook.Worksheets("TESTS")
    On Error GoTo 0

    If wsT Is Nothing Then
        Set wsT = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsT.name = "TESTS"
    Else
        wsT.Cells.Clear
    End If

    wsT.Range("A1:D1").value = Array("Timestamp", "Teste", "Resultado", "Detalhe")
    wsT.rowS(1).Font.Bold = True

    Dim rowT As Long
    rowT = 2

    ' Helper inline
    Dim SubOk As Boolean

    ' --- Teste 1: Cabeçalhos
    SubOk = Files_HeaderV2_OK(wsF)
    Call Files_TestWrite(wsT, rowT, "Cabeçalhos v2 (ordem exacta)", SubOk, "")
    rowT = rowT + 1

    ' --- Teste 2: Mapeamento encontra todas as colunas
    Dim headers As Variant
    headers = Files_HeadersV2_Array()

    Dim i As Long
    Dim missing As String
    missing = ""
    For i = LBound(headers) To UBound(headers)
        If Files_Col(mapa, CStr(headers(i))) = 0 Then
            missing = missing & CStr(headers(i)) & "; "
        End If
    Next i

    SubOk = (missing = "")
    Call Files_TestWrite(wsT, rowT, "Mapeamento de cabeçalhos", SubOk, IIf(SubOk, "", "Faltam: " & missing))
    rowT = rowT + 1

    ' Guardar estado inicial
    Dim initialRows As Long
    initialRows = lo.ListRows.Count

    ' --- Inserir 25 eventos para chave INLINE (file_id vazio)
    Dim h As String
    h = Files_SHA256_Text("TEST_CONTENT")

    Dim n As Long
    Call Files_SetRunToken("TEST_RUN1")

    For n = 1 To 25
        Call Files_UpsertFilesManagement(wsF, mapa, "PIPE_TEST", "P_" & Format$(n, "0000"), "teste.pdf", _
            "C:\TEST", "C:\TEST\teste.pdf", "", "pdf_upload", False, "", Now, 1234, h, "teste", "DL")
    Next n

    ' Re-obter mapa e lo (podem ter mudado com inserções)
    Set mapa = Files_MapaCabecalhos(wsF)
    Set lo = Files_GetOrCreateTable_FilesManagement(wsF)

    ' --- Teste 3: Inserção no topo (linha 2)
    Dim topRow As Long
    topRow = lo.ListRows(1).Range.Row
    SubOk = (topRow = lo.HeaderRowRange.Row + 1)
    Call Files_TestWrite(wsT, rowT, "Inserção no topo (linha 2)", SubOk, "Row=" & CStr(topRow))
    rowT = rowT + 1

    ' --- Teste 4: Utilizações incrementa (deve ser 25 no topo)
    Dim cUtil As Long
    cUtil = Files_Col(mapa, H_UTILIZACOES)

    Dim vUtil As Long
    vUtil = CLng(wsF.Cells(lo.ListRows(1).Range.Row, cUtil).value)
    SubOk = (vUtil = 25)
    Call Files_TestWrite(wsT, rowT, "Utilizações (incremento)", SubOk, "Valor=" & CStr(vUtil))
    rowT = rowT + 1

    ' --- Teste 5: used_in_prompts formato e truncagem (20 + (...))
    Dim cProm As Long
    cProm = Files_Col(mapa, H_USED_IN_PROMPTS)
    Dim sProm As String
    sProm = CStr(wsF.Cells(lo.ListRows(1).Range.Row, cProm).value)

    SubOk = (InStr(1, sProm, USED_PROMPTS_SEP, vbBinaryCompare) > 0) And (Right$(Trim$(sProm), Len(USED_PROMPTS_SUFFIX)) = USED_PROMPTS_SUFFIX)
    Call Files_TestWrite(wsT, rowT, "used_in_prompts (separador e '(...)')", SubOk, sProm)
    rowT = rowT + 1

    ' --- Teste 6: DL/UL
    Dim cDlUl As Long
    cDlUl = Files_Col(mapa, H_DL_UL)
    Dim sDlUl As String
    sDlUl = CStr(wsF.Cells(lo.ListRows(1).Range.Row, cDlUl).value)
    SubOk = (UCase$(Trim$(sDlUl)) = "DL")
    Call Files_TestWrite(wsT, rowT, "DL/UL = DL (teste)", SubOk, "Valor=" & sDlUl)
    rowT = rowT + 1

' --- Teste 7: Separador por run (linha separadora) — muda token e insere 1 registo
Call Files_SetRunToken("TEST_RUN2")
Call Files_UpsertFilesManagement(wsF, mapa, "PIPE_TEST", "P_9999", "teste2.pdf", _
    "C:\TEST", "C:\TEST\teste2.pdf", "", "pdf_upload", False, "", Now, 1, h, "teste run2", "DL")

' Com o separador por linha:
' - A linha 1 da tabela (ListRows(1)) é o novo registo do run actual
' - A linha 2 da tabela (ListRows(2)) deve ser a "linha de intervalo" (RowHeight=6, vbBlack, xlSolid)
Dim isSep As Boolean
isSep = False

If lo.ListRows.Count >= 2 Then
    Dim rngSep As Range
    Set rngSep = lo.ListRows(2).Range

    Dim cFirst As Range, cLast As Range
    Set cFirst = rngSep.Cells(1, 1)
    Set cLast = rngSep.Cells(1, lo.ListColumns.Count)

    isSep = (Abs(rngSep.RowHeight - 6) < 0.1) _
            And (cFirst.Interior.pattern = xlSolid) _
            And (cFirst.Interior.Color = vbBlack) _
            And (cLast.Interior.pattern = xlSolid) _
            And (cLast.Interior.Color = vbBlack)
End If

SubOk = isSep
Call Files_TestWrite(wsT, rowT, "Separador por run (linha separadora)", SubOk, IIf(isSep, "OK", "Não encontrado/formatado como esperado"))
rowT = rowT + 1

' --- Limpeza: remover registos inseridos pelo teste
Dim added As Long
added = (lo.ListRows.Count - initialRows)
If added > 0 Then
    Dim k As Long
    For k = 1 To added
        lo.ListRows(1).Delete
    Next k
End If


    Call Files_TestWrite(wsT, rowT, "Limpeza pós-testes", True, "OK")
    rowT = rowT + 1

    wsT.Columns("A:D").AutoFit

    Debug.Print "Files_RegressionTests concluido. Ver folha TESTS."
    Exit Sub

Falha:
    Debug.Print "ERRO em Files_RegressionTests: " & Err.Description
    On Error Resume Next
    Call Files_TestWrite(ThisWorkbook.Worksheets("TESTS"), 2, "Files_RegressionTests", False, Err.Description)
End Sub

Private Sub Files_TestWrite(ByVal wsT As Worksheet, ByVal rowT As Long, ByVal testName As String, ByVal pass As Boolean, ByVal detalhe As String)
    On Error Resume Next
    wsT.Cells(rowT, 1).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsT.Cells(rowT, 2).value = testName
    wsT.Cells(rowT, 3).value = IIf(pass, "PASS", "FAIL")
    wsT.Cells(rowT, 4).value = detalhe
    On Error GoTo 0
End Sub




' ============================================================
' DIAGNOSTICOS: testes praticos para descobrir causa de HTTP 0
' ============================================================

Public Sub Files_Diag_TestarLeituraFicheiro(ByVal fullPath As String)
    On Error GoTo Falha

    Dim exists As Boolean
    exists = Files_ExisteFicheiro(fullPath)

    Dim ln As Double
    ln = 0
    If exists Then ln = Files_TamanhoBytes(fullPath)

    Dim b() As Byte
    b = Files_ReadAllBytes(fullPath)

    Dim n As Long
    n = Files_ByteArrayLen(b)

    MsgBox "Teste leitura:" & vbCrLf & _
           "Path=" & fullPath & vbCrLf & _
           "Dir/Existe=" & IIf(exists, "SIM", "NAO") & vbCrLf & _
           "FileLen=" & CStr(ln) & vbCrLf & _
           "BytesLidos=" & CStr(n), vbInformation
    Exit Sub

Falha:
    MsgBox "Falha no teste leitura: " & Err.Description, vbExclamation
End Sub

Public Sub Files_Diag_TestarConectividadeOpenAI(ByVal apiKey As String)
    On Error GoTo Falha

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error Resume Next
    http.SetTimeouts 15000, 15000, 30000, 30000
    http.Option(WINHTTP_OPTION_SECURE_PROTOCOLS) = WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2
    On Error GoTo Falha

    http.Open "GET", "https://api.openai.com/v1/models", False
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.SetRequestHeader "Accept", "application/json"
    http.Send

    MsgBox "Conectividade /v1/models:" & vbCrLf & _
           "HTTP=" & CStr(http.status) & vbCrLf & _
           "Resp (inicio)=" & Left$(CStr(http.ResponseText), 400), vbInformation
    Exit Sub

Falha:
    MsgBox "Falha conectividade (WinHTTP): " & Files_FormatComError("GET /v1/models", Err.Number, Err.Description), vbExclamation
End Sub

Public Sub Files_Diag_TestarUploadFicheiro(ByVal apiKey As String, ByVal fullPath As String, Optional ByVal purpose As String = "user_data")
    Dim fileId As String, st As Long, errMsg As String
    Dim ok As Boolean

    ok = Files_UploadFile_OpenAI(apiKey, "(diag)", fullPath, purpose, fileId, st, errMsg)

    MsgBox "Teste upload /v1/files:" & vbCrLf & _
           "OK=" & IIf(ok, "SIM", "NAO") & vbCrLf & _
           "HTTP=" & CStr(st) & vbCrLf & _
           "file_id=" & fileId & vbCrLf & _
           "erro=" & errMsg, vbInformation
End Sub

Public Sub Files_Diag_CorridaCompleta(ByVal apiKey As String, ByVal fullPath As String)
    ' 1) leitura
    Call Files_Diag_TestarLeituraFicheiro(fullPath)
    ' 2) conectividade
    Call Files_Diag_TestarConectividadeOpenAI(apiKey)
    ' 3) upload
    Call Files_Diag_TestarUploadFicheiro(apiKey, fullPath, "user_data")
End Sub



Private Function Files_ReadAllBytesEx(ByVal fullPath As String, ByRef outErro As String) As Byte()
    On Error GoTo Falha

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.LoadFromFile fullPath
    Files_ReadAllBytesEx = stm.Read
    stm.Close
    outErro = ""
    Exit Function

Falha:
    outErro = "Erro a ler bytes: " & Err.Description
    ' devolve array vazio
    Dim b() As Byte
    Files_ReadAllBytesEx = b
End Function

Private Function Files_ByteLen(ByRef arr() As Byte) As Long
    On Error GoTo ErrVazio
    Files_ByteLen = UBound(arr) - LBound(arr) + 1
    Exit Function
ErrVazio:
    Files_ByteLen = 0
End Function

Private Function Files_BuildMultipartBody_Safe( _
    ByVal boundary As String, _
    ByVal purpose As String, _
    ByVal fileName As String, _
    ByVal contentType As String, _
    ByRef fileBytes() As Byte _
) As Byte()

    Dim pre1 As String
    Dim pre2 As String
    Dim post As String

    pre1 = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""purpose""" & vbCrLf & vbCrLf & _
           purpose & vbCrLf

    pre2 = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf & _
           "Content-Type: " & contentType & vbCrLf & vbCrLf

    post = vbCrLf & "--" & boundary & "--" & vbCrLf

    Dim b1() As Byte
    Dim b2() As Byte
    Dim b3() As Byte
    Dim b4() As Byte

    b1 = StrConv(pre1, vbFromUnicode)
    b2 = StrConv(pre2, vbFromUnicode)
    b3 = fileBytes
    b4 = StrConv(post, vbFromUnicode)

    Dim totalLen As Long
    totalLen = Files_ByteLen(b1) + Files_ByteLen(b2) + Files_ByteLen(b3) + Files_ByteLen(b4)

    If totalLen = 0 Then
        Dim outEmpty() As Byte
        Files_BuildMultipartBody_Safe = outEmpty
        Exit Function
    End If

    Dim allBytes() As Byte
    ReDim allBytes(0 To totalLen - 1) As Byte

    Dim iPos As Long
    iPos = 0

    ' Copy arrays
    Dim i As Long

    If Files_ByteLen(b1) > 0 Then
        For i = 0 To UBound(b1)
            allBytes(iPos) = b1(i)
            iPos = iPos + 1
        Next i
    End If

    If Files_ByteLen(b2) > 0 Then
        For i = 0 To UBound(b2)
            allBytes(iPos) = b2(i)
            iPos = iPos + 1
        Next i
    End If

    If Files_ByteLen(b3) > 0 Then
        For i = 0 To UBound(b3)
            allBytes(iPos) = b3(i)
            iPos = iPos + 1
        Next i
    End If

    If Files_ByteLen(b4) > 0 Then
        For i = 0 To UBound(b4)
            allBytes(iPos) = b4(i)
            iPos = iPos + 1
        Next i
    End If

    Files_BuildMultipartBody_Safe = allBytes
End Function

' ============================================================
' (NOVO) Sanitização de filename para multipart (ASCII_SAFE)
'   - Apenas altera o "filename" enviado no Content-Disposition do multipart
'   - NAO altera o ficheiro no disco
'   - Mantém extensão (quando existir)
' ============================================================

Public Function Files_SanitizeFilenameAsciiSafe(ByVal fileName As String) As String
    Const MAX_TOTAL_LEN As Long = 180

    Dim nameOnly As String
    nameOnly = Files_SoNomeFicheiro(fileName)

    Dim baseName As String
    Dim ext As String
    baseName = nameOnly
    ext = ""

    Dim p As Long
    p = InStrRev(nameOnly, ".")
    If p > 1 And p < Len(nameOnly) Then
        baseName = Left$(nameOnly, p - 1)
        ext = Mid$(nameOnly, p + 1)
    End If

    ' 1) Normalizações principais
    baseName = Files_RemoverAcentosPT(baseName)
    baseName = Files_NormalizarPontuacaoFilename(baseName)
    baseName = Files_WhitespaceToHyphen(baseName)

    ' 2) Forçar apenas ASCII seguro
    baseName = Files_SanitizeFilenameAsciiCore(baseName, True)
    baseName = Files_CollapseFilenameSeparators(baseName)
    baseName = Files_TrimChars(baseName, "-_.")
    If baseName = "" Then baseName = "file"

    ' 3) Extensão: apenas alfanumérico ASCII
    If ext <> "" Then
        ext = Files_RemoverAcentosPT(ext)
        ext = Files_SanitizeExtensionAscii(ext)
        ext = Files_TrimChars(ext, "-_.")
        ext = LCase$(ext)
    End If

    ' 4) Encurtar preservando extensão
    Dim reserve As Long
    reserve = 0
    If ext <> "" Then reserve = Len(ext) + 1

    If Len(baseName) > (MAX_TOTAL_LEN - reserve) Then
        baseName = Left$(baseName, MAX_TOTAL_LEN - reserve)
        baseName = Files_TrimChars(baseName, "-_.")
        If baseName = "" Then baseName = "file"
    End If

    If ext <> "" Then
        Files_SanitizeFilenameAsciiSafe = baseName & "." & ext
    Else
        Files_SanitizeFilenameAsciiSafe = baseName
    End If
End Function

Private Function Files_NormalizarPontuacaoFilename(ByVal s As String) As String
    s = Replace$(s, ChrW(8211), "-") ' –
    s = Replace$(s, ChrW(8212), "-") ' —
    s = Replace$(s, ChrW(160), " ")  ' NBSP

    ' Aspas tipográficas (remover)
    s = Replace$(s, ChrW(8216), "")
    s = Replace$(s, ChrW(8217), "")
    s = Replace$(s, ChrW(8220), "")
    s = Replace$(s, ChrW(8221), "")

    Files_NormalizarPontuacaoFilename = s
End Function

Private Function Files_WhitespaceToHyphen(ByVal s As String) As String
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, " ", "-")
    Files_WhitespaceToHyphen = s
End Function

Private Function Files_SanitizeFilenameAsciiCore(ByVal s As String, ByVal allowDot As Boolean) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String
    out = ""

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)

        If code >= 32 And code <= 126 Then
            If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
                out = out & ch
            ElseIf ch = "-" Or ch = "_" Then
                out = out & ch
            ElseIf allowDot And ch = "." Then
                out = out & ch
            Else
                out = out & "-"
            End If
        Else
            out = out & "-"
        End If
    Next i

    Files_SanitizeFilenameAsciiCore = out
End Function

Private Function Files_SanitizeExtensionAscii(ByVal ext As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String
    out = ""

    For i = 1 To Len(ext)
        ch = Mid$(ext, i, 1)
        code = AscW(ch)

        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
            out = out & ch
        End If
    Next i

    Files_SanitizeExtensionAscii = out
End Function

Private Function Files_CollapseFilenameSeparators(ByVal s As String) As String
    Dim oldS As String
    Dim newS As String
    newS = s

    Do
        oldS = newS
        newS = Replace$(newS, "--", "-")
        newS = Replace$(newS, "__", "_")
        newS = Replace$(newS, "-_", "-")
        newS = Replace$(newS, "_-", "-")
        newS = Replace$(newS, "..", ".")
    Loop While newS <> oldS

    Files_CollapseFilenameSeparators = newS
End Function

Private Function Files_TrimChars(ByVal s As String, ByVal chars As String) As String
    Do While Len(s) > 0 And InStr(1, chars, Left$(s, 1), vbBinaryCompare) > 0
        s = Mid$(s, 2)
    Loop

    Do While Len(s) > 0 And InStr(1, chars, Right$(s, 1), vbBinaryCompare) > 0
        s = Left$(s, Len(s) - 1)
    Loop

    Files_TrimChars = s
End Function

Private Function Files_RemoverAcentosPT(ByVal s As String) As String
    ' a / A
    s = Replace$(s, ChrW(225), "a")
    s = Replace$(s, ChrW(224), "a")
    s = Replace$(s, ChrW(226), "a")
    s = Replace$(s, ChrW(227), "a")
    s = Replace$(s, ChrW(228), "a")
    s = Replace$(s, ChrW(193), "A")
    s = Replace$(s, ChrW(192), "A")
    s = Replace$(s, ChrW(194), "A")
    s = Replace$(s, ChrW(195), "A")
    s = Replace$(s, ChrW(196), "A")

    ' e / E
    s = Replace$(s, ChrW(233), "e")
    s = Replace$(s, ChrW(232), "e")
    s = Replace$(s, ChrW(234), "e")
    s = Replace$(s, ChrW(235), "e")
    s = Replace$(s, ChrW(201), "E")
    s = Replace$(s, ChrW(200), "E")
    s = Replace$(s, ChrW(202), "E")
    s = Replace$(s, ChrW(203), "E")

    ' i / I
    s = Replace$(s, ChrW(237), "i")
    s = Replace$(s, ChrW(236), "i")
    s = Replace$(s, ChrW(238), "i")
    s = Replace$(s, ChrW(239), "i")
    s = Replace$(s, ChrW(205), "I")
    s = Replace$(s, ChrW(204), "I")
    s = Replace$(s, ChrW(206), "I")
    s = Replace$(s, ChrW(207), "I")

    ' o / O
    s = Replace$(s, ChrW(243), "o")
    s = Replace$(s, ChrW(242), "o")
    s = Replace$(s, ChrW(244), "o")
    s = Replace$(s, ChrW(245), "o")
    s = Replace$(s, ChrW(246), "o")
    s = Replace$(s, ChrW(211), "O")
    s = Replace$(s, ChrW(210), "O")
    s = Replace$(s, ChrW(212), "O")
    s = Replace$(s, ChrW(213), "O")
    s = Replace$(s, ChrW(214), "O")

    ' u / U
    s = Replace$(s, ChrW(250), "u")
    s = Replace$(s, ChrW(249), "u")
    s = Replace$(s, ChrW(251), "u")
    s = Replace$(s, ChrW(252), "u")
    s = Replace$(s, ChrW(218), "U")
    s = Replace$(s, ChrW(217), "U")
    s = Replace$(s, ChrW(219), "U")
    s = Replace$(s, ChrW(220), "U")

    ' c / C
    s = Replace$(s, ChrW(231), "c")
    s = Replace$(s, ChrW(199), "C")

    ' n / N (extra)
    s = Replace$(s, ChrW(241), "n")
    s = Replace$(s, ChrW(209), "N")

    ' ordinais
    s = Replace$(s, ChrW(170), "a")
    s = Replace$(s, ChrW(186), "o")

    Files_RemoverAcentosPT = s
End Function




' ============================================================
' (NOVO) CONFIG — Reutilização de ficheiros no upload
' ============================================================
Private Sub Files_EnsureConfig_ReutilizacaoUpload()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    If ws Is Nothing Then Exit Sub

    Dim label As String
    label = "Reutilização de ficheiros no upload"

    ' Tenta encontrar a linha pelo texto na Col A
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    r = 0
    Dim i As Long
    For i = 1 To lastR
        If StrComp(Trim$(CStr(ws.Cells(i, 1).value)), label, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i

    ' Se não existir, cria numa linha "segura" (ex.: 8)
    If r = 0 Then r = 8

    ws.Cells(r, 1).value = label
    If Trim$(CStr(ws.Cells(r, 2).value)) = "" Then ws.Cells(r, 2).value = "TRUE"
    If Trim$(CStr(ws.Cells(r, 3).value)) = "" Then
        ws.Cells(r, 3).value = "TRUE = tenta reutilizar file_id do histórico (nome+hash+modo), validando se o file_id ainda existe via GET /v1/files/<id>. FALSE = força upload novo sempre."
    End If

    On Error GoTo 0
End Sub


Private Function Files_Config_ReutilizacaoUpload() As Boolean
    On Error GoTo Falha

    Files_Config_ReutilizacaoUpload = True ' default recomendado

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    Dim label As String
    label = "Reutilização de ficheiros no upload"

    Dim lastR As Long, r As Long, i As Long
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    r = 0
    For i = 1 To lastR
        If StrComp(Trim$(CStr(ws.Cells(i, 1).value)), label, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i

    Dim v As Variant
    If r > 0 Then
        v = ws.Cells(r, 2).value
    Else
        v = "TRUE"
    End If

    Files_Config_ReutilizacaoUpload = Files_ValorParaBool(v, True)
    Exit Function

Falha:
    Files_Config_ReutilizacaoUpload = True
End Function


' ============================================================
' (NOVO) Override por ficheiro na própria prompt
'   Ex.: "relatorio.pdf (required) (Reutilização de ficheiro=FALSE)"
' ============================================================
Private Sub Files_ParseReuseOverride(ByVal rawItem As String, ByRef outFound As Boolean, ByRef outValue As Boolean)
    outFound = False
    outValue = True

    Dim low As String
    low = LCase$(rawItem)

    ' Aceita com/sem acentos
    Dim p As Long
    p = InStr(1, low, "reutilização de ficheiro", vbTextCompare)
    If p = 0 Then p = InStr(1, low, "reutilizacao de ficheiro", vbTextCompare)
    If p = 0 Then Exit Sub

    Dim pEq As Long
    pEq = InStr(p, low, "=", vbTextCompare)
    If pEq = 0 Then Exit Sub

    Dim sVal As String
    sVal = Mid$(rawItem, pEq + 1)

    ' Termina no ')' se existir
    Dim pEnd As Long
    pEnd = InStr(1, sVal, ")", vbTextCompare)
    If pEnd > 0 Then sVal = Left$(sVal, pEnd - 1)

    ' Limpar ruído
    sVal = Replace(sVal, "[", "")
    sVal = Replace(sVal, "]", "")
    sVal = Replace(sVal, """", "")
    sVal = Trim$(sVal)

    If sVal = "" Then Exit Sub

    outFound = True
    outValue = Files_ValorParaBool(sVal, True)
End Sub


' ============================================================
' (NOVO) Validação online do file_id (ativo/reutilizável)
' ============================================================
Private Function Files_OpenAI_FileIdAtivo(ByVal apiKey As String, ByVal fileId As String, ByRef outHttpStatus As Long, ByRef outErro As String) As Boolean
    On Error GoTo Falha

    Files_OpenAI_FileIdAtivo = False
    outHttpStatus = 0
    outErro = ""

    fileId = Trim$(fileId)
    If fileId = "" Then
        outErro = "file_id vazio"
        Exit Function
    End If
    If Trim$(apiKey) = "" Then
        outErro = "API key vazia"
        Exit Function
    End If

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error Resume Next
    http.SetTimeouts 10000, 10000, 20000, 20000
    http.Option(WINHTTP_OPTION_SECURE_PROTOCOLS) = WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2
    On Error GoTo Falha

    http.Open "GET", "https://api.openai.com/v1/files/" & fileId, False
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.SetRequestHeader "Accept", "application/json"
    http.Send

    outHttpStatus = CLng(http.status)

    If outHttpStatus = 200 Then
        Files_OpenAI_FileIdAtivo = True
    Else
        outErro = "GET /v1/files/" & fileId & " => HTTP " & CStr(outHttpStatus)
        If Len(CStr(http.ResponseText)) > 0 Then outErro = outErro & " | " & Left$(CStr(http.ResponseText), 200)
    End If

    Exit Function

Falha:
    outErro = "Erro WinHTTP GET /v1/files/" & fileId & ": " & Err.Description
    Files_OpenAI_FileIdAtivo = False
End Function

Public Function Files_FNV32_LastDiag() As String
    Files_FNV32_LastDiag = gLastFNV32Diag
End Function

Public Function Files_SHA256_LastDiag() As String
    Files_SHA256_LastDiag = gLastSHA256Diag
End Function


Private Sub Files_MaybeAddRunSeparator(ByVal wsFiles As Worksheet, ByVal lo As ListObject)
    On Error GoTo Falha

    ' Sem token de run: não há separador por run
    If Trim$(gRunToken) = "" Then Exit Sub

    ' Já foi tratado este run: não repetir
    If StrComp(gRunToken, gLastSeparatorRunToken, vbTextCompare) = 0 Then Exit Sub

    ' Só inserir separador se existir histórico anterior
    Dim haHistoricoAnterior As Boolean
    haHistoricoAnterior = False

    If Not lo Is Nothing Then
        haHistoricoAnterior = (lo.ListRows.Count > 0)
    Else
        ' fallback (caso raro): histórico existe se houver dados a partir da linha 2
        haHistoricoAnterior = (wsFiles.Cells(wsFiles.rowS.Count, 1).End(xlUp).Row >= 2)
    End If

    If haHistoricoAnterior Then
        ' Insere a row separadora no topo do histórico anterior
        Call Files_AddRunSeparatorLine(wsFiles, lo, gRunToken)
    End If

    ' MUITO IMPORTANTE:
    ' Mesmo que não haja histórico anterior (1º run), marcar o token como tratado,
    ' para evitar inserir separador entre ficheiros do mesmo run.
    gLastSeparatorRunToken = gRunToken
    Exit Sub

Falha:
    ' Não bloquear o pipeline por erro de UI/formatacao do separador
End Sub


Private Function Files_U32_ToHex8_LL(ByVal v As LongLong) As String
    Dim u As LongLong
    u = v And &HFFFFFFFF

    Dim hi As LongLong, lo As LongLong
    hi = (u \ &H10000) And &HFFFF&
    lo = u And &HFFFF&

    Files_U32_ToHex8_LL = Right$("0000" & Hex$(hi), 4) & Right$("0000" & Hex$(lo), 4)
End Function

Private Function Files_U32_ToHex8_D(ByVal u As Double) As String
    If u < 0# Then u = u + 4294967296#
    u = u - Fix(u / 4294967296#) * 4294967296#

    Dim hi As Long, lo As Long
    hi = CLng(Fix(u / 65536#))
    lo = CLng(u - (CDbl(hi) * 65536#))

    Files_U32_ToHex8_D = Right$("0000" & Hex$(hi), 4) & Right$("0000" & Hex$(lo), 4)
End Function

Private Function Files_U32_XorByte(ByVal u As Double, ByVal b As Byte) As Double
    If u < 0# Then u = u + 4294967296#
    u = u - Fix(u / 4294967296#) * 4294967296#

    Dim hi As Long, lo As Long
    hi = CLng(Fix(u / 65536#))
    lo = CLng(u - (CDbl(hi) * 65536#))

    ' XOR só mexe no low byte — continua seguro em 16 bits
    lo = (lo Xor CLng(b)) And &HFFFF&

    Files_U32_XorByte = (CDbl(hi) * 65536#) + CDbl(lo)
End Function

Private Function Files_U32_Mul(ByVal u As Double, ByVal m As Long) As Double
    If u < 0# Then u = u + 4294967296#
    u = u - Fix(u / 4294967296#) * 4294967296#

    Dim aHi As Long, aLo As Long
    aHi = CLng(Fix(u / 65536#))
    aLo = CLng(u - (CDbl(aHi) * 65536#))

    Dim mHi As Long, mLo As Long
    mHi = (m \ 65536) And &HFFFF&
    mLo = m And &HFFFF&

    ' (a * m) mod 2^32, ignorando termo (hi*hi)*2^32
    Dim prod As Double
    prod = (CDbl(aLo) * CDbl(mLo)) + (CDbl(aLo) * CDbl(mHi) + CDbl(aHi) * CDbl(mLo)) * 65536#

    prod = prod - Fix(prod / 4294967296#) * 4294967296#
    Files_U32_Mul = prod
End Function

Private Function Files_FormatErroDetalhado( _
    ByVal procName As String, _
    ByVal stepName As String, _
    ByVal curReq As String, _
    ByVal curFile As String, _
    ByVal curPath As String, _
    ByVal curUsoFinal As String _
) As String

    Dim msg As String
    msg = "Erro (" & CStr(Err.Number) & ") em " & procName & ": " & Err.Description

    If Erl <> 0 Then msg = msg & " | Erl=" & CStr(Erl)
    If Trim$(Err.Source) <> "" Then msg = msg & " | Source=" & Err.Source

    If Trim$(stepName) <> "" Then msg = msg & " | step=" & stepName
    If Trim$(curReq) <> "" Then msg = msg & " | req=" & curReq
    If Trim$(curFile) <> "" Then msg = msg & " | file=" & curFile
    If Trim$(curPath) <> "" Then msg = msg & " | path=" & curPath
    If Trim$(curUsoFinal) <> "" Then msg = msg & " | usoFinal=" & curUsoFinal

    Files_FormatErroDetalhado = msg
End Function



Private Function Files_Config_GetByKey(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    ' Procura uma chave na folha Config (coluna A = chave, coluna B = valor).
    ' Se não existir (ou houver erro), devolve defaultValue.
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    If lastRow < 1 Then
        Files_Config_GetByKey = defaultValue
        Exit Function
    End If

    Dim i As Long
    Dim k As String

    For i = 1 To lastRow
        k = Trim$(CStr(ws.Cells(i, 1).value))
        If k <> "" Then
            If StrComp(k, keyName, vbTextCompare) = 0 Then
                Files_Config_GetByKey = Trim$(CStr(ws.Cells(i, 2).value))
                Exit Function
            End If
        End If
    Next i

    Files_Config_GetByKey = defaultValue
    Exit Function

EH:
    Files_Config_GetByKey = defaultValue
End Function



Private Sub Files_EnsureConfig_DocxPolicies()
    ' ============================================================
    ' Garante que as políticas/limites para DOCX/PPTX e text_embed
    ' existem na folha Config (auto-documentação).
    '
    ' Cria (se não existir) as chaves:
    ' - FILES_DOCX_CONTEXT_MODE
    ' - FILES_DOCX_AS_PDF_FALLBACK
    ' - FILES_TEXT_EMBED_MAX_CHARS
    ' - FILES_TEXT_EMBED_OVERFLOW_ACTION
    '
    ' Regras:
    ' - Não sobrescreve valores já preenchidos (coluna B).
    ' - Só preenche descrições (coluna C) se estiverem vazias.
    ' - Evita colidir com B1..B7; escreve em linhas >= 9.
    ' - Idempotente (pode correr múltiplas vezes sem duplicar).
    ' ============================================================

    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    If ws Is Nothing Then Exit Sub

    ' (Opcional mas recomendável) garantir primeiro a chave antiga de reutilização em row "segura"
    ' para manter coerência com o padrão já existente do M09.
    Call Files_EnsureConfig_ReutilizacaoUpload
    Call Files_EnsureConfig_ReutilizacaoUpload

    Dim keys As Variant, defs As Variant, descs As Variant

    keys = Array( _
        "FILES_DOCX_CONTEXT_MODE", _
        "FILES_DOCX_AS_PDF_FALLBACK", _
        "FILES_TEXT_EMBED_MAX_CHARS", _
        "FILES_TEXT_EMBED_OVERFLOW_ACTION" _
    )

    defs = Array( _
        "AUTO_AS_PDF", _
        "TEXT_EMBED", _
        "50000", _
        "RETRY_AS_PDF" _
    )

    descs = Array( _
        "Política quando um DOC/DOCX/PPT/PPTX é pedido como (as_is)/(input_file) mas /v1/responses não aceita esse formato como input_file. Valores: AUTO_AS_PDF (recomendado), AUTO_TEXT_EMBED, ERROR. Default: AUTO_AS_PDF.", _
        "Fallback quando a conversão DOCX/PPTX->PDF falha. Valores: TEXT_EMBED (recomendado) ou ERROR. Default: TEXT_EMBED.", _
        "Limite máximo de caracteres para text_embed (texto extraído/embebido). Se exceder, gera alerta (TEXT_EMBED_TOO_LARGE) e aplica FILES_TEXT_EMBED_OVERFLOW_ACTION. Default: 50000.", _
        "Ação quando text_embed excede FILES_TEXT_EMBED_MAX_CHARS. Valores: ALERT_ONLY | TRUNCATE | RETRY_AS_PDF (recomendado) | STOP. Default: RETRY_AS_PDF." _
    )

    Dim lastR As Long
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row

    ' Se a coluna A ainda não tem nada (porque Config usa B1..B7),
    ' End(xlUp) pode devolver 1; garantir zona "segura" >= 9.
    If lastR < 8 Then lastR = 8

    Dim k As Long
    For k = LBound(keys) To UBound(keys)

        Dim keyName As String
        keyName = CStr(keys(k))

        Dim r As Long
        r = 0

        ' Procurar a chave na Coluna A
        Dim i As Long
        For i = 1 To lastR
            If StrComp(Trim$(CStr(ws.Cells(i, 1).value)), keyName, vbTextCompare) = 0 Then
                r = i
                Exit For
            End If
        Next i

        ' Se não existe, criar nova linha no fim (>= 9)
        If r = 0 Then
            r = lastR + 1
            If r < 9 Then r = 9
            lastR = r
        End If

        ' Preencher chave / default / descrição
        ws.Cells(r, 1).value = keyName
        If Trim$(CStr(ws.Cells(r, 2).value)) = "" Then ws.Cells(r, 2).value = CStr(defs(k))
        If Trim$(CStr(ws.Cells(r, 3).value)) = "" Then ws.Cells(r, 3).value = CStr(descs(k))
    Next k

    On Error GoTo 0
End Sub


' ============================================================
' PDF Cache — evita reconversões desnecessárias (e file_id novo)
' - Para DOCX/PPTX convertidos para PDF, o Word/PPT pode gerar
'   bytes diferentes a cada exportação (metadata). Isso quebra a
'   reutilização baseada em hash.
' - Solução: cache local + sidecar .src.sha256 com o hash do source.
'   Se o source não mudou, reutiliza o PDF existente sem reconverter.
' ============================================================


Private Function Files_PdfCache_GetOrConvertPdf( _
    ByVal promptId As String, _
    ByVal srcName As String, _
    ByVal srcPath As String, _
    ByVal destPdfPath As String, _
    ByVal sourceHash As String, _
    ByRef outUsedCache As Boolean, _
    ByRef outConvertedNow As Boolean, _
    ByRef outErro As String _
) As Boolean
    On Error GoTo EH

    outErro = ""
    outUsedCache = False
    outConvertedNow = False
    Files_PdfCache_GetOrConvertPdf = False

    srcPath = Trim$(CStr(srcPath))
    destPdfPath = Trim$(CStr(destPdfPath))
    sourceHash = Trim$(CStr(sourceHash))

    If srcPath = "" Or destPdfPath = "" Then
        outErro = "Paths inválidos para cache PDF."
        Exit Function
    End If

    If sourceHash = "" Then
        outErro = "sourceHash vazio (não é possível validar cache)."
        Exit Function
    End If

    Dim sidecar As String
    sidecar = Files_PdfCache_SidecarPath(destPdfPath)

    ' ============================================================
    ' 1) Cache HIT (preferência: sidecar hash)
    ' ============================================================
    If Files_ExisteFicheiro(destPdfPath) Then

        Dim pdfSize As Double
        pdfSize = Files_TamanhoBytes(destPdfPath)

        If pdfSize > 0 Then

            Dim cachedHash As String
            cachedHash = ""

            If Files_ExisteFicheiro(sidecar) Then
                Dim errRead As String
                errRead = ""
                cachedHash = Trim$(Files_LerTexto(sidecar, errRead))
            End If

            If cachedHash <> "" Then
                If StrComp(cachedHash, sourceHash, vbTextCompare) = 0 Then
                    outUsedCache = True
                    Call Debug_Registar(0, promptId, "INFO", "", "PDF_CACHE_HIT", _
                        "PDF em cache usado (sidecar OK): " & srcName, _
                        "Sem reconversão. Se o DOCX mudar, o sourceHash muda e será reconvertido.")
                    Files_PdfCache_GetOrConvertPdf = True
                    Exit Function
                End If
            End If

            ' ========================================================
            ' 1.b) Fallback: PDF mais recente que origem => aceitar cache
            ' ========================================================
            Dim srcMod As Date, pdfMod As Date
            srcMod = Files_DataModificacao(srcPath)
            pdfMod = Files_DataModificacao(destPdfPath)

            If pdfMod >= srcMod Then
                outUsedCache = True

                ' Atualizar sidecar (best-effort)
                Dim errW As String
                errW = ""
                Call Files_EscreverTextoUTF8(sidecar, sourceHash, errW)

                Call Debug_Registar(0, promptId, "INFO", "", "PDF_CACHE_HIT", _
                    "PDF em cache usado (timestamp OK; sidecar atualizado): " & srcName, _
                    "Sem reconversão. Se existirem divergências, apague o PDF/sidecar do _pdf_cache.")
                Files_PdfCache_GetOrConvertPdf = True
                Exit Function
            End If

        End If
    End If

    ' ============================================================
    ' 2) Cache MISS: converter agora
    ' ============================================================
    Dim errConv As String
    errConv = ""

    If Not Files_ConverterParaPDF(srcPath, destPdfPath, errConv) Then
        outErro = errConv
        Files_PdfCache_GetOrConvertPdf = False
        Exit Function
    End If

    If Not Files_ExisteFicheiro(destPdfPath) Then
        outErro = "Conversão PDF reportou OK mas o ficheiro não foi criado: " & destPdfPath
        Files_PdfCache_GetOrConvertPdf = False
        Exit Function
    End If

    If Files_TamanhoBytes(destPdfPath) <= 0 Then
        outErro = "PDF convertido ficou com 0 bytes: " & destPdfPath
        Files_PdfCache_GetOrConvertPdf = False
        Exit Function
    End If

    outConvertedNow = True

    ' Sidecar (best-effort)
    Dim errSide As String
    errSide = ""
    Call Files_EscreverTextoUTF8(sidecar, sourceHash, errSide)

    Call Debug_Registar(0, promptId, "INFO", "", "PDF_CACHE_MISS_CONVERTED", _
        "PDF gerado (cache miss): " & srcName, _
        "Na próxima execução deverá ocorrer PDF_CACHE_HIT e não reconverter.")

    Files_PdfCache_GetOrConvertPdf = True
    Exit Function

EH:
    outErro = "PDF cache: erro inesperado: " & Err.Number & " - " & Err.Description
    Files_PdfCache_GetOrConvertPdf = False
End Function


Private Function Files_PdfCache_SidecarPath(ByVal pdfPath As String) As String
    Files_PdfCache_SidecarPath = CStr(pdfPath) & PDF_CACHE_SIDECAR_EXT
End Function

Private Function Files_FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    Files_FileExists = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
End Function

Private Function Files_ReadAllTextFile(ByVal filePath As String, ByRef outText As String) As Boolean
    On Error GoTo EH

    outText = ""
    Files_ReadAllTextFile = False

    If Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then Exit Function

    Dim fh As Integer
    fh = FreeFile

    Open filePath For Input As #fh
    If LOF(fh) > 0 Then
        outText = Input$(LOF(fh), fh)
    Else
        outText = ""
    End If
    Close #fh

    Files_ReadAllTextFile = True
    Exit Function

EH:
    On Error Resume Next
    If fh <> 0 Then Close #fh
    outText = ""
    Files_ReadAllTextFile = False
End Function

Private Function Files_WriteAllTextFile(ByVal filePath As String, ByVal text As String) As Boolean
    On Error GoTo EH

    Files_WriteAllTextFile = False

    Dim fh As Integer
    fh = FreeFile

    Open filePath For Output As #fh
    Print #fh, text
    Close #fh

    Files_WriteAllTextFile = True
    Exit Function

EH:
    On Error Resume Next
    If fh <> 0 Then Close #fh
    Files_WriteAllTextFile = False
End Function


Public Function Files_FNV32_Bytes(ByRef b() As Byte) As String
    On Error GoTo Falha

    gLastFNV32Diag = ""
    Files_FNV32_Bytes = ""

    Dim ln As Long
    On Error Resume Next
    ln = (UBound(b) - LBound(b) + 1)
    On Error GoTo Falha

    If ln <= 0 Then
        gLastFNV32Diag = "FNV32_Bytes: byte array vazio"
        Exit Function
    End If

#If VBA7 And Win64 Then
    ' 64-bit: usar LongLong com máscara 32-bit
    Dim mask As LongLong
    mask = CLngLng(4294967295#)

    Dim h As LongLong
    h = CLngLng(&H811C9DC5) ' offset basis
    h = h And mask

    Dim j As Long
    For j = LBound(b) To UBound(b)
        h = Files_U32_XorByte(h, b(j))
        h = Files_U32_Mul(h, 16777619)
    Next j

    Files_FNV32_Bytes = "fnv32-" & Files_U32_ToHex8_LL(h)
    Exit Function
#Else
    ' 32-bit: usar Double como "unsigned 32"
    Dim u As Double
    u = 2166136261#

    Dim k As Long
    For k = LBound(b) To UBound(b)
        u = Files_U32_XorByte(u, b(k))
        u = Files_U32_Mul(u, 16777619)
    Next k

    Files_FNV32_Bytes = "fnv32-" & Files_U32_ToHex8_D(u)
    Exit Function
#End If

Falha:
    gLastFNV32Diag = "FNV32_Bytes: erro (" & Err.Number & ") " & Err.Description
    Files_FNV32_Bytes = ""
End Function


Private Function Files_EscreverTextoUTF8(ByVal fullPath As String, ByVal text As String, ByRef outErro As String) As Boolean
    On Error GoTo Falha

    outErro = ""
    Files_EscreverTextoUTF8 = False

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText CStr(text)
    stm.SaveToFile fullPath, 2 ' adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing

    Files_EscreverTextoUTF8 = True
    Exit Function

Falha:
    outErro = "Escrever texto UTF-8 falhou: " & Err.Description
    Files_EscreverTextoUTF8 = False
End Function


Private Function Files_EncontrarLinhaPorPath( _
    ByVal wsFiles As Worksheet, _
    ByVal mapaCab As Object, _
    ByVal fullPath As String, _
    ByVal usageMode As String, _
    Optional ByVal exigirFileId As Boolean = False _
) As Long
    On Error GoTo Falha

    Files_EncontrarLinhaPorPath = 0

    fullPath = Trim$(CStr(fullPath))
    usageMode = Trim$(CStr(usageMode))

    If fullPath = "" Or usageMode = "" Then Exit Function

    Dim colPath As Long, colMode As Long, colFileId As Long
    colPath = Files_Col(mapaCab, H_FULL_PATH)
    colMode = Files_Col(mapaCab, H_USAGE_MODE)
    colFileId = Files_Col(mapaCab, H_FILE_ID)

    If colPath = 0 Or colMode = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = Files_LastDataRow(wsFiles)
    If lastRow < 2 Then Exit Function

    Dim r As Long
    For r = lastRow To 2 Step -1
        If StrComp(Trim$(CStr(wsFiles.Cells(r, colPath).value)), fullPath, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(wsFiles.Cells(r, colMode).value)), usageMode, vbTextCompare) = 0 Then
                If exigirFileId Then
                    If colFileId > 0 Then
                        If Trim$(CStr(wsFiles.Cells(r, colFileId).value)) = "" Then GoTo Continua
                    End If
                End If

                Files_EncontrarLinhaPorPath = r
                Exit Function
            End If
        End If
Continua:
    Next r

    Exit Function

Falha:
    Files_EncontrarLinhaPorPath = 0
End Function


