Attribute VB_Name = "M01_Tipos"
Option Explicit

' =============================================================================
' Modulo: M01_Tipos
' Proposito:
' - Declarar tipos partilhados entre modulos para transporte de dados de prompts e respostas API.
' - Evitar estruturas ad-hoc e manter contratos internos explicitos.
'
' Atualizacoes:
' - 2026-03-12 | Codex | Explicita metrica de tentativas/retries no contrato da API
'   - Acrescenta `apiAttemptCount` e `retryCount` ao tipo `ApiResultado` para transportar contagem real de novas tentativas HTTP.
'   - Mantem compatibilidade retroativa: defaults continuam zero quando nao ha retry.
' - 2026-02-12 | Codex | Implementacao do padrao de header obrigatorio
'   - Adiciona proposito, historico de alteracoes e inventario de rotinas publicas.
'   - Mantem documentacao tecnica do modulo alinhada com AGENTS.md.
'
' Funcoes e procedimentos (inventario publico):
' - PromptDefinicao (Type): Estrutura publica partilhada entre modulos.
' - ApiResultado (Type): Estrutura publica partilhada entre modulos.
' =============================================================================

Public Type PromptDefinicao
    Id As String
    nomeFolha As String

    NomeCurto As String
    NomeDescritivo As String
    textoPrompt As String

    modelo As String
    modos As String
    storage As Boolean

    ConfigExtra As String

    Comentarios As String
    NotasDev As String
    HistoricoVersoes As String
End Type

Public Type ApiResultado
    httpStatus As Long
    responseId As String
    outputText As String
    rawResponseJson As String
    apiAttemptCount As Long
    retryCount As Long
    Erro As String
End Type

Public Type GH_HttpCallResult
    ok As Boolean
    status As Long
    body As String
    errorDetail As String
    stepName As String
End Type

Public Type GH_CommitOrchestratorResult
    ok As Boolean
    commitSha As String
    attemptCount As Long
    errorCode As String
    errorDetail As String
End Type
