Attribute VB_Name = "M01_Tipos"
Option Explicit

' =============================================================================
' Módulo: M01_Tipos
' Propósito:
' - Declarar tipos partilhados entre módulos para transporte de dados de prompts e respostas API.
' - Evitar estruturas ad-hoc e manter contratos internos explícitos.
'
' Atualizações:
' - 2026-02-12 | Codex | Implementação do padrão de header obrigatório
'   - Adiciona propósito, histórico de alterações e inventário de rotinas públicas.
'   - Mantém documentação técnica do módulo alinhada com AGENTS.md.
'
' Funções e procedimentos (inventário público):
' - PromptDefinicao (Type): Estrutura pública partilhada entre módulos.
' - ApiResultado (Type): Estrutura pública partilhada entre módulos.
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
    Erro As String
End Type
