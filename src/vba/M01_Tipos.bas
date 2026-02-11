Attribute VB_Name = "M01_Tipos"
Option Explicit

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
