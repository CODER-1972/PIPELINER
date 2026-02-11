Attribute VB_Name = "M11_DebugLogging"
Option Explicit

' Logging leve e compatível com folhas DEBUG/Seguimento.
' Escreve apenas texto e metadados; nunca escreve chaves nem bytes de ficheiros.

Public Enum LogSeverity
sevINFO = 0
sevALERTA = 1
sevERRO = 2
End Enum

Public Sub DebugLog(ByVal severity As LogSeverity, ByVal tag As String, ByVal message As String)
On Error GoTo EH
Dim ws As Object, nxt As Long
Set ws = EnsureSheetByName("DEBUG")
nxt = NextFreeRow(ws, 1)
ws.Cells(nxt, 1).value = Now
ws.Cells(nxt, 2).value = Switch(severity = sevINFO, "INFO", severity = sevALERTA, "ALERTA", True, "ERRO")
ws.Cells(nxt, 3).value = tag
ws.Cells(nxt, 4).value = message
Exit Sub
EH:
' Evitar falhar o fluxo por erro de logging
End Sub

Public Sub SeguimentoInfo(ByVal campo As String, ByVal valor As String)
On Error GoTo EH
Dim ws As Object, nxt As Long
Set ws = EnsureSheetByName("Seguimento")
nxt = NextFreeRow(ws, 1)
ws.Cells(nxt, 1).value = Now
ws.Cells(nxt, 2).value = "[INFO]"
ws.Cells(nxt, 3).value = campo
ws.Cells(nxt, 4).value = valor
Exit Sub
EH:
End Sub

Public Sub FilesManagementNote(ByVal note As String)
On Error GoTo EH
Dim ws As Object, nxt As Long
Set ws = EnsureSheetByName("FILES_MANAGEMENT")
nxt = NextFreeRow(ws, 1)
ws.Cells(nxt, 1).value = Now
ws.Cells(nxt, 2).value = "NOTE"
ws.Cells(nxt, 3).value = note
Exit Sub
EH:
End Sub

Private Function EnsureSheetByName(ByVal name As String) As Object
Dim ws As Object
For Each ws In ThisWorkbook.Worksheets
If StrComp(ws.name, name, vbTextCompare) = 0 Then
Set EnsureSheetByName = ws
Exit Function
End If
Next ws
Set EnsureSheetByName = ThisWorkbook.Worksheets.Add
EnsureSheetByName.name = name
End Function

Private Function NextFreeRow(ByVal ws As Object, ByVal colIndex As Long) As Long
Dim r As Long
r = ws.Cells(ws.rowS.Count, colIndex).End(-4162).Row ' xlUp
If Len(Trim$(ws.Cells(r, colIndex).value & "")) > 0 Then
NextFreeRow = r + 1
Else
NextFreeRow = r
End If
End Function
