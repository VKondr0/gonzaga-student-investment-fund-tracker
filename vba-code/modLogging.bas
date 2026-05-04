Option Explicit

Public Sub LogAction(ByVal macroName As String, _
                     ByVal actionName As String, _
                     ByVal statusText As String, _
                     ByVal messageText As String, _
                     ByVal sheetAffected As String)

    Dim wsLog As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("VBA_Log")
    On Error GoTo 0

    If wsLog Is Nothing Then Exit Sub

    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    If nextRow < 4 Then nextRow = 4

    wsLog.Cells(nextRow, "A").Value = Now
    wsLog.Cells(nextRow, "B").Value = macroName
    wsLog.Cells(nextRow, "C").Value = actionName
    wsLog.Cells(nextRow, "D").Value = statusText
    wsLog.Cells(nextRow, "E").Value = messageText
    wsLog.Cells(nextRow, "F").Value = "Manual button"
    wsLog.Cells(nextRow, "G").Value = sheetAffected

    wsLog.Cells(nextRow, "A").NumberFormat = "dd.mm.yyyy hh:mm:ss"

End Sub
