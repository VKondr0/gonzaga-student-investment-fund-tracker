Option Explicit

Sub SaveSnapshot()

    Dim wsHist As Worksheet
    Dim wsSum As Worksheet
    Dim wsDash As Worksheet
    Dim nextRow As Long
    
    On Error GoTo SnapshotError
    
    Set wsHist = ThisWorkbook.Worksheets("History")
    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    Application.ScreenUpdating = False
    Application.Calculate
    
    nextRow = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
    
    wsHist.Cells(nextRow, "A").Value = Now
    wsHist.Cells(nextRow, "B").Value = wsSum.Range("B4").Value
    wsHist.Cells(nextRow, "C").Value = wsSum.Range("B3").Value
    wsHist.Cells(nextRow, "D").Value = wsSum.Range("B6").Value
    wsHist.Cells(nextRow, "E").Value = wsSum.Range("B7").Value
    wsHist.Cells(nextRow, "F").Value = wsSum.Range("B8").Value
    wsHist.Cells(nextRow, "G").Value = wsSum.Range("B9").Value
    wsHist.Cells(nextRow, "H").Value = wsSum.Range("B10").Value
    wsHist.Cells(nextRow, "I").Value = wsDash.Range("E49").Value
    wsHist.Cells(nextRow, "J").Value = wsDash.Range("I49").Value
    
    wsHist.Cells(nextRow, "A").NumberFormat = "dd.mm.yyyy hh:mm"
    wsHist.Cells(nextRow, "B").NumberFormat = "[$$-en-US]#,##0.00"
    wsHist.Cells(nextRow, "C").NumberFormat = "[$$-en-US]#,##0.00"
    wsHist.Cells(nextRow, "D").NumberFormat = "[$$-en-US]#,##0.00"
    wsHist.Cells(nextRow, "E").NumberFormat = "0.00%"
    wsHist.Cells(nextRow, "F").NumberFormat = "0"
    wsHist.Cells(nextRow, "H").NumberFormat = "0.00%"
    wsHist.Cells(nextRow, "I").NumberFormat = "0.00%"
    wsHist.Cells(nextRow, "J").NumberFormat = "0.00%"
    
    Application.ScreenUpdating = True
    
    Call LogAction("SaveSnapshot", _
                   "Save portfolio snapshot", _
                   "Success", _
                   "Snapshot saved successfully", _
                   "History")
    
    MsgBox "Portfolio snapshot saved successfully.", vbInformation, "Portfolio Tracker"
    Exit Sub

SnapshotError:
    Application.ScreenUpdating = True
    
    Call LogAction("SaveSnapshot", _
                   "Save portfolio snapshot", _
                   "Error", _
                   "Snapshot could not be saved", _
                   "History")
    
    MsgBox "Snapshot could not be saved. Please check the History, Summary, and Dashboard sheets.", vbExclamation, "Portfolio Tracker"

End Sub
