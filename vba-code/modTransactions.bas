Option Explicit

Sub ClearFormButton()
    ClearForm True
End Sub

Private Sub ClearForm(Optional ByVal ShowMessage As Boolean = True)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Transactions")
    
    ws.Range("P3").ClearContents
    ws.Range("P4").ClearContents
    ws.Range("P6").ClearContents
    ws.Range("P7").ClearContents
    ws.Range("P9").ClearContents
    ws.Range("P10").ClearContents
    ws.Range("P11").ClearContents
    
    ws.Range("P3").Value = Date
    ws.Range("P9").Value = 0
    ws.Range("P10").Value = "Main Portfolio"
    
    If ShowMessage Then
        MsgBox "Input form cleared.", vbInformation, "Portfolio Tracker"
    End If

End Sub

Function GetAvailableQuantity(ByVal tickerSymbol As String) As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim currentTicker As String
    Dim actionType As String
    Dim qty As Double
    
    GetAvailableQuantity = 0
    
    Set ws = ThisWorkbook.Worksheets("Transactions")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        currentTicker = UCase(Trim(ws.Cells(i, "B").Value))
        actionType = UCase(Trim(ws.Cells(i, "D").Value))
        
        If currentTicker = UCase(Trim(tickerSymbol)) Then
            If IsNumeric(ws.Cells(i, "E").Value) Then
                qty = CDbl(ws.Cells(i, "E").Value)
                
                If actionType = "BUY" Then
                    GetAvailableQuantity = GetAvailableQuantity + qty
                ElseIf actionType = "SELL" Then
                    GetAvailableQuantity = GetAvailableQuantity - qty
                End If
            End If
        End If
    Next i

End Function

Function ValidateInput() As Boolean

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Transactions")
    
    ValidateInput = False
    
    If Trim(ws.Range("P3").Value) = "" Then
        MsgBox "Please enter the transaction date.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If Trim(ws.Range("P4").Value) = "" Then
        MsgBox "Please select a ticker.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If Trim(ws.Range("P5").Value) = "" Then
        MsgBox "Please select an asset type.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If Trim(ws.Range("P6").Value) = "" Then
        MsgBox "Please select an action.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If UCase(Trim(ws.Range("P6").Value)) <> "BUY" And UCase(Trim(ws.Range("P6").Value)) <> "SELL" Then
        MsgBox "Action must be Buy or Sell.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If Trim(ws.Range("P7").Value) = "" Or Not IsNumeric(ws.Range("P7").Value) Then
        MsgBox "Please enter a valid quantity.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If CDbl(ws.Range("P7").Value) <= 0 Then
        MsgBox "Quantity must be greater than 0.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If UCase(Trim(ws.Range("P6").Value)) = "SELL" Then
        If CDbl(ws.Range("P7").Value) > GetAvailableQuantity(Trim(ws.Range("P4").Value)) Then
            MsgBox "You cannot sell more than the available quantity for this ticker.", vbExclamation, "Portfolio Tracker"
            Exit Function
        End If
    End If
    
    If Trim(ws.Range("P8").Value) = "" Or Not IsNumeric(ws.Range("P8").Value) Then
        MsgBox "Please enter a valid price per unit.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If CDbl(ws.Range("P8").Value) <= 0 Then
        MsgBox "Price per unit must be greater than 0.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If Trim(ws.Range("P9").Value) = "" Then
        ws.Range("P9").Value = 0
    End If
    
    If Not IsNumeric(ws.Range("P9").Value) Then
        MsgBox "Broker fee must be a number.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If CDbl(ws.Range("P9").Value) < 0 Then
        MsgBox "Broker fee cannot be negative.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    If Trim(ws.Range("P10").Value) = "" Then
        MsgBox "Please enter an account name.", vbExclamation, "Portfolio Tracker"
        Exit Function
    End If
    
    ValidateInput = True

End Function

Sub TestValidation()

    If ValidateInput = True Then
        MsgBox "Input is valid.", vbInformation, "Portfolio Tracker"
    End If

End Sub

Sub AddTransaction()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    Dim colTotalAmount As Long
    Dim colNetCashFlow As Long
    Dim colAvgCostBeforeSale As Long
    Dim colRealizedGL As Long
    
    Dim rowNum As Long
    Dim prevRowNum As Long
    
    Set ws = ThisWorkbook.Worksheets("Transactions")
    Set tbl = ws.ListObjects("tblTransactions")
    
    If ValidateInput = False Then Exit Sub
    
    Set newRow = tbl.ListRows.Add
    
    On Error GoTo ColumnError
    
    newRow.Range.Cells(1, tbl.ListColumns("Date").Index).Value = ws.Range("P3").Value
    newRow.Range.Cells(1, tbl.ListColumns("Ticker").Index).Value = UCase(Trim(ws.Range("P4").Value))
    newRow.Range.Cells(1, tbl.ListColumns("Asset Type").Index).Value = Trim(ws.Range("P5").Value)
    newRow.Range.Cells(1, tbl.ListColumns("Action").Index).Value = WorksheetFunction.Proper(LCase(Trim(ws.Range("P6").Value)))
    newRow.Range.Cells(1, tbl.ListColumns("Quantity").Index).Value = CDbl(ws.Range("P7").Value)
    newRow.Range.Cells(1, tbl.ListColumns("Price per Unit").Index).Value = CDbl(ws.Range("P8").Value)
    newRow.Range.Cells(1, tbl.ListColumns("Broker Fee").Index).Value = CDbl(ws.Range("P9").Value)
    newRow.Range.Cells(1, tbl.ListColumns("Account").Index).Value = Trim(ws.Range("P10").Value)
    newRow.Range.Cells(1, tbl.ListColumns("Notes").Index).Value = Trim(ws.Range("P11").Value)
    
    colTotalAmount = tbl.ListColumns("Total Amount").Index
    colNetCashFlow = tbl.ListColumns("Net Cash Flow").Index
    colAvgCostBeforeSale = tbl.ListColumns("Avg Cost Before Sale").Index
    colRealizedGL = tbl.ListColumns("Realized G/L").Index
    
    newRow.Range.Cells(1, colTotalAmount).FormulaR1C1 = "=RC[-2]*RC[-1]"
    newRow.Range.Cells(1, colNetCashFlow).FormulaR1C1 = "=IF(UPPER(RC[-5])=""SELL"",RC[-2]-RC[-1],RC[-2]+RC[-1])"
    
    rowNum = newRow.Range.Row
    prevRowNum = rowNum - 1
    
    newRow.Range.Cells(1, colAvgCostBeforeSale).Formula2 = _
    "=IFERROR(IF($D" & rowNum & "=" & """Buy""" & "," & _
    "((SUMIFS($E$1:$E" & prevRowNum & ",$B$1:$B" & prevRowNum & ",$B" & rowNum & ",$D$1:$D" & prevRowNum & ",""Buy"")-" & _
    "SUMIFS($E$1:$E" & prevRowNum & ",$B$1:$B" & prevRowNum & ",$B" & rowNum & ",$D$1:$D" & prevRowNum & ",""Sell""))*" & _
    "XLOOKUP($B" & rowNum & ",$B$1:$B" & prevRowNum & ",$J$1:$J" & prevRowNum & ",0,0,-1)+$I" & rowNum & ")/" & _
    "((SUMIFS($E$1:$E" & prevRowNum & ",$B$1:$B" & prevRowNum & ",$B" & rowNum & ",$D$1:$D" & prevRowNum & ",""Buy"")-" & _
    "SUMIFS($E$1:$E" & prevRowNum & ",$B$1:$B" & prevRowNum & ",$B" & rowNum & ",$D$1:$D" & prevRowNum & ",""Sell""))+$E" & rowNum & ")," & _
    "XLOOKUP($B" & rowNum & ",$B$1:$B" & prevRowNum & ",$J$1:$J" & prevRowNum & ",0,0,-1)),0)"
    
    newRow.Range.Cells(1, colRealizedGL).Formula2 = _
    "=IF($D" & rowNum & "=" & """Sell""" & ",$I" & rowNum & "-($E" & rowNum & "*$J" & rowNum & "),"""")"
    
    On Error GoTo 0
    
    ClearForm False
    RefreshPortfolio False
    
    Call LogAction("AddTransaction", _
                   "Add new transaction", _
                   "Success", _
                   "Transaction added and portfolio refreshed successfully", _
                   "Transactions")
    
    MsgBox "Transaction added and portfolio refreshed successfully.", vbInformation, "Portfolio Tracker"
    Exit Sub

ColumnError:
    Call LogAction("AddTransaction", _
                   "Add new transaction", _
                   "Error", _
                   "Error while adding the transaction row or formulas", _
                   "Transactions")
    
    MsgBox "Error while adding the transaction row or formulas." & vbCrLf & _
           "Please check the Transactions table structure.", vbCritical, "Portfolio Tracker"

End Sub
