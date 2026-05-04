Option Explicit

Sub RefreshPortfolioButton()
    RefreshPortfolio True
End Sub

Public Sub RefreshPortfolio(Optional ByVal ShowMessage As Boolean = True)

    Dim wsDash As Worksheet
    Dim warningMessage As String
    
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    On Error GoTo CleanFail
    
    Application.ScreenUpdating = False
    
    Application.Calculate
    UpdateSummaryReport
    Application.Calculate
    
    warningMessage = BuildPortfolioWarnings()
    
    wsDash.Range("B3").Value = Now
    wsDash.Range("B3").NumberFormat = "dd.mm.yyyy hh:mm"
    
    Application.ScreenUpdating = True
    
    If warningMessage = "" Then
        Call LogAction("RefreshPortfolio", _
                       "Refresh dashboard and summary", _
                       "Success", _
                       "Portfolio refreshed successfully", _
                       "Dashboard")
    Else
        Call LogAction("RefreshPortfolio", _
                       "Refresh dashboard and summary", _
                       "Warning", _
                       "Portfolio refreshed with warnings", _
                       "Dashboard")
    End If
    
    If ShowMessage Then
        If warningMessage = "" Then
            MsgBox "Portfolio refreshed successfully.", vbInformation, "Portfolio Tracker"
        Else
            MsgBox warningMessage, vbExclamation, "Portfolio Tracker Warning"
        End If
    End If
    
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    
    Call LogAction("RefreshPortfolio", _
                   "Refresh dashboard and summary", _
                   "Error", _
                   "Error while refreshing the portfolio", _
                   "Dashboard")
    
    MsgBox "Error while refreshing the portfolio.", vbCritical, "Portfolio Tracker"

End Sub

Public Function BuildPortfolioWarnings() As String

    Dim msg As String
    Dim part As String
    
    msg = ""
    
    part = GetMissingPricesWarning()
    If part <> "" Then msg = msg & part & vbCrLf & vbCrLf
    
    part = GetNegativeHoldingsWarning()
    If part <> "" Then msg = msg & part & vbCrLf & vbCrLf
    
    part = GetConcentrationWarning()
    If part <> "" Then msg = msg & part & vbCrLf & vbCrLf
    
    If Right(msg, 4) = vbCrLf & vbCrLf Then
        msg = Left(msg, Len(msg) - 4)
    End If
    
    BuildPortfolioWarnings = msg

End Function

Public Function GetMissingPricesWarning() As String

    Dim wsPrices As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim missingList As String
    Dim tickerName As String
    Dim priceValue As Variant
    
    Set wsPrices = ThisWorkbook.Worksheets("Prices")
    
    lastRow = wsPrices.Cells(wsPrices.Rows.Count, "A").End(xlUp).Row
    missingList = ""
    
    For i = 2 To lastRow
        tickerName = Trim(wsPrices.Cells(i, "A").Value)
        priceValue = wsPrices.Cells(i, "C").Value
        
        If tickerName <> "" Then
            If Trim(priceValue & "") = "" Or priceValue = 0 Then
                missingList = missingList & "- " & tickerName & vbCrLf
            End If
        End If
    Next i
    
    If missingList <> "" Then
        GetMissingPricesWarning = "Missing current price for:" & vbCrLf & missingList
    Else
        GetMissingPricesWarning = ""
    End If

End Function

Public Function GetNegativeHoldingsWarning() As String

    Dim wsHold As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim negativeList As String
    Dim tickerName As String
    Dim qtyValue As Variant
    
    Set wsHold = ThisWorkbook.Worksheets("Holdings")
    
    lastRow = wsHold.Cells(wsHold.Rows.Count, "A").End(xlUp).Row
    negativeList = ""
    
    For i = 2 To lastRow
        tickerName = Trim(wsHold.Cells(i, "A").Value)
        qtyValue = wsHold.Cells(i, "E").Value
        
        If tickerName <> "" And IsNumeric(qtyValue) Then
            If CDbl(qtyValue) < 0 Then
                negativeList = negativeList & "- " & tickerName & " (" & qtyValue & ")" & vbCrLf
            End If
        End If
    Next i
    
    If negativeList <> "" Then
        GetNegativeHoldingsWarning = "Negative holding detected for:" & vbCrLf & negativeList
    Else
        GetNegativeHoldingsWarning = ""
    End If

End Function

Public Function GetConcentrationWarning() As String

    Dim wsHold As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalMarketValue As Double
    Dim assetValue As Double
    Dim concentrationList As String
    Dim tickerName As String
    
    Set wsHold = ThisWorkbook.Worksheets("Holdings")
    
    lastRow = wsHold.Cells(wsHold.Rows.Count, "A").End(xlUp).Row
    concentrationList = ""
    totalMarketValue = 0
    
    For i = 2 To lastRow
        If IsNumeric(wsHold.Cells(i, "H").Value) Then
            totalMarketValue = totalMarketValue + CDbl(wsHold.Cells(i, "H").Value)
        End If
    Next i
    
    If totalMarketValue <= 0 Then
        GetConcentrationWarning = ""
        Exit Function
    End If
    
    For i = 2 To lastRow
        tickerName = Trim(wsHold.Cells(i, "A").Value)
        
        If tickerName <> "" And IsNumeric(wsHold.Cells(i, "H").Value) Then
            assetValue = CDbl(wsHold.Cells(i, "H").Value)
            
            If assetValue / totalMarketValue > 0.4 Then
                concentrationList = concentrationList & "- " & tickerName & " (" & Format(assetValue / totalMarketValue, "0.0%") & ")" & vbCrLf
            End If
        End If
    Next i
    
    If concentrationList <> "" Then
        GetConcentrationWarning = "High concentration risk detected:" & vbCrLf & concentrationList
    Else
        GetConcentrationWarning = ""
    End If

End Function
