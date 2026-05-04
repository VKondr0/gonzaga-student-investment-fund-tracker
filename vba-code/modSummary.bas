Option Explicit

Public Sub UpdateSummaryReport()

    Dim wsDash As Worksheet
    Dim wsSum As Worksheet
    Dim wsHold As Worksheet
    Dim wsTrans As Worksheet
    
    Dim lastRowHold As Long
    Dim lastRowTrans As Long
    Dim i As Long
    
    Dim netInvestedCapital As Double
    Dim currentPortfolioValue As Double
    Dim totalCostBasis As Double
    Dim totalUnrealizedGL As Double
    Dim totalUnrealizedGLPct As Double
    
    Dim numberOfAssets As Long
    
    Dim largestHolding As String
    Dim largestHoldingPct As Double
    Dim bestPerformingAsset As String
    Dim worstPerformingAsset As String
    Dim topGainer As String
    Dim topLoser As String
    
    Dim largestValue As Double
    Dim bestPerfValue As Double
    Dim worstPerfValue As Double
    Dim topGainerValue As Double
    Dim topLoserValue As Double
    
    Dim tickerName As String
    Dim transactionType As String
    
    Dim marketValue As Double
    Dim costBasis As Double
    Dim unrealizedGL As Double
    Dim unrealizedGLPct As Double
    Dim dayChangePct As Double
    Dim weightPct As Double
    Dim netQty As Double
    
    Dim diversificationStatus As String
    Dim diversificationComment As String
    
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set wsHold = ThisWorkbook.Worksheets("Holdings")
    Set wsTrans = ThisWorkbook.Worksheets("Transactions")
    
    netInvestedCapital = 0
    currentPortfolioValue = 0
    totalCostBasis = 0
    totalUnrealizedGL = 0
    totalUnrealizedGLPct = 0
    numberOfAssets = 0
    
    largestHolding = ""
    largestHoldingPct = 0
    bestPerformingAsset = ""
    worstPerformingAsset = ""
    topGainer = ""
    topLoser = ""
    
    largestValue = -1
    bestPerfValue = -1E+99
    worstPerfValue = 1E+99
    topGainerValue = -1E+99
    topLoserValue = 1E+99
    
    ' Net Invested Capital from Transactions
    lastRowTrans = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRowTrans
        transactionType = Trim(wsTrans.Cells(i, "D").Value)
        
        If IsNumeric(wsTrans.Cells(i, "I").Value) Then
            If UCase(transactionType) = "BUY" Then
                netInvestedCapital = netInvestedCapital + CDbl(wsTrans.Cells(i, "I").Value)
            ElseIf UCase(transactionType) = "SELL" Then
                netInvestedCapital = netInvestedCapital - CDbl(wsTrans.Cells(i, "I").Value)
            End If
        End If
    Next i
    
    ' Holdings columns used:
    ' A = Ticker
    ' E = Net Quantity
    ' H = Market Value
    ' I = Unrealized Gain/Loss
    ' J = Cost Basis
    ' K = Unrealized Gain/Loss %
    ' N = Day Change %
    ' O = Weight %
    lastRowHold = wsHold.Cells(wsHold.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRowHold
        tickerName = Trim(wsHold.Cells(i, "A").Value)
        
        If tickerName <> "" Then
            
            netQty = 0
            marketValue = 0
            costBasis = 0
            unrealizedGL = 0
            unrealizedGLPct = 0
            dayChangePct = 0
            weightPct = 0
            
            If IsNumeric(wsHold.Cells(i, "E").Value) Then
                netQty = CDbl(wsHold.Cells(i, "E").Value)
            End If
            
            If IsNumeric(wsHold.Cells(i, "H").Value) Then
                marketValue = CDbl(wsHold.Cells(i, "H").Value)
            End If
            
            If IsNumeric(wsHold.Cells(i, "J").Value) Then
                costBasis = CDbl(wsHold.Cells(i, "J").Value)
            End If
            
            If IsNumeric(wsHold.Cells(i, "I").Value) Then
                unrealizedGL = CDbl(wsHold.Cells(i, "I").Value)
            End If
            
            If IsNumeric(wsHold.Cells(i, "K").Value) Then
                unrealizedGLPct = CDbl(wsHold.Cells(i, "K").Value)
            End If
            
            If IsNumeric(wsHold.Cells(i, "N").Value) Then
                dayChangePct = CDbl(wsHold.Cells(i, "N").Value)
            End If
            
            If IsNumeric(wsHold.Cells(i, "O").Value) Then
                weightPct = CDbl(wsHold.Cells(i, "O").Value)
            End If
            
            If netQty > 0 Then
                numberOfAssets = numberOfAssets + 1
            End If
            
            currentPortfolioValue = currentPortfolioValue + marketValue
            totalCostBasis = totalCostBasis + costBasis
            totalUnrealizedGL = totalUnrealizedGL + unrealizedGL
            
            If marketValue > largestValue Then
                largestValue = marketValue
                largestHolding = tickerName
            End If
            
            If weightPct > largestHoldingPct Then
                largestHoldingPct = weightPct
            End If
            
            If unrealizedGLPct > bestPerfValue Then
                bestPerfValue = unrealizedGLPct
                bestPerformingAsset = tickerName
            End If
            
            If unrealizedGLPct < worstPerfValue Then
                worstPerfValue = unrealizedGLPct
                worstPerformingAsset = tickerName
            End If
            
            If dayChangePct > topGainerValue Then
                topGainerValue = dayChangePct
                topGainer = tickerName
            End If
            
            If dayChangePct < topLoserValue Then
                topLoserValue = dayChangePct
                topLoser = tickerName
            End If
            
        End If
    Next i
    
    If totalCostBasis <> 0 Then
        totalUnrealizedGLPct = totalUnrealizedGL / totalCostBasis
    Else
        totalUnrealizedGLPct = 0
    End If
    
    If numberOfAssets = 0 Then
        largestHolding = "No active holdings"
        largestHoldingPct = 0
        bestPerformingAsset = "N/A"
        worstPerformingAsset = "N/A"
        topGainer = "N/A"
        topLoser = "N/A"
        diversificationStatus = "No Holdings"
        diversificationComment = "Portfolio currently has no active holdings."
    Else
        If largestHoldingPct <= 0.35 Then
            diversificationStatus = "Good"
            diversificationComment = "Portfolio is well diversified."
        ElseIf largestHoldingPct <= 0.5 Then
            diversificationStatus = "Moderate"
            diversificationComment = "Portfolio has moderate concentration risk."
        Else
            diversificationStatus = "Limited"
            diversificationComment = "Portfolio has limited diversification."
        End If
    End If
    
    wsSum.Range("B2").Value = wsDash.Range("B3").Value
    wsSum.Range("B3").Value = netInvestedCapital
    wsSum.Range("B4").Value = currentPortfolioValue
    wsSum.Range("B5").Value = totalCostBasis
    wsSum.Range("B6").Value = totalUnrealizedGL
    wsSum.Range("B7").Value = totalUnrealizedGLPct
    wsSum.Range("B8").Value = numberOfAssets
    wsSum.Range("B9").Value = largestHolding
    wsSum.Range("B10").Value = largestHoldingPct
    wsSum.Range("B11").Value = bestPerformingAsset
    wsSum.Range("B12").Value = worstPerformingAsset
    wsSum.Range("B13").Value = topGainer
    wsSum.Range("B14").Value = topLoser
    wsSum.Range("B15").Value = diversificationStatus
    wsSum.Range("B16").Value = diversificationComment
    
    wsSum.Range("B2").NumberFormat = "dd.mm.yyyy hh:mm"
    wsSum.Range("B3:B6").NumberFormat = "[$$-en-US]#,##0.00"
    wsSum.Range("B7").NumberFormat = "0.00%"
    wsSum.Range("B8").NumberFormat = "0"
    wsSum.Range("B10").NumberFormat = "0.00%"

End Sub
