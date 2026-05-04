Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim wsPrices As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim selectedTicker As String

    If Intersect(Target, Me.Range("P4")) Is Nothing Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False

    Me.Range("P5").ClearContents
    Me.Range("P8").ClearContents

    selectedTicker = UCase(Trim(Me.Range("P4").Value))

    If selectedTicker <> "" Then
        Set wsPrices = ThisWorkbook.Worksheets("Prices")
        lastRow = wsPrices.Cells(wsPrices.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If UCase(Trim(wsPrices.Cells(i, 1).Value)) = selectedTicker Then
                Me.Range("P5").Value = wsPrices.Cells(i, 2).Value
                Me.Range("P8").Value = wsPrices.Cells(i, 3).Value
                Exit For
            End If
        Next i
    End If

SafeExit:
    Application.EnableEvents = True

End Sub
