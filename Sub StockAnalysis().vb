Sub StockAnalysis()

Dim ticker As String
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentageChange As Double
Dim totalVolume As Double

Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double

Dim i As Long

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    ' Set initial values for greatest increase, decrease, and volume
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    For i = 2 To Range("A" & Rows.Count).End(xlUp).Row
        ticker = Cells(i, 1).Value
        openingPrice = Cells(i, 3).Value
        closingPrice = Cells(i, 6).Value
        yearlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentageChange = (yearlyChange / openingPrice) * 100
        Else
            percentageChange = 0
        End If
        totalVolume = Cells(i, 7).Value

        Cells(i, 10).Value = ticker
        Cells(i, 11).Value = yearlyChange
        Cells(i, 12).Value = percentageChange
        Cells(i, 13).Value = totalVolume

        If yearlyChange > 0 Then
            Cells(i, 11).Interior.Color = vbGreen
        ElseIf yearlyChange < 0 Then
            Cells(i, 11).Interior.Color = vbRed
        End If

        If percentageChange > greatestIncrease Then
            greatestIncrease = percentageChange
        End If
        If percentageChange < greatestDecrease Then
            greatestDecrease = percentageChange
        End If
        If totalVolume > greatestVolume Then
            greatestVolume = totalVolume
        End If
    Next i

    Cells(2, 16).Value = "Greatest % Increase"
    Cells(3, 16).Value = "Greatest % Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"

    Cells(2, 17).Value = Application.WorksheetFunction.Index(Range("K:K"), Application.WorksheetFunction.Match(greatestIncrease, Range("L:L"), 0))
    Cells(3, 17).Value = Application.WorksheetFunction.Index(Range("K:K"), Application.WorksheetFunction.Match(greatestDecrease, Range("L:L"), 0))
    Cells(4, 17).Value = Application.WorksheetFunction.Index(Range("K:K"), Application.WorksheetFunction.Match(greatestVolume, Range("M:M"), 0))
Next ws

End Sub
