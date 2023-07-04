Attribute VB_Name = "Module3"
Sub FindGreatestValues()
    Dim lastRow As Long
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through the data to find the max values and tickers
    For i = 2 To lastRow
        ticker = Cells(i, 9).Value ' Ticker in column I
        percentChange = Cells(i, 11).Value ' Percent change in column K
        totalVolume = Cells(i, 12).Value ' Total volume in column L
        
        ' Check for max percent increase
        If percentChange > maxPercentIncrease Then
            maxPercentIncrease = percentChange
            maxPercentIncreaseTicker = ticker
        End If
        
        ' Check for max percent decrease
        If percentChange < maxPercentDecrease Then
            maxPercentDecrease = percentChange
            maxPercentDecreaseTicker = ticker
        End If
        
        ' Check for max total volume
        If totalVolume > maxTotalVolume Then
            maxTotalVolume = totalVolume
            maxTotalVolumeTicker = ticker
        End If
    Next i
    
    ' Output the results
    Cells(2, 16).Value = maxPercentIncreaseTicker
    Cells(3, 16).Value = maxPercentDecreaseTicker
    Cells(4, 16).Value = maxTotalVolumeTicker
    Cells(2, 17).Value = Format(maxPercentIncrease, "0.00%")
    Cells(3, 17).Value = Format(maxPercentDecrease, "0.00%")
    Cells(4, 17).Value = maxTotalVolume
End Sub

