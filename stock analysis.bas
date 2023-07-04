Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    outputRow = 2 ' Starting row for output
    
    ' Set initial values
    ticker = Cells(2, 1).Value
    openingPrice = Cells(2, 3).Value
    totalVolume = 0
    
    ' Loop through the data
    For i = 2 To lastRow
        ' Check if the ticker symbol has changed
        If Cells(i, 1).Value <> ticker Then
            ' Output the results for the previous ticker
            closingPrice = Cells(i - 1, 6).Value
            yearlyChange = closingPrice - openingPrice
            percentChange = yearlyChange / openingPrice
            
            ' Output the results next to the ticker symbol
            Cells(outputRow, 9).Value = ticker
            Cells(outputRow, 10).Value = yearlyChange
            Cells(outputRow, 11).Value = Format(percentChange, "0.00%")
            Cells(outputRow, 12).Value = totalVolume
            
            ' Conditional formatting based on percent change
            If percentChange > 0 Then
                Cells(outputRow, 10).Interior.ColorIndex = 4 ' Green
            ElseIf percentChange < 0 Then
                Cells(outputRow, 10).Interior.ColorIndex = 3 ' Red
            End If
            
            ' Move to the next row for output
            outputRow = outputRow + 1
            
            ' Reset variables for the new ticker
            ticker = Cells(i, 1).Value
            openingPrice = Cells(i, 3).Value
            totalVolume = 0
        End If
        
        ' Accumulate the total stock volume
        totalVolume = totalVolume + Cells(i, 7).Value
    Next i
    
    ' Output the results for the last ticker
    closingPrice = Cells(lastRow, 6).Value
    yearlyChange = closingPrice - openingPrice
    percentChange = yearlyChange / openingPrice
    
    ' Output the results next to the ticker symbol
    Cells(outputRow, 9).Value = ticker
    Cells(outputRow, 10).Value = yearlyChange
    Cells(outputRow, 11).Value = Format(percentChange, "0.00%")
    Cells(outputRow, 12).Value = totalVolume
    
    ' Conditional formatting based on percent change
    If percentChange > 0 Then
        Cells(outputRow, 10).Interior.ColorIndex = 4 ' Green
    ElseIf percentChange < 0 Then
        Cells(outputRow, 10).Interior.ColorIndex = 3 ' Red
    End If
    
    ' Additional formatting or actions can be performed based on your requirements
    
End Sub



