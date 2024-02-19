Sub CalculateYearlyMetricsWithConditionalFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim volume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim yearlyChangeColumn As Range
    Dim percentageChangeColumn As Range
    
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Add columns for calculated values
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        ws.Cells(1, 11).Value = "Yearly Change ($)"
        ws.Cells(1, 12).Value = "Percent Change"
        
        ' Loop through each row
        For i = 2 To lastRow
            ' Read data from current row
            ticker = ws.Cells(i, 1).Value
            volume = ws.Cells(i, 7).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            
            ' Calculate yearly metrics
            yearlyChange = closePrice - openPrice
            percentageChange = yearlyChange / openPrice * 100
            totalVolume = totalVolume + volume
            
            ' Write data to columns
            ws.Cells(i, 9).Value = ticker
            ws.Cells(i, 10).Value = volume
            ws.Cells(i, 11).Value = yearlyChange
            ws.Cells(i, 12).Value = percentageChange
            
            ' Apply conditional formatting
            Set yearlyChangeColumn = ws.Cells(i, 11)
            Set percentageChangeColumn = ws.Cells(i, 12)
            ApplyConditionalFormatting yearlyChangeColumn
            ApplyConditionalFormatting percentageChangeColumn
            
            ' Track greatest increase, decrease, and volume
            If yearlyChange > greatestIncrease Then
                greatestIncrease = yearlyChange
                greatestIncreaseTicker = ticker
            End If
            If yearlyChange < greatestDecrease Then
                greatestDecrease = yearlyChange
                greatestDecreaseTicker = ticker
            End If
            If volume > greatestVolume Then
                greatestVolume = volume
                greatestVolumeTicker = ticker
            End If
        Next i
        
        ' Write calculated values to summary rows
        ws.Cells(lastRow + 2, 9).Value = "Greatest % Increase"
        ws.Cells(lastRow + 2, 10).Value = greatestIncreaseTicker
        ws.Cells(lastRow + 2, 11).Value = greatestIncrease
        ws.Cells(lastRow + 3, 9).Value = "Greatest % Decrease"
        ws.Cells(lastRow + 3, 10).Value = greatestDecreaseTicker
        ws.Cells(lastRow + 3, 11).Value = greatestDecrease
        ws.Cells(lastRow + 4, 9).Value = "Greatest Total Volume"
        ws.Cells(lastRow + 4, 10).Value = greatestVolumeTicker
        ws.Cells(lastRow + 4, 11).Value = greatestVolume
    Next ws
End Sub

Sub ApplyConditionalFormatting(rng As Range)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0) ' Green for positive change
    End With
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0) ' Red for negative change
    End With
End Sub
