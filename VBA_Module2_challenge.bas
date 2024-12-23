Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()

    ' Speed optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Initiate variables to roll through the Sheets
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim i As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Initialize variables for greatest values
    greatestIncrease = -1
    greatestDecrease = 1
    greatestVolume = 0

    ' Get active workbook
    Set wb = ThisWorkbook ' Ensure the workbook is set correctly
    
    ' Define specific sheets to loop through
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")

    ' Loop through specified worksheets
    For Each ws In wb.Worksheets
        If Not IsError(Application.Match(ws.Name, sheetNames, 0)) Then
            ' Initialize the summary table starting row
            summaryRow = 2
            
            ' Find the last row with data in the worksheet
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Set up headers in the worksheet
            With ws
                .Cells(1, 9).Value = "Ticker"
                .Cells(1, 10).Value = "Quarterly Change"
                .Cells(1, 11).Value = "Percent Change"
                .Cells(1, 12).Value = "Total Stock Volume"
                .Cells(1, 15).Value = "Ticker"
                .Cells(1, 16).Value = "Value"
                .Cells(2, 14).Value = "Greatest % Increase"
                .Cells(3, 14).Value = "Greatest % Decrease"
                .Cells(4, 14).Value = "Greatest Total Volume"
                .Columns("I:L").ColumnWidth = 14
                .Columns("I:L").HorizontalAlignment = xlCenter
            End With

            ' Increase the cell width of N2, N3, and N4 by 9 units
            ws.Range("N2:N4").ColumnWidth = ws.Range("N2:N4").ColumnWidth + 9

            ' Loop through each row in the worksheet
            For i = 2 To lastRow
                ' Check if the ticker symbol changes (indicating a new stock)
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    If i > 2 Then
                        ' Calculate quarterly change and percentage change for the previous stock
                        closePrice = ws.Cells(i - 1, 6).Value
                        quarterlyChange = closePrice - openPrice
                        If openPrice <> 0 Then
                            percentChange = ((quarterlyChange / openPrice) * 100) / 100 ' Adjusting the percentage by dividing by 100
                        Else
                            percentChange = 0
                        End If
                        
                        ' Output the results to the worksheet
                        ws.Cells(summaryRow, 9).Value = ticker
                        ws.Cells(summaryRow, 10).Value = quarterlyChange
                        ws.Cells(summaryRow, 11).Value = percentChange
                        ws.Cells(summaryRow, 12).Value = totalVolume
                        
                        ' Format the percentage change as a percentage with two decimals
                        ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                        
                        ' Move to the next row in the worksheet
                        summaryRow = summaryRow + 1
                        
                        ' Check for greatest percentage increase
                        If percentChange > greatestIncrease Then
                            greatestIncrease = percentChange
                            greatestIncreaseTicker = ticker
                        End If
                        
                        ' Check for greatest percentage decrease
                        If percentChange < greatestDecrease Then
                            greatestDecrease = percentChange
                            greatestDecreaseTicker = ticker
                        End If
                        
                        ' Check for greatest total volume
                        If totalVolume > greatestVolume Then
                            greatestVolume = totalVolume
                            greatestVolumeTicker = ticker
                        End If
                    End If
                    
                    ' Reset variables for the new stock
                    ticker = ws.Cells(i, 1).Value
                    openPrice = ws.Cells(i, 3).Value
                    totalVolume = 0
                End If
                
                ' Accumulate the total volume for the current stock
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            Next i
            
            ' Output the results for the last stock
            closePrice = ws.Cells(lastRow, 6).Value
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = ((quarterlyChange / openPrice) * 100) / 100 ' Adjusting the percentage by dividing by 100
            Else
                percentChange = 0
            End If
            
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = quarterlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = totalVolume
            
            ' Format the percentage change as a percentage with two decimals
            ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
            
            ' Check the last stock for greatest values
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
            
            ' Output the greatest values for the worksheet
            ws.Cells(2, 15).Value = greatestIncreaseTicker
            ws.Cells(2, 16).Value = greatestIncrease
            ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(3, 15).Value = greatestDecreaseTicker
            ws.Cells(3, 16).Value = greatestDecrease
            ws.Cells(3, 16).NumberFormat = "0.00%"
            ws.Cells(4, 15).Value = greatestVolumeTicker
            ws.Cells(4, 16).Value = greatestVolume
            
            ' Apply conditional formatting to column 10 (Quarterly Change)
            Dim rngQuarterlyChange As Range
            Set rngQuarterlyChange = ws.Range("J2:J" & summaryRow - 1)
            rngQuarterlyChange.FormatConditions.Delete
            With rngQuarterlyChange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
                .Interior.Color = RGB(0, 255, 0) ' Green
            End With
            With rngQuarterlyChange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
                .Interior.Color = RGB(255, 0, 0) ' Red
            End With
            
            ' Apply conditional formatting to column 11 (Percent Change)
            Dim rngPercentChange As Range
            Set rngPercentChange = ws.Range("K2:K" & summaryRow - 1)
            rngPercentChange.FormatConditions.Delete
            With rngPercentChange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
                .Interior.Color = RGB(0, 255, 0) ' Green
            End With
            With rngPercentChange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
                .Interior.Color = RGB(255, 0, 0) ' Red
            End With
        End If
    Next ws

    ' Re-enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

