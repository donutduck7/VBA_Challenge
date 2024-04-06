# VBA_Challenge

CODE:
Sub StockMarketAnalysis()
    ' Loop through each worksheet in the workbook
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        AnalyzeStockData ws
    Next ws
End Sub

Sub AnalyzeStockData(ByRef ws As Worksheet)
    ' Declare variables for calculations
    Dim lastRow As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim startPrice As Double
    Dim endPrice As Double
    Dim summaryTableRow As Integer
    
    ' Initialize summary table row
    summaryTableRow = 2
    
    ' Prepare headers for the summary table
    With ws
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Yearly Change"
        .Cells(1, 11).Value = "Percent Change"
        .Cells(1, 12).Value = "Total Stock Volume"
        
        ' Find the last row of data in column A
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the totalVolume
        totalVolume = 0
        
        ' Loop through all rows of data
        For i = 2 To lastRow
            ' Check if we are still within the same ticker symbol
            If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                ticker = .Cells(i, 1).Value
                endPrice = .Cells(i, 6).Value
                totalVolume = totalVolume + .Cells(i, 7).Value
                yearlyChange = endPrice - startPrice
                If startPrice <> 0 Then
                    percentChange = (yearlyChange / startPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Print the ticker and its metrics to the summary table
                .Cells(summaryTableRow, 9).Value = ticker
                .Cells(summaryTableRow, 10).Value = yearlyChange
                .Cells(summaryTableRow, 11).Value = percentChange
                .Cells(summaryTableRow, 12).Value = totalVolume
                
                ' Move to the next row in the summary table
                summaryTableRow = summaryTableRow + 1
                
                ' Reset totalVolume for the next ticker
                totalVolume = 0
                
                ' Set startPrice for the next ticker if not the last row
                If i + 1 <= lastRow Then
                    startPrice = .Cells(i + 1, 3).Value
                End If
            Else
                ' Add to the totalVolume if the same ticker
                totalVolume = totalVolume + .Cells(i, 7).Value
            End If
            
            ' Set the startPrice if it's the first row
            If i = 2 Then
                startPrice = .Cells(i, 3).Value
            End If
        Next i
    End With
    
    ' Call the subroutine to apply conditional formatting
    Call ApplyConditionalFormatting(ws)
    
    ' Call the subroutine to find greatest increase, decrease, and total volume
    Call FindGreatestMetrics(ws)
End Sub

Sub ApplyConditionalFormatting(ByRef ws As Worksheet)
    ' Apply your conditional formatting here
    ' Placeholder for your conditional formatting code
End Sub

Sub FindGreatestMetrics(ByRef ws As Worksheet)
    ' Calculate and find the greatest % increase, % decrease, and total volume here
    ' Placeholder for your calculations for greatest metrics code
End Sub
