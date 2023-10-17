Attribute VB_Name = "Module1"
Sub VbaChallenge()

For Each ws In Worksheets
        
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim SummaryRow As Long
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
                                
            ws.Cells(1, 9).Value = "Ticker"
            
            ws.Cells(1, 10).Value = "Yearly Change"
            
            ws.Cells(1, 11).Value = "Percent Change"
            
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            SummaryRow = 2
             
             For i = 2 To LastRow
             
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
                Ticker = ws.Cells(i, 1).Value
             
                OpeningPrice = ws.Cells(i, 3).Value
             
                ClosingPrice = ws.Cells(i, 6).Value
            
                YearlyChange = ClosingPrice - OpeningPrice
             
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice) * 100
             
             Else
                PercentChange = 0
             
             
             
             End If
                
                ws.Cells(SummaryRow, 9).Value = Ticker
                
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                
                ws.Cells(SummaryRow, 11).Value = PercentChange
                
                'ws.Cells(SummaryRow, 12).Value = Application.WorksheetFunction.Sum (ws.Range(ws.Cells(i - Volume + 1, 7), wsCells(i, 7)))
                
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                
                End If
                
                    SummaryRow = SummaryRow + 1
                
                TotalVolume = 0
                 
                 End If
                 
                 TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                 
                    Next i
                    
                ' Find the maximum percent increase
                Dim MaxPercentIncrease As Double
                Dim MaxPercentTicker As String
                
                MaxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K2:k" & SummaryRow))
                
                MaxPercentTicker = ws.Cells(Application.WorksheetFunction.Match(MaxPercentIncrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
                
                ws.Cells(2, 16).Value = MaxPercentTicker
                
                ws.Cells(2, 17).Value = MaxPercentIncrease
                
                ws.Cells(2, 17).NumberFormat = "0.00%"
                
                'Find the maxium percent decrease
                Dim MaxPercentDecrease As Double
                Dim MaxPercentDecreaseTicker As String
                
                'MaxPercentDecrease =
                Application.WorksheetFunction.Min (ws.Range("k2:k" & SummaryRow))
                'MaxPercentDecreaseTicker = ws.Cells(Application.WorksheetFunction.Match(MaxPercentDecrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
                ws.Cells(3, 16).Value = MaxPercentDecreaseTicker
                ws.Cells(3, 17).Value = MaxPercentDecrease
                ws.Cells(3, 17).NumberFormat = "0.00%"
                
                ' Find the stock with the greatest total volume
                Dim MaxTotalVolume As Double
                Dim MaxTotalVolumeTicker As String
                
                MaxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & SummaryRow))
                    'MaxTotalVolumeTicker = ws.Cells(Application.WorksheetFunction.Function.Match(MaxTotalVolume, ws.Range("L2:L" & SummaryRow), 0) + 1, 9).Value
                ws.Cells(4, 17).Value = MaxTotalVolume
                
                    MsgBox ("Got here")
                
                Exit For
                Next ws
                
                End Sub
                
                Sub AnalyzeStockData()
   
   ' Declare variables for spreadsheet, row index and various metrics
   Dim ws As Worksheet
   Dim last_row As Long
   Dim opening_price As Double
   Dim closing_price As Double
   Dim yearly_change As Double
   Dim percent_change As Double
   Dim total_volume As Double
   Dim max_increase As Double
   Dim max_decrease As Double
   Dim max_volume As Double
   Dim max_increase_ticker As String
   Dim max_decrease_ticker As String
   Dim max_volume_ticker As String
   Dim output_row As Long
   
   ' Initialize variables to trace lines of output
   output_row = 2
   
   ' Loop through all the worksheets in the workbook
   For Each ws In ThisWorkbook.Worksheets
       ' Find the last row of data in column A
       last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       
       ' Initialize variables for each worksheet
       opening_price = ws.Cells(2, 3).Value ' Assuming opening price is in column C (3)
       closing_price = ws.Cells(last_row, 6).Value ' Assuming closing price is in column F (6)
       total_volume = Application.WorksheetFunction.Sum(ws.Range("G2:G" & last_row)) ' Assuming volume is in column G (7)
       
       ' Calculate the change of year and the percentage change
       yearly_change = closing_price - opening_price
       If opening_price <> 0 Then
           percent_change = (yearly_change / opening_price) * 100
       Else
           percent_change = 0
       End If
       
       ' Output results
       ws.Cells(output_row, 9).Value = "Ticker"
       ws.Cells(output_row, 10).Value = "Yearly Change"
       ws.Cells(output_row, 11).Value = "Percent Change"
       ws.Cells(output_row, 12).Value = "Total Stock Volume"
       ws.Cells(output_row + 1, 9).Value = ws.Cells(2, 1).Value ' Assuming ticker is in column A (1)
       ws.Cells(output_row + 1, 10).Value = yearly_change
       ws.Cells(output_row + 1, 11).Value = percent_change
       ws.Cells(output_row + 1, 12).Value = total_volume
       
       ' Find and use conditional formatting for annual changes
       If yearly_change >= 0 Then
           ws.Cells(output_row + 1, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
       Else
           ws.Cells(output_row + 1, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
       End If
       
       ' Find the greatest % increase, % decrease, and total volume
       If percent_change > max_increase Then
           max_increase = percent_change
           max_increase_ticker = ws.Cells(2, 1).Value ' Assuming ticker is in column A (1)
       End If
       If percent_change < max_decrease Then
           max_decrease = percent_change
           max_decrease_ticker = ws.Cells(2, 1).Value ' Assuming ticker is in column A (1)
       End If
       If total_volume > max_volume Then
           max_volume = total_volume
           max_volume_ticker = ws.Cells(2, 1).Value ' Assuming ticker is in column A (1)
       End If
       
       ' Go to the next  row of results
       output_row = output_row + 2
   Next ws
   
   ' Print the greatest % increase, % decrease, and total volume in the last worksheet
   Dim summaryWs As Worksheet
   Set summaryWs = ThisWorkbook.Worksheets.Add
   summaryWs.Name = "Summary"
   summaryWs.Cells(1, 1).Value = "Greatest % Increase"
   summaryWs.Cells(2, 1).Value = "Greatest % Decrease"
   summaryWs.Cells(3, 1).Value = "Greatest Total Volume"
   summaryWs.Cells(1, 2).Value = max_increase_ticker
   summaryWs.Cells(2, 2).Value = max_decrease_ticker
   summaryWs.Cells(3, 2).Value = max_volume_ticker
   summaryWs.Cells(1, 3).Value = max_increase
   summaryWs.Cells(2, 3).Value = max_decrease
   summaryWs.Cells(3, 3).Value = max_volume
   
   
    
   
End Sub

                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                 
             
            
             
             

    
    

    
    

