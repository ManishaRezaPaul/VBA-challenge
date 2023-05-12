Attribute VB_Name = "Module1"
Sub VBAmodulechallenge()
    
    For Each ws In Worksheets
    
    
        'Set variables for the ticker symbol
        Dim ticker As String
        Dim ticker_greatest_percent_increase As String
        Dim ticker_greatest_percent_decrease As String
        Dim ticker_greatest_total_volume As String
    
        'Set variables for rows to identify opening price and closing price
        Dim firstrow As Long
        Dim i As Long
        Dim lastrow As Long
    
        'Set variables for the opening and closing year, year change
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        
    
        'Set variables for the year change percentage, greatest % increase and greatest % decrease
        Dim yearly_change_percent As Double
        Dim greatest_percent_increase As Double
        Dim greatest_percent_decrease As Double
        year_change_percent = 0
        greatest_percent_increase = 0
        greatest_percent_decrease = 0

    
        'Set a variable for the stock volume
        Dim stockvolume As Double
        Dim totalstockvolume As Double
        Dim greatest_total_volume As Double
        totalstockvolume = 0
    
        'Set location for the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
    
        'Set column header for the summary table
        ws.Range("J1").Value = "Ticker symbol"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Ticker symbol"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        
    
        'Consider the last row of the stock in that year
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
        'Loop through to check the yearly change in the stocks
        For i = 2 To lastrow
                'Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                'Check if it's a new ticker symbol
                If ws.Cells(i - 1, 1).Value <> ticker Then
                    
                    'Store the first row and opening price and stock volume for the new ticker symbol
                    firstrow = i
                    opening_price = ws.Cells(i, 3).Value
                 End If
                
                        
                'Check if it's the last row of the ticker symbol
                If ws.Cells(i + 1, 1).Value <> ticker Then
                    
                    'Store the closing price
                    closing_price = ws.Cells(i, 6).Value
                    
                    'Calculate the yearly change only when it's the last row of the ticker symbol
                    yearly_change = closing_price - opening_price
                    
                    'Calculate the year_change
                    yearly_change_percent = yearly_change / opening_price
                    
                    'Calculate stock volume
                    totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                    
                    
                    'Print the ticker in the Summary Table
                    ws.Range("J" & Summary_Table_Row).Value = ticker
        
                    'Print the yearly_change to the Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = yearly_change
                    ws.Range("K" & Summary_Table_Row).Style = "Currency"
                    
                    'Print the yearly_change_percent to the Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = yearly_change_percent
                    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                    'Print the stockvolume to the Summary Table
                    ws.Range("M" & Summary_Table_Row).Value = totalstockvolume
                    
                      'Apply conditional formatting when yearly_change is positive or negative
                        If yearly_change > 0 Then
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                        ElseIf yearly_change < 0 Then
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                        Else
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = xlNone
                        End If
                    
                        
                         'Adding functionality to return the stock with the "Greatest % increase" and "Greatest % decrease"
                        If yearly_change_percent > greatest_percent_increase Then
                            greatest_percent_increase = yearly_change_percent
                            ticker_greatest_percent_increase = ticker
                            
                        ElseIf yearly_change_percent < greatest_percent_decrease Then
                            greatest_percent_decrease = yearly_change_percent
                            ticker_greatest_percent_decrease = ticker
                        End If
                                
                        'Adding functionality to return the stock with the "Greatest total volume"
                        If totalstockvolume > greatest_total_volume Then
                            greatest_total_volume = totalstockvolume
                            ticker_greatest_total_volume = ticker
                        End If
                        
                    'Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    'Reset values to 0
                    yearly_change = 0
                    yearly_change_percent = 0
                    totalstockvolume = 0
                     
                Else
    
                    'Continue calculating the values
                    year_change = year_change + (closing_price - opening_price)
                    year_change_percent = year_change / opening_price
                    totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                         
                  End If
                  
                  
        Next i
        
        'Print the greatest values to the summary table
        ws.Range("Q2").Value = ticker_greatest_percent_increase
        ws.Range("R2").Value = greatest_percent_increase
        ws.Range("Q3").Value = ticker_greatest_percent_decrease
        ws.Range("R3").Value = greatest_percent_decrease
        ws.Range("Q4").Value = ticker_greatest_total_volume
        ws.Range("R4").Value = greatest_total_volume
        ws.Range("R2:R3").NumberFormat = "0.00%"
        ws.Range("J:R").Columns.AutoFit

    Next ws


End Sub

