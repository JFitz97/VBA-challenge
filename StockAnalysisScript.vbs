Sub StockAnalysis()
      
      'Set initial variable for holding ticker name as well as ticker names for greatest % inc, Greatest % dec, and greatest total volume
      Dim Ticker As String
      Dim Greatest_Percent_Inc_Ticker As String
      Dim Greatest_Percent_Dec_Ticker As String
      Dim Greatest_Tot_Vol_Ticker As String
            
        'Set variables for remaining columns in summary table, also added variables for open price and close price as these are needed in calculation, initialize values
        Dim Yearly_Change, Percent_Change, Total_Stock_Volume, Open_Price, Close_Price, Greatest_Percent_Inc, Greatest_Percent_Dec, Greatest_Tot_Vol As Double
        Yearly_Change = 0
        Percent_Change = 0
        Total_Stock_Volume = 0
        Open_Price = 0
        Close_Price = 0
        Greatest_Percent_Inc = 0
        Greatest_Percent_Dec = 0
        Greatest_Tot_Vol = 0
        
     'Loop through sheets
      For Each ws In Worksheets

        'Generate prescribed column and row headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
            'Count number of rows
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Locate each Ticker in summary table
            Dim Summary_Row As Integer
            Summary_Row = 2
                    
            'Set initial open price
            Open_Price = ws.Cells(2, 3).Value
            
                    ' Loop through each row
                    For i = 2 To lastrow
            
                        'Check if next row has the same ticker symbol, if different then...
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        
                            'Set close price
                            Close_Price = ws.Cells(i, 6).Value
                            
                            'Calculate yearly change
                            Yearly_Change = (Close_Price - Open_Price)
                            
                            'Print yearly change in summary table
                            ws.Range("J" & Summary_Row).Value = Yearly_Change
                            
                            'Conditional format on yearly change
                            If (Yearly_Change > 0) Then
                                ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                                ElseIf (Yearly_Change < 0) Then
                                ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                            End If
                            
                            'Calculate % change
                            Percent_Change = (Yearly_Change / Open_Price) * 100
                            
                            'Print % change in summary table
                            ws.Range("K" & Summary_Row).Value = (Percent_Change & "%")
                            
                            'Conditional format on % change
                             If (Percent_Change > 0) Then
                                ws.Range("K" & Summary_Row).Interior.ColorIndex = 4
                                ElseIf (Percent_Change < 0) Then
                                ws.Range("K" & Summary_Row).Interior.ColorIndex = 3
                            End If
                            
                            'Set Ticker name in summary table
                            Ticker = ws.Cells(i, 1).Value
                            
                            'Add stock volume of common tickers
                            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                            
                            'Print ticker name in summary table
                            ws.Range("I" & Summary_Row).Value = Ticker
                            
                            'Print stock volume in summary table
                            ws.Range("L" & Summary_Row).Value = Total_Stock_Volume
                            
                            'Add new row to summary table
                            Summary_Row = Summary_Row + 1
                            
                        
                            'Reset stock volume for new ticker symbols (new rows)
                            Total_Stock_Volume = 0
                            
                            'Reset open price
                            Open_Price = ws.Cells(i + 1, 3).Value
                            
                        'If next row ticker symbol is the same...
                        Else
                        
                            'Calculate total stock volume
                            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                        End If
                        
                    Next i
                    
                   'Set greatest percent inc, greatest percent dec, and greatest total vol to 0
                   Greatest_Percent_Inc = 0
                   Greatest_Percent_Dec = 0
                   Greatest_Tot_Vol = 0
                   
                   For i = 2 To lastrow
                   'Do best/worst % change and greatest total volume calcs
                            If (ws.Cells(i, 11).Value > Greatest_Percent_Inc) Then
                                Greatest_Percent_Inc = ws.Cells(i, 11).Value
                                Greatest_Percent_Inc_Ticker = ws.Cells(i, 9).Value
                            End If
                            
                            If (ws.Cells(i, 11).Value < Greatest_Percent_Dec) Then
                                Greatest_Percent_Dec = ws.Cells(i, 11).Value
                                Greatest_Percent_Dec_Ticker = ws.Cells(i, 9).Value
                            End If
                            
                             'Determine ticker with greatest total volume
                            If (ws.Cells(i, 12).Value > Greatest_Tot_Vol) Then
                                Greatest_Tot_Vol = ws.Cells(i, 12).Value
                                Greatest_Tot_Vol_Ticker = ws.Cells(i, 9).Value
                            End If
                            
                   Next i
                            
                           'Print Greatest Total Inc/Dec and Greatest Total Volume
                            ws.Range("P2").Value = Greatest_Percent_Inc_Ticker
                            ws.Range("P3").Value = Greatest_Percent_Dec_Ticker
                            ws.Range("P4").Value = Greatest_Tot_Vol_Ticker
                            ws.Range("Q2").Value = (Greatest_Percent_Inc * 100 & "%")
                            ws.Range("Q3").Value = (Greatest_Percent_Dec * 100 & "%")
                            ws.Range("Q4").Value = Greatest_Tot_Vol
                            
      Next ws
      
End Sub

