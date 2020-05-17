Attribute VB_Name = "Module11"


'loop through stocks for 1 year

'output for each stock: ticker symbol, change.actual & change.percent from year.open to year.close, total (summed) stock volume in a new list to right of table w/ 1 entry for each ticker

'conditionally format positive change in green & negative change in red

'identify stock with greatest percent increase, greatest percent decrease, and greatest total volume (output ticker & value)

Sub summary()

    For Each ws In Worksheets
        
        
        
        
        
        'Set summary headers
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        
        'Set Max/Min row & col headers
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        
        'set scope of loop
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'declare variables total volume and summary row counters, initial open/final close, & Max/Min buckets
        Dim total_volume As Double
        Dim summary_row As Integer
        
        Dim initial_open As Double
        Dim final_close As Double
        
        Dim max_increase_ticker As String
        Dim max_increase_value As Double
        Dim max_decrease_ticker As String
        Dim max_decrease_value As Double
        Dim max_volume_ticker As String
        Dim max_volume_value As Double
        
        'initialize variables
        initial_open = ws.Cells(2, 3)
        
        total_volume = 0
        summary_row = 2
        
        max_increase_value = 0
        max_decrease_value = 0
        max_volume_value = 0
        
        
        'loop through rows in ws
        For i = 2 To last_row
            
            'Check for new ticker
            ticker = ws.Cells(i, 1)
            next_ticker = ws.Cells(i + 1, 1)
            
            If ticker = next_ticker Then
                'add volume to total volume counter
                total_volume = total_volume + ws.Cells(i, 7)
                
            Else
                'set final_close
                final_close = ws.Cells(i, 6)
                
                'set ticker
                ws.Cells(summary_row, 9) = ticker
                
                'calc change
                ws.Cells(summary_row, 10) = final_close - initial_open
                

                'conditional format yearly change cell color
                If ws.Cells(summary_row, 10) >= 0 Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                End If
                
                'calc % change and handle div/0 error
                If initial_open = 0 Then
                    ws.Cells(summary_row, 11) = 0
                Else
                    ws.Cells(summary_row, 11) = Format(ws.Cells(summary_row, 10) / initial_open, "Percent")
                End If
                
                'set total volume
                total_volume = total_volume + ws.Cells(i, 7)
                ws.Cells(summary_row, 12) = total_volume
                
                'check primacy of current total_volume against max_volume_value
                If ws.Cells(summary_row, 12) >= max_volume_value Then
                    max_volume_value = ws.Cells(summary_row, 12)
                    max_volume_ticker = ticker
                Else
                    max_volume_value = max_volume_value
                    max_volume_ticker = max_volume_ticker
                End If
                
                'check primacy of current %change against current max_increase_value
                If ws.Cells(summary_row, 11) >= max_increase_value Then
                    max_increase_value = ws.Cells(summary_row, 11)
                    max_increase_ticker = ticker
                Else
                    max_increase_value = max_increase_value
                    max_increase_ticker = max_increase_ticker
                End If
                
                'check primacy of current %change against current max_decrease_value
                If ws.Cells(summary_row, 11) <= max_decrease_value Then
                    max_decrease_value = ws.Cells(summary_row, 11)
                    max_decrease_ticker = ticker
                Else
                    max_decrease_value = max_decrease_value
                    max_decrease_ticker = max_decrease_ticker
                End If
                
                'reset variables for next ticker
                total_volume = 0
                initial_open = ws.Cells(i + 1, 3)
                summary_row = summary_row + 1
                            
            End If
            
        Next i
        
        'Print max/mins to max/min table
        ws.Cells(2, 16) = Format(max_increase_ticker, "percent")
        ws.Cells(2, 17) = Format(max_increase_value, "percent")
        ws.Cells(3, 16) = Format(max_decrease_ticker, "percent")
        ws.Cells(3, 17) = Format(max_decrease_value, "percent")
        ws.Cells(4, 16) = max_volume_ticker
        ws.Cells(4, 17) = max_volume_value
        
        

    Next ws
    
End Sub
