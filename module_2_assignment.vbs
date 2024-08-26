Sub QuarterlyStocks()
    Dim ws As Worksheet
    For Each ws In Worksheets
        
        'set column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
            ws.Range("J:J").ColumnWidth = 15 'Column width
        ws.Range("K1").Value = "Percent Change"
            ws.Range("K:K").ColumnWidth = 14 'Column width
        ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("L:L").ColumnWidth = 17 'Column width
        ws.Range("J:J").NumberFormat = "0.00" 'number format
        ws.Range("O:O").ColumnWidth = 20 'Column width
        ws.Range("O2").value = "Greatest % Increase"
        ws.Range("O3").value = "Greatest % Decrease"
        ws.Range("O4").value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'List Ticker Symbols
        Dim i As Long
        Dim a_row_count As Long
            a_row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim unique_ticker As String
        Dim I_row_tracker As Long
             I_row_tracker = 2
        Dim current_open As Double, last_close As Double, quarterly_change As Double
        Dim total_volume As LongLong

        For i = 2 To a_row_count
            Dim current_ticker As String
            current_ticker = ws.Cells(i, 1).Value
            
            If current_ticker <> unique_ticker Then 'new cell value
                unique_ticker = current_ticker
                
                
                ws.Cells(I_row_tracker, 9).Value = unique_ticker 'Add each unique stock ticker
                If i > 2 Then
                    last_close = ws.Cells(i - 1, 6).value
                    quarterly_change = last_close - current_open
                    ws.Cells(I_row_tracker - 1, 10).Value = quarterly_change 'Quarterly Change column
                    
                    'change background color
                    if (quarterly_change) > 0 Then
                        ws.Cells(I_row_tracker - 1, 10).Interior.ColorIndex = 4 'Green Background
                    elseif (quarterly_change) < 0 Then
                        ws.Cells(I_row_tracker - 1, 10).Interior.ColorIndex = 3 'Red Background
                    End If
                     
                    'calculate percent change
                    if (quarterly_change > 0 OR quarterly_change < 0) Then
                        ws.Cells(I_row_tracker - 1, 11).Value = (WorksheetFunction.Round((quarterly_change/current_open)*100,2)) &"%" 'Percent Change column
                    else
                        ws.Cells(I_row_tracker - 1, 11).Value = 0 &"%"
                    End If
                    ws.Cells(I_row_tracker - 1, 12).Value = total_volume
                    total_volume = 0 'reset
                End If
                total_volume = total_volume + ws.Cells(i,7).value
                current_open = ws.Cells(i, 3).value
                I_row_tracker = I_row_tracker + 1 'increment
                 else
                     total_volume = total_volume + ws.Cells(i,7).value
            End If
        Next i 
        total_volume = 0 'reset

        Dim max_percent_increase As Double, max_percent_decrease As Double
            max_percent_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
            max_percent_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
        Dim max_total_volume As Double
            max_total_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))

        'Greatest % Increase
        ws.Cells(2, 17).Value = max_percent_increase *100 &"%"
   
        'Greatest % Decrease
        ws.Cells(3, 17).Value = max_percent_decrease *100 &"%"
   
        'Greatest Total Volume
        ws.Cells(4, 17).Value = max_total_volume

        'Get Ticker Values
        Dim k_row_count As Long
            k_row_count = ws.Cells(Rows.Count, 11).End(xlUp).Row

        For i = 2 To k_row_count
            if(ws.Cells(i,11).value = max_percent_increase) Then
                ws.Cells(2,16).value = ws.Cells(i,9).value
            End If
            if(ws.Cells(i,11).value = max_percent_decrease) Then
                ws.Cells(3,16).value = ws.Cells(i,9).value
            End If
            if(ws.Cells(i,12).value = max_total_volume) Then
                ws.Cells(4,16).value = ws.Cells(i,9).value
            End If
        Next i
    Next ws
End Sub