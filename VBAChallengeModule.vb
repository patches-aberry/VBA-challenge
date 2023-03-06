Sub data_output_full()

    Dim ws As Worksheet
    Dim ticker_in, ticker_last, gpi_name, gpd_name, gtv_name As String
    Dim open_in, close_in, volume_sum, gpi, gpd, gtv As Double
    Dim ticker_int As Integer
    
    
    For Each ws In ThisWorkbook.Worksheets
        'find lastrow
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Add columns to add data outputs
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Initialize ticker_last with first name in worksheet
        ticker_last = ws.Cells(2, 1).Value
        'initialize open_in with first open value in worksheet
        open_in = ws.Cells(2, 3).Value
        'initialize ticker_int value at 2 (row 2)
        ticker_int = 2
        
        'loop through all rows in ws
        For i = 2 To lastrow + 1
            'Set ticker in value as relevant ticker name
            ticker_in = ws.Cells(i, 1).Value
            
            'check if new ticker name matches the previous ticker name
            If ticker_in = ticker_last Then
                'Add up the next volume
                volume_sum = volume_sum + ws.Cells(i, 7).Value
            
            ElseIf ticker_in <> ticker_last Then
                'Set Close value for ticker in year
                close_in = ws.Cells(i - 1, 6).Value
                'Output Ticker Name
                ws.Cells(ticker_int, 9).Value = ticker_last
                'Output Total Volume
                ws.Cells(ticker_int, 12).Value = volume_sum
                'Calculate and output yearly change
                ws.Cells(ticker_int, 10).Value = close_in - open_in
                'Format Change Column
                    If (close_in - open_in) > 0 Then
                        ws.Cells(ticker_int, 10).Interior.ColorIndex = 4
                    ElseIf (close_in - open_in) < 0 Then
                        ws.Cells(ticker_int, 10).Interior.ColorIndex = 3
                    End If
                'Calculate and output percent change
                ws.Cells(ticker_int, 11).Value = FormatPercent((close_in - open_in) / open_in)
                'Format Percent Change Column
                    If ((close_in - open_in) / open_in) > 0 Then
                        ws.Cells(ticker_int, 11).Interior.ColorIndex = 4
                    ElseIf ((close_in - open_in) / open_in) < 0 Then
                        ws.Cells(ticker_int, 11).Interior.ColorIndex = 3
                    End If
                'Set new ticker last value to compare
                ticker_last = ticker_in
                'Iterate ticker_int
                ticker_int = ticker_int + 1
                
                'Set new Open for year value
                open_in = ws.Cells(i, 3).Value
                volume_sum = ws.Cells(i, 7).Value
                
                End If
                
            Next i
        'Secondary Data Table
        'Start by adding table titles
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        'find lastrow of table we just made
        lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'initialize gpi,gpi_name,gpd,gpd_name,gtv,gtv_name with first row of new table
        gpi = ws.Cells(2, 11).Value
        gpi_name = ws.Cells(2, 9).Value
        gpd = ws.Cells(2, 11).Value
        gpd_name = ws.Cells(2, 9).Value
        gtv = ws.Cells(2, 12).Value
        gtv_name = ws.Cells(2, 9).Value
        
        For j = 2 To lastrow2 + 1
            
            If gpi < ws.Cells(j, 11).Value Then
                gpi = ws.Cells(j, 11).Value
                gpi_name = ws.Cells(j, 9).Value
            ElseIf gpd > ws.Cells(j, 11).Value Then
                gpd = ws.Cells(j, 11).Value
                gpd_name = ws.Cells(j, 9).Value
            ElseIf gtv < ws.Cells(j, 12).Value Then
                gtv = ws.Cells(j, 12).Value
                gtv_name = ws.Cells(j, 9)
            End If
            
        Next j
        
        'Output new table data
        ws.Cells(2, 16).Value = gpi_name
        ws.Cells(3, 16).Value = gpd_name
        ws.Cells(4, 16).Value = gtv_name
        ws.Cells(2, 17).Value = FormatPercent(gpi)
        ws.Cells(3, 17).Value = FormatPercent(gpd)
        ws.Cells(4, 17).Value = gtv
        
        ws.Cells.EntireColumn.AutoFit
        ws.Cells.EntireRow.AutoFit
        
    Next ws

End Sub

