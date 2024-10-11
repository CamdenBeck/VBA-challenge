Attribute VB_Name = "Module1"
Sub StockInfo()
    For Each ws In ThisWorkbook.Worksheets
        ' Inserting the "Ticker" column
        ws.Cells(1, "I").Value = "Ticker"
        
        ' Inserting the "Quarterly Change" column
        ws.Cells(1, "J").Value = "Quarterly Change"
        
        ' Inserting the "Percent Change" column
        ws.Cells(1, "K").Value = "Percent Change"
        
        ' Inserting the "Total Stock Volume" column
        ws.Cells(1, "L").Value = "Total Stock Volume"
        
        ' Getting the total number of rows
        Last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        Info_row = 2
        
        ' Setting the quarterly change and the total stock volume to 0
        Quarterly_change = 0
        Stock_volume = 0
        
        ' Getting the opening value for the starting ticker
        Open_value = ws.Cells(2, "C").Value
    
        ' Looping through all the rows in the worksheet
        For r = 2 To Last_row
            ' Getting the value of the ticker columns
            current_ticker = ws.Cells(r, "A").Value
            next_ticker = ws.Cells(r + 1, "A").Value
    
            If current_ticker <> next_ticker Then
                ws.Cells(Info_row, "I").Value = current_ticker
                
                ' Calculating the quarterly change
                Close_value = ws.Cells(r, "F").Value
                Quarterly_change = Close_value - Open_value
                
                ' Displaying the quarterly change
                ws.Cells(Info_row, "J").Value = Quarterly_change
                
                ' Formating the quarterly change column
                If Quarterly_change < 0 Then
                    ws.Cells(Info_row, "J").Interior.ColorIndex = 3
                ElseIf Quarterly_change > 0 Then
                    ws.Cells(Info_row, "J").Interior.ColorIndex = 4
                End If
                
                ' Calculate and display the percent change
                Percent_change = (Quarterly_change / Open_value)
                ws.Cells(Info_row, "K").Value = Percent_change
                'ws.Cells(Info_row, "K").Style = "Percent"
                ws.Cells(Info_row, "K").NumberFormat = "0.00%"
                
                ' Formatting the percent change column
                If Percent_change < 0 Then
                    ws.Cells(Info_row, "K").Interior.ColorIndex = 3
                ElseIf Percent_change > 0 Then
                    ws.Cells(Info_row, "K").Interior.ColorIndex = 4
                End If
                
                ' Display the total stock volume
                ws.Cells(Info_row, "L").Value = Stock_volume
                
                ' Reset the stock volume to 0
                Stock_volume = 0
                
                ' Reset the Open_value to the next ticker's row
                Open_value = ws.Cells(r + 1, "C").Value
    
                Info_row = Info_row + 1
            Else
                Stock_volume = Stock_volume + ws.Cells(r, "G").Value
                
            End If
        Next r
        
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        
        Percent_Rng = ws.Range("K2").EntireColumn
        
        ' Getting and displaying the greatest percent increase
        max_percent_increase = WorksheetFunction.Max(Percent_Rng)
        ws.Cells(2, "Q").Value = max_percent_increase
        ws.Cells(2, "Q").NumberFormat = "0.00%"
        
        ' Finding the row where the max percent value is located
        r = 1
        For Each cell In Percent_Rng
            If cell = max_percent_increase Then
                max_percent_row = r
                Exit For
            Else
                r = r + 1
            End If
        Next cell
        
        ' Finding and displaying the corresponding ticker for the max percent decrease
        ws.Cells(2, "P").Value = ws.Cells(max_percent_row, "I").Value
        
        ' Getting and displaying the greatest percent decrease
        max_percent_decrease = WorksheetFunction.Min(Percent_Rng)
        ws.Cells(3, "Q").Value = max_percent_decrease
        ws.Cells(3, "Q").NumberFormat = "0.00%"
        
        ' Finding the row where the min percent value is located
        r = 1
        For Each cell In Percent_Rng
            If cell = max_percent_decrease Then
                min_percent_row = r
                Exit For
            Else
                r = r + 1
            End If
        Next cell
        
        ' Finding and displaying the corresponding ticker for the max percent decrease
        ws.Cells(3, "P").Value = ws.Cells(min_percent_row, "I").Value
        
        ' Finding the greatest total volume
        Volume_Rng = Range("L2").EntireColumn
        max_volume = WorksheetFunction.Max(Volume_Rng)
        
        ' Display the greatest total volume
        ws.Cells(4, "Q").Value = max_volume
        
        ' Finding the row where the max volume is located
        r = 1
        For Each cell In Volume_Rng
            If cell = max_volume Then
                max_volume_row = r
                Exit For
            Else
                r = r + 1
            End If
        Next cell
        
        ' Finding and displaying the corresponding ticker for the max volume
        ws.Cells(4, "P").Value = ws.Cells(max_volume_row, "I").Value
        
    Next ws

End Sub



