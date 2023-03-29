Sub AllStockData()

    'go through each worksheet
    For Each ws In Worksheets
    
    'name all variables
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim GreatTot As Double
        Dim GreatIncTick As String
        Dim GreatDecTick As String
        Dim GreatTotTick As String
            
        WorksheetName = ws.Name
        
        'add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Vol"
        
        
        TickCount = 2
        
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRowA
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    vol = ws.Cells(i, 7).Value
                    year_open = ws.Cells(i, 3).Value
                     
                     
                     
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    
                    
                    
                   
                    year_close = ws.Cells(i, 6).Value
                    yearly_change = year_close - year_open
                    percent_change = (year_close - year_open) / year_open
                    vol = vol + ws.Cells(i, 7).Value
                    
                    ws.Cells(TickCount, 9).Value = ticker
                    ws.Cells(TickCount, 10).Value = yearly_change
                    ws.Cells(TickCount, 11).Value = percent_change
                    ws.Cells(TickCount, 12).Value = vol

                    If yearly_change >= 0 Then
                        ws.Cells(TickCount, 10).Interior.Color = 5287936
                    Else
                        ws.Cells(TickCount, 10).Interior.Color = 255
                    End If
                        
                    TickCount = TickCount + 1
                Else
                    vol = vol + ws.Cells(i, 7).Value
                    
                End If
                               
                
            Next i
            
            LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            GreatInc = 0
            GreatDec = 0
            
            For i = 2 To LastRowI
                
                If ws.Cells(i, 11).Value > GreatInc Then
                    GreatInc = ws.Cells(i, 11).Value
                    GreatIncTick = ws.Cells(i, 9).Value
                End If
            
                If ws.Cells(i, 11).Value < GreatDec Then
                    GreatDec = ws.Cells(i, 11).Value
                    GreatDecTick = ws.Cells(i, 9).Value
                End If
                
                If ws.Cells(i, 12).Value > GreatTot Then
                    GreatTot = ws.Cells(i, 12).Value
                    GreatTotTick = ws.Cells(i, 9).Value
                End If
            
            Next i
            ws.Cells(2, 14).Value = "Greatest Increase"
            ws.Cells(3, 14).Value = "Greatest Decrease"
            ws.Cells(4, 14).Value = "Greatest Volume"
            ws.Cells(2, 15).Value = GreatIncTick
            ws.Cells(3, 15).Value = GreatDecTick
            ws.Cells(4, 15).Value = GreatTotTick
            ws.Cells(2, 16).Value = GreatInc
            ws.Cells(3, 16).Value = GreatDec
            ws.Cells(4, 16).Value = GreatTot
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            
        Next ws
              
        
    

End Sub
