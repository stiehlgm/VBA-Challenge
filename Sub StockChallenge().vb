Sub StockChallenge()

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Start row
        Dim i As Long
        
        'Next column
        Dim j As Long
        
        'Value for "Ticker"
        Dim Ticker As Long
        
        'Last filled row for column A
        Dim FinalRowA As Long
    
        ' Value for the percent calculation
        Dim Percent As Double
        
        'Last filled row for column I
        Dim FinalRowI As Long
        
        'Variable for greatest increase
        Dim MaxIncrease As Double
        
        'Variable for greatest decrease
        Dim Maxdecrease As Double
        
        'Variable for greatest volume
        Dim MaxVolume As Double
        
        'Naming news columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'Start count at second row
        Ticker = 2
        j = 2
        
        'The last filled row in column A
        FinalRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop for Ticker
        For i = 2 To FinalRowA
        
            'Check if ticker name changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Write ticker in column I
            ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
            
            'Calculate yearly chnage
            ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
                'Color formating
                If ws.Cells(Ticker, 10).Value > 0 Then
                
                'Change box to green
                ws.Cells(Ticker, 10).Interior.ColorIndex = 4
                
                Else
                ws.Cells(Ticker, 10).Interior.ColorIndex = 3
                
                End If
            
            'Write percentage change in column K
            If ws.Cells(j, 3).Value <> 0 Then
            
            'Calculate the percentage
            Percent = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
            
            'Put in column K
            ws.Cells(Ticker, 11).Value = Format(Percent, "Percent")
    
            End If
            
            'Write total stock in column L
            ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(i, 7), ws.Cells(j, 7)))
            
            
            'Start new row for ticker change
            Ticker = Ticker + 1
            
            'Start new cloumn for next data points
            j = i + 1
             
            End If
        
        Next i
        
        'The last filled row in column I
        FinalRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Percent high, low and total volume categories
        MaxIncrease = ws.Cells(2, 11).Value
        Maxdecrease = ws.Cells(2, 11).Value
        MaxVolume = ws.Cells(2, 12).Value
        
            'Loop for the values
            For i = 2 To FinalRowI
            
                'Scan for greatest increase
                If ws.Cells(i, 11).Value > MaxIncrease Then
            
                MaxIncrease = ws.Cells(i, 11).Value
            
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            
                Else
                MaxIncrease = MaxIncrease
                
                End If
        
              'Scan for greatest decrease
                If ws.Cells(i, 11).Value < Maxdecrease Then
            
                Maxdecrease = ws.Cells(i, 11).Value
            
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
                Else
                Maxdecrease = Maxdecrease
        
                End If
            
            'Scan for greatest volume
                If ws.Cells(i, 12).Value > MaxVolume Then
            
                MaxVolume = ws.Cells(i, 12).Value
            
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
                Else
                MaxVolume = MaxVolume
        
                End If
            
            'Add the values in column Q
            ws.Cells(2, 17).Value = Format(MaxIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(Maxdecrease, "Percent")
            ws.Cells(4, 17).Value = Format(MaxVolume, "Scientific")
            
            Next i
            
        'Auto fit all coulmns
        Columns("A:Q").EntireColumn.AutoFit
        
    
    Next

End Sub

