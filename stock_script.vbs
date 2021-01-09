Sub stock_data()
    
    'Loop through all the worksheets'
    For Each ws In Worksheets
        
        'Insert the titles for the values that we're going to calculate'
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("N2") = "Greatest % Increase"
        ws.Range("N3") = "Greatest % Decrease"
        ws.Range("N4") = "Greatest Total Volume"
        ws.Range("O1") = "Ticker"
        ws.Range("P1") = "Value"
        
        'Identify the last row of each sheet'
        LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Declare variables needed'
        Dim FRow As Integer
        Dim Ticker_Name As String
        Dim Ticker_Vol As Double 'Sum of volume'
        Dim OPrice As Double 'Opening Price'
        Dim CPrice As Double 'Closing Price'
        Dim YChange As Double 'Yearly Change'
        Dim PChange As Double 'PErcentage Change'
        
        'Initialize variables'
        FRow = 2
        Ticker_Vol = 0
        OPrice = ws.Cells(2, 3).Value
       
        'Loop through all the rows to identify the diverse values that we need to create the summarize table'
        For i = 2 To LRow
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Insert Ticker name'
                Ticker_Name = ws.Cells(i, 1).Value
                ws.Range("I" & FRow).Value = Ticker_Name
                
                'Add volume per Ticker'
                Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value
                ws.Range("L" & FRow).Value = Ticker_Vol
                
                'Identify the closing price by the end of the year'
                CPrice = ws.Cells(i, 6).Value
                
                'Calculate the percentage of change of the year even when OPrice is zero'
                If OPrice <> 0 Then
                
                    'Calculate Yearly Change'
                    YChange = CPrice - OPrice
                    
                  
                    If YChange < 0 Then
                        
                        'Add the Yearly change to the table'
                        ws.Range("J" & FRow).Value = YChange
                        
                        'Color format of Yearly Change'
                        ws.Range("J" & FRow).Interior.ColorIndex = 3
                        
                        'Calculate Percentage Change'
                        PChange = ((CPrice - OPrice) / OPrice)
                        
                        'Add the Percentage Change value to the table'
                        ws.Range("K" & FRow).Value = PChange
                        
                        
                    Else
                       'Add the Yearly change to the table'
                        ws.Range("J" & FRow).Value = YChange
                        
                        'Color format of Yearly Change'
                        ws.Range("J" & FRow).Interior.ColorIndex = 4
                        
                        'Calculate Percentage Change'
                        PChange = ((CPrice - OPrice) / OPrice)
                        
                        'Add the Percentage Change value to the table'
                        ws.Range("K" & FRow).Value = PChange
                 
                    End If
                
                Else
                    PChange = 0
                    ws.Range("K" & FRow).Value = PChange
                    
                End If
                
                'Restart and save some values'
                OPrice = ws.Cells(i + 1, 3)
                FRow = FRow + 1
                Ticker_Vol = 0
               
            Else
                Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value
                
            End If
        Next
    'Change to Percent Style the Percent Change value'
    ws.Columns("K").NumberFormat = "0.00%"
    
    'Change size of the cells to fit exaclty to its content'
    ws.Columns("A:P").AutoFit
    Next
    
    'Code for the Bonus'
    For Each ws In Worksheets
    
        'Declare variables needed'
        Dim MaxP As Double 'Value of the greatest % increase'
        Dim MinP As Double 'Value of the greatest % decrease'
        Dim MaxVol As Double 'Value of the greatest total volume'
        
        'Initialize varibles'
        MaxP = 0
        MinP = 0
        MaxVol = 0
        
        'Identify the last row of the summarize table of each sheet'
        SRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Loop for finding the values we are looking for'
        For i = 2 To SRow
        
            'Condition to find out the Greatest % Increase'
            If MaxP < ws.Cells(i, 11) Then
                MaxP = ws.Cells(i, 11)
                ws.Range("P2") = MaxP
                ws.Range("P2").NumberFormat = "0.00%"
                ws.Range("O2") = ws.Cells(i, 9)
            Else
                MaxP = MaxP
            End If
            
            'Condition to find out the Greatest % Decrease'
            If MinP > ws.Cells(i, 11) Then
                MinP = ws.Cells(i, 11)
                ws.Range("P3") = MinP
                ws.Range("P3").NumberFormat = "0.00%"
                ws.Range("O3") = ws.Cells(i, 9)
                
            Else
                MinP = MinP
            End If
            
            'Condition to find out the greatest total volume'
            If MaxVol < ws.Cells(i, 12) Then
                MaxVol = ws.Cells(i, 12)
                ws.Range("P4") = MaxVol
                ws.Range("O4") = ws.Cells(i, 9)
            Else
                MaxVol = MaxVol
            End If
        Next
        
    Next

End Sub