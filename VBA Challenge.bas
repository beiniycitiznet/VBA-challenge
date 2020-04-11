Attribute VB_Name = "Module1"
Sub MultiYear()

For Each ws In Worksheets

    Dim LastRow As Double
    Dim SummaryRow As Double
    Dim Ticker As String
    Dim BeginRow As Long
    Dim YearlyChange As Double
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim PercentChange As Double
    Dim TotalStockV As LongLong
    Dim GreatInc As Double
    Dim GreatDec As Double
    Dim GreatTotalV As LongLong
    Dim IncRow As Double
    Dim DecRow As Double
    Dim VRow As Double
    
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    SummaryRow = 2
    BeginRow = 2
    TotalStockV = 0
    GreatInc = 0
    GreatDec = 0
    GreatTotalV = 0
    
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volumn"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
    
    
    For i = 2 To LastRow
        If ws.Cells(BeginRow, 3) = 0 And ws.Cells(i, 3) <> 0 Then
            BeginRow = i
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Ticker
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(SummaryRow, 9).Value = Ticker
            
            'Yearly Change
            YearOpen = ws.Cells(BeginRow, 3).Value
            YearClose = ws.Cells(i, 6).Value
            YearlyChange = YearClose - YearOpen
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            
            If YearlyChange > 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                
            Else
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            
            End If
            
            
            '% Change
            If YearOpen > 0 Then
                PercentChange = YearlyChange / YearOpen
                ws.Cells(SummaryRow, 11).Value = FormatPercent(PercentChange)
                
                'Challenges
                'Greatest % Increase
                If GreatInc < PercentChange Then
                    GreatInc = PercentChange
                    'MsgBox (Creat
                    IncRow = i
                        
                End If
                
                
                'Challenges
                'Greatest % Decrease
                If GreatDec > PercentChange Then
                    GreatDec = PercentChange
                    DecRow = i
                        
                End If
                
            Else
                
                ws.Cells(SummaryRow, 11).Value = FormatPercent(1)
            
            End If
            
            

            
            
            'Total Stock Volume
            TotalStockV = TotalStockV + ws.Cells(i, 7).Value
            ws.Cells(SummaryRow, 12).Value = TotalStockV
            
            'Challenges
            'Greatest Total Volumn
            If GreatTotalV < TotalStockV Then
                GreatTotalV = TotalStockV
                VRow = i
                        
            End If
            
            
            'For next Tiker
            TotalStockV = 0
            SummaryRow = SummaryRow + 1
            BeginRow = i + 1
            
            

        
            
        Else
        
            TotalStockV = TotalStockV + ws.Cells(i, 7).Value
        
        End If
    
    Next i
    
    ws.Range("O2").Value = ws.Cells(IncRow, 1).Value
    ws.Range("P2").Value = FormatPercent(GreatInc)
    
    ws.Range("O3").Value = ws.Cells(DecRow, 1).Value
    ws.Range("P3").Value = FormatPercent(GreatDec)
  
    ws.Range("O4").Value = ws.Cells(VRow, 1).Value
    ws.Range("P4").Value = GreatTotalV

Next ws


End Sub

