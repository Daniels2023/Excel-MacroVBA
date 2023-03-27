Attribute VB_Name = "Module1"
Sub Ticker():
        
        Dim ws As Worksheet
        
        Dim GIncrease As Double
        
        Dim GIncreaseTicker As String
               
        Dim GDecrease As Double
        
        Dim GDecreaseTicker As String
              
        Dim GTotalVolume As Double
        
        Dim GTotalTicker As String
        
        Dim i As Double
        
        Dim LastRow As Double
        
        Dim counterTicker As Double
        
        Dim counterVolume As Double
        
        Dim counterYearly As Double
        
        Dim volume As Double
        
        Dim stockOpen As Double
        
        Dim stockClose As Double
        
        Dim yearlyChange As Double
        
        Dim percentageChange As Double
                
            
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------

        For Each ws In Worksheets
         
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
             
             counterTicker = 2
             counterVolume = 2
             counterYearly = 2
             stockOpen = 0
             stockClose = 0
             yearlyChange = 0#
             percentageChange = 0#
             volume = 0
             GDecrease = 0
             GDecreaseTicker = ""
             GTotalTicker = ""
             GTotalVolume = 0
             GIncreaseTicker = ""
             GIncrease = 0
            
            For i = 2 To LastRow
    
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    ws.Cells(counterTicker, 9).Value = ws.Cells(i, 1).Value
                    stockOpen = Cells(i, 3).Value
                    counterTicker = counterTicker + 1
                End If
                
                If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                    volume = volume + ws.Cells(i, 7).Value
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    volume = volume + ws.Cells(i, 7).Value
                    ws.Cells(counterVolume, 12).Value = Format(volume, "#,##0")
                    counterVolume = counterVolume + 1
                    If volume > GTotalVolume Then
                        GTotalVolume = volume
                        GTotalTicker = ws.Cells(i, 1).Value
                    End If
                    stockClose = ws.Cells(i, 6).Value
                    yearlyChange = stockClose - stockOpen
                    If stockOpen <> 0 Then
                        percentageChange = (yearlyChange / stockOpen) * 100
                        If percentageChange > GIncrease Then
                            GIncrease = percentageChange
                            GIncreaseTicker = ws.Cells(i, 1).Value
                        ElseIf percentageChange < GDecrease Then
                            GDecrease = percentageChange
                            GDecreaseTicker = ws.Cells(i, 1).Value
                        End If
                    Else
                        percentageChange = 0
                    End If
                    ws.Cells(counterYearly, 10).Value = Format(yearlyChange, "#.00")
                    ws.Cells(counterYearly, 11).Value = Format(percentageChange, "0.00") & "%"
                    If ws.Cells(counterYearly, 10).Value < 0 Then
                        ws.Cells(counterYearly, 10).Interior.ColorIndex = 3
                        ws.Cells(counterYearly, 11).Interior.ColorIndex = 3
                    Else
                        ws.Cells(counterYearly, 10).Interior.ColorIndex = 4
                        ws.Cells(counterYearly, 11).Interior.ColorIndex = 4
                    End If
                    
                    counterYearly = counterYearly + 1
                    volume = 0
                    yearlyChange = 0
                    percentageChange = 0
                End If
         
             Next i
             
        ws.Range("P2").Value = GIncreaseTicker
        ws.Range("P3").Value = GDecreaseTicker
        ws.Range("P4").Value = GTotalTicker
        ws.Range("Q2").Value = Format(GIncrease, "0.00") & "%"
        ws.Range("Q3").Value = Format(GDecrease, "0.00") & "%"
        ws.Range("Q4").Value = Format(GTotalVolume, "#,##0")
        Next ws
        
End Sub
