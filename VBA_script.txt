Sub Stocks()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        WorksheetName = ws.Name
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Opening price"
        ws.Range("K1").Value = "Closing price"
        ws.Range("L1").Value = "Yearly change"
        ws.Range("M1").Value = "Percentage Change"
        ws.Range("N1").Value = "Total Stock Volume"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        
        Dim lastrow, i, Headline As Long
        Dim Tickercolumn As String
        'Dim Volume As Long
        Headline = 2
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Getting "Yearly Change"
            For i = 2 To lastrow
             Tickercolumn = Cells(i, 1).Value
             'First case of each ticket
                If Cells(i - 1, 1).Value <> Tickercolumn Then
                    ws.Range("I" & Headline).Value = Tickercolumn
                    ws.Range("J" & Headline).Value = Cells(i, 3).Value
                    Volume = Cells(i, 7)
                    
                'Last row for same ticket
                ElseIf Cells(i + 1, 1) <> Tickercolumn Then
                    ws.Range("K" & Headline).Value = Cells(i, 6).Value
                    Volume = Volume + Cells(i, 7).Value
                    ws.Range("N" & Headline).Value = Volume
                    Headline = Headline + 1
                    
                'All the tickets that are not the first or the last
                Else: Volume = Volume + Cells(i, 7).Value
                
                End If
                
             Next i
        
      'Getting "Percent Change"
     
        Dim j As Integer
        'Dim max_volume As Long
        max_volume = 0
        Dim max_percentage As Double
        Dim min_percentage As Double
        Dim greatestincrease As Double
        greatestincrease = 0
        Dim greatestdecrease As Double
        greatestdecrease = 0
        'Dim valuecolumn As Long
        Dim Tickercolumn2 As String
        Dim Tickercolumn3 As String
        Dim Tickercolumn4 As String
        Dim lastrow2 As Long
        lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To lastrow2
            ws.Cells(j, 12) = Cells(j, 11) - Cells(j, 10)
            ws.Cells(j, 13).Value = Cells(j, 12) / Cells(j, 10)
            ws.Cells(j, 13).Value = FormatPercent(Cells(j, 13))
            
                
                If Cells(j, 12).Value < 0 Then
                    ws.Range("L" & j).Interior.ColorIndex = 3
                Else
                    ws.Range("L" & j).Interior.ColorIndex = 4
                End If
                
            valuecolumn = ws.Cells(j, 14).Value
                 
                If valuecolumn > max_volume Then
                    max_volume = valuecolumn
                    Tickercolumn2 = ws.Cells(j, 9).Value
                End If
                
            max_percentage = ws.Cells(j, 13).Value
                If max_percentage > greatestincrease Then
                    greatestincrease = max_percentage
                    Tickercolumn3 = ws.Cells(j, 9).Value
                End If
                
            min_percentage = ws.Cells(j, 13).Value
                If min_percentage < greatestdecrease Then
                    greatestdecrease = min_percentage
                    Tickercolumn4 = ws.Cells(j, 9).Value
                End If
                 
               
            Next j
            
        ws.Range("R4").Value = max_volume
        ws.Range("Q4").Value = Tickercolumn2
        ws.Range("R2").Value = greatestincrease
        ws.Cells(2, 18).Value = FormatPercent(Cells(2, 18))
        ws.Range("Q2").Value = Tickercolumn3
        ws.Range("R3").Value = greatestdecrease
        ws.Cells(3, 18).Value = FormatPercent(Cells(3, 18))
        ws.Range("Q3").Value = Tickercolumn4
              
        
            
             
                
       
     
            
            
            
  
   
                     
        
    Next ws
         
    
      
        
            
End Sub

