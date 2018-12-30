Attribute VB_Name = "Module1"
Sub HomeWork_VBA_part2()

For Each ws In Worksheets
    Dim WorksheetName As String
    Dim StockName As String
    Dim StockTotal As Double
    Dim RowReference As Integer
    
    Dim PreAmount As Double
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Single
    Dim PercentageChange As Double
    
    Dim LastRow As Long
    
    
    RowReference = 2
    StockTotal = 0
    PreAmount = 2
    
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    WorksheetName = ws.Name
    
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly_Change"
    ws.Cells(1, "K").Value = "Percentage_Change"
    ws.Cells(1, "L").Value = "Total_Stock_Volume"
    
    
    For i = 2 To LastRow
    
        StockTotal = StockTotal + ws.Cells(i, 7).Value
        StockName = ws.Cells(i, 1).Value
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        
            YearlyClose = ws.Range("F" & i)
            YearlyOpen = ws.Range("C" & PreAmount)
            YearlyChange = YearlyClose - YearlyOpen
            
            If YearlyOpen = 0 Then
                PercentageChange = 0
            Else
                PercentageChange = (YearlyChange / YearlyOpen)
            End If
            
            
            ws.Range("J" & RowReference).Value = YearlyChange
            ws.Range("L" & RowReference).Value = StockTotal
            ws.Range("I" & RowReference).Value = StockName
            ws.Range("K" & RowReference).Value = PercentageChange
        
        
                
                If ws.Range("J" & RowReference).Value >= 0 Then
                    ws.Range("J" & RowReference).Interior.ColorIndex = 4
                            
                Else
                    ws.Range("J" & RowReference).Interior.ColorIndex = 3
                End If
                
            
            RowReference = RowReference + 1
            PreAmount = i + 1
            
    
        End If
        
    Next i
    
    For i = 2 To LastRow
        ws.Range("k" & i).Style = "Percent"
    Next i
    
    

Next ws
End Sub
