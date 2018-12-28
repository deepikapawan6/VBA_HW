Attribute VB_Name = "Module1"
Sub VBA_HW()
    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim TickerSymbol As String
        Dim TotalStockVolume As Double
            TotalStockVolume = 0
        Dim SummarytableRow As Integer
            SummarytableRow = 2
        
        
        WorksheetName = ws.Name
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I1").Value = "<Ticker_Symbol>"
        ws.Range("J1").Value = "<Total_Stock_Volume>"
        
        
        For i = 2 To LastRow
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                TickerSymbol = ws.Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                
                ws.Range("I" & SummarytableRow).Value = TickerSymbol
                ws.Range("J" & SummarytableRow).Value = TotalStockVolume
                
                SummarytableRow = SummarytableRow + 1
                TotalStockVolume = 0
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
            End If
        Next i
    
    
    
     Next ws

End Sub

