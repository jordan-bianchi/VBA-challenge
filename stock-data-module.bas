Attribute VB_Name = "Module1"
Sub StockData():


For Each ws In Worksheets

    Dim WorksheetName As String
    
    Dim YearlyChange As Double
    
    Dim DecimalChange As Double
    Dim PercentChange As String
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim OpenCol As Double
    OpenCol = 3
    
    Dim CloseCol As Double
    CloseCol = 6
    
    ws.Activate
    WorksheetName = ws.Name
    Debug.Print (WorksheetName)
    
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Range("I2:I" & LastRow).Value = ws.Range("A2:A" & LastRow).Value
        
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
        For i = 2 To LastRow
            
            YearlyChange = Cells(i, 6).Value - Cells(i, 3).Value
       
            ws.Cells(i, 10).Value = YearlyChange
            
            If ws.Cells(i, 3).Value = 0 And ws.Cells(i, 6).Value = 0 Then
                ws.Cells(i, 11).Value = "0.00%"
                
            ElseIf ws.Cells(i, 3).Value <> 0 Then
            DecimalChange = (ws.Cells(i, 10).Value / ws.Cells(i, 3).Value)
            PercentChange = FormatPercent(DecimalChange, 2)
            ws.Cells(i, 11).Value = PercentChange
            
            End If
            
            
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(i, 11).Value < 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 3
                
            ElseIf ws.Cells(i, 11).Value = 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 0
            
            End If
            
           
        Next i
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Range("L2:L" & LastRow).Value = ws.Range("G2:G" & LastRow).Value
        
        ws.Columns("A:L").AutoFit
        
Next ws




End Sub
