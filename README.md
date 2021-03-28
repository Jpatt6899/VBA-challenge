# VBA-challenge

Sub test()

Dim I As Long
Dim summary_table_row As Long
Dim total_value As Double
Dim Counter_open As Double
Dim ticker As String
Dim Percent_finder As Double
Dim Yearly_change As Double


For Each ws In Worksheets

    summary_table_row = 2
    
    total_value = 0
    
    Counter_open = Cells(2, "c").Value
    
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ws.Activate
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 12).Value = "Total Value"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    
    For I = 2 To RowCount
    
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    
        ticker = ws.Cells(I, 1).Value
        
        total_value = total_value + ws.Cells(I, 7).Value
        
        Counter_open = Cells(2, "c").Value
        Counter_close = Cells(I, "F").Value
        
   
        Yearly_change = Counter_close - Counter_open
        
        
        Percent_finder = (Yearly_change / Counter_open)
        
        
         
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("L" & summary_table_row).Value = total_value
        ws.Range("J" & summary_table_row).Value = Yearly_change
        ws.Range("K" & summary_table_row).Value = Percent_finder
        ws.Range("K" & summary_table_row).NumberFormat = "0%"
        
        
             If ws.Range("J" & summary_table_row).Value < 0 Then
              ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
             
             End If
             
             
             If ws.Range("J" & summary_table_row).Value >= 0 Then
              ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
    
    
          End If
      
        
        summary_table_row = summary_table_row + 1
          
          total_value = 0
          
          Counter_open = ws.Cells(I + 1, "C").Value
          Counter_close = ws.Cells(I + 1, "F").Value
        
        Else
    
          total_value = total_value + ws.Cells(I, 7).Value
    
    
    End If
    
    Next I
        
             
    
    
    
    Next ws

End Sub

