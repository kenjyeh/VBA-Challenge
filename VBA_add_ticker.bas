Attribute VB_Name = "Module1"
Sub add_ticker():

For Each ws In Worksheets
    
 'variable creation
    Dim ticker_name As String
    Dim total_volume As Double
    Dim ticker_row As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    
    total_volume = 0
    ticker_row = 2
    open_price = ws.Cells(2, 3).Value
    
    
  'formatting table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
   'last row/column counter
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    last_column = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    
    
    For i = 2 To last_row
    
    'If values differ
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'store ticker name and volume
            ticker_name = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'print out ticker name and volume
            ws.Range("i" & ticker_row).Value = ticker_name
            ws.Range("l" & ticker_row).Value = total_volume
            
            
            'store close price
            close_price = ws.Cells(i, 6).Value
            'calculate year change
            yearly_change = (close_price - open_price)
            
            'print yearly change in chart
            ws.Range("j" & ticker_row).Value = yearly_change
            
            
            If open_price = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / open_price
            End If
                ws.Range("k" & ticker_row).Value = percent_change
                ws.Range("k" & ticker_row).NumberFormat = "0.00%"
                
                'next row on counter
                ticker_row = ticker_row + 1
                
                'reset total volume variable for next ticker
                total_volume = 0
                
                'reset open price for next ticker
                open_price = ws.Cells(i + 1, 3)
        Else
        
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If

    Next i
    
    'color format
    'create last row for summary table
     last_row_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    
    For i = 2 To last_row_summary_table
        If ws.Cells(i, 10).Value >= 0 Then
                'green
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            Else
                'red
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
    Next i
    
    
    'set up table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
    ' for loop to fill the table
    For i = 2 To last_row_summary_table
        
            'greatest percent increase
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'greatest percent decrease
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'greatest total volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
        
        If ws.Cells(2, 17).Value >= 0 Then
                'green
                ws.Cells(2, 17).Interior.ColorIndex = 10
            
            Else
                'red
                ws.Cells(2, 17).Interior.ColorIndex = 3
            
            End If
            
        If ws.Cells(3, 17).Value >= 0 Then
                'green
                ws.Cells(3, 17).Interior.ColorIndex = 10
            
            Else
                'red
                ws.Cells(3, 17).Interior.ColorIndex = 3
            
        End If
    
    
    
 Next ws
 

End Sub
