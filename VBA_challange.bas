Attribute VB_Name = "Module1"
Sub Stock_year()
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
            'add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Changed"
        ws.Cells(1, 12).Value = "Total Volume"

        'define last row
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        'set cariable for holding the ticker name
        Dim ticker_name As String
        
        'set stock open value
        Dim stock_open As Double
        
        'set variable for stock close
        Dim stock_close As Double
    
        'set variable for yearly change
        Dim yearly_change As Double
        yearly_change = 0
    
        'set variable for percent changed
        Dim percent_changed As Double
        percent_changed = 0
    
        'set variable for total stock volume
        Dim total_volume As Double
        total_volume = 0
    
    
        'define summary table
        Dim summary_table_row As Integer
        summary_table_row = 2
    
        'use for loop to collect all the info
        For i = 2 To LastRow
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                'add total volume
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                'grab stock open
                stock_open = ws.Cells(i, 6)
                

            'if ticker is different
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'grab ticker name
                ticker_name = ws.Cells(i, 1).Value
            
                'add to total volume
                total_volume = total_volume + ws.Cells(i, 7)
                
                    'grab year close
                stock_close = ws.Cells(i, 6)
                
                 'calculate yearly change
                yearly_change = stock_close - stock_open
                
                'convert yearly change to a percent
                percent_changed = yearly_change
                
                'add yearly change to summary table
                ws.Range("J" & summary_table_row).Value = yearly_change
                
                'add changed percent
                ws.Range("K" & summary_table_row).Value = percent_changed
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                    'Format the color
                    If ws.Range("K" & summary_table_row).Value > 0 Then
                        ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
                    Else
                        ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
                    End If
                    
                'add ticker to summary table
                ws.Range("I" & summary_table_row).Value = ticker_name
            
                'add total to summary table
                ws.Range("L" & summary_table_row).Value = total_volume
            
                'add a row to the summary table
                summary_table_row = summary_table_row + 1
            
            
            'If ticker is the same
            Else
                'add total volume
                total_volume = total_volume + ws.Cells(i, 7).Value
        
            End If
        Next i
    
    Next ws
    
End Sub


