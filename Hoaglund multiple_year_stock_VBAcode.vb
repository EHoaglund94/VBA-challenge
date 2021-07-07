Sub multiple_year_stock_data()

'For loop to loop through every worksheet
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

    'Creating Variables
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    close_price = 0
    Dim change_price As Double
    change_price = 0
    Dim percent_change_price As Double
    Dim volume_stock As Double
    volume_stock = 0

    'Creating variable for summary table rows
    Dim summary_table_row As Integer
    summary_table_row = 2

    'Set summary table headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    'Find last row of data
    last_row = Cells(Rows.Count, 1).End(xlUp).Row

    'Set first open price
    open_price = Cells(2, 3).Value

        'Loop through column 1 to find next ticker

        For i = 2 To last_row
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Set Ticker Variable
                ticker = Cells(i, 1).Value
        
                'set close price variable
                close_price = Cells(i, 6)
        
                'Set yearly change variable
                change_price = close_price - open_price
        
                    'set percent change variable
                    If (open_price = 0 And close_price = 0) Then
                        percent_change_price = 0
                    ElseIf (open_price = 0 And close_price <> 0) Then
                        percent_change_price = 1
                    Else
                        percent_change_price = change_price / open_price
                        Cells(summary_table_row, 11).Value = percent_change_price
                        Cells(summary_table_row, 11).NumberFormat = "0.00%"
                
                    End If
            
                
        
                'Set volume stock varaible
                volume_stock = volume_stock + Cells(i, 7).Value
        
                'Print the ticker variable in table
                Range("I" & summary_table_row).Value = ticker
        
                'Print yearly change in table
                Range("J" & summary_table_row).Value = change_price
        
       
        
    
                'Print the total stock volume in table
                Range("L" & summary_table_row).Value = volume_stock
        
                'Jump to next row
                summary_table_row = summary_table_row + 1
        
                'Reset counts
                volume_stock = 0
            
                'reset opening price
                open_price = Cells(i + 1, 3).Value
        
        
        
            Else
        
                'Keep adding stock volumes when stock ticker is the same
                volume_stock = volume_stock + Cells(i, 7).Value
            
            
            End If
        
        Next i
    
        'Find last row of created table
        last_row2 = Cells(Rows.Count, 9).End(xlUp).Row
    
        'Color formatting conditional
        For j = 2 To last_row2
            If Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
    
    
    
    'Set labels for bonus table
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Loop through created table
    
    For bonus = 2 To last_row2
    
        'Conditionals to find max/min percent changes and max total stock volume
        
        If Cells(bonus, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & last_row2)) Then
            Cells(2, 16) = Cells(bonus, 9).Value
            Cells(2, 17) = Cells(bonus, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
                  
        ElseIf Cells(bonus, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & last_row2)) Then
            Cells(3, 16) = Cells(bonus, 9).Value
            Cells(3, 17) = Cells(bonus, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
               
        ElseIf Cells(bonus, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & last_row2)) Then
            Cells(4, 16) = Cells(bonus, 9).Value
            Cells(4, 17) = Cells(bonus, 12).Value
        
        
        End If
        
    Next bonus
    
Next WS
        
        
End Sub
    
    