
Sub multiple_year_stock_data()
    
    'set an inital variable to count the number of worksheets
    Dim r As Integer
    
    'loop through all worksheets
    For r = 1 To Worksheets.Count
        
        Worksheets(r).Select
        
        'keep track of the location of each ticker in the summary table
        Dim ticker_counter As Integer
        summary_row_counter = 1
        
        'set an inital variable to hold the total stock volume for each ticker
        Dim total_stock_volume As LongLong
        total_stock_volume = 0
        
        'check the last row number
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'set an initial variable to count the number of rows for EACH ticker
        Dim each_ticker_counter As Long
        each_ticker_counter = 0
        
        'set an initial variable to calculate the yearly_change for each ticker
        Dim yearly_change As Double
        
        'set an initial variable to calculate the percentage change for each ticker
        Dim percentage_change As Double
        
        'instructs VBA to jump to the line errorhandler whenever an unexpected error occurs at runtime
        On Error GoTo errorhandler
        
        'Loop through all stock
        For i = 2 To lastrow
            
            'if the next ticker is not the same as the previous one
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'add to the total_stock_volumn
                total_stock_volume = total_stock_volume + Cells(i, 7).Value
                
                'add one to the summary table row
                summary_row_counter = summary_row_counter + 1
                
                'print the ticker name in the summary table
                Cells(summary_row_counter, 9).Value = Cells(i, 1).Value
                
                'print the total_stock_volume in the summary table
                Cells(summary_row_counter, 12).Value = total_stock_volume
                
                'reset total_stock_volume
                total_stock_volume = 0
                
                'yearly_change=(close-open)for EACH ticker
                yearly_change = Cells(i, 6).Value - Cells((i - each_ticker_counter), 3).Value
                
                'print the yearly_change in the summary table
                Cells(summary_row_counter, 10).Value = yearly_change
                
                'conditional formatting
                'if yearly_change is positive, highlight in green
                If yearly_change > 0 Then
                    Cells(summary_row_counter, 10).Interior.ColorIndex = 4
                    'if yearly_change is negative,highlight in red
                Else
                    Cells(summary_row_counter, 10).Interior.ColorIndex = 3
                End If
                
                'percentage_change= (close-open)/open
                percentage_change = yearly_change / Cells((i - each_ticker_counter), 3).Value
                'print percentage_change on the summary table
                Cells(summary_row_counter, 11).Value = percentage_change
                
                'reset the variable to 0
                each_ticker_counter = 0
                
                'if the next ticker is the same as the previous one
            Else
                
                'add to the total_stock_volume
                total_stock_volume = total_stock_volume + Cells(i, 7).Value
                
                'add 1 to the variable each_ticker_counter
                each_ticker_counter = each_ticker_counter + 1
                
            End If
            
        Next i
        
errorhandler:
        ' If a divide by zero error occurs when calculating the percentage change, override the variable to zero
        percentage_change = 0
        Resume Next
        
        'Formatting the column percentage change to percentage
        Range("K:K").Select
        Selection.NumberFormat = "0.00%"
        
        'Create Headers Name
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % increase"
        Range("O3").Value = "Greatest % decrease"
        Range("O4").Value = "Greatest total volume"
        Range("P1").Value = "ticker"
        Range("Q1").Value = "value"
        
        'Greatest % Increase
        'set an initial variable to find the greatest percentage increase
        Dim greatest_percentage_increase As Double
        'set the variable as 0 to compare with each value in the column percentage change
        greatest_percentage_increase = 0
        
        'Greatest % Decrease
        'set an initial variable to find the greatest percentage decrease
        Dim greatest_percentage_decrease As Double
        'set the variable as 0 to compare with each value in the column percentage change
        greatest_percentage_decrease = 0
        
        'Greatest Total Volume
        'set an initial variable to find the greatest total volume
        Dim max_totalvol As LongLong
        'set the variable as 0 to compare with each value in the column total stock volume
        max_totalvol = 0
        
        'set an initial variable for the last row of column K
        Dim lastrow_k As Integer
        'check the last row number for column K,Percentage Changed
        lastrow_k = Cells(Rows.Count, "K").End(xlUp).Row
        
        'loop through all column K,Percentage Changed
        For j = 2 To lastrow_k
            'Greatest % Increase
            'if the value of the cell is greater than the variable, then assign the value of the cell to the variable
            If greatest_percentage_increase < Cells(j, 11).Value Then
                greatest_percentage_increase = Cells(j, 11).Value
                'print the name of the ticker for the greatest percentage increase
                Range("P2").Value = Cells(j, 9).Value
            End If
        Next j
        
        For j = 2 To lastrow_k
            'Greatest % Decrease
            'if the value of the cell is smaller than the variable, then assign the value of the cell to the variable
            If greatest_percentage_decrease > Cells(j, 11).Value Then
                greatest_percentage_decrease = Cells(j, 11).Value
                'print the name of the ticker for the greatest percentage decrease
                Range("P3").Value = Cells(j, 9).Value
            End If
        Next j
        
        For j = 2 To lastrow_k
            'Greatest Total Volume
            'if the value of the cell is greater than the variable, then assign the value of the cell to the variable
            If max_totalvol < Cells(j, 12).Value Then
                max_totalvol = Cells(j, 12).Value
                'print the name of the ticker for the greatest total volume
                Range("P4").Value = Cells(j, 9).Value
            End If
        Next j
        
        'print the value of the greatest percentage increase
        Range("Q2").Value = greatest_percentage_increase
        'print the value of the greatest percentage decrease
        Range("Q3").Value = greatest_percentage_decrease
        'print the value of the greatest total volume
        Range("Q4").Value = max_totalvol
        
        'Formatting cell Q2 to percentage
        Range("Q2").Select
        Selection.NumberFormat = "0.00%"
        
        'Formatting cell Q3 to percentage
        Range("Q3").Select
        Selection.NumberFormat = "0.00%"
        
    'Next sheet
    Next r
    
End Sub

