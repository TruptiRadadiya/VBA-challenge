Attribute VB_Name = "Module1"
Sub stockChange()
    'Loop through each worksheets
    For Each ws In Worksheets
    
        'Declare a variable to store ticker symbol
        Dim Ticker_Symbol As String
        
        'Declare and iniitiate variables to store yearly change, percentage change and total stock volume for each ticker
        Dim Yearly_Change, Percentage_Change, Total_Stock_Vol As Double
        Total_Stock_Vol = 0
        
        'Set columns for Ticker summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Declare variables to store opening and closing value of stocks
        Dim Opening_Value, Closing_Value As Double
        
        'Set a varibale to keep track of each row of ticker summary table
        Dim Ticker_Table_Row As Integer
        Ticker_Table_Row = 2
        
        'Declare variable to store first row for each stock
        Dim First_Row_Of_Ticker As Long
        First_Row_Of_Ticker = 2
        
        'Declare a variable to store last row
        Dim Last_Row As Long
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Loop through each ticker
        Dim i As Long
        For i = 2 To Last_Row
        
            'Check if the the ticker is changed from next row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set the Ticker name
                Ticker_Symbol = ws.Cells(i, 1).Value
                
                'Set the opening value for the stock
                Opening_Value = ws.Cells(First_Row_Of_Ticker, 3).Value
                
                'Set the first row for next ticker
                First_Row_Of_Ticker = i + 1
                
                'Set the closing value for the stock
                Closing_Value = ws.Cells(i, 6).Value
                
                'Calculate yearly change
                Yearly_Change = Closing_Value - Opening_Value
                
                'Output the yearly change in the ticker summary table
                ws.Range("J" & Ticker_Table_Row).Value = Yearly_Change
                
                'Check if the yearly chage is positive
                If Yearly_Change > 0 Then
                
                    'Set the cell color to green
                    ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 4
                
                'If the yearly chage is negative
                Else
                
                    'Set the cell color to red
                    ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 3
                    
                End If
                
                'Calculate percentage change
                Percentage_Change = Yearly_Change / Opening_Value
                
                'Output the percentage change in the ticker summary table
                ws.Range("K" & Ticker_Table_Row).Value = Percentage_Change
                
                'Format the percentage change cell to percentage
                ws.Range("K" & Ticker_Table_Row).NumberFormat = "0.00%"
                
                'Add the Stock Volume
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                
                'Output the Ticker in the ticker summary table
                ws.Range("I" & Ticker_Table_Row).Value = Ticker_Symbol
                
                'Output the Total Stock Volume in the ticker summary table
                ws.Range("L" & Ticker_Table_Row).Value = Total_Stock_Vol
                
                'Format the ticker summary table using autofit
                ws.Range("I1:L" & Ticker_Table_Row).Columns.AutoFit
                
                'Move to the next row of ticker summary table
                Ticker_Table_Row = Ticker_Table_Row + 1
                
                'Reset the Total Tock Volume to Zero
                Total_Stock_Vol = 0
                
            'If the next row has the same ticker
            Else
            
                'Add the Stock Volume
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
            End If
        Next i
        
        '''Show Greatest % Increase, Greatest % Decrease and Greatest Total Volume
        
        'Headers for the table to show greatest % increase, greatest % decrease and greatest total volume
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'Set a variable to store last row of the ticker summary table
        Dim Last_Row_Ticker_Table As Long
        Last_Row_Ticker_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
        
        'Declare variables to store ticker, greatest % increase, greatest % decrease and greatest total volume
        Dim Ticker_Per_Increase, Ticker_per_Decrease, Ticker_total_vol As String
        Dim Greatest_Percent_Increase, Greatest_Percent_Decrease, Greatest_Total_Volume
        Greatest_Percent_Increase = 0
        Greatest_Percent_Decrease = 0
        Greatest_Total_Volume = 0
        
        'Loop through each ticker from the ticker summary table
        Dim j As Long
        For j = 2 To Last_Row_Ticker_Table
        
            'Check for greatest % increase record
            If ws.Cells(j, 11).Value > Greatest_Percent_Increase Then
                
                'Set the greates % increase value
                Greatest_Percent_Increase = ws.Cells(j, 11).Value
                
                'Set the ticker name
                Ticker_Per_Increase = ws.Cells(j, 9).Value
                
            'Check for greatest % decrease
            ElseIf ws.Cells(j, 11).Value < Greatest_Percent_Decrease Then
            
                 'Set the greates % decrease value
                Greatest_Percent_Decrease = ws.Cells(j, 11).Value
                
                'Set the ticker name
                Ticker_per_Decrease = ws.Cells(j, 9).Value
                
            End If
            
            'Check for greatest total volume
            If ws.Cells(j, 12).Value > Greatest_Total_Volume Then
            
                 'Set the greates total volume value
                Greatest_Total_Volume = ws.Cells(j, 12).Value
                
                'Set the ticker name
                Ticker_total_vol = ws.Cells(j, 9).Value
                
            End If
        Next j
        
        'Output the greatest % increase
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("O2").Value = Ticker_Per_Increase
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P2").Value = Greatest_Percent_Increase
        
        'Output the greatest % decrease
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("O3").Value = Ticker_per_Decrease
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P3").Value = Greatest_Percent_Decrease
        
        'Output the greatest total volume
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O4").Value = Ticker_total_vol
        ws.Range("P4").NumberFormat = "0.00E+00"
        ws.Range("P4").Value = Greatest_Total_Volume
        
        'Format the table using autofit
        ws.Range("N1:P4").Columns.AutoFit
        
    Next ws
    
End Sub


