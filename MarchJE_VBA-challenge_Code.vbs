Attribute VB_Name = "Module1"
Sub Stock_Calculations()
'Stock_Calculations Macro is a VBA script to analyze real stock market data

    'Declaration of Variables
    Dim myworksheet As Worksheet 'used to walk through the current and other worksheets later
    Dim rowcounter As Long 'used to loop through rows
    Dim ticker As String 'used to store current ticker
    Dim opening As Double 'used to store the opening price at the beginning of the year for a given stock
    Dim closing As Double 'used to store the closing price at the end of the year for a given stock
    Dim volume As Double 'used to store the total volume of stock
    Dim newrow As Long 'used to determine which row we are writing to on the new table
    Dim GIncrease As Double 'used to determine the greatest % increase
    Dim GDecrease As Double 'used to determine the greatest % decrease
    Dim GVolume As Double 'used to determine the greatest total volume
    
    'Perform the calculations on every worksheet in the excel file
    For Each myworksheet In Worksheets
    
        'Set up headers for stock calculations table
        myworksheet.Cells(1, 9).Value = "Ticker"
        myworksheet.Cells(1, 10).Value = "Yearly Change"
        myworksheet.Cells(1, 11).Value = "Percent Change"
        myworksheet.Cells(1, 12).Value = "Total Stock Volume"
        
        'Set Column Sizes
        myworksheet.Columns("I").ColumnWidth = 12
        myworksheet.Columns("J").ColumnWidth = 14
        myworksheet.Columns("K").ColumnWidth = 14
        myworksheet.Columns("L").ColumnWidth = 18
        
        'Set up x and y axis of table for greatest % increase, greatest % decrease, and greatest total volume
        myworksheet.Cells(1, 16).Value = "Ticker"
        myworksheet.Cells(1, 17).Value = "Value"
        myworksheet.Cells(2, 15).Value = "Greatest % Increase"
        myworksheet.Cells(3, 15).Value = "Greatest % Decrease"
        myworksheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Column Sizes
        myworksheet.Columns("O").ColumnWidth = 20
        myworksheet.Columns("P").ColumnWidth = 12
        myworksheet.Columns("Q").ColumnWidth = 12
    
        'Set newrow to first row to be written to in new table
        newrow = 2
    
        'Set ticker to first value in the list of tickers and copy this ticker to the stock calculations table
        ticker = myworksheet.Cells(2, 1).Value
        myworksheet.Cells(newrow, 9).Value = ticker
                
        'Set opening to the opening price at the beginning of the year for the current ticker
        opening = myworksheet.Cells(2, 3).Value

        'Set volume to the volume of the first date of the current ticker
        volume = myworksheet.Cells(2, 7).Value
        
        'Set rowcounter to the next row to be worked
        rowcounter = 3
        
        'Loop through all the stocks for one year
        Do Until IsEmpty(myworksheet.Cells(rowcounter, 1).Value)
            
            If myworksheet.Cells(rowcounter, 1).Value = ticker Then
                'Increment volume
                volume = volume + myworksheet.Cells(rowcounter, 7).Value
       
            Else
                'Determine the closing price for the year for the current ticker
                closing = myworksheet.Cells((rowcounter - 1), 6).Value
                
                'Output the Yearly Change from opening price at the beginning of the year to closing price at the end of the year
                myworksheet.Cells(newrow, 10).Value = closing - opening
                
                'Conditional Formatting to highlight positive changes in green and negative changes in red
                If myworksheet.Cells(newrow, 10).Value >= 0 Then
                    myworksheet.Cells(newrow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    myworksheet.Cells(newrow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                'Check if there was 0 Yearly Change
                If myworksheet.Cells(newrow, 10).Value = 0 Then
                    'Output 0% for the percent change from opening price at the beginning of the year to closing price at the end of the year
                    myworksheet.Cells(newrow, 11).Value = 0
                    myworksheet.Cells(newrow, 11).NumberFormat = "0.00%"
                ElseIf opening = 0 Then
                    'Output the Percent Change as if opening price was $0.01 at the beginning of the year
                    myworksheet.Cells(newrow, 11).Value = (myworksheet.Cells(newrow, 10) / 0.01)
                    myworksheet.Cells(newrow, 11).NumberFormat = "0.00%"
                Else
                    'Output the Percent Change from opening price at the beginning of the year to closing price at the end of the year
                    myworksheet.Cells(newrow, 11).Value = (myworksheet.Cells(newrow, 10) / opening)
                    myworksheet.Cells(newrow, 11).NumberFormat = "0.00%"
                End If
                
                'Output the Total Stock Volume of the stock
                myworksheet.Cells(newrow, 12).Value = volume
                
                'Increment newrow to set it to the next row
                newrow = newrow + 1
                                
                'Set ticker equal to new ticker
                ticker = myworksheet.Cells(rowcounter, 1).Value
                
                'Output new ticker
                myworksheet.Cells(newrow, 9).Value = ticker
                
                'Reset new opening price
                opening = myworksheet.Cells(rowcounter, 3)
                
                'Reset new volume
                volume = myworksheet.Cells(rowcounter, 7)
            
            End If
                     
            'Increment rowcounter
            rowcounter = rowcounter + 1
            
        Loop
    
        'Run final calculations for last row
                'Determine the closing price for the year for the current ticker
                closing = myworksheet.Cells((rowcounter - 1), 6).Value
                
                'Output the Yearly Change from opening price at the beginning of the year to closing price at the end of the year
                myworksheet.Cells(newrow, 10).Value = closing - opening
                
                'Conditional Formatting to highlight positive changes in green and negative changes in red
                If myworksheet.Cells(newrow, 10).Value >= 0 Then
                    myworksheet.Cells(newrow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    myworksheet.Cells(newrow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                'Check if there was 0 Yearly Change
                If myworksheet.Cells(newrow, 10).Value = 0 Then
                    'Output 0% for the percent change from opening price at the beginning of the year to closing price at the end of the year
                    myworksheet.Cells(newrow, 11).Value = 0
                    myworksheet.Cells(newrow, 11).NumberFormat = "0.00%"
                ElseIf opening = 0 Then
                    'Output the Percent Change as if opening price was $0.01 at the beginning of the year
                    myworksheet.Cells(newrow, 11).Value = (myworksheet.Cells(newrow, 10) / 0.01)
                    myworksheet.Cells(newrow, 11).NumberFormat = "0.00%"
                Else
                    'Output the Percent Change from opening price at the beginning of the year to closing price at the end of the year
                    myworksheet.Cells(newrow, 11).Value = (myworksheet.Cells(newrow, 10) / opening)
                    myworksheet.Cells(newrow, 11).NumberFormat = "0.00%"
                End If
                
                'Output the Total Stock Volume of the stock
                myworksheet.Cells(newrow, 12).Value = volume

        'Return the Ticker and Value of the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
        
        'Set greatest % increase and greatest % decrease to 0
        GIncrease = 0
        GDecrease = 0
        
        'Set rowcounter to the top of the newly made table and iterate through new table using it to find/fill in Greatest % Increase and Greatest % Decrease
        rowcounter = 2
        Do Until IsEmpty(myworksheet.Cells(rowcounter, 9).Value)
            
            'Check for Percent Change being greater than GIncrease and copy it if so into GIncrease
            If myworksheet.Cells(rowcounter, 11).Value > GIncrease Then
                GIncrease = myworksheet.Cells(rowcounter, 11).Value
                
                'Copy the Ticker into the Greatest % Increase Cell with plans to copy over it if a better answer is discovered
                myworksheet.Cells(2, 16).Value = myworksheet.Cells(rowcounter, 9).Value
                
                'Copy the Value into the Greatest % Increase Cell with plans to copy over it if a better answer is discovered
                myworksheet.Cells(2, 17).Value = myworksheet.Cells(rowcounter, 11).Value
                myworksheet.Cells(2, 17).NumberFormat = "0.00%"
                
            'Check for Percent Change being less than GDecrease and copy it if so into GDecrease
            ElseIf myworksheet.Cells(rowcounter, 11).Value < GDecrease Then
                GDecrease = myworksheet.Cells(rowcounter, 11).Value
                
                'Copy the Ticker into the Greatest % Decrease Cell with plans to copy over it if a better answer is discovered
                myworksheet.Cells(3, 16).Value = myworksheet.Cells(rowcounter, 9).Value
                
                'Copy the Value into the Greatest % Decrease Cell with plans to copy over it if a better answer is discovered
                myworksheet.Cells(3, 17).Value = myworksheet.Cells(rowcounter, 11).Value
                myworksheet.Cells(3, 17).NumberFormat = "0.00%"
                            
            Else
            End If
        
            'Increment rowcounter
            rowcounter = rowcounter + 1
        Loop
        
        'Set greatest total volume to first value in list and copy Ticker and Value to Greatest Total Volume with plans to copy over it if a better answer is discovered
        GVolume = myworksheet.Cells(2, 12).Value
        myworksheet.Cells(4, 16).Value = myworksheet.Cells(2, 9).Value
        myworksheet.Cells(4, 17).Value = myworksheet.Cells(2, 12).Value
        
        'Set rowcounter to one below the top of the same table and iterate through table using it to find/fill in Greatest Total Volume
        rowcounter = 3
        Do Until IsEmpty(myworksheet.Cells(rowcounter, 9).Value)
        
            'See if Volume is greater than previous greatest volume. If so, replace it and copy Ticker and Value to Greatest Total Volume
            If myworksheet.Cells(rowcounter, 12).Value > GVolume Then
                GVolume = myworksheet.Cells(rowcounter, 12).Value
                myworksheet.Cells(4, 16).Value = myworksheet.Cells(rowcounter, 9).Value
                myworksheet.Cells(4, 17).Value = myworksheet.Cells(rowcounter, 12).Value
            Else
            End If
        
            'Increment rowcounter
            rowcounter = rowcounter + 1
        Loop
    
    Next

End Sub
