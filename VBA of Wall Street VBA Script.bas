Attribute VB_Name = "Module1"
Sub Stonks()

'Run code across each sheet in the excel file

Dim ws As Worksheet

' Each cell/range must now have (ws.) in front of it so that it runs through each worksheet
For Each ws In Worksheets


' ---------------------------------------------------'
'           DECLARING INITIAL VARIABLES
'                               &
'           FORMATTING TABLES/CHARTS


' General vba script used to find the last row for any column
    Dim last_row As Double
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' Create a summary table that will eventually house the pulled info
    Dim summary_table As Integer
    summary_table = 2

'print summary table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"


' -----------------------------------------------------------------------------'

' Dim the variables that will be pulled when looping row by row
' Ticker symbol, volume of stock, opening value for the year, & closing value for the year
    Dim ticker As String
    Dim stock_vol As Double
    Dim yr_open As Double
    Dim yr_close As Double
    
' Dim the variables that will be used for totals or calculations
    Dim total_vol As Double

' Set for loop to go through the whole years stock info
' Data starts in row to and will go until the (last_row) function recognizes the final row of data
    For i = 2 To last_row
    
    ' This gave me issues for a long time of not calculating the total currently
    ' The debug helped me realize that stock_vol was not being pulled currently if it was not given its value in the if statements, moving it here solved the problem
    stock_vol = ws.Cells(i, 7)
    ' Debug.Print (stock_vol)
        
        ' Use "if" logic to collect information for each stock
        ' I found issues later in the assaignment if the open value = 0, starting with this if statement solved those issues from what I can tell
        If ws.Cells(i, 3).Value = 0 Then
    
        ' If the cell below does not have a matching ticker symbol, then we will record the ticker symbol
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                ticker = ws.Cells(i, 1)
           
            End If
                
        ' If the cell below is the same as the cell we are on, then we will record it's volume
        ElseIf ws.Cells(i + 1, 1) = ws.Cells(i, 1) Then

            ' To get our total volume for that stock, we will use the below function to add the recorded stock volume to a total volume variable
            total_vol = total_vol + stock_vol
                    
            ' If the previous cells ticker is not equal to the cell we are on, then we know that this is the opening value for that stock
            If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
                yr_open = ws.Cells(i, 3)
            
            End If
    
    ' Now we are outside of our nested if statements and back to our original if, which checks if there is a value in the 'open' column
    ' We can use the Else to collect the information we want from each row print it the summary table
        
        ' Define the values that will be pulled
        ' Using the names for easier readability in the code
        Else
        ticker = ws.Cells(i, 1).Value
        yr_close = ws.Cells(i, 6).Value
        
        ' Equation to calculate the total volume of each stock
        total_vol = total_vol + stock_vol
        
        ' Print the ticker symbol and and the total stock volume over to the summary table
        ws.Range("I" & summary_table).Value = ticker
        ws.Range("L" & summary_table).Value = total_vol
        
        ' Before doing the next part we need to make sure we are not going to divide by zero and get an error
            If total_vol > 0 Then
            
            ' Equation to calculate the change in the stocks value over the course of the year
            ' Then print the yearly change to the summary table
            yr_change = yr_close - yr_open
            
            ws.Range("J" & summary_table).Value = yr_change
        
            ' With all of the above information gathered we can run the equation for Percent change
            ' And print that to our summary table
            ws.Range("K" & summary_table).Value = yr_change / yr_open
    
            ' Going back to the highest level "if"
            Else
                'set yearly and % change to zero if no stock data
                ws.Range("J" & summary_table).Value = 0
                ws.Range("K" & summary_table).Value = 0
            
            End If
           
' --------------------------------------------------------------'
'           FORMATTING FOR THE SUMMARY TABLE SECTION
           
' Set the Percent Change column "K" of the summary table to display as a percent
    ws.Range("K" & summary_table).Style = "percent"
                        
' Add conditional formatting to the yearly change column depending on the results
    ' This if statement uses the built-in color index code and values
    ' 4 = Red, 3 = Green, 0 = No color
        If ws.Range("J" & summary_table).Value > 0 Then
            ws.Range("J" & summary_table).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & summary_table).Value < 0 Then
            ws.Range("J" & summary_table).Interior.ColorIndex = 3
        ' Unsure if the clear (Else) statement is needed, keep just incase a yearly change value happens to be zero
        Else
            ws.Range("J" & summary_table).Interior.ColorIndex = 0
                
        End If
                        
' Set the total volume back to zero before moving onto the next stock ticker
    total_vol = 0
        
' Set summary table to move to the next row
    summary_table = summary_table + 1
          
' This is the end of the original if and first for loop
    End If

Next i

'-----------------------------------------------------------------'
'                   BONUS SECTION

' Print bonus table headers
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Yearly Change"

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Yearly Change"

' Define the variables that will be pulled when looping through our summary table from above
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_vol As Double
    
 ' I will also set the current greatest variable levels to 0 so that loop starts working right away
    greatest_increase = 0
    greatest_decrease = 0
    greatest_vol = 0

' Define variables within our summary table for readablity in the code
    Dim ticker_sym As String
    Dim percent_change As Double
    Dim total_stock_vol As Double


' Our first for loop will look for the greatest increase and decrease in precentage change from the summary table
    For i = 2 To last_row

' Initilize the ticker symbol before doing the loops - Was printing wrong ticker before
    ' I am now just looking at the ticker values printed in the summary table, so that is why I am not using the same 'ticker' that I dim'd before
    ' This also helped me keep the two concepts seperate when writing my code, not sure if it is redundent
    ticker_sym = ws.Cells(i, 9).Value
    'Debug.Print (ticker_sym)
        
    ' Before running the if statement, I need to define a value to use
    percent_change = ws.Cells(i, 11).Value

        ' I will again use nested if statements to find and pull the information I need
        If percent_change > greatest_increase Then
            ' If the cell % change is bigger than our recorded greatest increase, then we update the greatest increase value and move to the next row
            greatest_increase = percent_change
            
            ' We can now grab the connnected ticker symbol from the summary table
            ' And we can move that information into the bonus table
            
            'THIS MAY NEED TO BE   UPDATED   TO BE CLEANED UP
            ws.Range("P2").Value = ticker_sym
            ws.Range("Q2").Value = greatest_increase
            
        ' We will use an else if now to basically look for the opposite information as before, and find biggest decrease
        ElseIf percent_change < greatest_decrease Then
            greatest_decrease = percent_change
            
            ' Same as before, print these values to our bonus section
            ws.Range("P3").Value = ticker_sym
            ws.Range("Q3").Value = greatest_decrease
            
        End If
    
    Next i

' (Hopefully) final for loop and if statement for this code
' Now we will find the stock with the greatest total volume from our summary table
    For i = 2 To last_row
        
        ' We will take the "total volume" information from the summary table and compare it row by row to get our stock with the greatest volume
        total_stock_vol = ws.Cells(i, 12).Value

        ' Initilize the ticker symbol again for this for loop - Was printing wrong ticker before
        ticker_sym = ws.Cells(i, 9).Value

        
        ' We will use the "total_vol" value that we defined in the original problem to do this
        ' As we loop down the rows of the summary table we test to see if the total stock volume is greater than our recorded greatest volume
        If total_stock_vol > greatest_vol Then
        
            ' If it is then we update the greatest volume
            greatest_vol = total_stock_vol
        
            ' Print the associated ticker symbol and the greatest stock volume to the bonus table
            ws.Range("P4").Value = ticker_sym
            ws.Range("Q4").Value = greatest_vol
        
        End If
    
    Next i

' --------------------------------------------------------------'
'           FORMATTING FOR THE BONUS SECTION
'                                   &
'                       FINAL FORMATTING
           
' Set the greatest increase and decrease value cells to display as a percent
    ws.Range("Q2").Style = "percent"
    ws.Range("Q3").Style = "percent"

' Googled around to find this VBA code that will auto adjust all of the columns to display properly
    ws.Columns("I:Q").AutoFit

' Have the change Percent column display to two decimals
    ws.Columns("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

End Sub
