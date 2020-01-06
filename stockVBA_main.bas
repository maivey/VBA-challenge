Attribute VB_Name = "stockVBA_main"
'*** Runs through one year of all stocks (Base Code)
' The method stockVBA_main loops through all the stocks for one year for each run and takes the following information.

  '* The ticker symbol.

  '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The total stock volume of the stock.

' Conditional formatting will highlight positive change in green and negative change in red.

Sub stockVBA_main()

    'Declare variables
    Dim tickerSym As String
    Dim total_Vol As Double
    Dim percent_Change As Double
    Dim yearlyChange As Double
    Dim closePrice As Double
    Dim openPrice As Double
    
    'Set total volume equal to zero at begining
    total_Vol = 0
    
    'Declare summary table row and set equal to 2 for summary output table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    ' Find number of rows in each sheet
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
     
    'Declare variables for greatest % increase, greatest % decrease, greatest volume
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Set first ticker's start price
    openPrice = Cells(2, 3).Value
    
    
    For i = 2 To last_row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set ticker symbol
            tickerSym = Cells(i, 1).Value
            
            ' Calculate close price
            closePrice = Cells(i, 6).Value

            ' Calculate yearly change
            yearlyChange = closePrice - openPrice
            
            ' Calculate percent change
            '***Account FOR 0 VALUES- WILL PRODUCE ERROR
            If openPrice <> 0 Then
                percent_Change = (yearlyChange / openPrice) * 100
            ElseIf closePrice = 0 Then
                percent_Change = 0
                yearlyChange = 0
            End If
            
           ' Calculate total volume
           total_Vol = total_Vol + Cells(i, 7).Value

           ' Print the Credit Card Brand in the Summary Table (col I)
           Range("I" & Summary_Table_Row).Value = tickerSym
    
          ' Print the Brand Amount to the Summary Table (col L)
           Range("L" & Summary_Table_Row).Value = total_Vol
          
           ' Print yearlyChage to summary table (col J)
           Range("J" & Summary_Table_Row).Value = yearlyChange
          
           ' Print percent Change to summary table (col K)
           Range("K" & Summary_Table_Row).Value = Str(percent_Change) & "%"

          ' Set cell color to red (index=3) for negative percent changes and green (index=4) for positive percent changes
           If (yearlyChange > 0) Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
           ElseIf (yearlyChange <= 0) Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
           End If
          
          
          'Add one to summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
        'Reset yearly change and and close price
          yearlyChange = 0
          closePrice = 0
          
          'Set next ticker open price
          openPrice = Cells(i + 1, 3).Value
        
        'Reset total_Vol
        total_Vol = 0
            
            
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
           
        
            total_Vol = total_Vol + Cells(i, 7).Value

            'Account for 0 open price: if 0 is the open price, set openPrice as next non-zero value'
            If openPrice = 0 Then
                openPrice = Cells(i + 1, 3).Value
            End If
            
        End If
    Next i
    
    ' AutoFit to format Summary Table
    Columns("I:Q").AutoFit

End Sub

