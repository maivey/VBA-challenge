Attribute VB_Name = "stockVBA_Chal_1"
'*********CHALLENGE 1***********
'*********ONE SHEET*************
Sub stockVBA_Chal_1()
' declare worksheet as ws
'Dim ws As Worksheet

'loop through each worksheet
'For Each ws In Worksheets

    'declare variables
    Dim tickerSym As String
    Dim total_Vol As Double
    Dim percent_Change As Double
    Dim yearlyChange As Double
    Dim closePrice As Double
    Dim openPrice As Double
    
    'set total volume = 0 at begining
    total_Vol = 0
    'declare summary table row and set = 2 for summary output table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    
    
    ' find number of rows in each sheet
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
     
     'declare variables for greatest % increase, greatest % decrease, greatest volume
    Dim maxPerc As Double
    Dim minPerc As Double
    Dim maxVol As Double
    Dim maxTicker As String
    'set maxPerc, minPerc, maxVol = 0 at first
    maxPerc = 0
    minPerc = 0
    maxVol = 0
    
    'all worksheets
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'set first ticker's start price
    openPrice = Cells(2, 3).Value
    
    
    For i = 2 To last_row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            tickerSym = Cells(i, 1).Value
            
            'calculate close price
            closePrice = Cells(i, 6).Value
            'calculate yearly change
            yearlyChange = closePrice - openPrice
            
            'calculate percent change
            '***ACCOUNT FOR 0 VALUES- WILL PRODUCE ERROR
            If openPrice <> 0 Then
                percent_Change = (yearlyChange / openPrice) * 100
           ' ElseIf (openPrice = 0) And (closePrice = 0) Then
            ElseIf closePrice = 0 Then
                percent_Change = 0
                yearlyChange = 0
            End If
            
        
           total_Vol = total_Vol + Cells(i, 7).Value
           ' Print the Credit Card Brand in the Summary Table (col I)
          Range("I" & Summary_Table_Row).Value = tickerSym
    
          ' Print the Brand Amount to the Summary Table (col L)
          Range("L" & Summary_Table_Row).Value = total_Vol
          
          'Print yearlyChage to summary table col J
          Range("J" & Summary_Table_Row).Value = yearlyChange
          
          'print percent Change to summary table col K
          Range("K" & Summary_Table_Row).Value = Str(percent_Change) & "%"
          'color red (index=3) for negative and green(index=4) for positive % changes
          If (yearlyChange > 0) Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf (yearlyChange <= 0) Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
          
          
          'add one to summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          'reset yearlyChange, and closePrice
          yearlyChange = 0
          closePrice = 0
          
          'set next ticker open price
          openPrice = Cells(i + 1, 3).Value
    
        ' loop to find maxPerc change
        If percent_Change > maxPerc Then
            maxPerc = percent_Change
            maxTicker = tickerSym
        End If
        If percent_Change < minPerc Then
            minPerc = percent_Change
            minTicker = tickerSym
        End If
        If total_Vol > maxVol Then
            maxVol = total_Vol
            maxVol_ticker = tickerSym
        End If
        
        'reset total_Vol
        total_Vol = 0
            
            
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
           
        
            total_Vol = total_Vol + Cells(i, 7).Value
            '*******************************
            '**account for 0 open price: if 0 open price, set openPrice as next non-zero value'
            If openPrice = 0 Then
                openPrice = Cells(i + 1, 3).Value
                'Debug.Print openPrice
            End If
                
             '*******************************
            
    
        End If
    
    Next i
    
    'MAX PERCENT TABLE
    '----------------
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    
    Dim last_row_sumTable As Long
    last_row_sumTable = Cells(Rows.Count, 9).End(xlUp).Row
    'Debug.Print last_row_sumTable
    'Dim maxPerc As Double
    'Dim minPerc As Double
    'Dim maxVol As Double
    'Dim maxTicker As String
    '
    'maxPerc = 0
    'minPerc = 0
    'maxVol = 0
    'For j = 2 To last_row_sumTable
    '    If Cells(j, 11).Value > maxPerc Then
    '        maxPerc = Cells(j, 11).Value
    '        maxTicker = Cells(j, 9).Value
    '    End If
    '    If Cells(j, 11).Value < minPerc Then
    '        minPerc = Cells(j, 11).Value
    '        minTicker = Cells(j, 9).Value
    '    End If
    '    If Cells(j, 12).Value > maxVol Then
    '        maxVol = Cells(j, 12).Value
    '        maxVol_ticker = Cells(j, 9).Value
    '    End If
    '
    '
    'Next j
    Range("P2").Value = maxTicker
    Range("Q2").Value = Str(maxPerc) & "%"
    Range("P3").Value = minTicker
    Range("Q3").Value = Str(minPerc) & "%"
    Range("P4").Value = maxVol_ticker
    Range("Q4").Value = maxVol
    
    Columns("I:Q").AutoFit
    
 '   Next ws

End Sub

