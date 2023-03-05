Attribute VB_Name = "VBA_challenge"
Sub stocks()
'Repeat for each sheet
For Each ws In Worksheets

    ' Determine the Last Row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Esablish variables
    Dim stock_ticker As String
    Dim stock_total As Double
    Dim row_num As Integer
    Dim first_open_price As Double
    Dim last_close_price As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim ticker_num As Integer

    row_num = 2
    stock_total = 0
    ticker_num = 0

    'Create Titles
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Percent Change"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Range("M1").Select

    For i = 2 To lastRow 'Move through rows 2 to 2000
    
        'Get values for first stock in group
        first_open_price = ws.Cells(i, 3)
    
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then 'Check to see if both ws.Cells are different. If So:
        
            'stock_ticker is set to the value of previous cell
            stock_ticker = ws.Cells(i, 1).Value
        
            'On Table, add value of stock_ticker
            ws.Cells(row_num, 9) = stock_ticker
        
            'Add Last cell to running total
            stock_total = stock_total + ws.Cells(i, 7).Value
        
            'Print total to Table
            ws.Cells(row_num, 12) = stock_total
        
            'Reset Total
            stock_total = 0
        
            'Get values or last stock in group
            last_close_price = ws.Cells(i, 6)
        
            'Calculate Values
            year_change = last_close_price - first_open_price
            'percentChange = (currentPrice - previousPrice) / previousPrice * 100
            percent_change = (last_close_price - first_open_price) / first_open_price * 100
            'Print values to Table
            ws.Cells(row_num, 10).Value = year_change
        
            ws.Cells(row_num, 11).Value = Round(percent_change, 2) / 100 'Round percent_change
            ws.Cells(row_num, 11).NumberFormat = "0.00%" 'Format cell
        
            'Color Values
            If year_change >= 0 Then
                ws.Cells(row_num, 10).Interior.Color = RGB(0, 255, 0)
        
            ElseIf year_change < 0 Then
                ws.Cells(row_num, 10).Interior.Color = RGB(255, 0, 0)
        
            End If
        
            If percent_change >= 0 Then
                ws.Cells(row_num, 11).Interior.Color = RGB(0, 255, 0)
            
            ElseIf percent_change < 0 Then
                ws.Cells(row_num, 11).Interior.Color = RGB(255, 0, 0)
        
            End If
        
            'Add to total number of tickers
            ticker_num = ticker_num + 1
            'Advance to the next row on Table
            row_num = row_num + 1
    
        Else
            'If ws.Cells are the same, add to stock total
            stock_total = stock_total + ws.Cells(i, 7).Value
        
        End If
    

    Next i

    'After creating table, find Greatest Values

    'Create Table
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "Greatest % Increase"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "Greatest % Decrease"
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "Greatest Total Volume"
    Range("N5").Select

    'Find Values
   

    'Find Max Value %
    Dim percent_lastRow As Long
    Dim maxVal As Double
    Dim maxRow As Long
    Dim val_ticker As Variant
    
    'Find the last row of data in column K
    percent_lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Find the maximum value in column A and its row number
    maxVal = WorksheetFunction.Max(Range("K2:K" & percent_lastRow))
    maxRow = WorksheetFunction.Match(maxVal, Range("K2:K" & percent_lastRow), 0)
    
   'Find the ticker symbol in column A for the row where the max value in column K is found
    val_ticker = ws.Cells(maxRow, "I").Value

    'Assign the value of the ticker symbol to cell O2
    Range("O2").Value = val_ticker
    
    'Print Values
    ws.Cells(2, 16).Value = maxVal
    
    'Format %
    ws.Cells(2, 16).NumberFormat = "0.00%" 'Format cell
    
    'Find Min Value %
    Dim minVal As Double
    Dim minRow As Long
    
    'Find the minimum value in column A and its row number
    minVal = WorksheetFunction.Min(Range("K2:K" & percent_lastRow))
    minRow = WorksheetFunction.Match(minVal, Range("K2:K" & percent_lastRow), 0)
    
    'Retrieve the value of ticker in the row where the minimum value is found
    val_ticker = ws.Cells(minRow, "I").Value
    
    'Print Values
    ws.Cells(3, 16).Value = minVal
    
    'Format %
    ws.Cells(3, 16).NumberFormat = "0.00%" 'Format cell
    
    'Print Ticker
    Range("O3").Value = val_ticker

    Dim totalVal As Double
    Dim totalRow As Long
    
    'Find the total value in column A and its row number
    totalVal = WorksheetFunction.Max(Range("L2:L" & percent_lastRow))
    totalRow = WorksheetFunction.Match(totalVal, Range("L2:L" & percent_lastRow), 0)
    
    'Retrieve the value of cell B1 in the row where the total value is found
    val_ticker = ws.Cells(totalRow, "I").Value
    
    'Print Values
    ws.Cells(4, 16).Value = totalVal
    
    'Print Ticker
    Range("O3").Value = val_ticker
Next ws
   

End Sub





