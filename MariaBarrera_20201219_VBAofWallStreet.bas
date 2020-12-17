Attribute VB_Name = "Module1"
'Subroutine: Summarize Stock Totals
'Author:  Maria Barrera
'Date created:  12/12/2020
'Decription:  A VBA script that will loop through all the stocks for one year and output the following information:
'1) The ticker symbol.
'2) Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'3) The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'4) The total stock volume of the stock.
'5) Also has conditional formatting that will highlight positive change in green and negative change in red.

Sub Summarize_Stock_Totals()

'Set headers for Ticker, Yearly Change, Percent Change & Total Stock Volume
Range("I" & 1).Value = "Ticker"
Range("J" & 1).Value = "Yearly Change"
Range("K" & 1).Value = "Percent Change"
Range("L" & 1).Value = "Total Stock Volume"

'TEMP headers & holders -- to be modified (no need to print?)
'Range("M" & 1).Value = "Date Last-4" -- to be deleted -- no longer needed
Range("N" & 1).Value = "Start Price"
Range("O" & 1).Value = "End Price"
Range("P" & 1).Value = "Diff"
Dim Start_Price As Double
Dim End_Price As Double
Dim Diff As Double

'Set a variable for holding the stock Ticker name
Dim Stock_Name As String

'Set initial variables for holding the Yearly Change, Percent Change & Total Stock Volume (per stock)
Dim Yearly_Change As Double
Yearly_Change = 0

Dim Percent_Change As Double
Percent_Change = 0

Dim Stock_Total As Double
Stock_Total = 0

Dim Stock_Date As String

'---------------------------------------------------------
'Set the temp variables to zero for every new stock
Start_Price = 0
End_Price = 0
Diff = 0
'---------------------------------------------------------


'Keep track of the location for each stock name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Determine last row of spreadsheet
lastrow = Cells(Rows.Count, 1).End(xlUp).Row


'Go through the rows of the spreadsheet
'Loop through all the stock volumes
For i = 2 To lastrow
    
    'Get the date, check if ending in "0101" ==> then update stock start price
    Stock_Date = Right((Range("B" & i).Value), 4)
    If Stock_Date = "0101" Then
        
        'Print last 4 of stock_date -- just for reference testing only -- to be deleted
        'Range("M" & i).Value = Stock_Date
        
        'Get the start price & print
        Start_Price = Cells(i, 3).Value
        Range("N" & Summary_Table_Row).Value = Start_Price
        
    End If
    
    
    'Check if the cell still has the same stock name, if not then track the stock name & add to the stock total, update the summary table:
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set the stock name
        Stock_Name = Cells(i, 1).Value
        
        'Add to the Stock Total
        Stock_Total = Stock_Total + Cells(i, 7).Value
        
        'Set End Price
        End_Price = Cells(i, 6).Value
        
        'Print the Stock name in the summary table
        Range("I" & Summary_Table_Row).Value = Stock_Name
        
        'Print the Stock_Total in the summary table
        Range("L" & Summary_Table_Row).Value = Stock_Total
        
        'Print the End Price in the summary table
        Range("O" & Summary_Table_Row).Value = Cells(i, 6).Value
        
        'Calculate price difference
        Diff = End_Price - Start_Price
        
        'Print the Difference in the summary table
        Range("P" & Summary_Table_Row).Value = Diff
        
        
        'Calculate the Yearly Change (aka Diff), Percent Change & print to summary table & change cell color
        Range("J" & Summary_Table_Row).Value = Diff
        If Diff < 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3  'Red
            Else
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4  'Green
        End If
            
        Percent_Change = (Diff / Start_Price) * 100
        Range("K" & Summary_Table_Row).Value = Round(Percent_Change, 2)

        '---------------------------------------
        'Increment row & reset Stock Total
        '---------------------------------------
            
        'Increment the summary table row by 1
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset the Stock Total
        Stock_Total = 0
       
       
        Else
    
        'If the cell immediately following a row is the same stock name....
        'Add to the Stock_Total
        Stock_Total = Stock_Total + Cells(i, 7).Value
                    
        End If
    Next i
    
    'Adjust column widths for summary table
    Range("I1").ColumnWidth = 10
    Range("J1").ColumnWidth = 14
    Range("K1").ColumnWidth = 14
    Range("L1").ColumnWidth = 18
     
    'Right align the column headers
    Range("J1:L1").HorizontalAlignment = xlRight
    
End Sub

