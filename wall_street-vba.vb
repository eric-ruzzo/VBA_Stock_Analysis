Sub Stocks()

'Declare variables to identify tickers and their locations
Dim ticker As String
Dim ticker_row As Integer
Dim first_row_ticker As Long
Dim last_row_ticker As Long
Dim ticker_total As Integer

'Declare variables to find values related to yearly/percent changes and volume
Dim yearly_change As Double
Dim year_open As Double
Dim year_close As Double
Dim percent_change As Double
Dim volume As Long

'Declare variable to find last row of original data set
Dim data_end_row As Long

'Declare variables to find min/max values and their locations
Dim max_percent As Double
Dim min_percent As Double
Dim max_volume As Long
Dim max_percent_row As String
Dim min_percent_row As String
Dim max_volume_row As String

'Declare variables used to create rows in a new table
Dim yearly_change_row As Integer
Dim percent_change_row As Integer
Dim last_change_row As Integer
Dim volume_row As Integer

'Declare variables used to loop through worksheets
Dim year As String
Dim ws As Worksheet

'Set variable for worksheet object
Set ws = ActiveWorkbook.ActiveSheet

For Each ws In ActiveWorkbook.Worksheets

year = ws.Name

'Set values for row variables
data_end_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
ticker_row = 1
yearly_change_row = 1
percent_change_row = 1
volume_row = 1

'Set initial values for yearly change, percent change and total volume
yearly_change = 0
percent_change = 0
volume = 0

'Set intitial value for ticker counter
ticker_count = 0

'Loop through rows to find totals
For i = 2 To data_end_row
    
    'Create conditional to determine if ticker in next row is the same as the current row
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

        'Set variable for ticker
        ticker = ws.Cells(i, 1).Value

        'Determine value of open price at start of year
        year_open = ws.Cells(i + 1, 3).Value

        'Find last row of data for each ticker
        last_row_ticker = ws.Cells(i + 1, 1).Row - 1

        'Set variables for year opening and year closing price
        year_open = ws.Cells(last_row_ticker - ticker_count, 3).Value
        year_close = ws.Cells(last_row_ticker, 6).Value
        
        'Create conditional to change year_open to 1 if it is actually equal to 0
        If year_open = 0 Then
            year_open = 1
            
        Else: year_open = year_open
        
        End If

        'Set variables for yearly change, percent change and total volume
        yearly_change = year_close - year_open
        percent_change = (yearly_change / year_open)
        volume = volume + ws.Cells(i, 7).Value

        'Enter ticker in a new cell
        ws.Cells(ticker_row + 1, 9) = ticker

        'Enter yearly change, percent change and volume for each ticker in a new cell
        ws.Cells(yearly_change_row + 1, 10) = yearly_change
        ws.Cells(percent_change_row + 1, 11) = percent_change
        ws.Cells(volume_row + 1, 12) = volume

        'Enter each ticker, yearly change, percent change & volume into a new row
        ticker_row = ticker_row + 1
        yearly_change_row = yearly_change_row + 1
        percent_change_row = percent_change_row + 1
        volume_row = volume_row + 1
        
        'Create conditional for if yearly change is positive or negative
        If yearly_change > 0 Then
            ws.Cells(yearly_change_row, 10).Interior.Color = vbGreen
            
        ElseIf yearly_change < 0 Then
            ws.Cells(yearly_change_row, 10).Interior.Color = vbRed
            
        End If
        
        'Format percentage change results as percentage
        ws.Cells(percent_change_row, 11).Style = "Percent"
        ws.Cells(percent_change_row, 11).NumberFormat = "0.00%"

        'Reset ticker counter & volume
        ticker_count = 0
        volume = 0
        
    Else
    
    ticker_count = ticker_count + 1

    End If
    
Next i
    
    'Find last row of new table
    last_change_row = ws.Cells(Rows.Count, 11).End(xlUp).Row

    'Use WorksheetFunction.Max to find max percent change & volume and min percent change
    max_percent = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_change_row))
    min_percent = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_change_row))
    max_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_change_row))
    
    'Use WorksheetFunction.Match to find row of corresponding ticker
    max_percent_row = Application.WorksheetFunction.Match(max_percent, ws.Range("K2:K" & last_change_row), 0)
    min_percent_row = Application.WorksheetFunction.Match(min_percent, ws.Range("K2:K" & last_change_row), 0)
    max_volume_row = Application.WorksheetFunction.Match(max_volume, ws.Range("L2:L" & last_change_row), 0)
    
    'Add headers to top row in new columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"

    'Add row titles for greatest % increase/decrease & greatest total volume
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Add ticker next to titles
    ws.Cells(2, 16).Value = ws.Cells(max_percent_row + 1, 9).Value
    ws.Cells(3, 16).Value = ws.Cells(min_percent_row + 1, 9).Value
    ws.Cells(4, 16).Value = ws.Cells(max_volume_row + 1, 9).Value
    
    'Add min/max percentage increase and volume next to ticker
    ws.Cells(2, 17).Value = max_percent
    ws.Cells(3, 17).Value = min_percent
    ws.Cells(4, 17).Value = max_volume
    
    'Format percentage change results as percentage
    ws.Range("Q2:Q3").Style = "Percent"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

    'Autofit to display headers and titles
    ws.Columns("A:Q").AutoFit
    
    Next

End Sub