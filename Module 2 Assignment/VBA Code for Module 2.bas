Attribute VB_Name = "Module1"
Sub Ticker_Tracker()

For Each ws In Worksheets

' creates headings for summary
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"

' headings for extra functionality
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    '  find last row for for loop
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' variable for the ticker symbol
    Dim Symbol As String

    ' variable for volume, to be used as counter
    Dim Volume As LongLong

    ' variable for the opening stock price on the first day of the year
    Dim FirstPrice As Double

    ' variable for the closing stock price on the last day of the year
    Dim LastPrice As Double

    ' variable for the table, set the first value to 2
    Dim TickerTable As Integer
    TickerTable = 2


    ' begin for loop to move through data
    For i = 2 To LastRow
        'resets the ticker symbol
        Symbol = ""
    
        'checks to see if the previous row is different than the subsequent in order to capture the opening price
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
         'sets first price
            FirstPrice = ws.Cells(i, 3).Value
        
         'volume count starts
            Volume = Volume + ws.Cells(i, 7).Value
        
        ' logic for capturing the close information
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' sets the new ticker symbol
            Symbol = ws.Cells(i, 1).Value
        
        ' adds to the volume
            Volume = Volume + ws.Cells(i, 7).Value
        
        'captures the close price on the last day of the year
            LastPrice = ws.Cells(i, 6).Value
        
        ' Puts the symbol into the sheet
            ws.Range("I" & TickerTable).Value = Symbol
        
        ' The yearly change in absolute terms
           ws.Range("J" & TickerTable).Value = LastPrice - FirstPrice
        
        'Calculates the percent change and accounts for any 0 values
            If FirstPrice <> 0 Then
                ws.Range("K" & TickerTable).Value = (LastPrice - FirstPrice) / FirstPrice
            End If
        
        'Puts the volume into the sheet
            ws.Range("L" & TickerTable).Value = Volume
        
        'Updates where new info is stored on the sheet
            TickerTable = TickerTable + 1
        
        'Resets variables
            Volume = 0
            LastPrice = 0
        
        Else
        ' when adjacent tickers are the same, only the volume will be updated
            Volume = Volume + ws.Cells(i, 7).Value
        End If
    
    Next i

'fill yearly change depending on positive or negative value or no change

    For n = 2 To TickerTable
        If ws.Cells(n, 10).Value > 0 Then
            ws.Cells(n, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(n, 10).Value < 0 Then
            ws.Cells(n, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(n, 10).Value = 0 Then
            ws.Cells(n, 10).Interior.ColorIndex = 44
        End If
        
    Next n
    
'change format of percent change to percentages
    Set Rng = ws.Range("K:K")
    Rng.NumberFormat = "0.00%"
    
' change format for greatest increase and decrease
    Set Rng2 = ws.Range("Q2:Q3")
    Rng2.NumberFormat = "0.00%"
    
' change format of volume to add commas
    Set Rng3 = ws.Range("L:L")
    Rng3.NumberFormat = "#,##0"
    
    'change format of greatest volume
    Cells(4, 17).NumberFormat = "#,##0"

' variable for the max increase
    Dim maximum As Double
    maximum = 0

'variable to store the max gain ticker
    Dim max_ticker As String
    max_ticker = ""

'variable for the min increase
    Dim minimum As Double
    minimum = 0

'variable to store the min gain ticker
    Dim min_ticker As String
    min_ticker = ""

'variable to store maximum volume value
    Dim max_vol As LongLong
    max_vol = 0

'variable to store maximum volume ticker
    Dim max_vol_ticker As String
    max_vol_ticker = ""

    'Dim new_ticker As Integer
    'new_ticker = TickerTable - 2
    
    For m = 2 To TickerTable

    ' compares value to find and store max and max ticker
        If ws.Cells(m, 11).Value > maximum Then
            maximum = ws.Cells(m, 11).Value
            max_ticker = ws.Cells(m, 9).Value
        End If
        
    ' compares value to find and store min and min ticker
        If ws.Cells(m, 11).Value < minimum Then
            minimum = ws.Cells(m, 11).Value
            min_ticker = ws.Cells(m, 9).Value
        End If
    
    ' compares values to find and store max vol value and ticker
        If ws.Cells(m, 12).Value > max_vol Then
            max_vol = ws.Cells(m, 12).Value
            max_vol_ticker = ws.Cells(m, 9).Value
        End If

    Next m

' displays the values in the sheet
    ws.Cells(2, 16).Value = max_ticker
    ws.Cells(2, 17).Value = maximum
    ws.Cells(3, 16).Value = min_ticker
    ws.Cells(3, 17).Value = minimum
    ws.Cells(4, 16).Value = max_vol_ticker
    ws.Cells(4, 17).Value = max_vol


Next ws

End Sub

