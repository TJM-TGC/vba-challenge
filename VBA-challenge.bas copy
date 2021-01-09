Attribute VB_Name = "Module1"
Sub StockTickers():

' Define Variables
Dim ws As Worksheet
Dim ticker As String
Dim num_of_tickers As Double
Dim lastRowState As Long
Dim opening As Double
Dim closing As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim highest_percent_increase As Double
Dim highest_increase_ticker As String
Dim highest_percent_decrease As Double
Dim highest_decrease_ticker As String
Dim highest_volume As Double
Dim highest_volume_ticker As String

'-----------------------
' Loop through worksheets
For Each ws In Worksheets

    'Activate WS

            ws.Activate

'Find last cell of worksheets
lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

' Assign Headers

ws.Cells(1, 9).Value = "Ticket"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

'Initialize Variables

number_of_tickers = 0
ticker = ""
yearly_change = 0
opening = 0
percent_change = 0
total_volume = 0

'Loop 1
For i = 2 To lastRowState

    'Find ticker
        ticker = Cells(i, 1).Value

    ' Opening Price
    If opening = 0 Then
        opening = Cells(i, 3).Value
    End If

    ' Sum stock volume for specific ticker
    total_volume = total_volume + Cells(i, 7).Value

    'Run if ticker is incorrect
    If Cells(i + 1, 1).Value <> ticker Then
        'Increment number of tickers when we get to a different ticker
    number_of_tickers = number_of_tickers + 1
    Cells(number_of_tickers + 1, 9) = ticker

    'Find closing price
    closing = Cells(i, 6)

    'absolute change
    yearly_change = closing - opening

'Fill cells
Cells(number_of_tickers + 1, 10).Value = yearly_change

'If change is positive go green
If yearly_change > 0 Then
    Cells(number_of_tickers + 1, 10).Interior.ColorIndex = 4
'Otherwise, shade it red
        ElseIf yearly_change < 0 Then
            Cells(number_of_tickers + 1, 10).Interior.ColorIndex = 3
  End If
           
 '% Change per Ticker
           
If opening = 0 Then
    percent_change = 0
Else
    percent_change = yearly_change / opening
    
End If

' Format as %
Cells(number_of_tickers + 1, 11).Value = Format(percent_change, "Percent")


' Reset opening price w/ new ticker
opening = 0

'Fill cells with total volume
Cells(number_of_tickers + 1, 12).Value = total_volume

'Reset opening price w/ new ticker again
total_volume = 0
    End If

Next i


'---------------
'Bonus Solution
Range("O2").Value = "Highest % Increase"
Range("O3").Value = "Highest % Decrease"
Range("O4").Value = "Highest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Find Last Row
lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row

' Zero-out variables
highest_percent_increase = Cells(2, 11).Value
highest_increase_ticker = Cells(2, 9).Value
highest_percent_decrease = Cells(2, 11).Value
highest_decrease_ticker = Cells(2, 9).Value
highest_volume = Cells(2, 12).Value
highest_volume_ticker = Cells(2, 9).Value

For i = 2 To lastRowState
    
    'Find highest growing ticker name
    If Cells(i, 11).Value > highest_percent_increase Then
        highest_percent_increase = Cells(i, 11).Value
        highest_increase_ticker = Cells(i, 9).Value
    End If
        
    'Find highest decrease ticker
    If Cells(i, 11).Value < highest_percent_decrease Then
    highest_percent_decrease = Cells(i, 12).Value
    highest_decrease_ticker = Cells(i, 9).Value
    End If
    
    'Find highest volume by ticker
    If Cells(i, 12).Value > highest_volume Then
    highest_Stock_volume = Cells(i, 12).Value
    highest_volume_ticker = Cells(i, 9).Value
    End If
    
Next i

' Add values
Range("P2").Value = Format(highest_increase_ticker, "Percent")
Range("Q2").Value = Format(highest_percent_increase, "Percent")
Range("P3").Value = Format(highest_decrease_ticker, "Percent")
Range("Q3").Value = Format(highest_percent_decrease, "Percent")
Range("P4").Value = highest_volume_ticker
Range("Q4").Value = highest_volume

Next ws
    

End Sub


