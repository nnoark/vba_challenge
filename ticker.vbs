Sub ticker()

'set ticker name as string
Dim ticker_name As String

'set number of tickers as an integer
Dim n_tick As Integer

'Set volume_total as long
Dim volume_total As Double

'set open as double
Dim opening As Double

'set closing as double
Dim closing As Double

'set yearly change as double
Dim year_change As Double

'set the percentage change as double
Dim percentage_change As Double

'Set last row as long
Dim lastrow As Long




'Set column names
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"



'Set variables to 0 to begin
ticker_name = ""
volume_total = 0
opening = 0
closing = 0
year_change = 0
percentage_change = 0



'Find and set last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 For i = 2 To lastrow
 
 'Find the ticker name
 ticker_name = Cells(i, 1).Value
 
 'Find opening price
If opening = 0 Then
opening = Cells(i, 3).Value
End If
 
 'Find volume total
 volume_total = Cells(i, 7).Value + volume_total
 
 'Check if the ticker name is the same
 If Cells(i + 1, 1).Value <> ticker_name Then
 n_tick = n_tick + 1
 Cells(n_tick + 1, 10) = ticker_name
 
 'Find the closing price
 closing = Cells(i, 6)
 
 'Calculate the change over the year and print in the summary table
 
 year_change = closing - opening
 Cells(n_tick + 1, 11).Value = year_change
 
  'Set conditional formatting
 If year_change > 0 Then
 Cells(n_tick + 1, 11).Interior.ColorIndex = 4
 ElseIf year_change < 0 Then
 Cells(n_tick + 1, 11).Interior.ColorIndex = 3
 Else
 End If
 
 'Calculate the percentage change
If opening = 0 Then
 percentage_change = 0
 Else: percentage_change = year_change / opening
End If
 
 'Format percentage change to an actual percentage and print
Cells(n_tick + 1, 12) = Format(percentage_change, "Percent")
 
 'Set conditional formatting
If percentage_change > 0 Then
 Cells(n_tick + 1, 12).Interior.ColorIndex = 4
 ElseIf percentage_change < 0 Then
 Cells(n_tick + 1, 12).Interior.ColorIndex = 3
 Else
 End If
 
 'Print Volume total
 Cells(n_tick + 1, 13).Value = volume_total
 
 'Reset variables
 opening = 0
 volume_total = 0
 
 End If
 
 Next i
 
 
 
 
 
 
 
 





End Sub