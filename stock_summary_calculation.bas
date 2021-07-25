Attribute VB_Name = "Module1"
Public Sub stockcalculation()

Dim last_row, i, j As Long
Dim summary_table_row As Long
Dim initial_value, GreatestPercentIncrease, GreatestPercentDecrease, GreatestTotalVolume As Double
'cant use LongLong data type on a 32 bit excel, so using double instead
Dim total_volume As Double
Dim tickerGreatestIncrease, tickerGreatestDecrease, tickerGreatestVolume As String
Dim ws As Worksheet




For Each ws In ThisWorkbook.Worksheets

With ws

last_row = .Cells(Rows.Count, 1).End(xlUp).Row

'setting up counters to their initial state

summary_table_row = 2
initial_value = .Cells(2, 3).Value
total_volume = 0





'setting up summary cells headers
.Cells(1, 9).Value = "Ticker"
.Cells(1, 10).Value = "Yearly Cheange"
.Cells(1, 11).Value = "Percent Change"
.Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To last_row

    If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
        total_volume = total_volume + .Cells(i, 7).Value
        .Cells(summary_table_row, 9).Value = .Cells(i, 1).Value  ' set Ticker
        .Cells(summary_table_row, 10).Value = .Cells(i, 6).Value - initial_value  ' Set Yearly change
        If .Cells(i, 6).Value - initial_value >= 0 Then
            .Cells(summary_table_row, 10).Interior.ColorIndex = 4
        Else
            .Cells(summary_table_row, 10).Interior.ColorIndex = 3
        End If
            If initial_value > 0 Then ' if we have stock listed in this year with some initial price
                .Cells(summary_table_row, 11).Value = (.Cells(i, 6).Value - initial_value) / initial_value  ' set percent change
            Else
                .Cells(summary_table_row, 11).Value = 0  ' if the stock value was zero for the whole year
            End If
            
            .Cells(summary_table_row, 11).NumberFormat = "0.00%"
            .Cells(summary_table_row, 12).Value = total_volume

            
            total_volume = 0
            summary_table_row = summary_table_row + 1
            initial_value = .Cells(i + 1, 3).Value
        Else
            total_volume = total_volume + .Cells(i, 7).Value
            If initial_value = 0 Then initial_value = .Cells(i, 3).Value ' if the stock is listed in this year but not at the beginning of the year
        End If

Next i


' Bonus calculation

GreatestPercentIncrease = 0
GreatestPercentDecrease = 0
GreatestTotalVolume = 0



last_row = .Cells(Rows.Count, 9).End(xlUp).Row

For summary_table_row = 2 To last_row

 ' find which stock has greatest increase, decrease and volume
            If .Cells(summary_table_row, 11).Value > GreatestPercentIncrease Then
                GreatestPercentIncrease = .Cells(summary_table_row, 11).Value
                tickerGreatestIncrease = .Cells(summary_table_row, 9).Value
            End If
            
            If .Cells(summary_table_row, 11).Value < GreatestPercentDecrease Then
                GreatestPercentDecrease = .Cells(summary_table_row, 11).Value
                tickerGreatestDecrease = .Cells(summary_table_row, 9).Value
            End If
            
            
            If .Cells(summary_table_row, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = .Cells(summary_table_row, 12).Value
                tickerGreatestVolume = .Cells(summary_table_row, 9).Value
            End If
            
Next summary_table_row

' set cell headers
.Cells(1, 16).Value = "Ticker"
.Cells(1, 17).Value = "Value"
.Cells(2, 15).Value = "Greatest % Increase"
.Cells(3, 15).Value = "Greatest % Decrease"
.Cells(4, 15).Value = "Greatest Total Volume"


' set cell values for greatest increase, decrease and volume

.Cells(2, 16).Value = tickerGreatestIncrease
.Cells(2, 17).Value = GreatestPercentIncrease
.Cells(2, 17).NumberFormat = "0.00%"
.Cells(3, 16).Value = tickerGreatestDecrease
.Cells(3, 17).Value = GreatestPercentDecrease
.Cells(3, 17).NumberFormat = "0.00%"
.Cells(4, 16).Value = tickerGreatestVolume
.Cells(4, 17).Value = GreatestTotalVolume

End With
Next ws

End Sub
