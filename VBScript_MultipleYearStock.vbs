Attribute VB_Name = "Module1"
Sub Stonks()

'Loop through all the worksheets
For Each ws In Worksheets

'Finds the number of rows in the sheet
Dim last_row As Long
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Dimension Variables
Dim TotalVol As Double
Dim t As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double

'Initiate Variables
OpenPrice = ws.Cells(2, 3).Value
t = 2

'Insert Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"


For i = 2 To last_row
    'Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        TotalVol = TotalVol + ws.Cells(i, 7).Value
    Else
        'Adds final row of each ticker to TotalVol
        TotalVol = TotalVol + ws.Cells(i, 7).Value
        'Gets value for ClosePrice and YearlyChange
        ClosePrice = ws.Cells(i, 6).Value
        YearlyChange = ClosePrice - OpenPrice
        
        'Insert Values into cells and format cells
        ws.Cells(t, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(t, 10).Value = YearlyChange
        
        If OpenPrice = 0 Then
            ws.Cells(t, 11).Value = 0
        Else
            ws.Cells(t, 11).Value = Format(YearlyChange / OpenPrice, "0.00%")
        End If
        
        ws.Cells(t, 12).Value = TotalVol
        
        'Conditional Formatting
        If ws.Cells(t, 10).Value > 0 Then
            ws.Cells(t, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(t, 10).Interior.ColorIndex = 3
        End If
        
        'Preparing for next i
        OpenPrice = ws.Cells(i + 1, 3).Value
        TotalVol = 0
        ClosePrice = 0
        t = t + 1
    End If
Next i

'Create second table with greatest and least values
ws.Cells(2, 16).Value = Format(WorksheetFunction.Max(ws.Range("K1:" & "K" & t)), "0.00%")
ws.Cells(3, 16).Value = Format(WorksheetFunction.Min(ws.Range("K1:" & "K" & t)), "0.00%")
ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L1:" & "L" & t))
Dim Index1 As Double
Index1 = WorksheetFunction.Match(ws.Cells(2, 16).Value, ws.Range("K1:" & "K" & t), 0)
ws.Cells(2, 15).Value = ws.Cells(Index1, 9).Value
Dim Index2 As Double
Index2 = WorksheetFunction.Match(ws.Cells(3, 16).Value, ws.Range("K1:" & "K" & t), 0)
ws.Cells(3, 15).Value = ws.Cells(Index2, 9).Value
Dim Index3 As Double
Index3 = WorksheetFunction.Match(ws.Cells(4, 16).Value, ws.Range("L1:" & "L" & t), 0)
ws.Cells(4, 15).Value = ws.Cells(Index3, 9).Value

'Make table look nicer
ws.Columns("I:P").AutoFit

last_row = 0
TotalVol = 0
t = 0
OpenPrice = 0
ClosePrice = 0
YearlyChange = 0
Index1 = 0
Index2 = 0
Index3 = 0

Next ws

End Sub

