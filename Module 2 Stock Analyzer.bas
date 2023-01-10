Attribute VB_Name = "Module1"
Sub WorksheetLoop()

For Each ws In Worksheets
    ws.Select
    Call StockAnalyzer
Next ws

End Sub
Sub StockAnalyzer()

' Define all variables
Dim ticker As String
Dim totalVol As LongLong
Dim Stock_Summary_Row As Long
Dim yearDelta, Percent As Double

Stock_Summary_Row = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Create the new summary table
Range("J1").Value = "Ticker"
Range("K1").Value = "Year-End Difference"
Range("L1").Value = "Percentage Change"
Range("M1").Value = "Total Stock Volume"

' Change all of the <date> values from text to numbers
With Range("B2:B" & LastRow)
    .NumberFormat = "General"
    .Value = .Value
End With

' Loop through the worksheet...
For i = 2 To LastRow
' Find the year-end values
    If (Cells(i, 2).Value = Application.WorksheetFunction.Min(Range("B2:B" & LastRow))) Then
        yearOpen = Cells(i, 3).Value
    ElseIf (Cells(i, 2).Value = Application.WorksheetFunction.Max(Range("B2:B" & LastRow))) Then
        yearEnd = Cells(i, 6).Value
    End If
' Loop through the database and extract/calculate the necessary data
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
            Range("J" & Stock_Summary_Row).Value = ticker
        totalVol = totalVol + Cells(i, 7).Value
            Range("M" & Stock_Summary_Row).Value = totalVol
        yearDelta = yearEnd - yearOpen
            Range("K" & Stock_Summary_Row).Value = yearDelta
        Percent = yearDelta / yearOpen
            Range("L" & Stock_Summary_Row).Value = FormatPercent(Percent)
        Stock_Summary_Row = Stock_Summary_Row + 1
            totalVol = 0
    Else
        totalVol = (totalVol + (Cells(i, 7).Value))
    End If
' Apply conditional formatting
    If IsEmpty(Cells(i, 11).Value) Then Exit For
    If Cells(i, 11).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        Cells(i, 11).Interior.ColorIndex = 4
    ElseIf Cells(i, 11).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        Cells(i, 11).Interior.ColorIndex = 3
    Else
        Cells(i, 10).Interior.ColorIndex = 0
        Cells(i, 11).Interior.ColorIndex = 0
    End If
Next i

Range("J:M").Columns.AutoFit

End Sub
