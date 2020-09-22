Attribute VB_Name = "Module1"
Sub TestScript_Please_Work()
    
    ' Loopping through sheets first
    
'Dim WS_Count As Integer
'Dim ws As Integer

'WS_Count = ActiveWorkbook.Worksheets.Count
    
'For ws = 1 To WS_Count
        
Dim ws As Worksheet

    For Each ws In Worksheets
    ws.Activate
                    
    ' Give me the last row
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Creating Headers for new rows
        
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
        
    ' Create Variables
        
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Ticker_Name As String
    Dim Percent_Change As Double
    Dim Volume As Double
        
    ' Variables to reset counters later on
    
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
'Set initial Open Price
    Open_Price = Cells(2, "C").Value
        
' Loop through all ticker names
        
    For i = 2 To LastRow
        
' Grabbing all the same tickers
         
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
            ' Grab that ticker name and fill the colummn with the value
                
            Ticker_Name = Cells(i, Column).Value
            Cells(Row, "I").Value = Ticker_Name
                
' Set Close Price
            Close_Price = Cells(i, "F").Value
                
' Add Yearly Change, going to sustract the close price from the open price nad fill in the value
                
            Yearly_Change = Close_Price - Open_Price
            Cells(Row, "J").Value = Yearly_Change
                
' Calculating the percentage change from the beginning of the year to the end of that year...

'If Cells(i, "C").Value <> 0 Then
'            Percent_Change = Yearly_Change / Open_Price
'            Cells("K").Value = Percent_Change
'            Cells("K").NumberFormat = "0.00%"
'End If

' Code above dind't work as inteded. This way I should be able to get and accurate percentage from opening to closing.

            If (Open_Price = 0 And Close_Price = 0) Then
                Percent_Change = 0
            ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                Percent_Change = 1
            Else
                Percent_Change = Yearly_Change / Open_Price
                Cells(Row, "K").Value = Percent_Change
                Cells(Row, "K").NumberFormat = "0.00%"
            End If
' Setting variable for total volume
            Volume = Volume + Cells(i, "G").Value

' Filling the volume collumn
            Cells(Row, "L").Value = Volume
' Keep growing the summary table rows and reset the opening price
            Row = Row + 1
            Open_Price = Cells(i + 1, Column + 2)


' Reset the volumen to 0 so I don't keep adding to the same volumnbe volume before moving on the next ticker.
            Volume = 0

' when the cells are the same ticker
        Else
            Volume = Volume + Cells(i, "G").Value
        End If
    
    Next i
        
        
' Lopp to set the last row of Column "Yearly Change" per worksheet and add the coloring.

YearChangeLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        
    For j = 2 To YearChangeLastRow
        If (ws.Cells(j, "J").Value > 0 Or ws.Cells(j, "J").Value = 0) Then
            ws.Cells(j, "J").Interior.ColorIndex = 10
        
        ElseIf ws.Cells(j, "J").Value < 0 Then
            ws.Cells(j, "J").Interior.ColorIndex = 3
        End If
        
    Next j
        
        
        
' Creating mini table for greatest increase and decrease percentage.
        
    Cells(1, "P").Value = "TICKER"
    Cells(1, "Q").Value = "VALUE"
    Cells(3, "O").Value = "Greatest % Increase"
    Cells(4, "O").Value = "Greatest % Decrease"
    Cells(5, "O").Value = "Greatest Total Volume"

' Loop to check each row to find the greatest value for the corresponding ticker. Using last Row of J because that where my tickers end and it's defined already.

        For x = 2 To YearChangeLastRow

            If ws.Cells(x, "K").Value = WorksheetFunction.Max(Range("K2:K" & YearChangeLastRow)) Then
                ws.Cells(3, "P").Value = Cells(x, "I").Value
                ws.Cells(3, "Q").Value = Cells(x, "J").Value
                ws.Cells(3, "J").NumberFormat = "0.00%"
        
            ElseIf Cells(x, "K").Value = WorksheetFunction.Min(Range("K2:K" & YearChangeLastRow)) Then
                ws.Cells(4, "P").Value = Cells(x, "I").Value
                ws.Cells(4, "Q").Value = Cells(x, "J").Value
                ws.Cells(4, "J").NumberFormat = "0.00%"
        
            ElseIf Cells(x, "L").Value = WorksheetFunction.Max(Range("L2:L" & YearChangeLastRow)) Then
                ws.Cells(5, "P").Value = Cells(x, "I").Value
                ws.Cells(5, "Q").Value = Cells(x, "L").Value
            End If
        
        Next x
        
    Next ws
        
End Sub





