Attribute VB_Name = "Module1"
Sub alphabet()

' define worksheet in order to run on each worksheet in the workbook.
For Each ws In Worksheets

' define variables
    Dim i As Long
    Dim TotalVolume As Double
    TotalVolume = 0
    Dim Tick As Long
    Tick = 0
    Dim counter As Double
    counter = 1
    
' add header titles for summary table
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

' define last row for end of for loop range
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' define for loop range
    For i = 2 To lastrow
' count each unique ticker as a ticker
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        Tick = Tick + 1
        End If

' use look ahead to find where ticker value changes
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        counter = counter + 1
        
' define the first and last row of each unique ticker in order to calcuate yearly change
        LastUniqueRow = i
        FirstUniqueRow = LastUniqueRow - Tick

' print ticker in column
        ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value

' print yearly change
        ws.Cells(counter, 10).Value = ws.Cells(LastUniqueRow, 6).Value - ws.Cells(FirstUniqueRow, 3)

' calculate percent change and account for division by 0
        If ws.Cells(FirstUniqueRow, 3).Value <> 0 Then
            ws.Cells(counter, 11).Value = (ws.Cells(LastUniqueRow, 6).Value - ws.Cells(FirstUniqueRow, 3).Value) / ws.Cells(FirstUniqueRow, 3).Value
        Else
            ws.Cells(counter, 11).Value = 0
        End If
' format percent change value to a percent
        ws.Cells(counter, 11).Value = Format(ws.Cells(counter, 11).Value, "percent")

' create nested for loop to add each row of volume for each unique ticker
            For j = FirstUniqueRow To LastUniqueRow
                TotalVolume = TotalVolume + ws.Cells(j, 7).Value
                Next j

' print total volume
        ws.Cells(counter, 12).Value = TotalVolume

' reset tick and totalvolume count
        Tick = 0
        TotalVolume = 0

        End If

' conditional formating to make positive values green and negative red
        If ws.Cells(counter, 10).Value > 0 Then
            ws.Cells(counter, 10).Interior.ColorIndex = 4
            ws.Cells(1, 10).Interior.ColorIndex = 0
        Else: ws.Cells(counter, 10).Interior.ColorIndex = 3

        End If

        Next i
   
   Next ws
   
End Sub

Sub Challenge()

' define variables
    Dim i As Long
    Dim max As Double
    max = 0
    Dim min As Double
    min = 0
    Dim maxvolume As Double
    maxvolume = 0
    
' add header titles for summary table
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
' define worksheet to run through each worksheet in workbook
For Each ws In Worksheets

' define last row for end of for loop range
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' loop to find largest value and make it equal "max"
    For i = 2 To lastrow
        If ws.Cells(i, 11).Value > max Then
        max = ws.Cells(i, 11).Value
        Cells(2, 16).Value = ws.Cells(i, 9) 'print ticker
        End If
    
' find smallest value and make it equal "min"
        If ws.Cells(i, 11).Value < min Then
        min = ws.Cells(i, 11).Value
        Cells(3, 16).Value = ws.Cells(i, 9) 'print ticker
        End If
    
' find largest volume value and make equal "maxvolume"
        If ws.Cells(i, 12).Value > maxvolume Then
        maxvolume = ws.Cells(i, 12).Value
        Cells(4, 16).Value = ws.Cells(i, 9) 'print ticker
        End If
        Next i
    Next ws
    
' print max and min values and format to percent
    Cells(2, 17).Value = max
    Cells(2, 17).Value = Format(Cells(2, 17).Value, "percent")
    Cells(3, 17).Value = min
    Cells(3, 17).Value = Format(Cells(3, 17).Value, "percent")
    Cells(4, 17).Value = maxvolume
    
End Sub
