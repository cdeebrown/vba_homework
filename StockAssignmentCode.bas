Attribute VB_Name = "Module2"
Sub WorksheetLoop()
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In Worksheets
    ws.Activate
        
        Dim Total As Double
            Total = 0
        Dim Ticker As String
        Dim PrintRow As Integer
            PrintRow = 2
        Dim Open_s As Double
        Dim Close_s As Double
        Dim counter As Integer
            counter = 0
        Dim Percent As Double
        Dim increase_num As Double
        Dim rowCount As Long
            rowCount = Cells(Rows.Count, 1).End(xlUp).Row
        Dim start As Double
            start = 2
        Dim tickerP As String
        Dim yearlyP As String
        Dim percP As String
        Dim totalV As String
        Dim MaxP As String
        Dim MinP As String
        Dim MaxT As String
        Dim Value As String
            MaxP = "Greatest % Increase"
            MinP = "Greatest % decrease"
            MaxT = "Greatest Total Volume"
            tickerP = "Ticker"
            yearlyP = "Yearly Change"
            percP = "Percent Change"
            totalV = "Total Volume"
            Value = "Value"
                Range("J1, P1").Value = tickerP
                Cells(1, 11).Value = yearlyP
                Cells(1, 12).Value = percP
                Cells(1, 13).Value = totalV
                Cells(2, 15).Value = MaxP
                Cells(3, 15).Value = MinP
                Cells(4, 15).Value = MaxT
                Range("Q1").Value = Value
                        
                                For i = 2 To rowCount
                                        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                        Total = Total + Cells(i, 7).Value
                                        Ticker = Cells(i, 1).Value
                                        Range("J" & PrintRow).Value = Ticker
                                        Range("M" & PrintRow).Value = Total
                                        counter = counter + 1
                                        Close_s = Cells(i, 6).Value
                                        Open_s = Cells(start, 3).Value
                                                    If Cells(i, 6).Value And Cells(start, 3).Value > 0 Then
                                                            Range("K" & PrintRow).Value = Close_s - Open_s
                                                            Range("L" & PrintRow).Value = (Range("K" & PrintRow).Value / Cells(start, 3))
                                                            Range("L2:L" & rowCount).NumberFormat = "0.00%"
                                                    End If
                                        PrintRow = PrintRow + 1
                                        Total = 0
                                        counter = 0
                                        start = i + 1
                                        Range("L2:L" & rowCount).NumberFormat = "0.00%"
                                        Range("Q2,Q3").NumberFormat = "0.00%"
                                        Range("A1:Q1").HorizontalAlignment = xlCenter
                                        Else
                                        Total = Total + Cells(i, 7).Value
                                        End If
                                Next i
Range("Q2").Value = WorksheetFunction.Max(Range("L2:L" & rowCount))
increase_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)
Range("P2") = Cells(increase_num + 1, 10)
Range("Q3").Value = WorksheetFunction.Min(Range("L2:L" & rowCount))
increase_num = WorksheetFunction.Match(WorksheetFunction.Min(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)
Range("P3") = Cells(increase_num + 1, 10)
Range("Q4").Value = WorksheetFunction.Max(Range("M2:M" & rowCount))
increase_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("M2:M" & rowCount)), Range("M2:M" & rowCount), 0)
Range("P4") = Cells(increase_num + 1, 10)
rowCount = Cells(Rows.Count, "K").End(xlUp).Row
    For i = 2 To rowCount
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        ElseIf Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.ColorIndex = 3
        Else
            Cells(i, 11).Interior.ColorIndex = 0
        End If
    Next
Next
starting_ws.Activate

End Sub
