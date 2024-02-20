Attribute VB_Name = "Module1"
Sub stock_data()

    ' Dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowTotal As Long
    Dim percent As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim average As Double
    Dim ws As Worksheet
    Dim increase As Long
    Dim decrease As Long
    Dim volume As Long

    For Each ws In Worksheets

        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Annual Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        rowTotal = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowTotal

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                total = total + ws.Cells(i, 7).Value

                If total = 0 Then
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                Else
                    If ws.Cells(start, 3).Value = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    change = (ws.Cells(i, 6).Value - ws.Cells(start, 3).Value)
                    percent = change / ws.Cells(start, 3).Value

                    start = i + 1

                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = percent
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = total

                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If

                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0

            Else
                total = total + ws.Cells(i, 7).Value
            End If

        Next i

        ws.Range("Q2") = "%" & WorksheetFunction.max(ws.Range("K2:K" & rowTotal)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowTotal)) * 100
        ws.Range("Q4") = WorksheetFunction.max(ws.Range("L2:L" & rowTotal))

        increase = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("K2:K" & rowTotal)), ws.Range("K2:K" & rowTotal), 0)
        decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowTotal)), ws.Range("K2:K" & rowTotal), 0)
        volume = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("L2:L" & rowTotal)), ws.Range("L2:L" & rowTotal), 0)

        ws.Range("P2") = ws.Cells(increase + 1, 9)
        ws.Range("P3") = ws.Cells(decrease + 1, 9)
        ws.Range("P4") = ws.Cells(volume + 1, 9)

    Next ws

End Sub

