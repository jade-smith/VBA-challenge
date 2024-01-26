Sub year_stock_analysis()

    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Long
    Dim start As Long
    Dim rows As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim sheet As Worksheet
    Dim rowCount As Long
    Dim find_value As Long
    Dim increase_number As Long
    Dim decrease_number As Long
    Dim volume_number As Long

    For Each sheet In Worksheets
    
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
        rowCount = sheet.Cells(sheet.rows.Count, "A").End(xlUp).Row
        
        sheet.Range("I1").Value = "Ticker"
        sheet.Range("J1").Value = "Yearly Change"
        sheet.Range("K1").Value = "Percent Change"
        sheet.Range("L1").Value = "Total Stock Volume"
        sheet.Range("P1").Value = "Ticker"
        sheet.Range("Q1").Value = "Value"
        sheet.Range("O2").Value = "Greatest % Increase"
        sheet.Range("O3").Value = "Greatest % Decrease"
        sheet.Range("O4").Value = "Greatest Total Volume"
        
        For i = 2 To rowCount
        
            If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
                total = total + sheet.Cells(i, 7).Value
                
                If total = 0 Then
                    sheet.Range("I" & 2 + j).Value = sheet.Cells(i, 1).Value
                    sheet.Range("J" & 2 + j).Value = 0
                    sheet.Range("K" & 2 + j).Value = "%" & 0
                    sheet.Range("L" & 2 + j).Value = 0
                Else
                    If sheet.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If sheet.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    change = (sheet.Cells(i, 6) - sheet.Cells(start, 3))
                    percentChange = change / sheet.Cells(start, 3)
                    
                    start = i + 1
                    
                    sheet.Range("I" & 2 + j).Value = sheet.Cells(i, 1).Value
                    sheet.Range("J" & 2 + j).Value = change
                    sheet.Range("J" & 2 + j).NumberFormat = "0.00"
                    sheet.Range("K" & 2 + j).Value = percentChange
                    sheet.Range("K" & 2 + j).NumberFormat = "0.00%"
                    sheet.Range("L" & 2 + j).Value = total
                    
                    Select Case change
                        Case Is > 0
                            sheet.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            sheet.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            sheet.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If
                
                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0
                
            Else
                total = total + sheet.Cells(i, 7).Value
            End If
    
        Next i
    
        sheet.Range("Q2").Value = "%" & WorksheetFunction.Max(sheet.Range("K2:K" & rowCount)) * 100
        sheet.Range("Q3").Value = "%" & WorksheetFunction.Min(sheet.Range("K2:K" & rowCount)) * 100
        sheet.Range("Q4").Value = WorksheetFunction.Max(sheet.Range("L2:L" & rowCount))

        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(sheet.Range("K2:K" & rowCount)), sheet.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(sheet.Range("K2:K" & rowCount)), sheet.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(sheet.Range("L2:L" & rowCount)), sheet.Range("L2:L" & rowCount), 0)
       
        sheet.Range("P2").Value = sheet.Cells(increase_number + 1, 9).Value
        sheet.Range("P3").Value = sheet.Cells(decrease_number + 1, 9).Value
        sheet.Range("P4").Value = sheet.Cells(volume_number + 1, 9).Value

    Next sheet

End Sub