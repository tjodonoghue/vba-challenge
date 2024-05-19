# vba-challenge

This is my mod 2 challenge assignment. For some reason, my code is not displaying when I save the sheets. For back up, the code I used to create the macro is:

Sub StockAnalysis()

    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        resultRow = 2 ' Start output in row 2

        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 15).Value = "Metric"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                ws.Cells(resultRow, 9).Value = ticker
                ws.Cells(resultRow, 10).Value = quarterlyChange
                ws.Cells(resultRow, 11).Value = percentChange
                ws.Cells(resultRow, 12).Value = totalVolume

                If quarterlyChange > 0 Then
                    ws.Cells(resultRow, 10).Interior.Color = RGB(0, 255, 0) 
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(resultRow, 10).Interior.Color = RGB(255, 0, 0) 
                End If

                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If

                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If

                totalVolume = 0
                resultRow = resultRow + 1
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    openPrice = ws.Cells(i, 3).Value
                End If
            End If
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease

        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume

    Next ws

End Sub

This project would not have been complete without the help of Youtube's finest coding community. Thanks to Youtube for allowing creators to educate us. 
