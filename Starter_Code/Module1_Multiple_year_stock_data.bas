Attribute VB_Name = "Module1"
Sub ProcessStockDataWithTable()
    Dim ws As Worksheet
    Dim tickerData As Object
    Dim ticker As String
    Dim quarterKey As String
    Dim row As Long
    Dim outputRow As Long
    Dim lastRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim firstDate As Variant
    Dim key As Variant
    Dim tickerArray As Variant
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Create dictionary for storing ticker data
        Set tickerData = CreateObject("Scripting.Dictionary")

        ' Determine the last row of the sheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

        ' Add headers for the output columns
        With ws
            .Cells(1, 8).Value = "Ticker"
            .Cells(1, 9).Value = "Quarter"
            .Cells(1, 10).Value = "Quarterly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
        End With

        ' Initialize variables for greatest values
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        increaseTicker = ""
        decreaseTicker = ""
        volumeTicker = ""

        ' Loop through the rows
        For row = 2 To lastRow
            ticker = ws.Cells(row, 1).Value
            firstDate = ws.Cells(row, 2).Value

            ' Validate the date
            If IsDate(firstDate) Then
                quarterKey = Year(firstDate) & " Q" & Application.WorksheetFunction.RoundUp(Month(firstDate) / 3, 0)

                If Not tickerData.exists(ticker & "|" & quarterKey) Then
                    tickerData(ticker & "|" & quarterKey) = Array(ws.Cells(row, 3).Value, 0, 0, firstDate)
                End If

                ' Update the close price and volume
                tickerArray = tickerData(ticker & "|" & quarterKey)

                ' Validate numeric values
                If IsNumeric(ws.Cells(row, 6).Value) And IsNumeric(ws.Cells(row, 7).Value) Then
                    tickerArray(1) = ws.Cells(row, 6).Value ' Close price
                    tickerArray(2) = tickerArray(2) + ws.Cells(row, 7).Value ' Volume
                    tickerData(ticker & "|" & quarterKey) = tickerArray
                End If
            End If
        Next row

        ' Write the summary data in the same worksheet
        outputRow = 2
        For Each key In tickerData.keys
            tickerArray = tickerData(key)

            ' Extract data
            ticker = Split(key, "|")(0)
            quarterKey = Split(key, "|")(1)
            openPrice = tickerArray(0)
            closePrice = tickerArray(1)
            totalVolume = tickerArray(2)

            ' Calculate changes
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If

            ' Write results to the current worksheet
            With ws
                .Cells(outputRow, 8).Value = ticker
                .Cells(outputRow, 9).Value = quarterKey
                .Cells(outputRow, 10).Value = Format(quarterlyChange, "0.00")
                .Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
                .Cells(outputRow, 12).Value = totalVolume
            End With

            ' Track greatest values
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                increaseTicker = ticker
            End If
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                decreaseTicker = ticker
            End If
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                volumeTicker = ticker
            End If

            outputRow = outputRow + 1
        Next key

        ' Add a summary table for greatest values
        With ws
            .Cells(2, 15).Value = "Summary Table"
            .Cells(3, 15).Value = "Metric"
            .Cells(3, 16).Value = "Ticker"
            .Cells(3, 17).Value = "Value"
            
            .Cells(4, 15).Value = "Greatest % Increase"
            .Cells(4, 16).Value = increaseTicker
            .Cells(4, 17).Value = Format(greatestIncrease, "0.00") & "%"
            
            .Cells(5, 15).Value = "Greatest % Decrease"
            .Cells(5, 16).Value = decreaseTicker
            .Cells(5, 17).Value = Format(greatestDecrease, "0.00") & "%"
            
            .Cells(6, 15).Value = "Greatest Total Volume"
            .Cells(6, 16).Value = volumeTicker
            .Cells(6, 17).Value = greatestVolume
        End With

        ' Apply conditional formatting to the Quarterly Change column
        Dim lastDataRow As Long
        lastDataRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).row ' Find the last row of data
        With ws.Range("J2:J" & lastDataRow) ' Column J contains Quarterly Change
            .FormatConditions.Delete ' Remove any existing conditional formatting

            ' Add green for positive changes
            With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Interior.Color = RGB(144, 238, 144) ' Light green
                .Font.Color = RGB(0, 100, 0) ' Dark green text
            End With

            ' Add red for negative changes
            With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Interior.Color = RGB(255, 182, 193) ' Light red
                .Font.Color = RGB(139, 0, 0) ' Dark red text
            End With
        End With
    Next ws

    MsgBox "Stock data processed, and summary tables added to each sheet."
End Sub

