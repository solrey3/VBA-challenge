Sub yearlyTickerSummary()

Dim yrwkb As Worksheet
Set yrwkb = Sheets.Add
yrwkb.Name = "Years"

' Add Active Worksheet As Object
Dim swks As Workbook
Set swks = ThisWorkbook
Dim swkb As Worksheet
'Set swkb = Sheets.Add

' Open Data Workbook
Dim wkb As Workbook
'Set wkb = Workbooks.Open("alphabetical_testing.xlsx")
Set wkb = Workbooks.Open("Multiple_year_stock_data.xlsx")

' Decalre Variables
Dim year As String
Dim lastRow As Double
Dim firstStockPriceHolder As Double
Dim firstStockDateOpenPrice As Double
Dim lastStockDateClosePrice As Double
Dim changeYearly As Double
Dim changePercentage As Double
Dim stockTicker As String
Dim stockTickerCounter As Double
Dim stockVolume As Double

'Initialize certain values

    
Dim Fruits As Collection
Set Fruits = New Collection
Dim Cell As Range
Dim CollectionCount As Integer
Dim f As Integer


For Each ws1 In wkb.Worksheets
    For Each Cell In ws1.UsedRange.Columns("B").Cells
        If Left(Cell.Value, 4) <> "<dat" Then
            CollectionAddUnique Fruits, Left(Cell.Value, 4)
        End If
    Next Cell
Next ws1

CollectionCount = Fruits.Count

For f = 1 To CollectionCount
    yrwkb.Cells(f, 1).Value = Fruits.Item(f)
    year = Fruits.Item(f)
    
    Set swkb = swks.Worksheets.Add

    'Add headers to Active Worksheet
    swkb.Name = year
    swkb.Cells(1, 1).Value = "Ticker"
    swkb.Cells(1, 2).Value = "Yearly Change"
    swkb.Cells(1, 3).Value = "Percentage Change"
    swkb.Cells(1, 4).Value = "Total Stock Volume"
    'swkb.Cells(1, 5).Value = "Open"
    'swkb.Cells(1, 6).Value = "Close"
    
    firstStockDatePriceHolder = 0
    stockTickerCounter = 1
    stockVolume = 0

    'For Each Worksheet
    For Each ws In wkb.Worksheets

        'Find last row of worksheet
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        'Cycle through each entry
        For i = 2 To lastRow
    
            'Check if Adjacent cell values for ticker do not match
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Check if appropriate year
                If Left(ws.Cells(i, 2).Value, 4) = year Then
                    stockTicker = ws.Cells(i, 1).Value
                    stockTickerCounter = stockTickerCounter + 1
                    swkb.Cells(stockTickerCounter, 1) = stockTicker
                    firstStockDateOpenPrice = ws.Cells(firstStockDatePriceHolder, 3).Value
                    lastStockDateClosePrice = ws.Cells(i, 6).Value
                    changeYearly = lastStockDateClosePrice - firstStockDateOpenPrice
                    swkb.Cells(stockTickerCounter, 2).Value = changeYearly
                    'swkb.Cells(stockTickerCounter, 5).Value = firstStockDateOpenPrice
                    'swkb.Cells(stockTickerCounter, 6).Value = lastStockDateClosePrice
                    
                    If changeYearly >= 0 Then
                        swkb.Cells(stockTickerCounter, 2).Interior.ColorIndex = 4
                    Else
                        swkb.Cells(stockTickerCounter, 2).Interior.ColorIndex = 3
                    End If
                
                    ' Calculate Change Percentage, Correct for division by zero
                    If firstStockDateOpenPrice = 0 Then
                        swkb.Cells(stockTickerCounter, 3).Value = 0
                    Else
                        changePercentage = changeYearly / firstStockDateOpenPrice
                        swkb.Cells(stockTickerCounter, 3).Value = changePercentage
                        swkb.Cells(stockTickerCounter, 3).NumberFormat = "0.00%"
                        firstStockPriceHolder = 1
                    End If
                    
                    stockVolume = stockVolume + ws.Cells(i, 7).Value
                    swkb.Cells(stockTickerCounter, 4).Value = stockVolume
                    stockVolume = 0
                    firstStockDatePriceHolder = i + 1
                End If
                
            ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And Left(ws.Cells(i, 2).Value, 4) = year Then
                stockVolume = stockVolume + Cells(i, 7).Value
                If firstStockDatePriceHolder = 0 Then
                    firstStockDatePriceHolder = i
                    firstStockDateOpenPrice = ws.Cells(firstStockDatePriceHolder, 3).Value
                End If
            End If
        Next i
    Next ws


    lastRow = swkb.Cells(Rows.Count, "A").End(xlUp).Row
    swkb.Cells(2, 7).Value = "Greatest % Increase"
    swkb.Cells(3, 7).Value = "Greatest % Decerease"
    swkb.Cells(4, 7).Value = "Greatest Total Volume"
    swkb.Cells(1, 8).Value = "Ticker"
    swkb.Cells(1, 9).Value = "Value"

    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double

    greatestPercentIncreaseTicker = "<ticker>"
    greatestPercentDecreaseTicker = "<ticker>"
    greatestTotalVolumeTicker = "<ticker>"
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0

    For i = 2 To lastRow
        If swkb.Cells(i, 3).Value > greatestPercentIncrease Then
            greatestPercentIncreaseTicker = swkb.Cells(i, 1).Value
            swkb.Cells(2, 8).Value = greatestPercentIncreaseTicker
            greatestPercentIncrease = swkb.Cells(i, 3).Value
            swkb.Cells(2, 9).Value = greatestPercentIncrease
            swkb.Cells(2, 9).NumberFormat = "0.00%"
        End If
        If swkb.Cells(i, 3).Value < greatestPercentDecrease Then
            greatestPercentDecreaseTicker = swkb.Cells(i, 1).Value
            swkb.Cells(3, 8).Value = greatestPercentDecreaseTicker
            greatestPercentDecrease = swkb.Cells(i, 3).Value
            swkb.Cells(3, 9).Value = greatestPercentDecrease
            swkb.Cells(3, 9).NumberFormat = "0.00%"
        End If
        If swkb.Cells(i, 4).Value > greatestTotalVolume Then
            greatestTotalVolumeTicker = swkb.Cells(i, 1).Value
            swkb.Cells(4, 8).Value = greatestTotalVolumeTicker
            greatestTotalVolume = swkb.Cells(i, 4).Value
            swkb.Cells(4, 9).Value = greatestTotalVolume
        End If
    Next i

Next f

'AutoFit Every Worksheet Column in a Workbook
For Each sht In ThisWorkbook.Worksheets
    sht.Cells.EntireColumn.AutoFit
Next sht

' Close Data Workbook
wkb.Close

End Sub

Public Function CollectionAddUnique(ByRef Target As Collection, Value As String) As Boolean

    Dim l As Long

    'SEE IF COLLECTION HAS ANY VALUES
    If Target.Count = 0 Then
        Target.Add Value
        Exit Function
    End If

    'SEE IF VALUE EXISTS IN COLLECTION
    For l = 1 To Target.Count
        If Target(l) = Value Then
            Exit Function
        End If
    Next l

    'NOT IN COLLECTION, ADD VALUE TO COLLECTION
    Target.Add Value
    CollectionAddUnique = True

End Function




