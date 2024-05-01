Sub runTickersEachWorksheet():
' Notes: I am a developer, but in all honesty, I haven't touch Basic, Visual Basic or VBA for several years and I'm rusty.
' During this challenge, I tried to stay within the commands and functions taught during class as much as possible.
' Saying so, I developed on alphabetical_testing.xlsm as instructed by challenge 2. To accomodate changes,
' I added quarter as part of the board to differentiate same tickets by different quarters.
'
' Sources of information:
' Loop through worksheets
' https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Call aggregateQuarterly(ws)
    Next
End Sub
Public Function getDate(ws As Worksheet, r As Long) As Date
' Sources of information:
' MyCheck = VarType(DateVar)  ' Returns 7.
' MyCheck = VarType(StrVar)   ' Returns 8.
' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function

    Dim potentialDate As Variant
    
    potentialDate = ws.Cells(r, 2).Value
    
    If (VarType(potentialDate) = vbDate) Then
        getDate = potentialDate
    ElseIf (VarType(potentialDate) = vbString) Then
        getDate = DateSerial(Left(potentialDate, 4), Mid(potentialDate, 5, 2), Right(potentialDate, 2))
    ElseIf (potentialDate = "") Then
        getDate = #1/1/1900#
    Else
        MsgBox "Date format not recognized on row " & r & ". Exiting program"
        End
    End If
End Function
Public Function CONVERT_YYYYMMDD_TO_DATE(strText As String) As Date
    ' Code by Dan Wagner (heavily modified)
    ' Source: https://danwagner.co/converting-numbers-like-yyyymmdd-to-dates-with-vba/
    
    ' Then we can slice and dice the number, feeding it all into DateSerial
    CONVERT_YYYYMMDD_TO_DATE = DateSerial(Left(Cells(r, 2).Value, 4), Mid(Cells(r, 2).Value, 5, 2), Right(Cells(r, 2).Value, 2))
End Function
Sub aggregateQuarterly(ws As Worksheet):
' Sources of information:
' Let Quarter = (MonthNumber - 1) \ 3 + 1
' https://stackoverflow.com/questions/26426648/access-vba-get-quarter-from-month-number
'
' Range("A1") = Excel.Application.WorksheetFunction.EoMonth(Range("A1").Value2, 0)
' https://stackoverflow.com/questions/50642569/vba-get-the-last-day-of-the-month
'
' Find last row
' https://www.thespreadsheetguru.com/last-row-column-vba/
'
' Autofit columns
' https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

    ' Declare variables
    Dim row As Long
    Dim rowLast As Long
    Dim rowBoard As Long
    Dim ticker As String
    Dim quarterlyChange As Double
    Dim quarterlychangePerc As Double
    Dim tickerPrev As String
    Dim dt As Date
    Dim dtNext As Date
    Dim mth As Integer
    Dim mthNext As Integer
    Dim quarter As Integer
    Dim quarterNext As Integer
    Dim openAmount As Double
    Dim high As Double
    Dim low As Double
    Dim closeAmount As Double
    Dim vol As Integer
    ' Had to declare totalStock as double as long would overflow
    Dim totalStock As Double
    Dim tickerGreatestInc As String
    Dim tickerGreatestDec As String
    Dim tickerGreatestTotalStock As String
    Dim percIncrease As Double
    Dim percDecrease As Double
    Dim totalStockGreatest As Double
        
    ' Place header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 13).Value = "Quarter"
    
    
    row = 2                             ' Start at row 2
    rowLast = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
                                        ' find the last row on the current worksheet
    rowBoard = 2                        ' initialize row where board will write
    
    ticker = ws.Cells(row, 1).Value     ' capture first ticker
    openAmount = ws.Cells(row, 3).Value ' capture open amount
    
    totalStock = 0                      ' initialize totalStock
    
    percIncrease = 0               ' initialize perIncrease beyond a potential min value it could be
    percDecrease = 0                 ' initialize perDecrease beyond a potential max value it could be
    totalStockGreatest = 0              ' initialize totalStockGreatest
    
    
    ' Loop until the last row
    For row = 2 To rowLast
        ' when same ticker
        ' Initialize date, month and quarter
        dt = getDate(ws, row)
        mth = Month(dt)
        quarter = (mth - 1) \ 3 + 1
        
        dtNext = getDate(ws, row + 1)
        mthNext = Month(dtNext)
        quarterNext = (mthNext - 1) \ 3 + 1
        
        totalStock = totalStock + ws.Cells(row, 7).Value
        
        If Not ((ticker = ws.Cells(row + 1, 1).Value) And (quarter = quarterNext) And (Year(dt) = Year(dtNext))) Then
            closeAmount = ws.Cells(row, 6).Value ' capture closeAmount
            ws.Cells(rowBoard, 9).Value = ticker
            quarterlyChange = closeAmount - openAmount
            ws.Cells(rowBoard, 10).Value = quarterlyChange
            ' highlight cell depending on quarterly change value
            If (quarterlyChange > 0) Then
                ws.Cells(rowBoard, 10).Interior.ColorIndex = 4
            ElseIf (quarterlyChange < 0) Then
                ws.Cells(rowBoard, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(rowBoard, 10).Interior.ColorIndex = 0
            End If
            ' Check if openAmount is 0 to avoid the error when dividing by zero
            If (openAmount = 0) Then
                quarterlychangePerc = 0
            Else
                quarterlychangePerc = (closeAmount - openAmount) / openAmount
            End If
            ws.Cells(rowBoard, 11).Value = quarterlychangePerc
            ws.Cells(rowBoard, 11).NumberFormat = "0.00%"
            ' highlight cell depending on quarterly change percent value
            If (quarterlychangePerc > 0) Then
                ws.Cells(rowBoard, 11).Interior.ColorIndex = 4
            ElseIf (quarterlychangePerc < 0) Then
                ws.Cells(rowBoard, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(rowBoard, 11).Interior.ColorIndex = 0
            End If
            ws.Cells(rowBoard, 12).Value = totalStock
            ws.Cells(rowBoard, 13).Value = CStr(Year(dt)) + "Q" + CStr(quarter)
            rowBoard = rowBoard + 1
            
            ' capture if greatest increase, decrease and total volume
            If (quarterlychangePerc > percIncrease) Then
                tickerGreatestInc = ticker
                percIncrease = quarterlychangePerc
            End If
            If (quarterlychangePerc < percDecrease) Then
                tickerGreatestDec = ticker
                percDecrease = quarterlychangePerc
            End If
            If (totalStock > totalStockGreatest) Then
                tickerGreatestTotalStock = ticker
                totalStockGreatest = totalStock
            End If
            
            ' Initialize for the next group variables
            ticker = ws.Cells(row + 1, 1).Value       ' capture next ticker of the group
            openAmount = ws.Cells(row + 1, 3).Value   ' capture open amount
            totalStock = 0
        End If
    Next row
    
    ' Place labels on greatest section
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    
    ' Place greatest values
    ws.Range("Q2").Value = tickerGreatestInc
    ws.Range("Q3").Value = tickerGreatestDec
    ws.Range("Q4").Value = tickerGreatestTotalStock
    ws.Range("R2").Value = percIncrease
    ws.Range("R3").Value = percDecrease
    ws.Range("R4").Value = totalStockGreatest
    
    ' Autofit the columns
    ws.Columns("I:R").AutoFit
    ws.Range("R2").NumberFormat = "0.00%"
    ws.Range("R3").NumberFormat = "0.00%"
    ws.Range("R4").NumberFormat = "General"
End Sub


