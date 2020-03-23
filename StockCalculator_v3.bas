Attribute VB_Name = "Module1"
Sub StockCalculator():

    ' Loop through all the stocks for one year and output the following info:
    ' The ticker symbol
    ' Yearly change from opening price to closing price
    ' % change
    ' Total stock volume
    ' Conditional formatting (green for positive change, red for negative change)
    ' Challenge: Return the stock with greatest % increase, greatest % decrease, and greatest total volume
    
    ' Part 1: ticker symbol
    ' Use CreditCardChecker as reference
    ' Create var to hold the total number of rows
    Dim lastRow As Double
    ' Create var to hold the row of the ticker
    Dim tickerRow As Double
    ' Create var to hold ticker symbol
    Dim ticker As String
    ' Create vars to hold the first open price and last close price
    Dim firstOpen As Double
    Dim lastClose As Double
    ' Create var to hold the yearly change
    Dim yearChange As Double
    ' Create var to hold percent change
    Dim perChange As Double
    ' Create var to hold total stock volume
    Dim volume As Double
     
    ' Assign var initial values
    ' This formula finds the last row somehow
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row - 1
    ' Ticker row begins at 2 for Range("I2")
    tickerRow = 2
    ' Ticker symbol begins at Range("A2")
    ticker = Range("A2").Value
    ' Initial open is Cells(2,3)
    firstOpen = Cells(2, 3).Value
    ' Initial close is Cells(2,6)
    lastClose = Cells(2, 6).Value
    ' Calculate change
    yearChange = lastClose - firstOpen
    ' Calculate and assign percent change
    ' Use loop to check if firstOpen = 0 to avoid dividing by 0
    If firstOpen = 0 Then
        perChange = 0
    Else
        perChange = yearChange / firstOpen
    End If
    ' Assign Cells(2,7) value to volume
    volume = Cells(2, 7).Value
    ' Display on spreadsheet the initial ticker symbol
    Cells(tickerRow, 9).Value = ticker
    ' Display on spreadsheet the initial yearly change
    Cells(tickerRow, 10).Value = yearChange
    ' Display on spreadsheet the inital percent change
    Cells(tickerRow, 11).Value = perChange
    ' Total is initially G2, display on spreadsheet
    Cells(tickerRow, 12).Value = volume
    
    ' This section formats the output section of the spreadsheet
    ' Can use .HorizontalAlignment = xlCenter to center align text for a cell or column(s)
    Columns("I:P").HorizontalAlignment = xlCenter
    ' Fill in headers for columns I (Ticker), J (Yearly Change), K (Percent Change), and L (Total Stock Volume)
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    ' Challenge: Fill in headers for Ticker (O) and Value (P)
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    ' Challenge: Fill in various cells
    Cells(2, 14).Value = "Greatest % increase"
    Cells(3, 14).Value = "Greatest % decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    ' Can use .EntireColumn.AutoFit to change column width to autofit text
    Range("I1:P4").EntireColumn.AutoFit
    
    ' First, loop compares current row to next row to see if ticker symbols are different
    For i = 2 To lastRow
    
        ' When different...
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            ' Update to new ticker symbol. This means updating tickerRow and reassigning firstOpen, lastClose, yearChange,
            ' and perChange. Display new ticker symbol, new yearly change, new percent change, and new total volume.
            ' Increment the tickerRow
            tickerRow = tickerRow + 1
            ' Change ticker symbol to the new one
            ticker = Cells(i + 1, 1).Value
            ' Put the new ticker symbol in the new ticker column
            Cells(tickerRow, 9).Value = ticker
            ' Assign new firstOpen and lastClose values
            firstOpen = Cells(i + 1, 3).Value
            lastClose = Cells(i + 1, 6).Value
            ' Calculate new year change
            yearChange = lastClose - firstOpen
            ' Calculate new percent change
            ' Need to use loop to check if firstOpen = 0 to avoid dividing by 0
            If firstOpen = 0 Then
                perChange = 0
            Else
                perChange = yearChange / firstOpen
            End If
            ' Assign new total volume
            volume = Cells(i + 1, 7).Value
            ' Display yearly change for the new ticker symbol
            Cells(tickerRow, 10).Value = yearChange
            ' Display % change for the new ticker symbol
            Cells(tickerRow, 11).Value = perChange
            ' Display new total stock volume
            Cells(tickerRow, 12).Value = volume
            ' For debugging purposes only
            ' Exit For
            
        ' When not different...
        Else
            ' Update lastClose
            lastClose = Cells(i + 1, 6).Value
            ' Calculate the yearly change
            yearChange = lastClose - firstOpen
            ' Calculate the percent change
            ' Use loop to check if firstOpen = 0 to avoid dividing by 0
            If firstOpen = 0 Then
                perChange = 0
            Else
                perChange = yearChange / firstOpen
            End If
            ' Update total volume
            volume = volume + Cells(i + 1, 7).Value
            ' Display updated yearly change
            Cells(tickerRow, 10).Value = yearChange
            ' Display updated percent change
            Cells(tickerRow, 11).NumberFormat = "0.00%"
            Cells(tickerRow, 11).Value = perChange
            ' Display updated total volume
            Cells(tickerRow, 12).Value = volume
            ' Conditional formatting
            ' Use grader.xlsm as reference
            ' ColorIndex 4 is green, ColorIndex 3 is red
            ' When the yearly change is less than 0, set the interior to red (3)
            If Cells(tickerRow, 10).Value < 0 Then
                Cells(tickerRow, 10).Interior.ColorIndex = 3
            ' When it's greater than 0, set interior to green (4)
            ElseIf Cells(tickerRow, 10).Value > 0 Then
                Cells(tickerRow, 10).Interior.ColorIndex = 4
            End If
        
        End If
    
    Next i
    
    ' Challenge: Create vars for greatest % incr, greatest % decr, greatest total vol, incTicker, decTicker, and volTicker
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestVol As Double
    Dim incTicker As String
    Dim decTicker As String
    Dim volTicker As String
    
    ' Challenge: Set initial values for each variable from the first ticker symbol. Use Cells(2,9) through Cells(2,12)'s values.
    incTicker = Cells(2, 9).Value
    decTicker = Cells(2, 9).Value
    volTicker = Cells(2, 9).Value
    greatestInc = Cells(2, 11).Value
    greatestDec = Cells(2, 11).Value
    greatestVol = Cells(2, 12).Value
    
    ' Challenge: Run a loop to update the values and ticker symbols of the variables
    ' Have an Until loop go down column I until it encounters a blank cell
    Dim x As Long
    x = 2
    
    Do Until IsEmpty(Cells(x, 9).Value) = True
        If Cells(x, 11).Value > greatestInc Then
            greatestInc = Cells(x, 11).Value
            incTicker = Cells(x, 9).Value
        End If
        If Cells(x, 11).Value < greatestDec Then
            greatestDec = Cells(x, 11).Value
            decTicker = Cells(x, 9).Value
        End If
        If Cells(x, 12).Value > greatestVol Then
            greatestVol = Cells(x, 12).Value
            volTicker = Cells(x, 9).Value
        End If
        x = x + 1
    Loop
    
    ' Challenge: After loop is done, display greatestInc, greatestDec, greatestVol
    Range("O2").Value = incTicker
    Range("O3").Value = decTicker
    Range("O4").Value = volTicker
    Range("P2:P3").NumberFormat = "0.00%"
    Range("P2").Value = greatestInc
    Range("P3").Value = greatestDec
    Range("P4").Value = greatestVol
    Columns("P").EntireColumn.AutoFit

End Sub
