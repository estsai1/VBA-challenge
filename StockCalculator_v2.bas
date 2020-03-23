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
    ' Create var to hold the current opening price for a row
    Dim curOpen As Double
    ' Create var to hold the current ending price for a row
    Dim curClose As Double
    ' Create var to hold the change for one row
    Dim change As Double
    ' Create var to hold percent change
    Dim perChange As Double
    ' Create vars to hold the first open price and last close price
    Dim firstOpen As Double
    Dim lastClose As Double
    ' Create var to hold volume for one row
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
    ' Assign dummy value to lastClose, it will change later
    lastClose = Cells(2, 6).Value
    ' Assign starting values for current open, current close
    curOpen = Cells(2, 3).Value
    curClose = Cells(2, 6).Value
    ' Calculate change
    change = curClose - curOpen
    ' Calculate and assign percent change
    perChange = change / curOpen

    ' Display on spreadsheet the initial ticker symbol
    Cells(tickerRow, 9).Value = ticker
    ' Display on spreadsheet the initial yearly change
    Cells(tickerRow, 10).Value = change
    ' Display on spreadsheet the inital percent change
    Cells(tickerRow, 11).Value = perChange
    ' Total is initially G2, display on spreadsheet
    Cells(tickerRow, 12).Value = Range("G2").Value
    
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
            
            ' Update to new ticker symbol. This means updating tickerRow and reassigning firstOpen, lastClose, curOpen, curClose,
            ' change, and perChange. Display new ticker symbol, new yearly change, new percent change, and new total volume.
            ' Increment the tickerRow
            tickerRow = tickerRow + 1
            ' Change ticker symbol to the new one
            ticker = Cells(i + 1, 1).Value
            ' Put the new ticker symbol in the new ticker column
            Cells(tickerRow, 9).Value = ticker
            ' Assign new curOpen, curClose
            curOpen = Cells(i + 1, 3).Value
            curClose = Cells(i + 1, 6).Value
            ' Calculate new change
            change = curClose - curOpen
            ' Calculate new percent change
            perChange = change / curOpen
            ' Display yearly change for the new ticker symbol
            Cells(tickerRow, 10).Value = change
            ' Display % change for the new ticker symbol
            Cells(tickerRow, 11).Value = perChange
            ' Assign new total stock volume
            Cells(tickerRow, 12).Value = Cells(i + 1, 7).Value
            ' Assign new firstOpen and lastClose values
            firstOpen = curOpen
            lastClose = curClose
            ' For debugging purposes only
            ' Exit For
            
        ' When not different...
        Else
            ' Add values from next row to running totals for current ticker symbol
            ' Assign next row to curOpen and curClose
            curOpen = Cells(i + 1, 3).Value
            curClose = Cells(i + 1, 6).Value
            ' Update lastClose to curClose
            lastClose = curClose
            ' Calculate the change
            change = curClose - curOpen
            ' Calculate the percent change
            perChange = change / curOpen
                       
            ' Calculate and update the yearly change
            Cells(tickerRow, 10).Value = lastClose - firstOpen
            ' Update percent change
            Cells(tickerRow, 11).NumberFormat = "0.00%"
            Cells(tickerRow, 11).Value = (lastClose - firstOpen) / firstOpen
            ' Read volume and add it to total volume
            Cells(tickerRow, 12).Value = Cells(tickerRow, 12).Value + Cells(i + 1, 7).Value

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
