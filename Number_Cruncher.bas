Attribute VB_Name = "Module1"
Sub Calculator():

    ' Loop through all the stocks for one year and output the following info:
    ' The ticker symbol
    ' Yearly change from opening price to closing price
    ' % change
    ' Total stock volume
    ' Conditional formatting (green for positive change, red for negative change)
    ' Challenge: Return the stock with greatest % increase, greatest % decrease, and greatest total volume
    
    ' Part 1: ticker symbol
    ' Use CreditCardChecker as reference
    ' Create var to hold the number of rows
    Dim lastRow As Double
    ' Create var to hold the row of the ticker
    Dim tickerRow As Long
    ' Create var to hold ticker symbol
    Dim ticker As String
    ' Create var to hold total stock volume
    Dim totalVolume As Double
    ' Create var to hold the current opening prices per ticker symbol
    Dim currentOpen As Double
    ' Create var to hold the current ending prices per ticker symbol
    Dim currentClose As Double
    ' Create var to hold the change for one entry
    Dim change As Double
    ' Create var to hold the net change per ticker symbol
    Dim netChange As Double
    ' Create vars to hold the first open price and last close price
    Dim firstOpen As Double
    Dim lastClose As Double
    ' Challenge: Create vars for greatest % incr, greatest % decr, greatest total vol, incTicker, decTicker, and volTicker
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestVol As Double
    Dim incTicker As String
    Dim decTicker As String
    Dim volTicker As String
    
    ' Assign var initial values
    ' This formula finds the last row somehow
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row - 1
    ' Ticker row begins at 2 for Range("I2")
    tickerRow = 2
    ' Ticker symbol begins at Range("A2")
    ticker = Range("A2").Value
    ' Initial open is Cells(2,3)
    firstOpen = Cells(2, 3).Value
    ' Assign dummy value to lastClose, won't be known until later
    lastClose = 0
    ' Assign values for current open, current close
    currentOpen = Cells(2, 3).Value
    currentClose = Cells(2, 6).Value
    ' Calculate change
    change = currentClose - currentOpen
    ' Assign change to netChange
    netChange = change
    ' Total is initially 0
    totalVolume = 0
    ' Challenge: Set greatestInc and greatestDec to (currentClose - currentOpen)/currentOpen for now
    greatestInc = (currentClose - currentOpen) / currentOpen
    greatestDec = (currentClose - currentOpen) / currentOpen
    ' Challenge: Set greatest total vol to G2 for now
    greatestVol = Range("G2").Value
    ' Challenge: Set ticker values to initial ticker for now
    incTicker = ticker
    decTicker = ticker
    volTicker = ticker
    
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
    
    ' Put initial ticker symbol into Cells(2,9)
    Cells(tickerRow, 9).Value = Range("A2").Value
    ' Set initial yearly change, total stock volume values
    Cells(tickerRow, 10).Value = netChange
    Cells(tickerRow, 12).Value = totalVolume
    
    ' First, loop compares current row to next row to see if ticker symbols are different
    For i = 2 To lastRow
    
        ' When different...
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ' First calculate and display final yearly change, percent change, and total stock volume
            ' Read current open and close values
            currentOpen = Cells(i, 3).Value
            currentClose = Cells(i, 6).Value
            ' Calculate change
            change = currentClose - currentOpen
            ' Challenge: If change / currentOpen > greatestInc, update greatestInc
            If (change / currentOpen) > greatestInc Then
                incTicker = ticker
                greatestInc = change / currentOpen
            End If
            ' Challenge: If change / currentOpen < greatestDec, update greatestDec
            If (change / currentOpen) < greatestDec Then
                decTicker = ticker
                greatestDec = change / currentOpen
            End If
            ' Challenge: Check if volume > greatestVol, if so, update greatestVol
            If Cells(i, 7).Value > greatestVol Then
                volTicker = ticker
                greatestVol = Cells(i, 7).Value
            End If
            ' Update net change
            netChange = netChange + change
            ' Read volume and add to total volume
            totalVolume = totalVolume + Cells(i, 7).Value
            ' Assign value of the final close
            lastClose = Cells(i, 6).Value
            ' Output yearly change
            Cells(tickerRow, 10).Value = netChange
            ' Output total stock volume
            Cells(tickerRow, 12).Value = totalVolume
            ' Can use .NumberFormat = "0.00%" to format a cell to percentage
            Cells(tickerRow, 11).NumberFormat = "0.00%"
            ' Output the percent change, (lastClose - firstOpen) / firstOpen
            Cells(tickerRow, 11).Value = (lastClose - firstOpen) / firstOpen
            ' For conditional formatting, use grader.xlsm as reference
            ' ColorIndex 4 is green, ColorIndex 3 is red
            ' When the yearly change is less than 0, set the interior to red
            If Cells(tickerRow, 10).Value < 0 Then
                Cells(tickerRow, 10).Interior.ColorIndex = 3
            ' When it's greater than 0, set interior to green
            ElseIf Cells(tickerRow, 10).Value > 0 Then
                Cells(tickerRow, 10).Interior.ColorIndex = 4
            End If
            ' For debugging purposes only
            'Cells(tickerRow, 13).Value = firstOpen
            'Cells(tickerRow, 14).Value = lastClose
            
            ' Then update to new ticker symbol
            ' Increment the tickerRow
            tickerRow = tickerRow + 1
            ' Change ticker symbol to the new one
            ticker = Cells(i + 1, 1).Value
            ' Put the new ticker symbol in the new ticker column
            Cells(tickerRow, 9).Value = ticker
            ' Reset total stock volume
            totalVolume = 0
            ' Assign new currentOpen, currentClose
            currentOpen = Cells(tickerRow, 3).Value
            currentClose = Cells(tickerRow, 6).Value
            ' Calculate new change
            change = currentClose - currentOpen
            ' Assign it to netChange
            netChange = change
            ' Assign new firstOpen value
            firstOpen = Cells(tickerRow, 3).Value
            ' For debugging purposes only
            ' Exit For
            
        ' When not different...
        Else
            ' Read open and close values and add them to current open/close
            currentOpen = Cells(i, 3).Value
            currentClose = Cells(i, 6).Value
            ' Calculate the change
            change = currentClose - currentOpen
            ' Challenge: If change / currentOpen > greatestInc, update greatestInc
            If (change / currentOpen) > greatestInc Then
                incTicker = ticker
                greatestInc = change / currentOpen
            End If
            ' Challenge: If change / currentOpen < greatestDec, update greatestDec
            If (change / currentOpen) < greatestDec Then
                decTicker = ticker
                greatestDec = change / currentOpen
            End If
            ' Challenge: Check if volume > greatestVol, if so, update greatestVol
            If Cells(i, 7).Value > greatestVol Then
                volTicker = ticker
                greatestVol = Cells(i, 7).Value
            End If
            ' Calculate and update the net change
            netChange = netChange + change
            ' Read volume and add to total volume
            totalVolume = totalVolume + Cells(i, 7).Value
            ' Update yearly change
            Cells(tickerRow, 10).Value = netChange
            ' Update total stock volume
            Cells(tickerRow, 12).Value = totalVolume
            ' Update percent change
            Cells(tickerRow, 11).NumberFormat = "0.00%"
            Cells(tickerRow, 11).Value = (lastClose - firstOpen) / firstOpen
            ' When the yearly change is less than 0, set the interior to red
            If Cells(tickerRow, 10).Value < 0 Then
                Cells(tickerRow, 10).Interior.ColorIndex = 3
            ' When it's greater than 0, set interior to green
            ElseIf Cells(tickerRow, 10).Value > 0 Then
                Cells(tickerRow, 10).Interior.ColorIndex = 4
            End If
        
        End If
    
    Next i
    
    ' Weird bug where last column isn't filled in, so...
    
    
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
