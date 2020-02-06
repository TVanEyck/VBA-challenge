Attribute VB_Name = "Module1"
Sub MultiYrStckData()

' find last row with a TickerSymbol in it in column 1
Dim MaxRow As Double
  '    Find the last non-blank cell in column 1
    MaxRow = Cells(Rows.Count, 1).End(xlUp).row

' create and define variables as needed...

' create and define row-process variables
Dim col As Integer
col = 1
Dim row As Double
Dim rowTot As Double
rowTot = 0

' create and define Report variables
Dim ticker As Integer
ticker = 2

' create and define Summary variables
Dim GreatestPrctIncrTicker As String
Dim GreatestPrctIncr As Double
Dim GreatestPrctIncrVol As Double
Dim GreatestPrctDecrTicker As String
Dim GreatestPrctDecr As Double
Dim GreatestPrctDecrVol As Double

' Prime the Headings

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "YrlyChg"
Cells(1, 11).Value = "PrcntChg"
Cells(1, 12).Value = "TotStockVol"
Cells(1, 14).Value = "Greatest"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Volume"
Cells(2, 14).Value = "% Increase"
Cells(3, 14).Value = "% Decrease"
Cells(4, 14).Value = "Total Volume"

' prime the pump...
Dim openDol As Double
openDol = Cells(2, 3).Value

  ' Loop through rows in the column
  For row = 2 To MaxRow
    rowTot = rowTot + Cells(row, 7).Value

'  Searches for when the value of the next cell is different than that of the current cell
    If Cells(row + 1, col).Value <> Cells(row, col).Value Then
        
        Cells(ticker, 9).Value = Cells(row, 1).Value
        Cells(ticker, 10).Value = Cells(row, 6).Value - openDol
' ============================   for testing only ================
'     If Cells(row, col).Value = "PLNT" Then
'        MsgBox ("Ticker: " & Cells(row, col).Value & " chg to " & Cells(row + 1, col).Value)
'        MsgBox ("Ticker " & Cells(ticker, 9).Value & " rowTot " & rowTot)
'        MsgBox ("DiffPrct " & Cells(ticker, 11).Value & " openDol " & openDol & " DiffVal>> " & Cells(ticker, 10).Value)
'     End If
' ============================
        If Cells(ticker, 10).Value < 0 Then
             ' set to red
             Cells(ticker, 10).Interior.ColorIndex = 3
        ElseIf Cells(ticker, 10).Value = 0 Then
             ' set to no fill
             Cells(ticker, 10).Interior.ColorIndex = 2
        ElseIf Cells(ticker, 10).Value > 0 Then
             ' set to green
             Cells(ticker, 10).Interior.ColorIndex = 4
        End If
        If openDol <> 0 Then
            Cells(ticker, 11).Value = Cells(ticker, 10).Value / openDol
            Cells(ticker, 11).Value = Format(Cells(ticker, 11).Value, "Percent")
        Else
            Cells(ticker, 11).Value = 0
            Cells(ticker, 11).Value = Format(Cells(ticker, 11).Value, "Percent")
        End If
        Cells(ticker, 12).Value = rowTot
'        MsgBox ("Ready for Greatest")
'  Test for Greatest...
        If Cells(ticker, 11).Value > 0 And Cells(ticker, 11).Value > GreatestPrctIncr Then
            GreatestPrctIncrTicker = Cells(ticker, 9).Value
            GreatestPrctIncrVol = Cells(ticker, 12).Value
            GreatestPrctIncr = Cells(ticker, 11).Value
        ElseIf Cells(ticker, 11).Value < 0 And Cells(ticker, 11).Value < GreatestPrctDecr Then
            GreatestPrctDecrTicker = Cells(ticker, 9).Value
            GreatestPrctDecrVol = Cells(ticker, 12).Value
            GreatestPrctDecr = Cells(ticker, 11).Value
        End If
        If Cells(ticker, 12).Value > GreatestTotVol Then
            GreatestTotTicker = Cells(ticker, 9).Value
            GreatestTotVol = Cells(ticker, 12).Value
        End If
'        MsgBox ("GreatestIncr " & GreatestPrctIncrTicker & "+" & GreatestPrctIncrVol & "+" & GreatestPrctIncr)
'        MsgBox ("GreatestDecr " & GreatestPrctDecrTicker & "+" & GreatestPrctDecrVol & "+" & GreatestPrctDecr)
'        MsgBox ("GreatestTot " & GreatestTotTicker & "+" & GreatestTotVol)
        
'  Advance for next ticker
        ticker = ticker + 1
        rowTot = 0
        
'  What if cells(row+1) > MaxRow  ?????????????
        openDol = Cells(row + 1, 3).Value

    End If
'  ==============This Code for Testing Purposes =======
'        If ticker > 5 Then
'            Exit Sub
'
'        End If
'  ==============This Code for Testing Purposes =======

  Next row
        
'  Output Greatest Values
  Cells(2, 15).Value = GreatestPrctIncrTicker
  Cells(2, 16).Value = GreatestPrctIncrVol
  Cells(3, 15).Value = GreatestPrctDecrTicker
  Cells(3, 16).Value = GreatestPrctDecrVol
  Cells(4, 15).Value = GreatestTotTicker
  Cells(4, 16).Value = GreatestTotVol
 
End Sub


