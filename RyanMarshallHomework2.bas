Attribute VB_Name = "Module1"
Sub LoopThroughStocks()

'Declare variables to loop through worksheets
Dim CurrentSheet As Worksheet

'Loop through all worksheets
For Each CurrentSheet In ThisWorkbook.Worksheets

    CurrentSheet.Columns("J").ColumnWidth = 18
    CurrentSheet.Columns("K").ColumnWidth = 14
    CurrentSheet.Columns("L").ColumnWidth = 14
    CurrentSheet.Columns("M").ColumnWidth = 14
    CurrentSheet.Columns("O").ColumnWidth = 24
    CurrentSheet.Columns("P").ColumnWidth = 15

    'Begin Code for creating Summary Table of Stock Data
    
    'Declare variables
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim LastRow As Long
    Dim StockSummaryTable As Long
    Dim StockVolumeCounter As LongLong
    Dim PercentChange As Double
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim YearlyChange As Double
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentIncreaseStock As String
    Dim GreatestPercentDecrease As Double
    Dim GreatestPercentDecreaseStock As String
    Dim GreatestStockVolume As LongLong
    Dim GreatestStockVolumeStock As String
    
    'Find the last row in the sheet
    LastRow = CurrentSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    'Set StockSummaryTable
    CurrentSheet.Range("J2").Value = "Stock Ticker Symbol"
    CurrentSheet.Range("J2").Font.FontStyle = "Bold"
    CurrentSheet.Range("K2").Value = "Yearly Change"
    CurrentSheet.Range("K2").Font.FontStyle = "Bold"
    CurrentSheet.Range("K3:K20000").NumberFormat = "0.00"
    CurrentSheet.Range("L2").Value = "Percent Change"
    CurrentSheet.Range("L2").Font.FontStyle = "Bold"
    CurrentSheet.Range("L3:L20000").NumberFormat = "0.00"
    CurrentSheet.Range("M2").Value = "Total Volume"
    CurrentSheet.Range("M2").Font.FontStyle = "Bold"
    StockSummaryTable = 3
    
    'Set Greatest Increase And Decrease Table
    CurrentSheet.Range("O2").Value = "Greatest Percent Increase"
    CurrentSheet.Range("O2").Font.FontStyle = "Bold"
    CurrentSheet.Range("O3").Value = "Stock Ticker Symbol"
    CurrentSheet.Range("O3").Font.FontStyle = "Bold"
    CurrentSheet.Range("P3").Value = "Percent Increase"
    CurrentSheet.Range("P3").Font.FontStyle = "Bold"
    CurrentSheet.Range("O6").Value = "Greatest Percent Decrease"
    CurrentSheet.Range("O6").Font.FontStyle = "Bold"
    CurrentSheet.Range("O7").Value = "Stock Ticker Symbol"
    CurrentSheet.Range("O7").Font.FontStyle = "Bold"
    CurrentSheet.Range("P7").Value = "Percent Decrease"
    CurrentSheet.Range("P7").Font.FontStyle = "Bold"
    CurrentSheet.Range("O10").Value = "Greatest Total Volume"
    CurrentSheet.Range("O10").Font.FontStyle = "Bold"
    CurrentSheet.Range("O11").Value = "Stock Ticker Symbol"
    CurrentSheet.Range("O11").Font.FontStyle = "Bold"
    CurrentSheet.Range("P11").Value = "Total Volume"
    CurrentSheet.Range("P11").Font.FontStyle = "Bold"
    
    'Set StockVolumeCounter to first value
    StockVolumeCounter = CurrentSheet.Cells(2, 7).Value
    
    'Set OpenValue to opening value of stock
    OpenValue = CurrentSheet.Cells(2, 3).Value
    
    'Loop through the stocks
    For i = 2 To LastRow
        
        If CurrentSheet.Cells(i, 1).Value <> CurrentSheet.Cells(i + 1, 1).Value Then
            
            'Set closing value of stock
            CloseValue = CurrentSheet.Cells(i, 6).Value
            
            'Calculate Change from Open to Close
            YearlyChange = CloseValue - OpenValue
            
            'Calculate Percent Change from Open to Close
            PercentChange = ((CloseValue / OpenValue) * 100) - 100
            
           'Print out values in Summary Table
            CurrentSheet.Cells(StockSummaryTable, 10).Value = CurrentSheet.Cells(i, 1).Value
            CurrentSheet.Cells(StockSummaryTable, 11).Value = YearlyChange
            CurrentSheet.Cells(StockSummaryTable, 12).Value = PercentChange
            CurrentSheet.Cells(StockSummaryTable, 13).Value = StockVolumeCounter
            
           
            'If statement for formatting
            If CurrentSheet.Cells(StockSummaryTable, 11).Value > 0 Then
                CurrentSheet.Cells(StockSummaryTable, 11).Interior.Color = RGB(7, 235, 110)
                CurrentSheet.Cells(StockSummaryTable, 12).Interior.Color = RGB(7, 235, 110)
            ElseIf CurrentSheet.Cells(StockSummaryTable, 11).Value < 0 Then
                CurrentSheet.Cells(StockSummaryTable, 11).Interior.Color = RGB(247, 144, 141)
                CurrentSheet.Cells(StockSummaryTable, 12).Interior.Color = RGB(247, 144, 141)
            End If
            
            'Increment table position
            StockSummaryTable = StockSummaryTable + 1
            
            'Sum Volume of stock
            StockVolumeCounter = CurrentSheet.Cells(i + 1, 7).Value
            
            'Set opening value of stock for next ticker symbol
            OpenValue = CurrentSheet.Cells(i + 1, 3).Value
        Else
            StockVolumeCounter = StockVolumeCounter + CurrentSheet.Cells(i + 1, 7).Value
        End If
    
    Next i

   'Find Greatest Percent Increase, Decrease, and Total Volume
   GreatestPercentIncreaseStock = CurrentSheet.Range("J3").Value
   GreatestPercentIncrease = CurrentSheet.Range("L3").Value
   GreatestPercentDecreaseStock = CurrentSheet.Range("J3").Value
   GreatestPercentDecrease = CurrentSheet.Range("L3").Value
   GreatestStockVolumeStock = CurrentSheet.Range("J3").Value
   GreatestStockVolume = CurrentSheet.Range("M3").Value
   
   'Set Format for Percent Increase and Decrease Cells
   CurrentSheet.Range("P4").NumberFormat = "0.00"
   CurrentSheet.Range("P8").NumberFormat = "0.00"
   CurrentSheet.Range("P12").NumberFormat = "0"
    
    'Loop through to find Greatest Percent Increase and Decrease
    For j = 3 To 500
        If CurrentSheet.Cells(j + 1, 12).Value <> "" Then
            If CurrentSheet.Cells(j, 12).Value > GreatestPercentIncrease Then
                GreatestPercentIncrease = CurrentSheet.Cells(j, 12).Value
                GreatestPercentIncreaseStock = CurrentSheet.Cells(j, 10).Value
            ElseIf Cells(j, 12).Value < GreatestPercentDecrease Then
                GreatestPercentDecrease = CurrentSheet.Cells(j, 12).Value
                GreatestPercentDecreaseStock = CurrentSheet.Cells(j, 10).Value
            End If
            CurrentSheet.Range("O4").Value = GreatestPercentIncreaseStock
            CurrentSheet.Range("P4").Value = GreatestPercentIncrease
            CurrentSheet.Range("O8").Value = GreatestPercentDecreaseStock
            CurrentSheet.Range("P8").Value = GreatestPercentDecrease
        End If
    Next j
    
    'Loop through to find the Greatest Stock Volume
    For k = 3 To 500
        If CurrentSheet.Cells(k + 1, 13).Value <> "" Then
            If CurrentSheet.Cells(k, 13).Value > GreatestStockVolume Then
                GreatestStockVolumeStock = CurrentSheet.Cells(k, 10).Value
                GreatestStockVolume = CurrentSheet.Cells(k, 13).Value
            End If
            CurrentSheet.Range("O12").Value = GreatestStockVolumeStock
            CurrentSheet.Range("P12").Value = GreatestStockVolume
        End If
    Next k
            
'Go to next worksheet
Next CurrentSheet

End Sub

