Sub WallStreetSingle()

'Designate all variables

'Variable used for counting all rows in the sheet to find upper limit of i
Dim LastRow As Double
'String used to record ticker name
Dim TickerName As String
'Variable used to store sum of Stock Volume
Dim StockVol As Double
'Variable Used to Count Rows in Summary Table
Dim Summary_Table_Row As Double
'Variable used to identify row where a stock's first open price is found
Dim OpenRow As Double
'Variable to hold maximum value of Stock Volume to then place in Bonus Table
Dim MaxStockVol As Double
'Variable to hold value of Maximum % Increase to then place in Bonus Table
Dim MaxInc As Double
'Variable to hold Maximum % Decrease to then place in Bonus Table
Dim MaxDec As Double

'Open Row and Summary Table Row variables designated inside For loop so they reset on each new worksheet
OpenRow = 2
Summary_Table_Row = 2

'Identifies upper limit of i as last row number on each sheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set Summary Headers
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

'Test of Last Row Output
'MsgBox (LastRow)

        For i = 2 To LastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          
                ' Set the Ticker Name
                TickerName = Cells(i, 1).Value

                ' Add to the Stock Volume Total for the final row
                StockVol = StockVol + Cells(i, 7).Value
    
                'Set Open Price Variable using OpenRow variable to set starting point
                Dim OpenPrice As Double
                OpenPrice = Cells(OpenRow, 3).Value
          
                'Set Close Price Variable, inside this if conditional, i will aways be last row of a stock
                Dim ClosePrice As Double
                ClosePrice = Cells(i, 6).Value
      
                'Set Price Difference as Variable and calculate difference between open and close price for the stock
                Dim PriceDif As Double
                PriceDif = ClosePrice - OpenPrice
                
                'Set variable for Percentage Change in value
                Dim PercentChange As Double
      
                'If conditional will resolve any attempts by the loop to divide by zero and just output a zero into that space
                'This is required to avoid any instances of the PercentChange being unable to calculate
                If OpenPrice = 0 Then
      
                    PercentChange = 0
      
                ElseIf OpenPrice > 0 Then
                    
                    'Percent Change in value is calculated as price difference divided by open price
                    'This value will be converted to a percentage in decimal form later
                    PercentChange = PriceDif / OpenPrice
      
                End If
            
                ' Print the Ticker in the Summary Table
                Range("J" & Summary_Table_Row).Value = TickerName
      
                'Print Price Difference in Summary Table
                Range("K" & Summary_Table_Row).Value = PriceDif
      
                'Print Percentage Change in Summary Table
                Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                Range("L" & Summary_Table_Row).Value = PercentChange
      
                'Conditional formatting here sets those with no loss or a profit based on percentage to green squares
                'Squares with a negative percentage are set to red
                If PriceDif >= 0 Then
      
                    Range("K" & Summary_Table_Row).Interior.Color = vbGreen
        
                ElseIf PriceDif < 0 Then
        
                    Range("K" & Summary_Table_Row).Interior.Color = vbRed
        
                End If

                ' Print the Total Stock Volume Amount to the Summary Table
                Range("M" & Summary_Table_Row).Value = StockVol
      
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                'Set OpenRow for next value of i so that it moves down 1 row from close of last stock
                OpenRow = i + 1
      
                ' Reset the Stock Volume Total
                StockVol = 0
      
    
            Else

                ' Add to the Stock Volume Total while ticker is the same
                StockVol = StockVol + Cells(i, 7).Value

            End If

        Next i
        

End Sub