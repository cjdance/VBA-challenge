Sub WallStreetLoop()

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
'Variable to count each worksheet
Dim ws As Integer
'Variable holding upper limit of worksheet count
Dim ws_num As Integer
'Designate object worksheet
Dim starting_ws As Worksheet

'Set start point
Set starting_ws = ActiveSheet
ws_num = ThisWorkbook.Worksheets.Count

'For loop to have Macro run on every sheet in Workbook
For ws = 1 To ws_num
    ThisWorkbook.Worksheets(ws).Activate

'Open Row and Summary Table ROw variables designated inside For loop so they reset on each new worksheet
OpenRow = 2
Summary_Table_Row = 2

'Identifies upper limit of i as last row number on each sheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set Summary Headers
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

'Set Bonus Summary headers
Range("P2").Value = "Greatest % Increase"
Range("P3").Value = "Greatest % Decrease"
Range("P4").Value = "Greatest Total Volume"
Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"

'Test of Last Row Output
'MsgBox (LastRow)

        For i = 2 To LastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          
                ' Set the Ticker Name
                TickerName = Cells(i, 1).Value

                ' Add to the Stock Volume Total for the final row
                StockVol = StockVol + Cells(i, 7).Value

                    'Compares total stock volume for a stock to current max stcok volume value
                    'If conditional checks for a higher value to replace greatest total volume
                    'Stores ticker with the highest volume for output to Bonus Table
                    If StockVol > MaxStockVol Then
      
                    MaxStockVol = StockVol
                    Range("Q4").Value = TickerName
        
                    End If
    
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
            
                
                'If conditional checks for the highest and lowest % in the summary table as the loop runs
                'Highest and lowest values are stored for output later
                If PercentChange > MaxInc Then
      
                    MaxInc = PercentChange
                    Range("Q2").Value = TickerName
        
                ElseIf PercentChange < MaxDec Then
    
                    MaxDec = PercentChange
                    Range("Q3").Value = TickerName
    
                End If
      
                ' Print the Ticker in the Summary Table
                Range("J" & Summary_Table_Row).Value = TickerName
      
                'Print Price Difference in Summary Table
                Range("K" & Summary_Table_Row).Value = PriceDif
      
                'Print Percentage Change in Summary Table
                Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                Range("L" & Summary_Table_Row).Value = PercentChange
                
                'Print Greatest % Increase, converted to percentage with 2 decimal places
                Range("R2").NumberFormat = "0.00%"
                Range("R2").Value = MaxInc
      
                'Print Greatest % Decrease, converted to percentage with 2 decimal places
                Range("R3").NumberFormat = "0.00%"
                Range("R3").Value = MaxDec
                
                'Print Greatest Total Stock VOlume to Bonus Summary table
                Range("R4").Value = MaxStockVol
      
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
        
    'Reset Max Values for next worksheet
    MaxDec = 0
    MaxInc = 0
    MaxStockVol = 0

Next ws

End Sub
