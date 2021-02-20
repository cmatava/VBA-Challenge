Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()


Dim ticker As String
Dim Percentage_Change As Double
Dim Total_Stock_Volume As Double
Dim Yearly_Change As Double
Dim Open_Price As Double
Dim Close_Price As Double

Range("H1").Value = "Ticker"
Range("I1").Value = "Yearly_Change"
Range("J1").Value = "Percentage_Change"
Range("K1").Value = "Total_Stock_Volume"
Range("C1").Value = "Open_Price"
Range("F1").Value = "Closing_Price"

Yearly_Change = Open_Price - Closing_Price

'1)find the first value for open price - C2 (row 2, column 3)
    'declare a variable that will contain my "open price"
    '1b. set an initial value for running total (0)
    
Dim runningTotal As Double
runningTotal = 0
    
'2. Remembered the value for the open price to reference later
'set the value for that variable to equal the value in C2 Cells(2,3).Value

Dim openPrice As Double
openPrice = Cells(2, 3).Value

Dim rowCount As Long
rowCount = Cells(Rows.Count, "A").End(x1up).Row
    
    Dim stockTicker As String
    Dim nextstockTicker
'3. scroll down until i find a value in column A that is different from the first value
'
'   iterate over rows with a for loop
For i = 2 To rowCount
    stockTicker = Cells(i, 1).Value
    nextstockTicker = Cells(i + 1, 1).Value
    If stockTicker = nextstockTicker Then
       
     'add the volume from each row with the running total of the volume
    dailyVolume = Cells(i, 7).Value
    'As i go through each row, check the value in column A with the row below it
     runningTotal = runningTotal + dailyVolume
     'until I find a value in column A that is different from the first value
     Else
         'add the volume from each row
    dailyVolume = Cells(i, 7).Value
    'with the running total of the volume
    
         runningTotal = runningTotal + dailyVolume
         'set the value in my yearly total volume column to equal runningTotal
         'calculate yearly change
         
'4.Once i find a different value, I get the value for closing price - F263(Row 263, column 6)
        For i = 263 To rowCount
        If nextstockTicker2 = Cells(i + 263, 6).Value Then
    
' 5. i can then find yearly change by subtracting the C2 from F263
    For i = 9 To Yearly_Change
        Yearly_Change = Cells(i + 9, 1).Value
        Yearly_Change = Cells(F2, F70926).Value - Cells(C2, C70926).Value
'5b. I can derermine the final
'6. now i move to the next row and look for the value in column A
' 7. I find the first value for that stocks open price
' 8. remembered the value for that open price to reference later
'repeat steps 3 to 8 until i run out of rows
'Percent Change
    For i = 10 To "Percentage_Change"
        Percentage_Change = (Cells(F2, F70926).Value - Cells(C2, C70926)) / Cells(C2, C70926).Value
        
    
End If
Next i
End Sub
