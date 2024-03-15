Sub multiYearStocks()

'Loop through worksheets
For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'Declaring variables for yearly total, percent change, and total volume by ticker
  Dim i As Long
  Dim openPrice_Row As Long
  Dim tickerName As String
  Dim openYearlyPrice As Double
  Dim totalStockVolume As Double
      totalStockVolume = 0
  Dim yearlyChange As Double
      yearlyChange = 0
  Dim yearlyPercentChange As Double
  Dim printRow As Long
      printRow = 2
  Dim lastRow As Long
  
      lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      openPriceRow = 2
      
'Pulling open value of the stock at beginning of the year
  
       openYearlyPrice = ws.Cells(openPriceRow, 3).Value

'Loop through the work sheet starting at row 2

  For i = 2 To lastRow
  

'If  next ticker cell is not equal to prev. ticker cell then print the value in colomn "I"

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       tickerName = ws.Cells(i, 1).Value
       ws.Range("I" & printRow).Value = tickerName
       
'Calculating Yearly Change of stock price.
       yearlyChange = (ws.Cells(i, 6).Value - openYearlyPrice)
        
       ws.Range("J" & printRow).Value = yearlyChange
    
'Calculating percent change of yearly stock price
        yearlyPercentChange = (yearlyChange / openYearlyPrice)
        ws.Range("K" & printRow).Value = yearlyPercentChange
        ws.Range("K" & printRow).Style = "Percent"
    
'Calculating Total Stock Volume
        totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
        ws.Range("L" & printRow).Value = totalStockVolume
        
'Resetting the variable
        printRow = printRow + 1
        yearlyChange = 0
        totalStockVolume = 0
        openYearlyPrice = ws.Cells(i + 1, 3).Value
        
    Else
'Always add stock volume to the total
        totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
    End If
Next i

'Declaring variables for cell formatting
  Dim yearLastRow As Long

      yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Adding Loop for cell formatting
For i = 2 To yearLastRow

'Adding Conditional for cell formatting
    If ws.Cells(i, 10).Value >= 0 Then
       ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
    
'Declaring variables to find max & min
  
  Dim percentLastRow As Long
      percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
  Dim percentMax As Double
      percentMax = 0
  Dim percentMin As Double
      percentMin = 0

'Adding Loop for finding max & min
For i = 2 To percentLastRow

'Add Conditional for max & min
    If percentMax < ws.Cells(i, 11).Value Then
        percentMax = ws.Cells(i, 11).Value
        ws.Cells(2, 17).Value = percentMax
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ElseIf percentMin > ws.Cells(i, 11).Value Then
        percentMin = ws.Cells(i, 11).Value
        ws.Cells(3, 17).Value = percentMin
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    End If
Next i

'Declaring variable for greatest total volume

  Dim totalStockVolumeRow As Long
      totalStockVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
  Dim totalStockVolumeRowMax As Double
      totalStockVolumeRowMax = 0

'Adding Loop for finding greatest total volume
 
 For i = 2 To totalStockVolumeRow

'Adding Conditional for greatest total volume
    If totalStockVolumeRowMax < ws.Cells(i, 12).Value Then
       totalStockVolumeRowMax = ws.Cells(i, 12).Value
       ws.Cells(4, 17).Value = totalStockVolumeRowMax
       ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
       
    End If
Next i
    
Next ws

End Sub


