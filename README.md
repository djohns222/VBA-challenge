# VBA-challenge - VBA code found in "This Workbook" - below sheet 6 (F)

- needed assistance from stack overflow to help create a rule for searching between 1 quarter dates https://stackoverflow.com/questions/9311699/max-value-in-worksheet
 'Check if the date falls within the first quarter (January 1st to March 31st)
  If ws.Cells(j, 1).Value = Ticker And Month(ws.Cells(j, 2).Value) >= 1 And Month(ws.Cells(j, 2).Value) <= 3 Then
  total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value 

- needed assistance from stack overflow to create the max_volume_row variable and = it to the greatest_vol to output correctly in (4,17) https://stackoverflow.com/questions/29531742/unable-to-get-match-property-of-the-worksheetfunction-class
 'Corresponding ticker symbol
  max_volume_row = WorksheetFunction.Match(Greatest_vol, ws.Range("L2:L" & last_row), 0)  
  Tag3 = ws.Cells(max_volume_row + 1, 9).Value 
  
