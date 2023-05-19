# VBA---Challenge

## **Challenge Module 2**

- First, create the ``Sub Stocks()`` function to start working with the Visual Basic (VBA) Lenguage.
   - _Sub Stock()_
- Second, to make sure that our script performs the same on every sheet, we declare the variable "ws" to represent a worksheet object.
  -  _Dim ws As Worksheet_
- Third,  ``For Each ws In ThisWorkBook.Worksheets`` sets up a loop to iterate through each worksheet in the current workbook.

- Fourth, the following lines set the headers for various columns in the worksheet using the ``ws.Range("...")``.Value syntax. These lines define the column labels for Ticker, Opening price, Closing price, Yearly change, Percentage Change, Total Stock Volume, and other categories.

- Fifth, several variables are declared using the Dim statement: 
  - lastrow: declare the last row.
  - Headline: defined as 2.
  - i: iterator from the 2nd row up to the last row. 
  - Tickercolumn:it is going to ve the first value on the second row for the column 1 (A), but it will change in each iteration because it is insede the loop.
  
  These variables will be used in the subsequent loop.

- Sixth, Loop to get the *Yearly Change*.The first loop (For i = 2 To lastrow) iterates through the rows of data in the worksheet, starting from the second row.

- Seventh, inside the loop, the code checks if the ticker symbol in the current row is different from the previous row ``If Cells(i - 1, 1).Value <> Tickercolumn Then``. If it is, it means it's the first occurrence of a new ticker symbol ``ws.Range("I" & Headline).Value = Tickercolumn``. In this case, the ticker symbol, opening price, and initial volume are recorded in the appropriate columns: 
 ``ws.Range("I" & Headline).Value = Tickercolumn``
  ``ws.Range("J" & Headline).Value = Cells(i, 3).Value``
   ``Volume = Cells(i, 7)``

- Eighth, if the ticker symbol in the current row is different from the next row ``ElseIf Cells(i + 1, 1) <> Tickercolumn Then``, it means it's the last occurrence of the ticker symbol. The closing price and total volume are recorded, and the headline variable is incremented to move to the next row:
   ``ElseIf Cells(i + 1, 1) <> Tickercolumn Then``
      ``ws.Range("K" & Headline).Value = Cells(i, 6).Value``
      ``Volume = Volume + Cells(i, 7).Value``
      ``ws.Range("N" & Headline).Value = Volume``
      ``Headline = Headline + 1``
      

- Ninth,  if the ticker symbol in the current row is the same as the next row, it means it's not the last occurrence of the ticker symbol. The volume is accumulated, but the opening price and other details remain unchanged.
  
  - All the tickets that are not the first or the last

    ``Else: Volume = Volume + Cells(i, 7).Value``

- Tenth, after the first loop, the code enters another loop (For j = 2 To lastrow2) to calculate *The Percentage Change* and perform conditional formatting based on the change value.

- Eleventh, the percentage change is calculated by subtracting the opening price from the closing price and dividing it by the opening price ``ws.Cells(j, 12) = Cells(j, 11) - Cells(j, 10)`` & ``ws.Cells(j, 13).Value = Cells(j, 12) / Cells(j, 10)``. The result is stored in the respective column.

- Twelfth, conditional formatting is applied to the "Yearly change" column based on whether the change value is positive or negative.
  ``If Cells(j, 12).Value < 0 Then``
    ``ws.Range("L" & j).Interior.ColorIndex = 3``
  ``Else``
    ``ws.Range("L" & j).Interior.ColorIndex = 4``
  ``End If``
``valuecolumn = ws.Cells(j, 14).Value``

- Thirteenth, the code also identifies the ticker symbol with the greatest percentage increase, greatest percentage decrease, and greatest total volume by comparing the values in the respective columns.

      ``If valuecolumn > max_volume Then``

        ``max_volume = valuecolumn``

        ``Tickercolumn2 = ws.Cells(j, 9).Value``

      ``End If``

  ``max_percentage = ws.Cells(j, 13).Value``

      ``If max_percentage > greatestincrease Then``

        ``greatestincrease = max_percentage``

        ``Tickercolumn3 = ws.Cells(j, 9).Value``

      ``End If``

  ``min_percentage = ws.Cells(j, 13).Value``

      ``If min_percentage < greatestdecrease Then``

      ``greatestdecrease = min_percentage``

      ``Tickercolumn4 = ws.Cells(j, 9).Value``

      ``End If``


- Fourteenth, finally, the maximum volume and corresponding ticker symbol, greatest percentage increase and corresponding ticker symbol, and greatest percentage decrease and corresponding ticker symbol are recorded in the summary section of the worksheet.

