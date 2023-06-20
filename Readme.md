# VBA_STOCK_MARKET_ANALYSIS-2018-2020

In this project VBA is used to analyse the stock market from 2018-2020

# GOAL
Write a script in VBA that loops through all the stocks of one year and outputs: 
1. Ticker labels
2. Find the yearly change
3. Percent change
4. Total stock volume
5. Greatest % increase
6. Greatest % decrease
7. Greatest Total volume

The script should run through all the given worksheets, i.e 2018, 2019 and 2020 at the same time

#CODE
- Create the headers in the Worksheet
   Ticker, Yearly change, Percent change, total stock volume,ticker, value, greatest % increase, greatest % decrease, Greatest total volume
   
- Declare the variables
  Ticker_Label, Volume_Total, Yearly_change, Percent_change, greatest_percent_increase, Greatest_percent_decrease, greatest_total_volume, ticker_max, ticker_min, ticker_volume
  
- Since the script needs to run through all worksheets, 
    For Each ws In Worksheets
    
- Assign value to the variables
          Volume_Total = 0
          Yearly_Change = 0
          Percent_Change = 0

- Define the row of 1st open price of the year
           j = 2
           
- Determine the last row for the provided data table
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
- Keep track of the location for each Ticker in the Ticker summary table
           Dim TickerSummary_Table_Row As Integer
           TickerSummary_Table_Row = 2

- Use For loop to loop through the stock data

- Check where the Ticker label is changing
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

(ws. is used to run through all the worksheets)   
 
- Set the ticker label
        Ticker_Label = ws.Cells(i, 1).Value
      
- Set yearly change (Final close - Initial open)
     Yearly_Change = Yearly_Change + ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
      
- Set percent change (Yearly change - Initial open)
    Percent_Change = Yearly_Change / ws.Cells(j, 3)
 
- Add to the Total Stock volume
    Volume_Total = Volume_Total + ws.Cells(i, 7).Value

- Print the Ticker_label, Yearly_change, Percent_change, Volume_total in the J, K, L, M columns respectively in the Ticker Summary Table
                
- Display the percent change in % format
    ws.Range("L:L").NumberFormat = "0.00%"
       
- Add one increment to the Ticker summary table row so that the next Ticker label gets stored there.
    TickerSummary_Table_Row = TickerSummary_Table_Row + 1
      
- Reset the Total Stock Volume, Yearly_Change, Percent_Change to 0
      
- Reset initial open value for the next Ticker name
                   j = i + 1
 
                Else

If the cell immediately following a row is the same label then add to the Total Stock volume
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value

- End If
    
- Conditional formatting of the Yearly_change and Percent_change columns. Positive values are in Green and negative values are in Red. 

    For i = 2 To LastRow
               For j = 11 To 12
                 If (ws.Cells(i, j).Value < 0) Then
                     ws.Cells(i, j).Interior.ColorIndex = 3
                     Else
                     ws.Cells(i, j).Interior.ColorIndex = 4
                    If (ws.Cells(i, j).Value = 0) Then
                      ws.Cells(i, j).Interior.ColorIndex = White
                    End If
                End If
              Next j
         Next i
         
*(Used conditional formatting for both the columns because in the assignment it is described to do so. But the image shown doesn't have in both the columns.)

- Copy the headers from sheet 1 to other sheets

- Define Last row of the New Summary Table (Final Table)
    LastRowSummary = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
- Select the values for the Greatest % increase, Greatest % decrease and Greatest total volume, using the inbuilt function i.e WorksheetFunction.Max & WorksheetFunction.Min

        Greatest_Percent_Increase = WorksheetFunction.Max((Range(ws.Cells(2, 12), ws.Cells(LastRowSummary, 12))))
        Greatest_Percent_Decrease = WorksheetFunction.Min((Range(ws.Cells(2, 12), ws.Cells(LastRowSummary, 12))))
        Greatest_Total_Volume = WorksheetFunction.Max((Range(ws.Cells(2, 13), ws.Cells(LastRowSummary, 13))))

- Print those values to its respective cell numbers in the Summary Table

- Get the respective Ticker labels 

    For k = 2 To LastRowSummary
                If ws.Cells(k, 12).Value = ws.Cells(2, 17).Value Then
                Ticker_Max = ws.Cells(k, 10)
                End If
          Next k
          
          For k = 2 To LastRowSummary
               If ws.Cells(k, 12).Value = ws.Cells(3, 17).Value Then
                Ticker_Min = ws.Cells(k, 10)
              End If
          Next k
              
           For k = 2 To LastRowSummary
               If ws.Cells(k, 13).Value = ws.Cells(4, 17).Value Then
                Ticker_Volume = ws.Cells(k, 10)
              End If
          Next k
           
- Print the respective Ticker labels in the Summary Table in the cells (2,16), (3,16) and (4,16)

- Autofit the column width 

- Next ws (that ends For Each ws In Worksheets)

# ANALYSIS
- In the Year 2018

  THB has the Highest % increase of 141.42% and RKS has the greatest % decrease of -90.02%. The greatest total volume was obtained by QKN with the value of 1.69E+12
  
- In the Year 2019

    RYU has the Highest % increase of 190.03% and RKS has the greatest % decrease of -91.60%. The greatest total volume was obtained by ZQD with the value of 4.37E+12
    
- In the Year 2020

    YDI has the Highest % increase of 188.76% and VNG has the greatest % decrease of -89.05%. The greatest total volume was obtained by QKN with the value of 3.45E+12
