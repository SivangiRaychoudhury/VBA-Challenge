Attribute VB_Name = "Module1"

Sub Ticker_symbol()


   'Create the headers
    Range("J1") = "Ticker"
    Range("K1") = "Yearly Change"
    Range("L1") = "Percent Change"
    Range("M1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % increase"
    Range("O3") = "Greatest % decrease"
    Range("O4") = "Greatest Total Volume"

    ' Declare the variables
     Dim Ticker_Label As String
     Dim Volume_Total As LongLong
     Dim Yearly_Change As Double
     Dim Percent_Change As Double
     Dim Greatest_Percent_Increase As Double
     Dim Greatest_Percent_Decrease As Double
     Dim Greatest_Total_Volume As LongLong
     Dim Ticker_Max As String
     Dim Ticker_Min As String
     Dim Ticker_Volume As String
     
     
    ' Loops through each worksheet
      For Each ws In Worksheets

        ' Assign value to each variable
          Volume_Total = 0
          Yearly_Change = 0
          Percent_Change = 0

        ' Define the row of 1st open price of the year
           j = 2

         'determine the last row for the provided data table
          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

         ' Keep track of the location for each Ticker in the Ticker summary table
           Dim TickerSummary_Table_Row As Integer
           TickerSummary_Table_Row = 2
  
  
          ' Loop through the stock data
           For i = 2 To LastRow
  

               ' Check where the Ticker label is changing
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

               ' Set the ticker label
                 Ticker_Label = ws.Cells(i, 1).Value
      
               'set yearly change (Final close - Initial open)
                Yearly_Change = Yearly_Change + ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
      
               'set percent change
                Percent_Change = Yearly_Change / ws.Cells(j, 3)
 
                ' Add to the Total Stock volume
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value

                ' Print the Ticker label in the Ticker Summary Table
                 ws.Range("J" & TickerSummary_Table_Row).Value = Ticker_Label
      
                ' Print the Yearly change to the Ticker Summary Table
                 ws.Range("K" & TickerSummary_Table_Row).Value = Yearly_Change
      
                ' Print the Percent change to the Ticker Summary Table
                 ws.Range("L" & TickerSummary_Table_Row).Value = Percent_Change

               ' Display the percent change in % format
                ws.Range("L:L").NumberFormat = "0.00%"
       
               ' Print the Total stock volume to the Ticker Summary Table
                 ws.Range("M" & TickerSummary_Table_Row).Value = Volume_Total
      
               ' Add one increment to the Ticker summary table row
                   TickerSummary_Table_Row = TickerSummary_Table_Row + 1
      
               ' Reset the Total Stock Volume
                 Volume_Total = 0
      
                ' Reset Yearly Change
                 Yearly_Change = 0
      
                 ' Reset Percent change
                  Percent_Change = 0
      
                 ' Reset initial open value
                   j = i + 1
 
                Else

       ' If the cell immediately following a row is the same label then add to the Total Stock volume
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value

              End If
    
       Next i
  
         ' Conditional formatting of the Yearly change and Percent change columns
  
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
            
              ' Copy the headers from sheet 1 to other sheets
               ws.Range("J1:Q1").Value = Sheets(1).Range("J1:Q1").Value
               ws.Range("O2:O4").Value = Sheets(1).Range("O2:O4").Value

            ' Define Last row of the New Summary Table (Final Table)
             LastRowSummary = ws.Cells(Rows.Count, 10).End(xlUp).Row
             
       ' Select the values for the Greatest % increase, Greatest % decrease and Greatest total volume
        Greatest_Percent_Increase = WorksheetFunction.Max((Range(ws.Cells(2, 12), ws.Cells(LastRowSummary, 12))))
        Greatest_Percent_Decrease = WorksheetFunction.Min((Range(ws.Cells(2, 12), ws.Cells(LastRowSummary, 12))))
        Greatest_Total_Volume = WorksheetFunction.Max((Range(ws.Cells(2, 13), ws.Cells(LastRowSummary, 13))))
    
      'Assign/print the values
        ws.Cells(2, 17).Value = Greatest_Percent_Increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 17).Value = Greatest_Percent_Decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 17).Value = Greatest_Total_Volume
        
        ' Get the respective Ticker label
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
          
          
         ' Print the Ticker name in the summary table
           ws.Cells(2, 16).Value = Ticker_Max
           ws.Cells(3, 16).Value = Ticker_Min
           ws.Cells(4, 16).Value = Ticker_Volume
     
      ' Autofit the column width throughout the Worksheets
        ws.Columns("A:Q").AutoFit
      
    
Next ws

End Sub












