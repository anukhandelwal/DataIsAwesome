'Moderate
'   Create a script that will loop through all the stocks and take the following info.
'   Yearly change from what the stock opened the year at to what the closing price was.
'   The percent change from the what it opened the year at to what it closed.
'   The total Volume of the stock
'   Ticker symbol
'   You should also have conditional formatting that will highlight positive change in green and negative change in red.


Sub ticker_data_moderate()
  
  For Each ws In ActiveWorkbook.Worksheets
  
        ws.Select
    ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRow)
        
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        
        ' MsgBox WorksheetName
  
        ' Set an initial variable for holding the ticker symbol
        Dim ticker As String

        ' Set an initial variable for holding the total volume per ticker symbol
        Dim volume As LongLong
        volume = 0

        ' location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Set Opening and Closing Volume of a ticker
        Dim OpeningValue As Double
        Dim ClosingValue As Double
        OpeningValue = 0
        ClosingValue = 0
   
        '  Establish the header
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        ' Set the Initial Counter to obtain OpeningVOlume of a ticker
        Dim Counter As Integer
        Counter = 1
        
        
        ' Loop through all tickerdata in  sheet
        For I = 2 To (LastRow)
        
            ' Check if  same ticker symbol..
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

                ' Set the ticker symbol
                ticker = Cells(I, 1).Value
                
                'Set Closing Volume
                ClosingValue = Cells(I, 6).Value
                
                ' MsgBox (ClosingValue)
                ' MsgBox (InitialValue)

                ' Add to the volume
                 volume = volume + Cells(I, 7).Value
            
                ' Print the ticker name in the Summary Table
                 Range("I" & Summary_Table_Row).Value = ticker
                 
                 
                 ' Print the Opening Value in the Summary Table
                 'Range("M" & Summary_Table_Row).Value = OpeningValue
                
                ' Print the Closing value in the Summary Table
                 'Range("N" & Summary_Table_Row).Value = ClosingValue
                
                 ' Print the yearly change in the Summary Table
                 Range("J" & Summary_Table_Row).Value = ClosingValue - OpeningValue
                 
                 ' Check if the Difference of Opening and CLosing Value is Positive
                 ' Then Fill Green in Cell
                 ' If Negative Difference then fill Red
                 
                 If (ClosingValue - OpeningValue) <= 0 Then
                 
                    ' Set the Cell Colors to Red
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                 
                  ElseIf (ClosingValue - OpeningValue) > 0 Then
                  
                    ' Set the Cell Colors to Green
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                  End If
        
                   
                 ' Print the Percentage change in the Summary Table
                 Range("K" & Summary_Table_Row).Value = (Abs(ClosingValue - OpeningValue) / OpeningValue)
                 
                 ' Set the Number format to % type
                 Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
                ' Print the volume against the ticker
                  Range("L" & Summary_Table_Row).Value = volume
        
                ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
              
                ' Reset the volume...getting ready for the next ticker
                  volume = 0
                  
                ' Reset the Counter for Opening volume...getting ready for the next ticker
                  Counter = 1

                ' If the cell immediately following a row is the same ticker...
            Else

                ' Add to the volume
                volume = volume + Cells(I, 7).Value
                
                ' Get the Opening Value for ticker
                If Counter = 1 Then
                    OpeningValue = Cells(I, 3).Value
                    Counter = Counter + 1
                End If
                                

            End If

        Next I

   Next ws

End Sub
