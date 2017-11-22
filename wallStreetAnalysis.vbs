'Easy
' Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
' You will also need to display the ticker symbol to coincide with the total volume.

Sub ticker_data_easy()
  
  For Each ws In ActiveWorkbook.Worksheets
  
        ws.Select
    ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRow)
        
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        MsgBox WorksheetName
  
        ' Set an initial variable for holding the ticker symbol
        Dim ticker As String

        ' Set an initial variable for holding the total volume per ticker symbol
        Dim volume As Long
        volume = 0

        ' location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
       
        '  Establish the header for ticker and total stock volume
        Cells(1, 9).Value = "Ticker"
        
        Cells(1, 10).Value = "Total Stock Volume"
        
        
        ' Loop through all tickerdata in  sheet
        For I = 2 To (LastRow)
        
            ' Check if  same ticker symbol..
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

                ' Set the ticker symbol
                ticker = Cells(I, 1).Value
                
                
                ' Add to the volume
                 volume = volume + Cells(I, 7).Value
            
                ' Print the ticker name in the Summary Table
                 Range("I" & Summary_Table_Row).Value = ticker
            
                ' Print the volume against the ticker
                  Range("L" & Summary_Table_Row).Value = volume
        
                ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
              
                ' Reset the volume...getting ready for the next ticker
                  volume = 0

                ' If the cell immediately following a row is the same ticker...
            Else

               
                ' Add to the volume
                volume = volume + Cells(I, 3).Value
                

            End If

        Next I

   Next ws

End Sub
