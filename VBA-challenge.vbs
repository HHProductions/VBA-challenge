Sub assigment_full_working_copy2()

For Each ws In Worksheets


'define the first trading day of the year
Dim start_date As Long
'define the last trading day of the year
Dim end_date As Long
'variable to capture unique tickers
Dim ticker As String
'Dim ticker_counter As Integer
Dim i As Long
'define opening price
Dim openvalue As Double
'define closing price
Dim closevalue As Double
'define sum of trading volume
Dim volcount As Double
'establish a variable for the row number in which the last new unique ticker is entered
Dim tc As Long
Dim numrows As Long
Dim WorksheetName As Integer

' Establish the WorksheetName
WorksheetName = ws.Name

'define start and end date as combination of worhseet naem (year) and defined first and last day of trade
start_date = WorksheetName & "0102"
end_date = WorksheetName & "1231"

'Define the initial variables for the data summary
tc = 2
volcount = 0
ticker = ""

ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"



      ' Identify the number of rows of data. Set numrows = number of rows of data.
        numrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
      ' Establish "For" loop to loop "numrows" number of times.

'loop through all the rows of data
      For i = 2 To numrows
      'what to do if we encounter an already recorded ticker.  This method only works if column A is sorted
      If ws.Cells(i, 1).Value = ticker Then

        If ws.Cells(i, 2).Value = start_date Then
  '         ws.Cells(tc - 1, 11).Value = ws.Cells(i, 3).Value
            openvalue = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i, 2).Value = end_date Then
 '              ws.Cells(tc - 1, 12).Value = ws.Cells(i, 6).Value
                closevalue = ws.Cells(i, 6).Value
        End If
         volcount = volcount + ws.Cells(i, 7).Value
 
      Else
   'what to do if we encounter a new ticker.  This method only works if column A is sorted
        If tc > 2 Then
            ws.Cells(tc - 1, 10).Value = (closevalue - openvalue)
            'conditional formatting
            If ws.Cells(tc - 1, 10).Value < 0 Then
                ws.Cells(tc - 1, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(tc - 1, 10).Value > 0 Then
                    ws.Cells(tc - 1, 10).Interior.ColorIndex = 4
            End If
      
      
            ws.Cells(tc - 1, 11).Value = FormatPercent((closevalue / openvalue) - 1)
            ws.Cells(tc - 1, 12).Value = volcount
            volcount = 0
            'resetting open /close values.  Not necessary as they will get overwritten again.
            openvalue = 0
            closevalue = 0
        End If
      
            ws.Cells(tc, 9).Value = ws.Cells(i, 1)
            ticker = ws.Cells(tc, 9).Value
        If ws.Cells(i, 2).Value = start_date Then
            openvalue = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i, 2).Value = end_date Then
                closevalue = ws.Cells(i, 6).Value
        End If
        volcount = ws.Cells(i, 7).Value
        tc = tc + 1
   
      End If

      Next i
'enter data for the last row, thsi  is needed as the above formula stops at the last row of data
 ws.Cells(tc - 1, 10).Value = (closevalue - openvalue)
         If ws.Cells(tc - 1, 10).Value < 0 Then
      ws.Cells(tc - 1, 10).Interior.ColorIndex = 3
      ElseIf ws.Cells(tc - 1, 10).Value > 0 Then
      ws.Cells(tc - 1, 10).Interior.ColorIndex = 4
      End If
      ws.Cells(tc - 1, 11).Value = FormatPercent((closevalue / openvalue) - 1)
      ws.Cells(tc - 1, 12).Value = volcount
      ws.Columns("i:l").EntireColumn.AutoFit

Next ws

MsgBox ("Finished Main Assignment")

'Bonus section
For Each ws In Worksheets
'define ticker that has highest value increase
    Dim tickermax As String
'define max increase
    Dim phmaxincrease As Double
'counter to identify  if multiple ctickers have the same max increase
    Dim multiplemaxflag As Integer
'define ticker that has highest value decrease
    Dim tickermin As String
'define max increase
    Dim phminincrease As Double
'counter to identify  if multiple ctickers have the same max increase
    Dim multipleminflag As Integer
'define ticker that has highest trading volume
    Dim tickermaxvolume As String
'define max volume
    Dim phmaxvolume As Double
'counter to identify  if multiple ctickers have the same max increase
    Dim multiplemaxvolume As Integer
    
    tickermax = ws.Cells(2, 9).Value
    tickermin = ws.Cells(2, 9).Value
    tickermaxvolume = ws.Cells(2, 9).Value
    phmaxincrease = ws.Cells(2, 11).Value
    phminincrease = ws.Cells(2, 11).Value
    phmaxvolume = ws.Cells(2, 12).Value
    multiplemaxflag = 0
    multipleminflag = 0
    multiplemaxvolume = 0

      
        numrows = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To numrows
            If ws.Cells(i, 11).Value > phmaxincrease Then
                phmaxincrease = ws.Cells(i, 11).Value
                tickermax = ws.Cells(i, 9).Value
                multiplemaxflag = 1
                ElseIf ws.Cells(i, 11).Value = phmaxincrease Then
                    multiplemaxflag = multiplemaxflag + 1
            End If
                
                
            If ws.Cells(i, 11).Value < phminincrease Then
                phminincrease = ws.Cells(i, 11).Value
                tickermin = ws.Cells(i, 9).Value
                multipleminflag = 1
           
                ElseIf ws.Cells(i, 11).Value = phminincrease Then
                    multipleminflag = multipleminflag + 1
            End If
             

            
            If ws.Cells(i, 12).Value > phmaxvolume Then
                phmaxvolume = ws.Cells(i, 12).Value
                tickermaxvolume = ws.Cells(i, 9).Value
                multiplemaxvolume = 1
                ElseIf ws.Cells(i, 12).Value = phmaxvolume Then
                    multiplemaxvolume = multiplemaxvolume + 1
            End If
            
        Next i
     
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    

    ws.Cells(2, 15).Value = "greatest % increase"
    ws.Cells(2, 17).Value = FormatPercent(phmaxincrease)
    If multiplemaxflag > 1 Then
        ws.Cells(2, 16).Value = multiplemaxflag & " tickers shared the greatest % increase"
        Else
            ws.Cells(2, 16).Value = tickermax
    End If
    
    ws.Cells(3, 15).Value = "greatest % decrease"
    ws.Cells(3, 17).Value = FormatPercent(phminincrease)
    If multipleminflag > 1 Then
        ws.Cells(3, 16).Value = multipleminflag & " tickers shared the greatest % decrease"
        Else
            ws.Cells(3, 16).Value = tickermin
    End If
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 17).Value = phmaxvolume
    If multiplemaxvolume > 1 Then
        ws.Cells(4, 16).Value = multiplemaxvolume & " tickers shared the greatest total volume"
        Else
            ws.Cells(4, 16).Value = tickermaxvolume
    End If

 ws.Columns("O:Q").EntireColumn.AutoFit

Next ws
MsgBox ("Finished Assignment Bonus")
End Sub

