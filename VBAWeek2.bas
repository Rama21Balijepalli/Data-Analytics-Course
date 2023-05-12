Attribute VB_Name = "Module1"
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")
    MsgBox (yearvalue)
 
    Worksheets(yearvalue).Activate
    startTime = Timer
         
    'Worksheets("AllStocksAnalysis").Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    
    'Create a header row
       Worksheets(yearvalue).Range("i1").Value = "Ticker"
       Worksheets(yearvalue).Range("j1").Value = "Yearly Change"
       Worksheets(yearvalue).Range("k1").Value = "Percentage Change"
       Worksheets(yearvalue).Range("l1").Value = "Total Stock Volume"
       
       Worksheets(yearvalue).Range("p1").Value = "Ticker"
       Worksheets(yearvalue).Range("q1").Value = "Value"
              
       Worksheets(yearvalue).Range("O2").Value = "Greatest % Increase"
       Worksheets(yearvalue).Range("O3").Value = "Greatest % Decrease"
       Worksheets(yearvalue).Range("O4").Value = "Greatest Total Volume"
       
Dim tickerindex As Integer
        tickerindex = 0
        
        Dim tickerCount As Integer
                 
 
    'Get the number of rows to loop over
    RowCount = Worksheets(yearvalue).Range("A2").End(xlDown).Row - 1 'Worksheets(yearValue).Rows.Count.End(xlDown)
    MsgBox ("Row count:" & RowCount)
    
   ' Copy the Unique ticker values in column I
   Range("A2:A" & RowCount).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("I2"), Unique:=True
  
    
   'Get the unique value count of tickers
   tickerCount = Worksheets(yearvalue).Range("I2").End(xlDown).Row - 1
   MsgBox ("Ticker Count :" & tickerCount)
Dim tickerStartingPrices() As Single
        Dim tickerEndingPrices() As Single
      
          Dim tickers() As String
          Dim tickerValue()  As Single
          Dim tickerReturn() As Single
          Dim tickervolume() As Variant
          
          ReDim tickers(tickerCount)
          ReDim tickerValue(tickerCount)
          ReDim tickerReturn(tickerCount)
          ReDim tickervolume(tickerCount)
          ReDim tickerStartingPrices(tickerCount)
          ReDim tickerEndingPrices(tickerCount)
       
          Dim ran As String
            
          Dim cnt_first, cnt_last As Long
       For tickerindex = 0 To tickerCount
          
        tickers(tickerindex) = Worksheets(yearvalue).Range("I" & tickerindex + 2).Value
            If cnt_last < RowCount Then
            
            cnt_first = WorksheetFunction.Match(tickers(tickerindex), Range("A1:A" & RowCount), 0)
            cnt_last = WorksheetFunction.Match(tickers(tickerindex), Range("A1:A" & RowCount), 0) + (WorksheetFunction.CountIf(Range("A1:A" & RowCount), tickers(tickerindex)) - 1)
          
     
        tickerStartingPrices(tickerindex) = CSng(Cells(cnt_first, 3).Value)
        tickerEndingPrices(tickerindex) = CSng(Cells(cnt_last, 6).Value)
        
        ' Calculate the ticker Return Value
        tickerValue(tickerindex) = tickerEndingPrices(tickerindex) - tickerStartingPrices(tickerindex)
                 
        ' Calculate the ticker Return Percentage
        tickerReturn(tickerindex) = (tickerEndingPrices(tickerindex) - tickerStartingPrices(tickerindex)) * 100 / tickerStartingPrices(tickerindex)
               
        ' Calculate the Ticker index volume
        ran = "G" + CStr(cnt_first) + ":G" + CStr(cnt_last)
        tickervolume(tickerindex) = CVar(WorksheetFunction.Sum(Range(ran)))
        
         Else
                Exit For
        End If
       Next
   
' Worksheets("All Stocks Analysis").Activate
   
    Dim i As Integer
   
    For i = 0 To tickerCount - 1
        Worksheets(yearvalue).Range("I" & i + 2).Value = tickers(i)
        Worksheets(yearvalue).Range("J" & i + 2).Value = tickerValue(i)
        Worksheets(yearvalue).Range("K" & i + 2).Value = tickerReturn(i)
        Worksheets(yearvalue).Range("L" & i + 2).Value = tickervolume(i)
      
      '  Worksheets(yearValue).Range("E" & i + 4).Value = tickerEndingPrices(i)

    Next i
    
       
  ' Get the tickers with Greatest % Increase, Decrease and total Volume
        Dim sumTickers As String
        Dim sumMax, sumMin As Double
        Dim sumTickerVol As Variant
        Dim rowNumber As Long
    
    rnSummary = Worksheets(yearvalue).Range("I2:L" & tickerCount + 1)
    
    
    'Get the Row number of the Largest % Change
    sumMax = WorksheetFunction.Max(tickerReturn())
    rowNumber = WorksheetFunction.Match(sumMax, Range("K2:K" & tickerCount + 1), 0) + 1
    sumTickers = Worksheets(yearvalue).Range("I" & rowNumber + 1).Value
    ' Display the ticker and the values on the Sheet
    Worksheets(yearvalue).Range("P2").Value = sumTickers
    Worksheets(yearvalue).Range("Q2").Value = sumMax
    
    
    ' Get the Min value and display
    sumMin = WorksheetFunction.Min(tickerReturn())
    rowNumber = WorksheetFunction.Match(sumMin, Range("K2:K" & tickerCount + 1), 0) + 1
    sumTickers = Worksheets(yearvalue).Range("I" & rowNumber + 1).Value
    ' Display the ticker and the values on the Sheet
    Worksheets(yearvalue).Range("P3").Value = sumTickers
    Worksheets(yearvalue).Range("Q3").Value = sumMin
   
    ' Get the Max Total Volume and display
    sumTickerVol = WorksheetFunction.Max(tickervolume())
    rowNumber = WorksheetFunction.Match(sumTickerVol, Range("L2:L" & tickerCount + 1), 0) + 1
    sumTickers = Worksheets(yearvalue).Range("I" & rowNumber + 1).Value
    ' Display the ticker and the values on the Sheet
    Worksheets(yearvalue).Range("P4").Value = sumTickers
    Worksheets(yearvalue).Range("Q4").Value = sumTickerVol
   

    i = 0
    For i = 2 To tickerCount + 1
       
        If Worksheets(yearvalue).Cells(i, 10) > 0 Then
           
            Worksheets(yearvalue).Cells(i, 10).Interior.Color = vbGreen
           
        Else
       
            Worksheets(yearvalue).Cells(i, 10).Interior.Color = vbRed
           
        End If
       
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub

       

