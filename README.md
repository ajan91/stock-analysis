# stock-analysis

Overview

The Project asked me to evaluate stocks for a client. The client requested that I use visual basic to discover which stock was benefical to give to his parents. The Excel VBA was used as a means of sorting, cleaning, and analsying which stocks benefical between the years 2018 and 2017. The total stocks value were also addressed in addition of using different methods (i.e. the volume array, price array and a simple ticker array). 

Results

The Following code was used for analysis

    Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = 2 To RowCount

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    End Sub

    Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
       
       '1) Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
        Range("A1").Value = "All Stocks (2018)"
        'Create a header row
         Cells(3, 1).Value = "Ticker"
         Cells(3, 2).Value = "Total Daily Volume"
         Cells(3, 3).Value = "Return"

         '2) Initialize array of all tickers
         Dim tickers(12) As String
         tickers(0) = "AY"
     tickers(1) = "CSIQ"
     tickers(2) = "DQ"
     tickers(3) = "ENPH"
     tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    '3a) Initialize variables for starting price and ending price
    Dim startingPrice As Single
    Dim endingPrice As Single
    '3b) Activate data worksheet
    Worksheets("2018").Activate
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through tickers
    For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

     Next i
   
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

    Sub formatAllStocksAnalysisTable()

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    'Conditioning
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

    
    End Sub

    Sub ClearWorksheet()

    Cells.Clear

    End Sub

    Sub yearValueAnalysis()

    yearValue = InputBox("What year would you like to run the analysis on?")

    Range("A1").Value = "All Stocks (2018)"
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    Worksheets("All Stocks Analysis").Activate
    Worksheets(yearValue).Activate

    End Sub


Conclusion

Refactoring is cleaner, more efficient and more robust. The code lines tend to be faster as seen in the screenshots. This decrease in time is the best seen in the run time difference between the non-refactored code and the refactored code. 
