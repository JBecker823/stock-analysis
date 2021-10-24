# stock-analysis
Deliverable 1

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    'The tickerIndex is set equal to zero before looping over the rows
    For i = 0 To 11
        tickerIndex = tickers(i)
       Next i
       
    '1b) Create three output arrays
        'arrays are created for tickerVolumes, tickerStartingPrices, and tickerEndingPrices
        Dim tickerVolumes As Long
        Dim tickerStartingPrices As Single, tickerEndingPrices As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        Worksheets(yearValue).Activate
        tickerVolumes = 0
        
        For j = 2 To RowCount
        
        
            'If the next row's ticker doesn't match, increase the tickerIndex.
            If Cells(j, 1).Value = tickerIndex Then
            
      
            End If
        
    ''2b) Loop over all the rows in the spreadsheet.
        If Cells(j, 1).Value = tickerIndex Then
        
    
        '3a) Increase volume for current ticker
            tickerVolumes = tickerVolumes + Cells(j, 8).Value
            End If
            
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
                'Store Starting Price Value
                tickerStartingPrices = Cells(j, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            'Store Ending Price Value
            tickerEndingPrics = Cells(j, 6).Value
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + Cells(j, 8).Value
            
            
        End If
    
        Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For c = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next c
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    


    For h = dataRowStart To dataRowEnd
        
        If Cells(h, 3) > 0 Then
            
            Cells(h, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(h, 3).Interior.Color = vbRed
            
        End If
        
    Next h
    
    Worksheets("All Stocks Analysis").Activate
    Range("A31:C31").Font.FontStyle = "Bold"
    Range("A31:C31").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B32:B43").NumberFormat = "#,##0"
    Range("C32:C32").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 32
    dataRowEnd = 43
    
        For k = dataRowStart To dataRowEnd
        
        If Cells(k, 3) > 0 Then
            
            Cells(k, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(k, 3).Interior.Color = vbRed
            
        End If
        
    Next k
   endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

-------------------------------------------------------------------------------------------------------------------------------------------------------------------

Deliverable 2

Background: Steve’s parents are looking to purchase stock from DQ. They are wondering about the activity of the stock that was actively traded in 2018. Finding was that in 2018 there was a -63% return on stock Daqo. 

Purpose: The purpose of this was to help Steve and his parents create a workbook to determine the best stock to invest in. Steve needed help analyzing some stock data. He wants to find the total daily volume and yearly return for each stock along with the yearly return of the stock from the beginning of the year to the end. 

Results:
Stock Performance: With the selected stocks that Steve and his parents were interested in, there were only 2 that were successful 2 years in a row. After having huge increases of 100% or more in 2017, stocks DQ, FSLR, and SEDG were all negative for 2018. This could mean that the company made changes or that the company was not as desirable as it was in 2017. 

Execution time
Summary: Before making any moves into the stock market, Steve and his parents might want to look into some of the companies more. Most of the stocks that they were analyzing were very volatile from 2017 to 2018. The yearly return, the percentage difference in price from beginning of the year to the end of the year, shows a decrease in most of the stock. This could mean that the stock would be a risk to invest in at this point in time.

About refactoring in general
Pros:	
•	Debugged 
•	Easier to read
•	Improved security and scalability, along with enhanced performance.
•	Can be easier to extend and maintain code
	
Cons: 
•	Can take a while to clean up and debug	
•	Can take extra work to maintain and comprehend which can lead to a complete re-development of a software system

About refactored VBA script
Pros: 	
•	The run time was much faster because information was already stored in the computer’s memory. 
•	The refractured VBS script is easier to read due to reduced complexity
•	Improved source code’s maintainability 
	
Cons: 	
•	Have to make sure everything is correct before entering 
•	For a little refactoring a great deal of attention is paid to expediently adding new features 
