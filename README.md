# Analysis of Stocks

## Overview of Project
The purpose of this project is to develop a macro using VBA that will allow the client, Steve, to quickly and easily analyze a dataset of stock information.

## Results
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

An initial macro was created to evaluate the total daily volume and return for a list of 12 tickers for one of two years. A timer was used in this macro to evaluate performance. This code used several nested loops to go through the the rows and place the data into the outputs.

```
Sub yearvalueanalysis()

Dim startime As Single
Dim endTime As Single


yearvalue = InputBox("what year would you like to run the analysis on?")

startTime = Timer

       
    'Format Output Sheet
    
        'Create Title
            Range("A1").Value = "All Stocks (" + yearvalue + ")"
        
        'Create Header Row
            Cells(3, 1).Value = "Ticker"
            Cells(3, 2).Value = "Total Daily Volume"
            Cells(3, 3).Value = "Return"
     Worksheets("All Stocks Analysis").Activate
       
        'Define array of Tickers
            Dim tickers(11) As String
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
  

        'Define prices

Worksheets(yearvalue).Activate

        'Count rows to loop over
            RowStart = 2
            RowEnd = Cells(Rows.Count, "A").End(xlUp).Row
            
    'Loop through each ticker
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
            
            Dim totaldailyvolume(11) As Long
            Dim endingPrice(11) As Double
            Dim startingPrice(11) As Double
            
            
            Worksheets(yearvalue).Activate
            
                For j = RowStart To RowEnd
                
                    'get total volume for current ticker
                    
                        If Cells(j, 1).Value = ticker Then

                        totalVolume = totalVolume + Cells(j, 8).Value

                        End If
                        
                        totaldailyvolume(i) = totalVolume
                        
                    'get starting price for current ticker
                        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                        startingPrice(i) = Cells(j, 6).Value
                        
                        End If
                        
                    'get ending price for current ticker
                        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
                        endingPrice(i) = Cells(j, 6).Value
                        
                        End If
           
           Next j
        Next i
        
    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = totaldailyvolume(i)
    Cells(4 + i, 3).Value = endingPrice(i) / startingPrice(i) - 1
    Next i


    
    'Format
    Worksheets("all stocks analysis").Activate
        Range("A3:C3").Font.Bold = True
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("C4:C15").NumberFormat = "0.0%"
        Range("B4:B15").NumberFormat = "#,##0"
        Columns("B").AutoFit

        dataRowStart = 4
        dataRowEnd = 15
        For i = dataRowStart To dataRowEnd
            If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        
            ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
            
            Else: Cells(i, 3).Interior.Color = xlNone

            End If
        Next i
endTime = Timer

MsgBox ("this code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue))

End Sub
```
This macro provided outputs in about 1.2 seconds for the year 2017, and about 1.5 seconds for the year 2018.

![Image](2017_01)

![Image](2018_01)


To improve the function, specifically the time it takes to run the macro. I refactored it by using arrays instead of looping multiple times. To do this I started with the following code preserved from the initial analysis:

```
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(11) As String
    
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
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  
```

Then, I created a variable named "tickerindex" and set that value to 0. I selected the integer data type for this variable, because I knew we would be using the numbers 0 to 11

```
    '1a) Create a ticker Index variable and set it equal to zero
    
          Dim tickerindex As Integer
          tickerindex = 0

```

Then, I created three output arrays for ticker volumes, ticker starting prices, and ticker ending prices. I knew that these arrays would contain the same number of variables as our array of all the tickers.

```
    '1b) Create three output arrays
  
          Dim tickervolumes(11) As Long
          Dim tickerstartingprices(11) As Single
          Dim tickerendingprices(11) As Single

```

Once my output arrays were created, I initialized all the values in the tickervolumes array to zero.

```
    ''2a) Create a for loop to initialize the tickerVolumes to zero. set all the values within that array to zero
            For i = 0 To 11
            tickervolumes(i) = 0
            Next i
    
```


Once my ticker volumes were initialized to zero, I created a for loop to loop over all the rows in the spreadsheet. For this I used an iterator I called "j", and defined my first row as row 2 (since there is a title row), and my final row to be the "RowCount" variable defined earlier in the code

```
    ''2b) Loop over all the rows in the spreadsheet.
            For j = 2 To RowCount

```

The first action I wanted the code to perform within this loop was to sum the ticker volumes for each ticker by increasing the volume for the current ticker. I created an if action in which the macro first checks if the value in the cell of the current loop is equal to the ticker that the tickerindex is accessing. If this is true, then the value in the "Volume" column for that ticker is added to the tickervolumes for that tickerindex. 

```
        '3a) Increase volume for current ticker
            If Cells(j, 1).Value = tickers(tickerindex) Then
            
            tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(j, 8).Value
            
            End If
```
The second action I wanted to perform within this loop was to find the ticker starting prices for each ticker within the ticker index. To do this, I added another if function in which the macro checks if the previous cell is not equal to the current ticker, and the current cell is equal to the current ticker. If this is true, then the value in the "Close" column is assigned to the tickerstartingprices array for the current ticker.

```
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> tickers(tickerindex) And Cells(j, 1).Value = tickers(tickerindex) Then

            tickerstartingprices(tickerindex) = Cells(j, 6).Value
                        
            End If
```

I also wanted to assign ticker ending prices to the tickerendingprices array for each ticker. To do this, I added an if function where the macro checks if the next cell in the tickers colum is not equal to the current ticker and if the current cell in the current tickers column is equal to the current ticker. If this is true, then the value in the "Close" column for the current ticker becomes the tickerendingprice for that ticker.

```
            If Cells(j + 1, 1).Value <> tickers(tickerindex) And Cells(j, 1) = tickers(tickerindex) Then
            
            tickerendingprices(tickerindex) = Cells(j, 6).Value
                        
            End If
```

Once these actions were performed, I wanted to increase the tickerindex value to go through the same steps for the following ticker. To do this, I created another if statement that checks if the tickerindex value is not equal to the current tickerindex value. If this is true, then the tickerindex value will be increased by 1

```
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        
            If Cells(j + 1, 1).Value <> tickers(tickerindex) Then
            
            '3d Increase the tickerIndex.
             tickerindex = tickerindex + 1
            
            End If
```
Once my code within the loops was complete, I ended the loop

```
next j
```

At this point, I had all my values stored within the tickers, tickervolumes, tickerstartingprices, and tickerendingprices arrays. I then wanted to output this information into a different worksheet. To do this, I looped through the tickerindex for all the arrays, outputting each value.

```
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For tickerindex = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + tickerindex, 1).Value = tickers(tickerindex)
        Cells(4 + tickerindex, 2).Value = tickervolumes(tickerindex)
        Cells(4 + tickerindex, 3).Value = tickerendingprices(tickerindex) / tickerstartingprices(tickerindex) - 1
       
    Next tickerindex
```

I also wanted my output to be formatted for visual pleasure. This is also preserved from the initial subroutine

```
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
```
I kept the ending timer and the message box as well to be able to evaluate performance and compare.

```
endTime = Timer

MsgBox ("this code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue))

End Sub

```

Once I completed refactoring the code, I re-ran the analysis for both years to compare how the refactored code performed compared to the initial. Both years analyses ran in under 0.3 seconds.


![Image](2017_02)

![Image](2018_02)

This was about a 1 second improvement on the original code for each year.

## Summary 

What are the advantages or disadvantages of refactoring code?

One of the advantages of refactoring code is the possibility of improving the code. In refactoring this code, the time to run the script was greatly improved. Refactoring code also allows the coder to maintain flexibility in understanding patterns by figuring out different ways to do things.

There are also potentially disadvantages of refactoring code. First, it could be difficult to determine where to start. Starting over may not be necessary, but then an important step may be missed. Another disadvanteage could be that the refactored code could perform worse than the initial code. 

How do these pros and cons apply to refactoring the original VBA script?
The original script was less efficient because it looped through several loops and outputted into the output worksheet as it went. An advantage of this method is that it seemed relatively straightforward once I learned loops. When I was refactoring the code, I ended up with several versions that were far less efficient (run time >50 seconds), and some versions that crashed excel. At this point, I was not quite sure if there was much benefit to refactoring the code since < 2 seconds seemed far better than what I was getting. However once I successfully refactored the code with arrays, I found that it was a bit more straightforward to read and understand as well as being significantly quicker.
