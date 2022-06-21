# Stock-Analysis
## Prefeorming a detailed analysis on stock data to help determine the total daily volume and returns
## Overview of Project: 
The entire over-arching goal of this project, was to analyze stock data for the years 2017 and 2018. The goal of this specific challenge however, was to refactor the VBA code to both determine if the stock is worth investing in, and to run more efficiently.
## Results: 
### The Code:
To begin, I started by adding the stater code to my stock analysis VBA document, then followed each prompt as seen below to format the improved code.


    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i
        
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

            '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
        
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

## Outputs and Results:
### 2017:

<img width="229" alt="2017 Original Run Time" src="https://user-images.githubusercontent.com/106495685/174711670-637e7265-1623-4cb8-8d10-abc3790e0575.PNG">

This Image is an example of the original runtime of the code for the 2017 stocks data.

<img width="219" alt="2017 Run Time Second Run" src="https://user-images.githubusercontent.com/106495685/174711847-a3abb031-2af3-4a94-86a8-dfe62474529b.PNG">

This image, is an example of the new and improved run time using the more efficient code for the 2017 stocks data.

<img width="266" alt="2017 Results" src="https://user-images.githubusercontent.com/106495685/174712077-d22edd1a-c35d-4346-b129-31f92b789f2d.PNG">

This final image is the results of the 2017 stocks data analysis.

### 2018:
<img width="231" alt="2018 Orginal Run Time" src="https://user-images.githubusercontent.com/106495685/174712283-0f6acab5-5280-4996-89ea-8782aa2b68bc.PNG">

Original run time of 2018 stocks data.

<img width="229" alt="2018 Run Time Second Run" src="https://user-images.githubusercontent.com/106495685/174712341-cc4a2242-4941-4d6b-8738-460c6ec8f898.PNG">

Improved run time of 2018 stocks data.

<img width="264" alt="2018 Results" src="https://user-images.githubusercontent.com/106495685/174712371-957b5874-42f2-44e3-a614-c10f64e6009c.PNG">

Results of 2018 stocks data analysis.

## Summary:
### What are the advantages or disadvantages of refactoring code?
### How do they apply to this VBA code?
Refactoring code, in many cases has more advantages than deiadvantages. It allows us to be more organized within our code as well as helping the code to run more efficiently. It also creates less room for error, and if error does occur it is easier to debug and pin-point. With that being said though, there will always be disadvantages. Simple things like the size of the data being analyized or the time it would take to fine an issue if the refactor broke the code are all possible disadvantages. In the case of this specific code though, the advantages outweigh easily. The data being analyised was relativly small and made refactoring simple. I did run into a formatting bug in my code, but it was easily fixed due to the scale. Refactoring just made sense for this data set, and made it better all the way around.

