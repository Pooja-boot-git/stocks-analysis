# Stock Analysis
## Project Overview: 
We are here to get an overview of Net Change in stock prices traded in years 2017 and 2018. This analysis will help our client Steve decide on which would be best stocks to invest in.

**Stocks/Tickers Analyzed**
1. AY
2. CSIQ
3. DQ
4. ENPH
5. FSLR
6. HASI
7. JKS
8. RUN
9. SEDG
10. SPWR
11. TERP
12. VSLR

## Goal of this challenge: To make the code more efficient.

## Description of data
1. **2017**  :
   - contains data for above mentioned tickers for year 2017 
2. **2018**  :
   - contains data for above mentioned tickers for year 2018
3. **All Stocks Analysis**  :
   - The sheet may/may not contain data when you first open it. 
   - Click **Clear Worksheet** button to clear the sheet of any existing data
   - Click **Stock Analysis** button to get a high level overview of each of the stocks for a given year.
## How it works
- Click on the **Stock Analysis** button. It prompts you to input the year. If you put an incorrect year or no year then it gives you an error message.

   ```YearValue = InputBox("What year would you like to run the analysis on?") 
    If YearValue <> "" Then
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    On Error GoTo ErrorHandler
    ErrorHandler:
    MsgBox ("Incorrect parameter entered")
     ```
     
   ![Incorrect parameter](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Images/Incorrect%20parameter.png)
- Click on **Cancel** button, to exit.
- Put a valid year and click on OK for the analysis to run for that year. The macro provides three output fields:
   1. Ticker - Unique list of all the tickers from the year being considered.
   2. Total Daily Volume - total of daily volume traded for each of the tickers in a given year. 
   3. Return -  (Last transaction (End of the year) - First transaction (start of the year))/First transaction (start of the year) * 100. It gives you the percetage increase or decrease in the price of a ticker in that year.

## Code performance 
- The macro ends with a pop up message stating the time it took for the macro to run. It will be beneficial to know if we decide to consider much larger number of stocks in future to know how our current macro is performing. 
```
    startTime = Timer
    <CODE BLOCK>
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)
 ```
![Time Elapsed](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Images/Time%20Elapsed.png)

- We have refactored the code to improve its performace. While calculating the Total Daily Volume and Return fields, instead of looping through each of the tickers and running code again and again for each of them, we have instead used arrays. By storing data in array, we reduced the amount of times our loop was running. This improved our code's performace.

```
For tickerIndex = 0 To 11
    ticker = tickers(tickerIndex)
        totalVolume = 0
             ''2b) Loop over all the rows in the spreadsheet.
            For i = 2 To rowcount
    
                    '3a) Increase volume for current ticker
                    
                    If Cells(i, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(i, 8).Value
                    End If
                    '3b) Check if the current row is the first row with the selected tickerIndex.
            
                        If Cells(i, 1).Value = ticker _
                        And Cells(i - 1, 1).Value <> ticker Then
                        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                    End If

                    
                    '3c) check if the current row is the last row with the selected ticker
                     'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
                         If Cells(i, 1).Value = ticker _
                    And Cells(i + 1, 1).Value <> ticker Then
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                    End If
'
   Next i

               '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

               tickerVolumes(tickerIndex) = totalVolume
               
              '3d Increase the tickerIndex.
    Next tickerIndex
    
For j = LBound(tickers) To UBound(tickers) - 1
               Cells(AllStocksRow, 1).Value = tickers(j)
               Cells(AllStocksRow, 2).Value = tickerVolumes(j)
               ReturnValue = tickerEndingPrices(j) / tickerStartingPrices(j) - 1
               Cells(AllStocksRow, 3).Value = ReturnValue
               AllStocksRow = AllStocksRow + 1
               Next j 
 ```
 
[Original performance for year 2017](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Resources/green_stocks_2017.png)

[Performance after refactoring](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Resources/VBA_Challenge_2017.png)

[Original performance for year 2018](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Resources/green_stocks_2018.png)

[Performance after refactoring](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Resources/VBA_Challenge_2018.png)

DISCLAIMER : The project currently only contains data for years 2017 and 2018 but it can work for any year as long as the supporting data is present in a separate sheet named after the year. Please note that the code has hardcoded values for tickers but can be improved to make it run for any ticker. Use the below code snippet to achieve that.
```
'unique value calculation taken from
': https://stackoverflow.com/questions/5890257/populate-unique-values-into-a-vba-array-from-excel
'
'If Not Selection Is Nothing Then
'   For Each cell In Selection
'      If (cell <> "") And (InStr(tmp, cell) = 0) Then
'        tmp = tmp & cell & "|"
'      End If
'   Next cell
'End If
'
'If Len(tmp) > 0 Then
'tmp = Left(tmp, Len(tmp) - 1)
'
'tickers = Split(tmp, "|")
'
'End If```
```

## Results
Results for years 2017 and 2018 are attached below.
![2017](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Images/2017.png)
![2018](https://github.com/Pooja-boot-git/stocks-analysis/blob/main/Module2_Challenge/Images/2018.png)

## What did we learn
- Total Daily Volume/investor demand for DQ has increased from 2017 to 2018 but has seen a marked drop in prices. So if we look at just 2 years of data then DQ might not be our best option to invest.
- Tickers ENPH & RUN have both showed positive returns in year 2017 as well as 2018. Their Total Daily volume is also high so high volume with increase in returns give us a good indication that they may be a good choices to invest in.

## Summary
1. Advantages and disadvantages of refactoring code in general
   - Pros
   
      -- The main goal of code refactoring is to make it easy to enhance and maintain in the future.
      
      -- Code size may be reduced improving code readability.
      
      -- Tight couplings are removed making lesser chances of bugs in future versions.
      
      -- Removal of duplicate code again making it easier to maintain.
   
   - Cons
      
      -- Sometimes, it may not be the best time to change the code especially when we have a tight deadline.
      
      -- Any code changes have to be followed by thorough testing which may not happen properly if there's a time crunch and that may inturn induce bugs.
      
2. Advantages and disadvantages of the original and refactored VBA script 

   - Pros
   
      -- We saw improvement in the time it took the code to run after we introduced arrays. Refer to Results section. The new code runs faster than the older one.
      
   - Cons
      
      -- Not applicable. I do not see any disadvantage as there is a high probability of adding many more stocks in future as the client becomes more advanced              stock traders. 
