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
#### Sheets for your reference
1. **2017**  :
   - contains data for above mentioned tickers for year 2017 
2. **2018**  :
   - contains data for above mentioned tickers for year 2018
3. **All Stocks Analysis**  :
   - The sheet may/may not contain data when you first open it. 
   - Click **Clear Worksheet** button to clear the sheet of any existing data
   - Click **Stock Analysis** button to get a high level overview of each of the stocks for a given year.
#### Understanding the macro
- The project currently only contains data for years 2017 and 2018 but it can work for any year as long as the supporting data is present in a separate sheet named after the year.
- When you click on the **Stock Analysis** button, it prompts you to input the year. If you put an incorrect year or no year then it gives you an error message.

   ```YearValue = InputBox("What year would you like to run the analysis on?") 
    If YearValue <> "" Then
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    On Error GoTo ErrorHandler
    ErrorHandler:
    MsgBox ("Incorrect parameter entered")
     ```
     
- Click on **Cancel** button, to exit.
- Put a valid year and click on OK for the analysis to run for that year. The macro runs to output the total of daily volume traded for each of the tickers in a given year. It also gives you the yearly return for each of the tickers. Return is calculated 
- There is timer that runs which gives you the time taken for the code to run
```
    startTime = Timer
    <CODE BLOCK>
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)
 ```
#### Results
    
    
