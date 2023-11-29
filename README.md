# JWH_VBA-challenge

This challenge has provided me with a Workbook, Multiple_year_stock_data.xlsx, and a support workbook, alphabeticaltesting.xlsx. 
The former has three worksheets, each for 2018, 2019, and 2020. Each of the three worksheets has about 750,000 rows of data. 
The latter has six worksheets, A, B, C, D, E, F. All 6 worksheets have data from 2018, with each sheet containing 
about 22,700 rows of data.

The data sheets comprise the daily trading details of stocks, including their Ticker, Date, Open Price, High Price, Low Price, and Closing Price.
It is suggested that we develop our code on the alphabeticaltesting.xlsx worksheet.

To start the work I opened the alphabeticaltesting.xlsx worksheet and saved it as alphabeticaltesting.xlsm file.

Next, I reviewed the instructions: 
Create a script that loops through all the stocks for one year and outputs the following information. Using the 
alphabeticaltesting workbook, I set out to write the code to process Page A.

 I planned my work around three PARTS. 
 PART 1 - Preparation work. Declaring and initialising variables;
 PART 2 - Processing the data in preparation for posting on Summary Array. I had to locate the boundaries
 between the 12-month data set for each Ticker to process the data. I did this by checking if the next Ticker value 
 was the same as the current Ticker value.
 The start the processing, for the first Ticker, I extracted the opening price and initialised the 
 Total Volume to 0. For each step down the column (column A) the Volume figure is incremented by the daily volume.
 Once the break in the Ticker occurs the data for the current Ticker is finalised, name Ticker, Price change over the year, 
 % price change over the year, and total volume for the year.
 This continues until all the data in Column A are processed.
 
 PART 3 - Output results to the Summary Array.
 Colour Formatting - I colour-formatted the Yearly Change in Price by testing whether the value was positive or 
 negative. If positive I coloured the cell green, if negative I coloured the cell red.

 BONUS - the instructions provided for a BONUS if I could add functionality to my script to return the stock with the 
 "Greatest % Increase, the "Greatest % Decrease, and "Greatest Total Volume. I realised that we could obtain the values using
 the MAX and MIN functions. However, I realised that unless I did considerable research, or borrowed some code from ChatGPT, 
 The MAX and MIN functions on their own would not necessarily tell me which row these values were in. I needed to know 
 their location to relate the values to their associate Ticker.

In each case I constructed a For IF Loop, comparing the appropriate values, and keeping the higher or lower value, 
depending on whether doing increasing or decreasing values, while storing the row number each time the value change.

 I used format commands to format the data accordingly.

 Once I had the code working for Sheet "A" in the alphbeticaltesting.xlsm worksheet. I tested the B, C, D, E, and F. sheets. 
 
 To do so:   

 Sub AnalyzeStockSummary()
 ' PART 1
    '-------------------------------------------------------------
     ' Declare the worksheet as ws
     Dim ws As Worksheet
     ' Change the sheet name for each worksheet.
      Set ws = ThisWorkbook.Sheets("A")    <-------  I changed this value to the worksheet name.
    ' Calculate the last row of the worksheet
        
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row  <-------- This "A" refers to Column A in the ws.

    After confirming the coding worked for all 6 worksheets I worked on the Project worksheet Multiple-year_stock_data.xlsm. 
    All I had to do was to change the "A" to "2018"

    Once I had confirmed the code worked for worksheet "2018" I repeated the process for worksheets "2019" and "2020"

    I have made appropriate screenshots showing the results for each worksheet.

    I have uploaded AnalyzeStockSummary.txt. This file contains the code. I have also
    uploaded AnalyzeStockSummary.vbs

    I tried to upload the workbook Multiple_year_stock_data.xlsm, however, Git Hub Rejected the file as being too large (77MB as Git Hubs maximum of 25.

    You can find my workbook online at https://d.docs.live.net/22a4cee72c1fe257/Desktop/VBA Homework/Multiple-year-stock-data/Multiple_year_stock_data.xlsm. My username is jerry.hooper@bigpond.com, password: JWHmicrosoft$. Note - both are case-sensitive.
    
