# Stock-Analysis Using VBA And Excel

## Overview of Project: VBA of Wall Street analysis 1. 2.

This project is based on **Wall Street** project analysis of _green energy_, becuase it was believed that if fosil fuel get used up, there will be more reliance on alternative energy production. There are many forms of green energy to invest in including, Hydro electricity,Wind energy, Geothermal energy and Bio. energy. Steve who was a new graduate in the field of finance had approached me to help with this project.

Steve's parent that needs investing in green energy had not done any research and have decided to invest all there money on **_DAQO New Energy Corp_** . A company that makes silicon wind force for solar panels. Steve is consigned in diversifying his parents funds, he has decided to analyse a handful of green energy stocks in addition to DAQO stocks. He also created an excel containing the stock data that he wants me to analyse.

In this project, i shall be using an extesion build in excel, usually refered to as VBA. VBA is a programming language that leaves in excel , it can read and write cells in worksheets and make calculations.Using code to do analysis will allow steve to use the script with any stocks and reduces the chances of errors.

Results:

1. The tickerIndex is set equal to zero before looping over the rows.

Created a tickerIndex variable and set it equal to zero before iterating over all the rows. I used this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.

### **_Refer to resouces file for code images_**

2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

### **Refer to resouces file for code images**

3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.

## **Refer to resouces file for code images**

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

## **Refer to resouces file for code images**

5. Code for formatting the cells in the spreadsheet is working.

We make positive returns green and negative returns red, to be a lot easier to determine which stocks did well and which ones didn't. Added some formatting based on the values of the returns.

6. There are comments to explain the purpose of the code.

Adding Comments is requiered, as a Best Practices for Writing Super Readable Code such,

-Commenting & Documentation,
-Consistent Indentation,
-Avoid Obvious Comments.
-Code Grouping,
-Consistent Naming Scheme,
-DRY (Don't Repeat Yourself) Principle,
-Avoid Deep Nesting,
-Limit Line Length, etc...

7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

Finally, we run the stock analysis, to confirm that our stock analysis outputs for 2017 and 2018 are the same as dataset example provided (as shown in the images below, named Dataset Examples Provided). In adition, in our resources folder and below you can see the final Stock Analysis Results named, Final VBA Analysis 2017 and 2018 save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook..

Dataset Examples Provided

![data set provided examples of analysis for 2017 & 2018](https://github.com/femiimam001/stock-analysis/blob/main/Resources/data%20set%20provided%20examples%20of%20analysis%20for%202017%20%26%202018.PNG)

SUMMARY:
Deliverable with detail analysis:

1. What are the advantages or disadvantages of refactoring code?

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

Disadvantages:

A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
A complex unstructured code is usually best to split in several functions.
Refactoring process can affect the testing outcomes.
Advantages:

Logical errors easily appear in well structure code that contains nested conditionals and loops.
In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source. 2. How do these pros and cons apply to refactoring the original VBA script?

Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. Now, let's think about something, What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand? If yes then definitely you didn’t pay attention to improve your code or to restructure your code.

We need to consider the code refactoring process as cleaning up the orderly house. Unnecessary clutter in a home can create a chaotic and stressful environment. - The same goes for written code.

A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.
