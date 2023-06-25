# JPMorgan-Excel-Vitural-Programme

_*Note: There are 5 tasks in total for this virtual experience programme. However, the first task is an online multiple choice quiz, and no files avaliable for me to upload. So I omit the descriptions of Task 1 and only include descriptions for the other tasks._


### Task 2: Conditional Formatting

Your task is to use Excel’s conditional formatting tools to explore and visualize the characteristics of the data in the dataset provided in the Additional Resources section below.  

First, familiarize yourself with Excel’s conditional formatting tools by watching the introductory videos using the links provided in Additional Resources. If you are already comfortable with how to use conditional formatting, feel free to refresh your memory with the videos or move on to the exercise.  

Then, open the spreadsheet and familiarize yourself with the data. What kind of data is there? What information do the columns contain? What kind of trends could you see with this kind of data?  

Then, use the conditional formatting tools (either the menu-based tools or write your own conditional formatting formulas, whichever you prefer) to do the following explorations of the data:

* Highlight any cells with formula errors in purple with white text.
* Highlight any cells with missing values in yellow.
* Identify accounts that have not been cross-sold with Product 2 by highlighting the appropriate Product 2 cells in orange.
* Identify accounts that have a 5-year sales CAGR of at least 100% by highlighting the appropriate CAGR cells in green and any accounts with a negative CAGR in red with white text.
* Identify accounts in the top 10% of unit sales for 2021 by highlighting the appropriate 2021 unit sales cells in blue.


### Task 3: Visual Basic for Applications (VBA) Macros

Your task is to familiarize yourself with recording and using simple macros in Excel, and then create two macros using the same spreadsheet you modified with conditional formatting from Task 1. A clean version of that spreadsheet is available in Additional Resources below so you can work from a fresh copy.  

You will create two macros and associated buttons:

* A macro to sort the entire spreadsheet by 5 YR CAGR in descending order to see which accounts have the highest overall 5-year sales growth
* A macro to sort the entire spreadsheet by 2021 unit sales in descending order to see which accounts have the highest overall unit sales in 2021

When you are finished, you will have two buttons that let you very quickly and easily see two ways of analyzing account sales data to inform account planning and other operational decision-making and quickly switch between them. 


### Task 4: Data Visualization in Excel

Your task is to create a simple dashboard using the account sales dataset you worked with in prior tasks. A clean copy of that spreadsheet is available in Additional Resources below.

First, do your background learning using the links in Additional Resources below, particularly for the videos on the basics of building a dashboard in Excel.

Then, consider your dataset. What charts and graphs would be useful related to this data? Unit sales by year? Top 10 accounts by unit sales or CAGR? Effectiveness of different marketing programs by the number of sales driven? Sales by account type? There are a variety of different ways you could gain insight from this dataset. Pick the ones you find most compelling, and use those to create your dashboard.  

Next, consider how you may need to transform the data in the dataset to simplify your analyses. Raw data is as you find it and often not in the ideal form for analysis. You may need to alter the spreadsheet structure or add calculations to support your analysis. Hint: disaggregating the raw data by building a new sheet that has a row per sales year per account, rather than a row per account that combines sales data for all five years, will make it much easier to use pivot tables for some types of analysis. Could you use one or more macros to make constructing that new sheet easier? You may want to filter the data into different views. You will want to add pivot tables to support some kinds of charts you could create. Feel free to change the dataset in any way that supports your analysis.  

Make your data an Excel table (rather than a range). Remember the shortcut for that? It is Ctrl-T. Some of Excel’s more useful capabilities work with data designated as a Table in Excel, including dynamic updating of charts and graphs and much of the pivot table functionality. It is a best practice to use Excel Tables when doing data analysis.

_(Data Source: Forage)_
