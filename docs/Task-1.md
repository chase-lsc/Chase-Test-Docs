---
layout: default
title: Creating a Macro for Organizing and Sorting Imported Data
nav_order: 2
---

# Creating a Macro for Organizing and Sorting Imported Data

Engineering, accounting, inventory and other enterprise software with huge databases usually export their data as TSV or CSV files. In the case of operation data, multiple table files with dozens of columns and hundreds of rows can be generated daily. Macros can automate the daily processing of this type of files completely. This section will go through the steps of creating a macro that can automatically hide the irrelevant columns, adjust selected columns to specified widths, and sort the rows by a chosen field/column.

## Create a New Macro

First we will create and name a new Macro.

1\. **Download** [Sample Data.csv](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/sampleData/Sample_Data.csv), as linked here.

2\. **Open** *"Sample Data.csv"* with Excel.

3\. [View] -> [Macro (expand arrow)] -> [Record Macro].

![1.1.3](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.1.3.png?raw=true)

4\. In the opened up “Record Macro” dialogue box, **type** in the name of your Macro.

![1.1.4](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.1.4.PNG?raw=true)

|![Caution Icon.](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/caution.PNG?raw=true) |**Caution**: Check if the shortcut key you are choosing overwrites any existing shortcut keys.|
|-----|:------|

5\. In the “Record Macro” dialogue box, **type** in the shortcut key you want to bind the Macro to.

![1.1.5](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.1.5.PNG?raw=true)

|![Caution Icon.](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/caution.PNG?raw=true) |**Caution**: Macros saved in “New Workbook” or “This Workbook” only work in those respective workbooks. Macros saved in “Personal Macro Workbook” can be used in all spreadsheets.|
|-----|:------|

6\. In the “Record Macro” dialogue box, select the work space to store the Macro in.

![1.1.6](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.1.6.PNG?raw=true)

|![Caution Icon.](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/caution.PNG?raw=true) |**Caution**: After clicking &lt;OK&gt;>, Excel will now be recording every single action you take and it will not be able to tell idle clicking apart from intended operations. Make sure that you are only taking intended steps inside the Excel window.|
|-----|:------|

7\. Click <OK> to create a new Macro.
  
![1.1.7](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.1.7.PNG?raw=true)

## Hide Columns

After confirming and creating a Macro, Excel will now be recording your keystrokes. In this task, we will hide unwanted columns and meanwhile Excel will record all perform into the Macro we just created.

1\. **Hold Ctrl** and **Click** the letter buttons of the columns you want to hide. In this case we are hiding columns &lt;A&gt;, &lt;B&gt;, and &lt;D&gt;.

![1.2.1](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.2.1.PNG?raw=true)

2\. **Right Click** any of the selected letter buttons, here we clicked B.

![1.2.2](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.2.2.PNG?raw=true)

3\. **Click** &lt;Hide&gt;.

## 3Sort by Field

We will now create a custom sort to sort the table by a field of our choosing.

1\. **Ctrl+A** to select all data in this Excel sheet.

2\. [Home] -> [Sort & Filter] -> [Custom Sort]

![1.3.2](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.3.2.PNG?raw=true)

3\. In the “Sort” dialogue box, **select** the column you want to sort by, values you want to sort on, and the order you want to sort with. In this case we went with sorting by the values of Total in decreasing order.

![1.3.3](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.3.3.PNG?raw=true)

4\. **Click** &lt;OK&gt;.

![1.3.4](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.3.4.PNG?raw=true)

5\. [View] -> [Macro (expand arrow)] -> [Stop Recording].

![1.3.5](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.3.5.PNG?raw=true)

## Test the Macro

Your Macro has been created. In this Task we will exit without saving changes to the CSV file and reopen the file to test that it works.

1\. **Close** the Excel window.

2\. **Click** &lt;Don't Save&gt;.

![1.4.2](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.4.2.PNG?raw=true)

3\. **Click** &lt;Save&gt;.

![1.4.3](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.4.3.PNG?raw=true)

4\. **Open** *“Sample Data.csv”*.

5\. [View] -> [Macro (expand arrow)] -> [View Macros].

![1.4.5](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.4.5.PNG?raw=true)

6\. **Check** that the created Macro has been saved.

![1.4.6](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.4.6.PNG?raw=true)

7\. **Click** &lt;Run&gt;.

![1.4.7](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.4.7.PNG?raw=true)

8\. **Check** that the columns have been hidden and the rows have been sorted.

![1.4.8](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.4.8.PNG?raw=true)


## Completion Checklist

This is the end of this section. You should now be able to create a Macro that hides specified columns and then sort the entire table by a specified field. You should also be able to open Excel and run a saved Macro. 







