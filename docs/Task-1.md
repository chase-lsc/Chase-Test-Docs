---
layout: default
title: Creating a Macro for Organizing and Sorting Imported Data
nav_order: 2
---

# Creating a Macro for Organizing and Sorting Imported Data

Engineering, accounting, inventory and other enterprise software with huge databases usually export their data as TSV or CSV files. In the case of operation data, multiple table files with dozens of columns and hundreds of rows can be generated daily. Macros can automate the daily processing of this type of files completely. This section will go through the steps of creating a macro that can automatically hide the irrelevant columns, adjust selected columns to specified widths, and sort the rows by a chosen field/column.

## 1. Create a New Macro

* 1.1\. **Download** [Sample Data.csv](sampleData/Sample_Data.csv), as linked here.

* 1.2\. **Open** *"Sample Data.csv"* with Excel.

* 1.3\. [View] -> [Macro (expand arrow)] -> [Record Macro].

![1.1.3](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/1.1.3.png?raw=true)

* 1.4\. In the opened up “Record Macro” dialogue box, **type** in the name of your Macro.

![1.1.4](images/1.1.4.png)

|![Caution Icon.](images/caution.png) |**Caution**: Check if the shortcut key you are choosing overwrites any existing shortcut keys.|
|-----|:------|

* 1.5\. 5.	In the “Record Macro” dialogue box, **type** in the shortcut key you want to bind the Macro to.

![1.1.5](images/1.1.5.png)

|![Caution Icon.](images/caution.png) |**Caution**: Macros saved in “New Workbook” or “This Workbook” only work in those respective workbooks. Macros saved in “Personal Macro Workbook” can be used in all spreadsheets.|
|-----|:------|

* 1.6\. In the “Record Macro” dialogue box, select the work space to store the Macro in.

![1.1.6](images/1.1.6.png)

|![Caution Icon.](images/caution.png) |**Caution**: After clicking &lt;OK&gt;>, Excel will now be recording every single action you take and it will not be able to tell idle clicking apart from intended operations. Make sure that you are only taking intended steps inside the Excel window.|
|-----|:------|

* 1.7\. Click <OK> to create a new Macro.
  
![1.1.7](images/1.1.7.png)

## 2. Hide Columns

After confirming and creating a Macro, Excel will now be recording your keystrokes. In this task, we will hide unwanted columns and meanwhile Excel will record all perform into the Macro we just created.

* 2.1\. **Hold Ctrl** and **Click** the letter buttons of the columns you want to hide. In this case we are hiding columns &lt;A&gt;, &lt;B&gt;, and &lt;D&gt;.

![1.2.1](images/1.2.1.png)

* 2.2\. **Right Click** any of the selected letter buttons, here we clicked B.

![1.2.2](images/1.2.2.png)

* 2.3\. **Click** &lt;Hide&gt;.

## 3. Sort by Field

* 3.1\. **Ctrl+A** to select all data in this Excel sheet.

* 3.2\. [Home] -> [Sort & Filter] -> [Custom Sort]

![1.3.2](images/1.3.2.png)

* 3.3\. In the “Sort” dialogue box, **select** the column you want to sort by, values you want to sort on, and the order you want to sort with. In this case we went with sorting by the values of Total in decreasing order.

![1.3.3](images/1.3.3.png)

* 3.4\. **Click** &lt;OK&gt;.

![1.3.4](images/1.3.4.png)

* 3.5\. [View] -> [Macro (expand arrow)] -> [Stop Recording].

![1.3.5](images/1.3.5.png)

## 4. Test the Macro

Your Macro has been created. In this Task we will exit without saving changes to the CSV file and reopen the file to test that it works.

* 4.1\. **Close** the Excel window.

* 4.2\. **Click** &lt;Don't Save&gt;.

![1.4.2](images/1.4.2.png)

* 4.3\. **Click** &lt;Save&gt;.

![1.4.3](images/1.4.3.png)

* 4.4\. **Open** *“Sample Data.csv”*.

* 4.5\. [View] -> [Macro (expand arrow)] -> [View Macros].

![1.4.5](images/1.4.5.png)

* 4.6\. **Check** that the created Macro has been saved.

![1.4.6](images/1.4.6.png)

* 4.7\. **Click** &lt;Run&gt;.

![1.4.7](images/1.4.7.png)

* 4.8\. **Check** that the columns have been hidden and the rows have been sorted.

![1.4.8](images/1.4.8.png)


## Completion Checklist

This is the end of this section. You should now be able to create a Macro that hides specified columns and then sort the entire table by a specified field. You should also be able to open Excel and run a saved Macro. 







