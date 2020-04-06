---
layout: default
title: Macro for Splitting and Combining Text Columns
nav_order: 3
---

# Macro for Splitting and Combining Text Columns


Long lists are commonly found in one column, which has a lack of readability. This macro does take one second by calculating the number of the address, extension size of the block and divide one column in three columns to increase the visibility. 



 ## Split address from one column into three columns
 
 1\. Press [CTROL+Home] -> [CTROL+A] to place curser at the A1 and select whole columns.
 
 ![2.1.1](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.1.png?raw=true)
 
 
 
 2\. Click [DATA menu] to select **TEXT to Columns**
 
 ![2.1.2](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.2.png?raw=true)
 
 
 
 3\. In the dialogue box, select the [**Delimited**] -> [Press NEXT].
 
 ![2.1.3](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.3.png?raw=true)
 
 
 
 4\. In the second dialogue box, [Unclick the tab] -> [Click the **space**] -> [Press next].
 
 ![2.1.4](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.4.png?raw=true)
 
 
 
 5\. In the last diaglogue box, Change [General] -> [**Text**] -> [Press Finish].
 
 ![2.1.5](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.5.png?raw=true)
 
 
 
 6\. Hold **shift+right arrow twice** to select all columns. Click [Home menu] -> [Click Sort &Filter] ->[Custom sort].
 
 In the custom sort dialogue, click **Add level twice** then choose **Column B** for [**sort by**] and **Column A** 
 
 for [**Then by**] -> [Press OK].
 
 ![2.1.6](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.6.png?raw=true)
    
    
    
 7\. Press CTROL+HOME then hold **shift+right arrow twice**. Click [Home menu] to select [Format] -> [Column Width]
  -> [type 15] then press OK.   
 
 ![2.1.8](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.8.png?raw=true)
 
 
 
 8\. Press [CTROL+HOME] -> [Home menu] to select [Insert] -> [Insert Sheet] to add one more row.
 
 Type [**Toal address**] -> [**Tab**] to move to **B1**. Type [**=COUNTA(**] then press **Shift+Down arrow** to calculate the 
 
 number of the address.
 
 
 
 ![2.1.9](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.9.png?raw=true)
 
 
 9\. Press stop recording, save the worksheet. Open the macro file and **test** it.
 
 
 
 ## Adjust for two columns into one
 
 
 What if you find that the half of the list have city name and rest of the list does not? This macro will create new columns of each 
 list and rejoin them in the original column.
 
 
 
 1\.  Press [CTROL+HOME] -> [CONTROL+A] to place the cursor at A1. Click **DATA menu** to select **Text to Columns**. 
 
 In the dialogue box, choose [**DELIMITED**]-> [Press NEXT]. In the second dialogue box, choose [**space**] ->[Press Next]. In the last 
 
 dialogue box, unclick the **tab** then click [**space**] -> [Press Finish].
 
 
 ![2.2.1](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.2.1.png?raw=true)
 
 
 
 2\. Press [CTROL+HOME] -> [CONTROL+A]. Click [Home menu] to select [SORT &Filter] -> [column C]
 
 ![2.1.6](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.1.6.png?raw=true)
 
 
 
 
 3\. Press **End** to move the bottom of the column, then press **Right arrow thrice+Down arrow once** and type **STOP**
 
 
 4\. Press [End]->[Up arrow] to move to the **D1**, then type **=A1&" "&B1+Enter**
 
 
 ![2.2.2](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.2.2.png?raw=true)
 
 
 
 |Icon|Explanation|
|-----|:------|
|![Caution Icon.](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/caution.png?raw=true) |**Caution**: Check the dotted line range. If only one block is selected, reselect whole **D** column except **Stop** block|
 
 
 
 5\. Press **CTROL+C**, then press [End] -> [Down arrow]-> [Up arrow] -> [CTROL+V] -> [CTROL+C]. Select [Home menu] -> [paste]
 
 ->[Paste Value]
 
 
 ![2.2.7](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.2.7.png?raw=true)
 
 
 
 6\. Press [CTROL+Home] -> [End] -> [Down arrow] -> [CTROL+V]
 
 
 7\. Press **Right arrow three times** then, [End] -> [Shift+Down arrow] -> [Delete] -> [Left arrow twice] -> [End] ->
 
 [Shift+Down arrow] ->[Delete]
 
 
 ![2.2.9](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/2.2.9.png?raw=true)
 
 
 
 8\. Press stop recording, save the worksheet. Open the macro file and test it.
 

## Completion Checklist

This is the end of this section. Now, you are able to split the one column to three individual columns.
Besides, you can combine the split columns.

