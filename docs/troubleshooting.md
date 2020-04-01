---
layout: default
title: Troubleshooting
nav_order: 6
---

# Troubleshooting

## Error saving Macro to Personal Workbook, "Microsoft Excel cannot access the file 'C:\Users\USERNAME\AppData\Roaming\Microsoft\Excel\XLSTART'"

### AppData is hidden

Excel's default folder for saving Macros is inside AppData\Roaming\Microsoft\Excel\XLSTART, a folder hidden by the system by default. Open up a File Explorer and navigate to [File]; -> [Options] -> [View] -> [Hidden files and folders] -> **Check** &ltShow hidden files, folders and drives&gt;.

![t1](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/t1.png?raw=true) 

**Click** &lt;Okay&gt;.

### You do not have the administrator's permissison to access /Roaming

If you are not authorized to access the Roaming folder, use any Excel window and navigate to [View] -> [Window] -> [Unhide Window] to bring up any currently hidden Excel windows.

![t2](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/t2.png?raw=true) 

Once unhidden, go to the Workbook named Personal and save it at a folder you are authorized to access.

Then return to any Excel window, navigate to [File] -> [Options] -> [Advanced] -> [General] -> At startup, open all files in:. **Type** in the folder path of the Personal Workbook you just saved to. Excel will now treat that location as the default location for your Personal Macro Workbook.

![t2.2](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/t2.2.png?raw=true) 

### Personal Macro Workbook disabled

The Personal Macro Workbook feature might have been disabled by something. Use any Excel window and navigate to [File] -> [Options] -> [Add-ins]. Browse through your Add-ins list to see if a Personal.xlsb file has been disabled. If so, **set** the file to enabled.

![t3](https://github.com/chase-lsc/Task-Automation-With-Excel-Macros/blob/gh-pages/images/t3.png?raw=true) 
