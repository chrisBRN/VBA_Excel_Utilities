# VBA_Excel_Utilities
A series of small utility functions/macros that can be used to speed up tedious tasks.

### clearFormatting - plug & play macro that clears cell formatting in a given workbook 

Straightforward macro that loops through all sheets in a workbook removing cell formatting (colour, border, etc), leaves existing number format in place to avoid clearing date formats to numbers.

Macro contains options to adjust/remove conditional formatting & change cell alignment (commented out by default).

### clearNonDefaultStyles - plug & play macro that deletes all non-default styles. 

Non default styles can propagate through a workbook when there is a lot of copy/pasting between different workbooks. This in turn can lead to a slowdown (and file size increase), this macro can help eliminate one cause of slowdown.

### clearConditionalFormatting - plug & play macro that deletes all conditional formatting

Whilst conditional formats are a real time saver, they also tend to self-propagate when users copy data from one workbook to another (inadvertently copying the hidden conditional formatting as well), this can cause unexpected behaviour and contribute to workbook slowing down or its size increasing. 

### speedUp - pre and post code functions that will generally speed up a macro

Turns off various excel features, such as screen updating and calculation to improve macro speed.
