# VBA_Excel_Utilities
A series of small utility functions/subroutines that can be used to speed up tedious tasks, and improve workbook speed/size. 
Provided as .txt & .bas files 

### clear_formatting - Clears all cell formatting in a given workbook (excluding number formats)

Straightforward macro that loops through all sheets in a workbook removing cell formatting (colour, border, etc), leaves existing number format in place to avoid clearing date formats to numbers.

Macro contains options to adjust/remove conditional formatting & change cell alignment (commented out by default).

### clear_non_default_styles - Deletes all non-default styles. 

Non default styles can propagate through a workbook when there is a lot of copy/pasting between different workbooks. This in turn can lead to a increased file size and associated slowdown. "clear_non_default_styles" helps to eliminate one cause of slowdown.

### clear_conditional_formatting - Deletes all conditional formatting

Whilst conditional formats are a real time saver, they also tend to self-propagate when users copy data from one workbook to another (inadvertently copying the hidden conditional formatting as well), this can cause unexpected behaviour and contribute to workbook file size increases.  

### speed_up - pre and post code subroutines that will generally speed up any macro
Turns off various excel features, such as screen updating and calculation to improve macro speed.
Also includes a separate subroutine with a boolean argument toggle to control these features with one line of code.
