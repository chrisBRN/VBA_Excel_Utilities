# VBA_Excel_Utilities
A series of small utility functions/macros that can be used to speed up tedious tasks.

### clearFormatting - Plug & play macro that clears cell formatting in a given workbook 

Straightforward macro that loops through all sheets in a workbook removing cell formatting (colour, border, etc), leaves existing number format in place to avoid clearing date formats to numbers.

Macro contains options to adjust to remove conditional formatting & change cell alignment (commented out by default).

### clearNonDefaultStyles - Plug & play macro that deletes all non default styles. 

Non default styles can propagate through a workbook when there is a lot of copy/pasting between different workbooks. This in turn can can lead to a slowdown (and file size increase), this macro can help eliminate one cause of slowdown.
