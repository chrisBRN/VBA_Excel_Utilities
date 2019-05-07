Attribute VB_Name = "Module1"
Sub clearConditionalFormatting()

'   A simple macro that deletes all conditional formatting in a given workbook

' ***********************************************************************************
' It is strongly recommended that you create a backup of your work before using any VBA.
' Deleting conditional formats in this way cannot be un-done
' ***********************************************************************************

Dim wSheet As Worksheet
    
    For Each wSheet In ActiveWorkbook.Worksheets
    
        wSheet.Cells.FormatConditions.Delete
        
    Next wSheet
    
End Sub
