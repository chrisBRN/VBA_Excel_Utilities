Attribute VB_Name = "Module1"
Sub RemoveStyles()

' Non-default styles can propagate when there is a lot of copying between workbooks.
' Individually there isn't much of an impact to this, however over time a workbook can accumulate hundreds of styles
' This invariably leads to a slow down (and file size increase).

' The below removes all such non-default styles eliminating one cause of a slow workbook.

' For workbooks with a lot of styles the below can take quite a while to finish and may appear to crash or hang.
' If you do experience 'hanging' closing the workbook and re-opening should fix the issue (with a part way style clearance completed).

'   Turning off calculation & screen updating can increase macro speed
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

Dim customStyle As Style

    On Error Resume Next
    For Each customStyle In ActiveWorkbook.Styles
    
        If Not customStyle.BuiltIn Then
        
             If customStyle.Name <> "1" Then customStyle.Delete
             
        End If
        
     Next customStyle
     
'   //Turning calculation & screen updating back on
    Application.Calculation = xlSemiautomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
     
End Sub
