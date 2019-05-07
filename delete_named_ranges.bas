Attribute VB_Name = "Module1"
Sub delete_names()

' Removes named ranges (reduces file size and improves performance)

Dim i As Long

    On Error Resume Next
    
    For i = ThisWorkbook.Names.Count To 1 Step -1
    
        ThisWorkbook.Names(i).Delete
        
    Next
    
End Sub
