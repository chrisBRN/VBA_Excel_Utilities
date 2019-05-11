Attribute VB_Name = "simple_profiler"
Sub simple_profiler()

' A not very accurate function to time the duration of a piece of code, useful for quick comparisons

' stores the current system time
Dim startTime As Date
startTime = Time()

' your code here
' your code here
' your code here
' your code here

' Calculates the time taken in seconds & outputs this to the immediate window
' Time function supplies the time in days, so we multiply 86400 to get the seconds
Debug.Print Format(CDbl((Time - startTime) * 86400), "0.0000"); " seconds elapsed"

End Sub
