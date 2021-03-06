Attribute VB_Name = "factorial"
Sub test()

factorial_calculator (10000000)

End Sub

Sub factorial_calculator(Optional n As Variant)

    ' Calculates factorial numbers up until 100 n time
    ' A useful example to use when benchmarking or when testing different profiling tools
    ' Printing to the immediate window increases the overall time considerably
    ' Using values for n of anything less than ~1M (when not printing to debug) results in near instant completion.

    Dim i As Integer
    Dim j As Variant
    Dim fact As Variant
    
    ' Defaults n to 1 if the optional argument is not provided.
    If IsMissing(n) = True Then
        n = 1
    End If
    
    For j = 1 To n
    
        fact = 1
        
        For i = 1 To 100 'overflows @ 170
            fact = fact * i
            'Uncomment below to show output in the debug immediate window
            'Debug.Print "factorial of "; Format(i, "00"); ":"; fact
        Next i
        
    Next j
    
    Debug.Print "done"
    
End Sub
