Option Explicit

' Basic approach to the application of recursion in VBA.

Sub factorial()

Dim intNumber As Integer

intNumber = 3
Debug.Print get_factorial(intNumber)

End Sub

' =============================================================
Function get_factorial(N)
  
  If N <= 1 Then
    ' Base case: N = 0
    get_factorial = 1
  Else:
    ' Recursive Case: N > 0.
    get_factorial = get_factorial(N - 1) * N
 End If
End Function
