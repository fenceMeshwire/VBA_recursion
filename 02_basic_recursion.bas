Option Explicit

Sub start_recursion()

Dim intNumber As Integer

intNumber = 5
Call recursion(intNumber)

End Sub

' =============================================
Sub recursion(intNumber As Integer)

If intNumber = 1 Then
  ' Base Case
  Exit Sub
Else
  ' Recursive Case
  intNumber = intNumber - 1
  recursion intNumber
End If
  
End Sub
