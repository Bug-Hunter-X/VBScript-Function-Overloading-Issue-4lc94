Instead of relying on function overloading, use different function names or add parameter checking to differentiate function behavior.  For instance:

' Original buggy code (bug.vbs)
Function MyFunction(a)
  MsgBox "One parameter: " & a
end Function
Function MyFunction(a, b)
  MsgBox "Two parameters: " & a & ", " & b
end Function

' Calling the function will only use the last defined version.
MyFunction 1, 2 ' Outputs "Two parameters: 1, 2"

' Solution (bugSolution.vbs) - Use different function names
Function MyFunction1(a)
  MsgBox "One parameter: " & a
end Function
Function MyFunction2(a, b)
  MsgBox "Two parameters: " & a & ", " & b
end Function

MyFunction1 1
MyFunction2 1, 2

'Solution (bugSolution.vbs) - Add parameter checking
Function MyFunction(a, b)
  If IsMissing(b) Then
    MsgBox "One parameter: " & a
  Else
    MsgBox "Two parameters: " & a & ", " & b
  End If
end Function

MyFunction 1
MyFunction 1,2