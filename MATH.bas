Attribute VB_Name = "MATH"
Public Function integral_trapezie(args As Range, funcs As Range)
 If args.Count <> funcs.Count Then
    MsgBox ("Размеры диапозонов не равны")
  Else
    cnt = args.Count
    Sum = 0
    For i = 2 To cnt
      Sum = Sum + (funcs.Cells(i - 1) + funcs.Cells(i)) * (args.Cells(i) - args.Cells(i - 1)) / 2
    Next
 End If
 integral_trapezie = Sum
 
End Function

