Attribute VB_Name = "CAMERA"
Public Function qt_from_T(T2 As Double, T3 As Double, eta As Double) As Double
 Dim CpTg As Double, CpTk As Double, ig As Double, CpT0 As Double, Hu As Double
 Hu = 42900
 
 CpTg = -0.000000022186 * T3 ^ 3 + 0.00015686 * T3 ^ 2 + 0.89113 * T3 + 20.36
 CpTk = -0.000000022186 * T2 ^ 3 + 0.00015686 * T2 ^ 2 + 0.89113 * T2 + 20.36
 ig = -0.00000014913 * T3 ^ 3 + 0.001093 * T3 ^ 2 + 1.4251 * T3 - 59.587
 CpT0 = -0.00000014913 * 288 ^ 3 + 0.001093 * 288 ^ 2 + 1.4251 * 288 - 59.587
 
 qt_from_T = (CpTg - CpTk) / (Hu * eta - ig + CpT0)
End Function

Public Function T3_from_T2(T2 As Double, alpha As Double, eta As Double) As Double
 Dim Hu As Double, Lo As Double, T3L As Double, eps As Double, T3M As Double, T3R As Double, eta1 As Double

 Lo = 14.93
 Hu = 42900
 T3R = 2900
 T3L = 300
 eps = 0.00001
 
     qt = 1 / alpha / Lo
     
 If (alpha < 1) Then
    'MsgBox ("Альфа должна быть больше 1")'
    qt = 1 / alpha / Lo
     a = 1
     eta1 = alpha * eta
        Do While (Abs(T3R - T3L) >= eps)
           T3M = (T3R + T3L) / 2
           qt_M = qt_from_T(T2, T3M, eta1) - qt
           qt_L = qt_from_T(T2, T3L, eta1) - qt
               If (qt_M * qt_L <= 0) Then
                   T3R = T3M
               Else
                   T3L = T3M
               End If
          a = Abs(T3M - T3L)
          T3M = (T3R + T3L) / 2
        Loop
    
     T3_from_T2 = T3M
 Else
     qt = 1 / alpha / Lo
     a = 1
        Do While (Abs(T3R - T3L) >= eps)
           T3M = (T3R + T3L) / 2
           qt_M = qt_from_T(T2, T3M, eta) - qt
           qt_L = qt_from_T(T2, T3L, eta) - qt
               If (qt_M * qt_L <= 0) Then
                   T3R = T3M
               Else
                   T3L = T3M
               End If
          a = Abs(T3M - T3L)
          T3M = (T3R + T3L) / 2
        Loop
    
     T3_from_T2 = T3M
 End If

End Function
