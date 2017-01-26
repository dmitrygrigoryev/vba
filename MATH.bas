Attribute VB_Name = "BASIS"
'Международная атмосфера'

'БЛОК Термодинамические функции '
Public Function R_molar(MolarMass As Single) As Double

    R_molar = 8314.46 / MolarMass
    
End Function

Public Function CVa(Temp As Single, MolarMass As Single) As Double

    CVa = CPa(Temp) - R_molar(MolarMass)
    
End Function

Public Function CVg(Temp As Single, MolarMass As Single) As Double

    CVg = CPg(Temp) - R_molar(MolarMass)
    
End Function

'КОНЕЦ БЛОКА Термодинмические функции'

'БЛОК Теплофизические свойства воздуха и газа(теплоемкость, теплопроводность,вязкость) от температуры'
Public Function DENSa(Temp As Single) As Double
    DENSa = 0
End Function


Public Function CPa(Temp As Single) As Double

 CPa = -0.000000000000334 * (Temp ^ 5) + 0.00000000168 * (Temp ^ 4) - 0.00000333 * (Temp ^ 3) + 0.00313 * (Temp ^ 2) - 1.17 * Temp + 1150#
 
 '-3.34E-13   1.68E-09    -3.33E-06   3.13E-03    -1.17E+00   1.15E+03'
End Function


Public Function LAMa(Temp As Single) As Double

 LAMa = -0.0000000163 * (Temp ^ 2) + 0.0000823 * Temp + 0.0034
 
 '-1.63E-08   8.23E-05    3.40E-03'
End Function

Public Function MUa(Temp As Single) As Double

 MUa = -0.000000000012 * (Temp ^ 2) + 0.0000000504 * Temp + 0.0000045
 
End Function

Public Function DENSg(Temp As Single) As Double
 DENSg = 0
End Function


Public Function CPg(Temp As Single) As Double

 CPg = -4.01E-14 * (Temp ^ 5) + 0.000000000291 * (Temp ^ 4) - 0.000000785 * (Temp ^ 3) + 0.000892 * (Temp ^ 2) - 0.143 * Temp + 1030#
 
End Function


Public Function LAMg(Temp As Single) As Double

 LAMg = -0.00000000187 * (Temp ^ 2) + 0.0000655 * Temp + 0.00482
 
 '-1.87E-09   6.55E-05    4.82E-03'
End Function

Public Function MUg(Temp As Single) As Double

 MUg = -0.0000000000046 * (Temp ^ 2) + 0.0000000395 * Temp + 0.00000604

 '-4.60E-12   3.95E-08    6.04E-06'
End Function
'КОНЕЦ БЛОКА Теплофизические свойства воздуха и газа(теплоемкость, теплопроводность,вязкость) от температуры'

'Газодинамические функции БЛОК'
Public Function q(Lambda As Single, k As Single) As Double
        Dim q1 As Double, q2 As Double
        q1 = Lambda * ((k + 1) / 2) ^ (1 / (k - 1))
        q2 = (1 - ((k - 1) * Lambda ^ 2) / (k + 1)) ^ (1 / (k - 1))
        q = q1 * q2
End Function

Public Function pi(Lambda As Single, k As Single) As Double

    pi = (1 - ((k - 1) * Lambda ^ 2) / (k + 1)) ^ (k / (k - 1))
    
End Function

Public Function tau(Lambda As Single, k As Single) As Double

    tau = (1 - ((k - 1) * Lambda ^ 2) / (k + 1))
    
End Function

Public Function eps(Lambda As Single, k As Single) As Double

    eps = (1 - ((k - 1) * Lambda ^ 2) / (k + 1)) ^ (1 / (k - 1))
    
End Function

Public Function lam_from_mach(Mach As Single, k As Single) As Double

    lam_from_mach = Sqr((k + 1) / 2) * Mach * (1 + (k - 1) * Mach ^ 2 / 2) ^ (-0.5)
    
End Function


'Газодинамические функции КОНЕЦ БЛОКА'

