Attribute VB_Name = "COMPRESSOR"
Public Function sigma_intake(Mach As Single) As Double

    If Mach <= 1 Then
        sigma_intake = 0.97
    Else
        sigma_intake = 0.97 - 0.1 * (Mach - 1) ^ 1.5
    End If
    
End Function

Public Function pressure_sec1(Mach As Single, pressure_free As Single, k As Single) As Double
    
    Dim Lam As Single
    Lam = BASIS.lam_from_mach(Mach, 1.4)
    pressure_sec1 = pressure_free * sigma_intake(Mach) / BASIS.pi(Lam, k)
    
    
End Function

Public Function temperature_sec1(Mach As Single, temperature_free As Single, k As Single) As Double
    
    Dim Lam As Single
    Lam = BASIS.lam_from_mach(Mach, k)
    temperature_sec1 = temperature_free / BASIS.tau(Lam, k)
    
    
End Function
Public Function pressure_sec2(piK As Single, pressure_sec1 As Single, k As Single) As Double
    
    pressure_sec2 = pressure_sec1 * piK
    
    
End Function
