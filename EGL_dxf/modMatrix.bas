Attribute VB_Name = "modMatrix"
Option Explicit

'Public Const sPI As Single = 3.14159
'Public Const sPIDiv180 As Single = sPI / 180
Private IdentityMatrix As Matrix

Public Type Matrix
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type

    
Public sngCos(-359 To 359) As Single
Public sngSin(-359 To 359) As Single

Public Sub LookUpTable()
    
    Dim idx As Integer
    Dim radian As Double
    
    For idx = -359 To 359
        radian = idx * (3.14159 / 180)
        sngCos(idx) = Round(Cos(radian), 3)
        sngSin(idx) = Round(Sin(radian), 3)
    Next
    
End Sub

Public Sub SetIdentity()
    
    IdentityMatrix = MatrixIdentity()

End Sub

'Private Function DegToRad(Degress As Single) As Single
'
'    DegToRad = Degress * (sPIDiv180)
'
'End Function

Private Function MatrixIdentity() As Matrix
    
    With MatrixIdentity
        .rc11 = 1: .rc12 = 0: .rc13 = 0: .rc14 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0: .rc24 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1: .rc34 = 0
        .rc41 = 0: .rc42 = 0: .rc43 = 0: .rc44 = 1
    End With

End Function

Public Function MatrixMultiply(M1 As Matrix, M2 As Matrix) As Matrix

    Dim M1t As Matrix
    Dim M2t As Matrix
    
    M1t = M1
    M2t = M2
    
    MatrixMultiply = IdentityMatrix
    
    With MatrixMultiply
        .rc11 = (M1t.rc11 * M2t.rc11) + (M1t.rc21 * M2t.rc12) + (M1t.rc31 * M2t.rc13) + (M1t.rc41 * M2t.rc14)
        .rc12 = (M1t.rc12 * M2t.rc11) + (M1t.rc22 * M2t.rc12) + (M1t.rc32 * M2t.rc13) + (M1t.rc42 * M2t.rc14)
        .rc13 = (M1t.rc13 * M2t.rc11) + (M1t.rc23 * M2t.rc12) + (M1t.rc33 * M2t.rc13) + (M1t.rc43 * M2t.rc14)
        .rc14 = (M1t.rc14 * M2t.rc11) + (M1t.rc24 * M2t.rc12) + (M1t.rc34 * M2t.rc13) + (M1t.rc44 * M2t.rc14)
        
        .rc21 = (M1t.rc11 * M2t.rc21) + (M1t.rc21 * M2t.rc22) + (M1t.rc31 * M2t.rc23) + (M1t.rc41 * M2t.rc24)
        .rc22 = (M1t.rc12 * M2t.rc21) + (M1t.rc22 * M2t.rc22) + (M1t.rc32 * M2t.rc23) + (M1t.rc42 * M2t.rc24)
        .rc23 = (M1t.rc13 * M2t.rc21) + (M1t.rc23 * M2t.rc22) + (M1t.rc33 * M2t.rc23) + (M1t.rc43 * M2t.rc24)
        .rc24 = (M1t.rc14 * M2t.rc21) + (M1t.rc24 * M2t.rc22) + (M1t.rc34 * M2t.rc23) + (M1t.rc44 * M2t.rc24)
        
        .rc31 = (M1t.rc11 * M2t.rc31) + (M1t.rc21 * M2t.rc32) + (M1t.rc31 * M2t.rc33) + (M1t.rc41 * M2t.rc34)
        .rc32 = (M1t.rc12 * M2t.rc31) + (M1t.rc22 * M2t.rc32) + (M1t.rc32 * M2t.rc33) + (M1t.rc42 * M2t.rc34)
        .rc33 = (M1t.rc13 * M2t.rc31) + (M1t.rc23 * M2t.rc32) + (M1t.rc33 * M2t.rc33) + (M1t.rc43 * M2t.rc34)
        .rc34 = (M1t.rc14 * M2t.rc31) + (M1t.rc24 * M2t.rc32) + (M1t.rc34 * M2t.rc33) + (M1t.rc44 * M2t.rc34)
        
        .rc41 = (M1t.rc11 * M2t.rc41) + (M1t.rc21 * M2t.rc42) + (M1t.rc31 * M2t.rc43) + (M1t.rc41 * M2t.rc44)
        .rc42 = (M1t.rc12 * M2t.rc41) + (M1t.rc22 * M2t.rc42) + (M1t.rc32 * M2t.rc43) + (M1t.rc42 * M2t.rc44)
        .rc43 = (M1t.rc13 * M2t.rc41) + (M1t.rc23 * M2t.rc42) + (M1t.rc33 * M2t.rc43) + (M1t.rc43 * M2t.rc44)
        .rc44 = (M1t.rc14 * M2t.rc41) + (M1t.rc24 * M2t.rc42) + (M1t.rc34 * M2t.rc43) + (M1t.rc44 * M2t.rc44)
    End With

End Function

Public Function MatrixMultVector(M As Matrix, V As Vector) As Vector
      
    MatrixMultVector.X = (M.rc11 * V.X) + (M.rc12 * V.Y) + (M.rc13 * V.Z) + (M.rc14 * V.W)
    MatrixMultVector.Y = (M.rc21 * V.X) + (M.rc22 * V.Y) + (M.rc23 * V.Z) + (M.rc24 * V.W)
    MatrixMultVector.Z = (M.rc31 * V.X) + (M.rc32 * V.Y) + (M.rc33 * V.Z) + (M.rc34 * V.W)
    MatrixMultVector.W = (M.rc41 * V.X) + (M.rc42 * V.Y) + (M.rc43 * V.Z) + (M.rc44 * V.W)
   
End Function

Public Function MatrixScale(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Matrix

    MatrixScale = IdentityMatrix
    MatrixScale.rc11 = X
    MatrixScale.rc22 = Y
    MatrixScale.rc33 = Z
   
End Function

Public Function MatrixTranslation(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Matrix

    MatrixTranslation = IdentityMatrix
    MatrixTranslation.rc14 = X
    MatrixTranslation.rc24 = Y
    MatrixTranslation.rc34 = Z
   
End Function

Public Function MatrixRotationX(ByVal angle As Single) As Matrix

    MatrixRotationX = IdentityMatrix
    MatrixRotationX.rc22 = sngCos(angle)
    MatrixRotationX.rc23 = -sngSin(angle)
    MatrixRotationX.rc32 = sngSin(angle)
    MatrixRotationX.rc33 = sngCos(angle)

End Function

Public Function MatrixRotationY(ByVal angle As Single) As Matrix

    MatrixRotationY = IdentityMatrix
    MatrixRotationY.rc11 = sngCos(angle)
    MatrixRotationY.rc31 = -sngSin(angle)
    MatrixRotationY.rc13 = sngSin(angle)
    MatrixRotationY.rc33 = sngCos(angle)
   
End Function

Public Function MatrixRotationZ(ByVal angle As Single) As Matrix

    MatrixRotationZ = IdentityMatrix
    MatrixRotationZ.rc11 = sngCos(angle)
    MatrixRotationZ.rc21 = sngSin(angle)
    MatrixRotationZ.rc12 = -sngSin(angle)
    MatrixRotationZ.rc22 = sngCos(angle)
   
End Function

Public Function WorldMatrix() As Matrix
 
    With ObjPart
        WorldMatrix = IdentityMatrix
        WorldMatrix = MatrixMultiply(WorldMatrix, MatrixTranslation(.Position.X, .Position.Y, .Position.Z))
        WorldMatrix = MatrixMultiply(WorldMatrix, MatrixRotationX(.Direction.X))
        WorldMatrix = MatrixMultiply(WorldMatrix, MatrixRotationY(.Direction.Y))
        WorldMatrix = MatrixMultiply(WorldMatrix, MatrixRotationZ(.Direction.Z))
    End With
    
End Function

 
