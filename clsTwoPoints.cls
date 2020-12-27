VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTwoPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IsValid As Boolean

'CX1,CY1,CZ1 are the cartesian coordinates of point 1
'CX2,CY2,CZ2 are the cartesian coordinates of point 2
Public Property Get CX1() As Double
    CX1 = gX1
End Property
Public Property Let CX1(ByVal vNewValue As Double)
    IsValid = (gX1 = vNewValue) And IsValid
    gX1 = vNewValue
End Property

Public Property Get CY1() As Double
    'If Not IsValid Then CartChange
    CX1 = gY1
End Property
Public Property Let CY1(ByVal vNewValue As Double)
    IsValid = (gY1 = vNewValue) And IsValid
    gY1 = vNewValue
End Property

Public Property Get CZ1() As Double
    'If Not IsValid Then CartChange
    CZ1 = gZ1
End Property
Public Property Let CZ1(ByVal vNewValue As Double)
    IsValid = (gZ1 = vNewValue) And IsValid
    gZ1 = vNewValue
End Property

Public Property Get CX2() As Double
    'If Not IsValid Then CartChange
    CX2 = gX2
End Property
Public Property Let CX2(ByVal vNewValue As Double)
    IsValid = (gX2 = vNewValue) And IsValid
    gX2 = vNewValue
End Property

Public Property Get CY2() As Double
    'If Not IsValid Then CartChange
    CY2 = gY2
End Property
Public Property Let CY2(ByVal vNewValue As Double)
    IsValid = (gY2 = vNewValue) And IsValid
    gY2 = vNewValue
End Property

Public Property Get CZ2() As Double
    'If Not IsValid Then CartChange
    CZ2 = gZ2
End Property
Public Property Let CZ2(ByVal vNewValue As Double)
    IsValid = (gZ2 = vNewValue) And IsValid
    gZ2 = vNewValue
End Property

Public Property Get Alpha() As Double
    If Not IsValid Then CartChange
    Alpha = gAlpha
End Property

Public Property Get Beta() As Double
    If Not IsValid Then CartChange
    Beta = gBeta
End Property

Public Property Get Gamma() As Double
    If Not IsValid Then CartChange
    Gamma = gGamma
End Property

Public Property Get Mag() As Double
    If Not IsValid Then CartChange
    Mag = gMag
End Property

Private Sub CartChange()
    'Cartesian coordinates changed
    'given x1,y1,z1 and x2,y2,z2 find the distance between points
    Dim XDist As Double
    Dim YDist As Double
    Dim ZDist As Double
    Dim p(0 To 2) As Boolean
    Dim i As Integer
    Dim PFlag As Long
    Dim Pi As Double
    
    Dim CosXa As Double, CosXb As Double
    Dim CosYa As Double, CosYb As Double
    Dim CosZa As Double, CosZb As Double
    
    Pi = 4 * Atn(1)
    
    For i = 0 To 2
        p(i) = 1
    Next i
    
    XDist = gX2 - gX1
    YDist = gY2 - gY1
    ZDist = gZ2 - gZ1
    
    gMag = Sqr(XDist ^ 2 + YDist ^ 2 + ZDist ^ 2)
    
    p(0) = XDist < 0
    p(1) = YDist < 0
    p(2) = ZDist < 0
    
    If p(0) Then PFlag = 1
    If p(1) Then PFlag = PFlag + 2
    If p(2) Then PFlag = PFlag + 4
    
    'Now calculate the direction angles (in radians)
    If gMag = 0 Then
        gAlpha = 0
        gBeta = 0
        gGamma = 0
    Else
        'gAlpha is the angle between the X-axis and the current line
        gAlpha = ACos(XDist / gMag)
        gBeta = ACos(YDist / gMag)
        gGamma = ACos(ZDist / gMag)
        
        Select Case PFlag
            Case 0 '+++ octant
                'gAlpha = gAlpha
                'gBeta = gBeta
                'gGamma = gGamma
            Case 1 '-++ octant
                'gAlpha = gAlpha
                'gGamma = gGamma
                gBeta = ((Pi / 2) - gBeta) + (3 * Pi / 2)
            Case 2 '+-+ octant
                'gAlpha = gAlpha
                'gBeta = gBeta
                'gGamma = gGamma
            Case 3 '--+ octant
                'gAlpha = gAlpha
                gBeta = ((Pi / 2) - gBeta) + (3 * Pi / 2)
                'gGamma = gGamma
            Case 4 '++- octant
                'gAlpha = gAlpha
                'gBeta = gBeta
                'gGamma = gGamma
            Case 5 '-+- octant
                'gAlpha = gAlpha
                'gBeta = gBeta
                'gGamma = gGamma
            Case 6 '+-- octant
                'gAlpha = gAlpha
                'gBeta = gBeta
                'gGamma = gGamma
            Case 7 '--- octant
                'gAlpha = gAlpha
                'gBeta = gBeta
                'gGamma = gGamma
        End Select
    End If
    
    IsValid = True
    
End Sub

Private Function ACos(X As Double) As Double
    'return the Arccosine of angle X
    
    'Arccos(X) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    
    'X is in Radians
    Dim SqrX As Double
    
    Select Case X
        Case 0
            ACos = 2 * Atn(1)
        Case -1
            ACos = 4 * Atn(1)
        Case 1
            ACos = 0
        Case Else
            SqrX = Sqr(1 - (X * X))
            'ACos = Atn(-X / SqrX) + 2 * Atn(1)
            ACos = (2 * Atn(1)) - Atn(X / SqrX)
    End Select
    
End Function

Private Sub Class_Initialize()
    IsValid = False
End Sub
