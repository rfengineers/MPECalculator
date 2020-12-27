VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lAntName As String
Private lAntMaxGain As Double
Private lAntHGain(0 To 359) As Double
Private lAntVGain(0 To 180) As Double

Private lMaxERP As Double
Private lXPosn As Double, lYPosn As Double, lZPosn As Double
Private lName As String
Private lFreq As Single
Private lDutyCycle As Single

Public Property Get AntName() As String
    AntName = lAntName
End Property
Public Property Get AntMaxGain() As String
    AntMaxGain = lAntMaxGain
End Property
Public Property Get AntVerGain(Elevation As Integer) As Double
    AntVerGain = lAntVGain(Elevation)
End Property
Public Property Get AntHorGain(Azimuth As Integer) As Double
    AntHorGain = lAntHGain(Azimuth)
End Property
'Private Sub LoadAntenna(AntDat As AntennaFile_Type)
'    'set the antenna parameter values
'    'convert pattern values to dB
'    Dim i As Integer
'    lAntName = AntDat.Name
'    lAntMaxGain = AntDat.MaxGain
'    If AntDat.Format = "LOG" Then
'        For i = 0 To 359
'            lAntHGain(i) = AntDat.HGain(i)
'        Next i
'        For i = 0 To 180
'            lAntVGain(i) = AntDat.VGain(i)
'        Next i
'    ElseIf AntDat.Format = "FIELD" Then
'        For i = 0 To 359
'            lAntHGain(i) = 20 * Log10(AntDat.HGain(i))
'        Next i
'        For i = 0 To 180
'            lAntVGain(i) = 20 * Log10(AntDat.VGain(i))
'        Next i
'    Else 'the only other possibility is "POWER"
'        For i = 0 To 359
'            lAntHGain(i) = 10 * Log10(AntDat.HGain(i))
'        Next i
'        For i = 0 To 180
'            lAntVGain(i) = 10 * Log10(AntDat.VGain(i))
'        Next i
'    End If
'End Sub

Public Property Get MaxERP() As Double
    MaxERP = lMaxERP
End Property
Public Property Let MaxERP(ByVal vNewValue As Double)
    lMaxERP = vNewValue
End Property

Public Property Get Name() As String
    Name = lName
End Property
Public Property Let Name(ByVal vNewValue As String)
    lName = vNewValue
End Property

Public Property Get XPosn() As Double
    XPosn = lXPosn
End Property
Public Property Let XPosn(ByVal vNewValue As Double)
    lXPosn = vNewValue
End Property

Public Property Get YPosn() As Double
    YPosn = lYPosn
End Property
Public Property Let YPosn(ByVal vNewValue As Double)
    lYPosn = vNewValue
End Property

Public Property Get ZPosn() As Double
    ZPosn = lZPosn
End Property
Public Property Let ZPosn(ByVal vNewValue As Double)
    lZPosn = vNewValue
End Property

Public Property Get Freq() As Single
    Freq = lFreq
End Property

Public Property Let Freq(ByVal vNewValue As Single)
    lFreq = vNewValue
End Property

Public Property Get DutyCycle() As Single
    DutyCycle = lDutyCycle
End Property

Public Property Let DutyCycle(ByVal vNewValue As Single)
    lDutyCycle = vNewValue
End Property

Public Function uWperCM2(X As Double, Y As Double, Z As Double) As Double
    'return power density for this emitter at point x,y,z
    Dim TP As New clsTwoPoints
    Dim lERP As Double
    Dim H As Integer, V As Integer
    Dim HdB As Double, VdB As Double
    Dim S As Double
    Dim GB As Double 'Antenna Gain (combined V and H)
    Dim PDens As Double 'Power Density
    
    With TP
        'location of the emitter
        .CX1 = lXPosn
        .CY1 = lYPosn
        .CZ1 = lZPosn
        'xy location of the point to test
        .CX2 = X
        .CY2 = Y
        
        .CZ2 = lZPosn 'delta Z is zero, so the alpha angle is in the XY Plane
        H = CInt(.Alpha)
        If H > 359 Then H = 0
        .CZ2 = Z 'reset elevation angle
        
        .CY2 = lYPosn 'delta Y is zero, so the gamma angle is in the XZ Plane
        V = CInt(.Gamma)
        If V > 359 Then V = 0
        .CY2 = Y 'reset azimuth angle
        
        HdB = lAntHGain(H) 'horizontal gain in dB
        VdB = lAntVGain(V) 'vertical gain in dB(must be zero or less!)
        GB = (HdB + VdB) / 10
        
        lERP = lMaxERP * lDutyCycle * (10 ^ GB)
        If .Mag > 0 Then
            PDens = (33.4 * lERP) / (.Mag ^ 2)
        Else
            PDens = lERP
        End If
    
        uWperCM2 = PDens
        
    End With
        
End Function







