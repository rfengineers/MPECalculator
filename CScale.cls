VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mMaxScale As Single, mMinScale As Single
Dim mMinValue As Long, mMaxValue As Long
Dim mZeroResult As Long
Dim ElevationColors() As Long

Private Function ColorEquation(ByVal Value As Single) As Single
    
    ColorEquation = Sqr(Value - mMinValue)

End Function
Private Function ElevationToScale(ByVal Value As Long) As Integer
    
    Dim Value2 As Single
    Dim iValue As Integer
    If Value < mMinValue Then Value = mMinValue
    If Value > mMaxValue Then Value = mMaxValue
    Value2 = ColorEquation(Value)
    iValue = Int((Value2 - mMinScale) / (mMaxScale - mMinScale) * 640)
    ElevationToScale = iValue

End Function
Public Function getBlue(ByVal Value As Long) As Integer
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getBlue = (mZeroResult And &H1FE0000) / &H10000
    Else
        If Value < mMinValue Then Value = mMinValue
        If Value > mMaxValue Then Value = mMaxValue
        getBlue = (ElevationColors(Value) And &H1FE0000) / &H10000
    End If

End Function
Public Function getColor(ByVal Value As Long) As Long
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getColor = mZeroResult
    Else
        If Value < mMinValue Then Value = mMinValue
        If Value > mMaxValue Then Value = mMaxValue
        getColor = ElevationColors(Value)
    End If

End Function
Public Function getGrayColorD(ByVal Value As Long) As Long
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getGrayColorD = mZeroResult
        Exit Function
    End If
    Dim comp As Integer
    comp = getGrayComponentD(Value)
    getGrayColorD = RGB(comp, comp, comp)

End Function
Public Function getGrayColorI(ByVal Value As Long) As Long
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getGrayColorI = mZeroResult
        Exit Function
    End If
    Dim comp As Integer
    comp = getGrayComponentI(Value)
    getGrayColorI = RGB(comp, comp, comp)

End Function
Public Function getGrayComponentD(ByVal Value As Long) As Integer
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getGrayComponentD = mZeroResult And &H10000FE
        Exit Function
    End If
    getGrayComponentD = 255 - getGrayComponentI(Value)

End Function
Public Function getGrayComponentI(ByVal Value As Long) As Integer
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getGrayComponentI = mZeroResult And &H10000FE
        Exit Function
    End If
    If Value < mMinValue Then Value = mMinValue
    If Value > mMaxValue Then Value = mMaxValue
    getGrayComponentI = ElevationToScale(Value) / 5

End Function
Public Function getGreen(ByVal Value As Long) As Integer
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getGreen = (mZeroResult And &H100FE00) / &H100
        Exit Function
    End If
    If Value < mMinValue Then Value = mMinValue
    If Value > mMaxValue Then Value = mMaxValue
    getGreen = (ElevationColors(Value) And &H100FE00) / &H100

End Function
Public Function getMax() As Long
    
    getMax = mMaxValue

End Function
Public Function getMin() As Long
    
    getMin = mMinValue

End Function
Public Function getRed(ByVal Value As Long) As Integer
    
    If Value = -9999 Or Value = -500 Or Value = 0 Then
        getRed = mZeroResult And &H10000FE
        Exit Function
    End If
    If Value < mMinValue Then Value = mMinValue
    If Value > mMaxValue Then Value = mMaxValue
    getRed = ElevationColors(Value) And &H10000FE

End Function
Public Sub Init(ByVal MinValue As Long, _
                ByVal MaxValue As Long, _
                ByVal ZeroResult As Long)
    
    
    mMinValue = MinValue
    mMaxValue = MaxValue
    mZeroResult = ZeroResult
    mMinScale = ColorEquation(mMinValue)
    mMaxScale = ColorEquation(mMaxValue)
    ReDim ElevationColors(mMinValue To mMaxValue) As Long
    Dim i As Integer
    For i = mMinValue To mMaxValue
        ElevationColors(i) = ScaleToColor(ElevationToScale(i))
    Next i

End Sub
Private Function ScaleToColor(ByVal Value As Long) As Long
    
    Dim iRed As Integer
    Dim iGreen As Integer
    Dim iBlue As Integer
    
    If Value < 128 Then
        'iRed = 0
        iGreen = Value
        iBlue = 64 - Value / 2
    ElseIf Value < 256 Then
        iRed = Value - 128
        iGreen = 127
        'iBlue = 0
    ElseIf Value < 384 Then
        iRed = 127
        iGreen = 128 - (Value - 256)
        'iBlue = 0
    ElseIf Value < 512 Then
        iRed = 127
        'iGreen = 0
        iBlue = Value - 384
    Else
        iRed = 127
        iGreen = Value - 512
        iBlue = 127
    End If
    ScaleToColor = RGB(iRed, iGreen, iBlue)
        

End Function

Public Sub setZero(ByVal ZeroResult As Long)
    
    mZeroResult = ZeroResult

End Sub

