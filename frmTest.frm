VERSION 5.00
Begin VB.Form frmMPECalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MPE Calculator:"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame20 
      Caption         =   "MultiStation Bitmaps:"
      Height          =   3135
      Left            =   6120
      TabIndex        =   67
      Top             =   2520
      Width           =   4695
      Begin VB.Frame Frame23 
         Caption         =   "Position Map"
         Height          =   1695
         Left            =   1680
         TabIndex        =   88
         Top             =   1440
         Width           =   1335
         Begin VB.Frame Frame25 
            Caption         =   "Horizontal:"
            Height          =   735
            Left            =   0
            TabIndex        =   93
            Top             =   960
            Width           =   1335
            Begin VB.TextBox tbxHOffset 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   480
               TabIndex        =   95
               Text            =   "0"
               Top             =   240
               Width           =   495
            End
            Begin VB.CommandButton cmdHPolarity 
               Caption         =   "E"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "m"
               Height          =   195
               Left            =   1080
               TabIndex        =   96
               Top             =   240
               Width           =   120
            End
         End
         Begin VB.Frame Frame24 
            Caption         =   "Vertical:"
            Height          =   735
            Left            =   0
            TabIndex        =   89
            Top             =   240
            Width           =   1335
            Begin VB.CommandButton cmdVPolarity 
               Caption         =   "N"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox tbxVOffset 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   480
               TabIndex        =   90
               Text            =   "0"
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "m"
               Height          =   195
               Left            =   1080
               TabIndex        =   91
               Top             =   240
               Width           =   120
            End
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Select Maps"
         Height          =   1695
         Left            =   120
         TabIndex        =   81
         Top             =   1440
         Width           =   1455
         Begin VB.OptionButton optRange 
            Caption         =   "1 to 5%"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optRange 
            Caption         =   "2 to 12%"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   86
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optRange 
            Caption         =   "5 to 25%"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   85
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optRange 
            Caption         =   "10 to 50%"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   84
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton optRange 
            Caption         =   "20 to 100%"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   83
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton optRange 
            Caption         =   "Make All"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   82
            Top             =   1440
            Width           =   1095
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Make Bitmap"
         Height          =   1695
         Left            =   3120
         TabIndex        =   77
         Top             =   1440
         Width           =   1455
         Begin VB.CommandButton cmdMakeBMP 
            Caption         =   "Public BMP"
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdKey 
            Caption         =   "Make Key"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdMakeOccBMP 
            Caption         =   "Occu BMP"
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdFromFile 
         Caption         =   "Add from File"
         Height          =   375
         Left            =   1800
         TabIndex        =   74
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdInit 
         Caption         =   "Clear/Init"
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox tbxCmpP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   71
         Text            =   "100"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbxPixels 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   69
         Text            =   "1000"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdAddStation 
         Caption         =   "Add Station"
         Height          =   375
         Left            =   1800
         TabIndex        =   68
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "cm/Pixel"
         Height          =   195
         Left            =   960
         TabIndex        =   72
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Pixels"
         Height          =   255
         Left            =   960
         TabIndex        =   70
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdBMP 
      Caption         =   "single BMP"
      Height          =   375
      Left            =   9720
      TabIndex        =   66
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   9720
      TabIndex        =   61
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   9720
      TabIndex        =   60
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame16 
      Caption         =   "Antenna Pattern:"
      Height          =   2415
      Left            =   8040
      TabIndex        =   57
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton cmdAzPat0 
         Caption         =   "Zero Az"
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdVerPat0 
         Caption         =   "Zero El"
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdVerPat 
         Caption         =   "Elevation"
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAzPat 
         Caption         =   "Azimuth"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "Scan Output:"
      Height          =   2415
      Left            =   3360
      TabIndex        =   54
      Top             =   2520
      Width           =   2775
      Begin VB.CommandButton cmdScanX 
         Caption         =   "Scan X"
         Height          =   375
         Left            =   1560
         TabIndex        =   65
         Top             =   720
         Width           =   1095
      End
      Begin VB.Frame Frame19 
         Caption         =   "Steps  :  Incr"
         Height          =   615
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   1335
         Begin VB.TextBox tbxScanX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   64
            Text            =   "25"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox tbxScanIncrX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            MaxLength       =   3
            TabIndex        =   63
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdScanZ 
         Caption         =   "Scan Z"
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Frame Frame18 
         Caption         =   "Steps  :  Incr"
         Height          =   615
         Left            =   120
         TabIndex        =   59
         Top             =   1800
         Width           =   1335
         Begin VB.TextBox tbxScanZ 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   13
            Text            =   "25"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox tbxScanIncrZ 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Steps  :  Incr"
         Height          =   615
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   1335
         Begin VB.TextBox tbxScanIncrXY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox tbxScanXY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   11
            Text            =   "25"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdScanXY 
         Caption         =   "Scan XY"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdRasterGP 
         Caption         =   "RasterGP"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRasterOcc 
         Caption         =   "RasterOcc"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Source ID:"
      Height          =   615
      Left            =   120
      TabIndex        =   53
      Top             =   120
      Width           =   3135
      Begin VB.TextBox tbxSourceID 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   9720
      TabIndex        =   21
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   615
      Left            =   9720
      TabIndex        =   22
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame13 
      Caption         =   "Position in Meters:"
      Height          =   2415
      Left            =   120
      TabIndex        =   44
      Top             =   840
      Width           =   3135
      Begin VB.Frame Frame8 
         Caption         =   "Target Height"
         Height          =   615
         Left            =   1680
         TabIndex        =   50
         Top             =   1680
         Width           =   1335
         Begin VB.TextBox TbxZ2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Target North"
         Height          =   615
         Left            =   1680
         TabIndex        =   49
         Top             =   240
         Width           =   1335
         Begin VB.TextBox tbxY2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Target East"
         Height          =   615
         Left            =   1680
         TabIndex        =   48
         Top             =   960
         Width           =   1335
         Begin VB.TextBox tbxX2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Source East"
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   1335
         Begin VB.TextBox tbxX1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Source North"
         Height          =   615
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1335
         Begin VB.TextBox tbxY1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   1
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Source Height"
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Width           =   1335
         Begin VB.TextBox tbxZ1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Elevation Gain:"
      Height          =   615
      Left            =   1800
      TabIndex        =   41
      Top             =   3360
      Width           =   1455
      Begin VB.TextBox tbxVGain 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dB"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Azimuth Gain:"
      Height          =   615
      Left            =   1800
      TabIndex        =   40
      Top             =   3960
      Width           =   1455
      Begin VB.TextBox tbxAzGain 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dB"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame10 
      Height          =   2415
      Left            =   3360
      TabIndex        =   27
      Top             =   120
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Adj ERP:"
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   2415
         Begin VB.Label lblAdjERP 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "% Occupational Exposure:"
         Height          =   855
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   2415
         Begin VB.Label lblPctOcc 
            Caption         =   "0"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "% General Public Exposure:"
         Height          =   855
         Index           =   5
         Left            =   2640
         TabIndex        =   36
         Top             =   1440
         Width           =   1815
         Begin VB.Label lblPctGP 
            Caption         =   "0"
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Power Density:"
         Height          =   615
         Index           =   4
         Left            =   2640
         TabIndex        =   34
         Top             =   840
         Width           =   1815
         Begin VB.Label lblPwrDens 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Azimuth:"
         Height          =   615
         Index           =   0
         Left            =   1440
         TabIndex        =   32
         Top             =   240
         Width           =   1095
         Begin VB.Label lblAzimuth 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Elevation:"
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
         Begin VB.Label lblGamma 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Distance:"
         Height          =   615
         Left            =   2640
         TabIndex        =   28
         Top             =   240
         Width           =   1815
         Begin VB.Label lblMag 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frequency (MHz)"
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   1575
      Begin VB.TextBox tbxFreq 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "100"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Max ERP (Watts)"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   1575
      Begin VB.TextBox tbxERP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "250"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Beta"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   8040
      Width           =   1455
      Begin VB.Label lblBeta 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "MPE Calculator by Joseph M. DiPietro, joed@rfengineers.com. "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   56
      Top             =   5040
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright 2004 by rfEngineers, Inc.  All Rights Reserved."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   55
      Top             =   5280
      Width           =   5175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMPECalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TP As New clsTwoPoints
Private dFreq As Double
Private dGP As Double
Private dERP As Double
Private NetGain As Double
'Private outfile As String
'Private PiOvr180 As Double

Private Type MPEType
    ID As String 'Descriptive string
    SN As Single 'Source North in meters
    SE As Single 'Source East in meters
    SH As Single 'Source Height in meters
    TN As Single 'Target North in meters
    TE As Single 'Target East in meters
    TH As Single 'Target Height in meters
    Freq As Double 'Source Frequency
    VGain As Double 'Antenna vertical gain in dB
    AGain As Double 'Antenna Azimuthal gain in dB
    MERP As Double 'Max ERP (with a gain of zero)
    AERP As Double 'Gain Adjusted ERP
    VAng As Double 'Vertical angle below the tangental plane
    AAng As Double 'Azimuthal angle from True North
    Dist As Double 'Distance, source to target in meters
    Pden As Double 'Power density in mW/cm^2
    Pocc As Double 'Percent Occupational exposure
    Pgp As Double 'Percent Public exposure
    Locc As Double 'Limit of Occupational exposure
    Lgp As Double 'Limit of Public exposure
End Type

'Private AntPat As AntPat_Type

'multi vars
Private MultiiH As Long, MultiiW As Long
Private Multiiexp() As Double
Private MultiiOcc() As Double
Private MultiIncr As Integer

Private MPE As MPEType

Private Sub cmdAzPat_Click()
    Dim HFileName As String
    
    HFileName = GetHPatFileForOpen
    If gFSO.FileExists(HFileName) Then
        If IsHPatFile(HFileName) Then
            Frame11.Enabled = False
            LoadHPat HFileName
            UpdateLabels
        ElseIf LCase(gFSO.GetExtensionName(HFileName)) = "gh" Then
            Frame11.Enabled = False
            LoadGenericHPat HFileName
            UpdateLabels
        Else
            Frame11.Enabled = True
            tbxAzGain.Text = 0
            Erase HGain
        End If
    Else
            Frame11.Enabled = True
            tbxAzGain.Text = 0
            Erase HGain
    End If
End Sub

Private Sub SilentUpdate()
    Dim dTmp As Double
    Dim Ang As Double
    Dim PDens As Double
    Dim adjERP As Double
    Dim iAng As Integer
    
    With TP
        'reposition the Target
        .CY2 = MPE.TN
        .CX2 = MPE.TE
    
        dTmp = .CZ2
            .CZ2 = .CZ1 'Make delta Z zero, so the azimuth angle is in the XY Plane
            Ang = CDbl(.Beta) / PiOvr180
            If Ang > 359.5 Then Ang = Ang - 360
            MPE.AAng = Ang
'            lblAzimuth.Caption = Format(Ang, "##0.0")
        .CZ2 = dTmp 'reset elevation angle
        
        Ang = CDbl(.Gamma) / PiOvr180
        MPE.VAng = 90 - Ang
'        lblGamma.Caption = Format(MPE.VAng, "##0.0")
        
        MPE.Dist = .Mag
'        lblMag.Caption = Format(.Mag, "#,##0.00 meters")
        
'        If Not Frame12.Enabled Then 'read elevation gain from array
            iAng = CInt(MPE.VAng)
            MPE.VGain = VGain(iAng)
'        End If
        
'        If Not Frame11.Enabled Then
            iAng = CInt(MPE.AAng)
            MPE.AGain = HGain(iAng)
'        End If
        
        NetGain = MPE.AGain + MPE.VGain
        
        If NetGain <> 0 Then
            MPE.AERP = (10 ^ (NetGain / 10)) * MPE.MERP
        Else
            MPE.AERP = MPE.MERP
        End If
        
        On Error Resume Next
            MPE.Pden = ((33.4 * MPE.AERP) / (.Mag ^ 2)) / 1000
        On Error GoTo 0
'        lblAdjERP.Caption = Format(MPE.AERP, "###,###,##0.### Watts")
    End With
'    MPE.Freq = CSng(tbxFreq.Text)
'    lblPwrDens.Caption = Format(MPE.Pden * 1000, "#,##0.0#### uW/cm^2")
    MPE.Pocc = MPE.Pden / MPE.Locc
'    lblPctOcc.Caption = Format(MPE.Pocc, "##0.00%") & " of " & Format(MPE.Locc, "##0.00# mW/cm^2")
    MPE.Pgp = MPE.Pden / MPE.Lgp
'    lblPctGP.Caption = Format(MPE.Pgp, "##0.00%") & " of " & Format(MPE.Lgp * 1000, "#,##0.00 uW/cm^2")
    
End Sub


Private Sub cmdAzPat0_Click()
    Frame11.Enabled = True
    tbxAzGain.Text = 0
    Erase HGain

End Sub

Private Sub cmdBMP_Click()

    Dim iH As Long, iW As Long
    Dim iexp() As Double
    Dim i As Long, j As Long
    Dim Incr As Integer
    Dim OffSet As Long
    Dim MinPct As Double, MinExp As Double
    Dim MaxPct As Double, MaxExp As Double
    Dim MaxX As Long, MaxY As Long
    
    'boundaries
    Dim bNorth As Double, bSouth As Double, bEast As Double, bWest As Double
    'area of analysis in pixels
    iH = 1000
    iW = 1000
    ReDim iexp(1 To iW, i To iH)

    Incr = 1
    '20 = 5cm steps to 25 meters
    '10 = 10cm steps to 50 meters
    '1 = 1 meter steps to 500-meters
    '0.800 = 1.25 meter steps to 625 meters
    '0.6667 = 1.5 meter steps to 750 meters
    '0.5 = 2-meter steps to 1000-meters
    
    OffSet = 500 'to center the source in the bitmap
    'meters/incr north, south, east and west of source
    bNorth = 500
    bSouth = -499
    bEast = 500
    bWest = -499
    'init min's and maxes
    MinPct = 100
    MaxPct = 0
    MinExp = 100000
    MaxExp = 0
    
    'set the source location
    MPE.SH = CDbl(tbxZ1.Text)
    TP.CZ1 = MPE.SH
    MPE.SE = CDbl(tbxX1.Text)
    TP.CX1 = MPE.SE
    MPE.SN = CDbl(tbxY1.Text)
    TP.CY1 = MPE.SN
        
    'get save-to filename
    outfile = GetbmpFileForSave
    
    Screen.MousePointer = vbHourglass

        If Len(outfile) > 0 Then
                
            For i = bSouth To bNorth 'south to north
                MPE.TN = CSng(i / Incr)
                For j = bWest To bEast 'west to east
                    MPE.TE = CSng(j / Incr)
                    'update calculations
                    SilentUpdate
                    'save percent Public to array
                    iexp(j + OffSet, i + OffSet) = MPE.Pgp * 100
                    'update min/max values
                    If MPE.Pgp > MaxPct Then
                        MaxPct = MPE.Pgp
                        MaxExp = MPE.Pden
                        MaxX = i
                        MaxY = j
                    End If
                    If MPE.Pgp < MinPct Then
                        MinPct = MPE.Pgp
                        MinExp = MPE.Pden
                    End If
                    DoEvents
                Next j
            Next i
            
            
'BackHere:
            'make a bitmap from the mpe.pgp
            MakeExposureBMP iH, iW, outfile, iexp(), CLng(MinExp), CLng(MaxExp * 6000)
            
'            Stop
'            GoTo BackHere
            
        End If
    
        Dim TS As TextStream
        
        Set TS = gFSO.OpenTextFile(outfile & "Info.txt", ForWriting, True)
        
            TS.writeline "RFR Analysis of " & MPE.ID
            TS.writeline "Created " & Format(Now, "ddmmmyyyy")
            TS.WriteBlankLines 1
            TS.writeline "The RF source was centered in the analysis at " & MPE.SH & " meters AGL."
            TS.writeline "The maximum ERP was set to " & Format(MPE.MERP, "##,###,##0.0") & " Watts."
            TS.writeline "Calculations were made out " & bNorth / Incr & " meters."
            TS.WriteBlankLines 1
            TS.writeline "The highest exposure calculated was " & Format(MaxExp * 1000, "###,##0.00 uW/cm^2")
            TS.writeline "Which represents " & Format(MaxPct, "##0.0#%") & " of the maximum permited Public exposure for " & Format(MPE.Freq, "##,###.###") & "MHz."
            TS.writeline "This point was calculated at " & MaxX / Incr & " meters north and " & MaxY / Incr & " meters east of the source."
        TS.Close
    Screen.MousePointer = vbNormal


End Sub

Private Sub cmdClose_Click()
    Dim f As Form
    
    For Each f In Forms
        Unload f
    Next f
End Sub

Private Sub cmdFromFile_Click()
    Dim infile As String
    Dim TS As TextStream
    Dim inLine As String
    Dim fields() As String
    Dim i As Integer
    Dim ic As Integer
    Dim peak As Single
    Dim VTilt As Integer, HRotate As Integer
    Dim Posn As Single
    
    Screen.MousePointer = vbHourglass
    
    infile = GetStationFileForOpen
    
    If infile = "" Then Exit Sub
    
    '  0    1    2     3     4    5    6    7    8     9
    'Name|North|East|Height|ERP|Freq|AntV|AntH|VTilt|HRotate
    
    Set TS = gFSO.OpenTextFile(infile, ForReading)
    
    Do
        inLine = TS.ReadLine
        If Left$(Trim(inLine), 1) <> ";" Then
            Erase fields
            fields() = Split(inLine, "|")
            ic = UBound(fields)
            If ic >= 7 Then
                tbxSourceID = fields(0)
                
                Posn = CSng(val(tbxHOffset.Text))
                If Posn <> 0 Then
                    If cmdHPolarity.Caption = "E" Then
                        Posn = -Posn
                    End If
                End If
                tbxX1.Text = CStr(val(fields(2)) + Posn)
                
                Posn = CSng(val(tbxVOffset.Text))
                If Posn <> 0 Then
                    If cmdHPolarity.Caption = "N" Then
                        Posn = -Posn
                    End If
                End If
                tbxY1.Text = CStr(val(fields(1)) + Posn)
                
                tbxZ1.Text = fields(3)
                tbxERP.Text = fields(4)
                tbxFreq.Text = fields(5)
                'these are optional
                If ic >= 8 Then
                    VTilt = CInt(val(fields(8)))
                Else
                    VTilt = 0
                End If
                If ic >= 9 Then
                    HRotate = CInt(val(fields(9)))
                Else
                    HRotate = 0
                End If
                
                If gFSO.FileExists(fields(6)) Then
                    If IsVPatFile(fields(6)) Then
                        Frame12.Enabled = False
                        LoadVPat fields(6), VTilt
                        UpdateLabels
                    ElseIf LCase(gFSO.GetExtensionName(fields(6))) = "gv" Then
                        Frame12.Enabled = False
                        LoadGenericVPat fields(6), VTilt
                    Else
                        Frame12.Enabled = True
                        tbxVGain.Text = 0
                        Erase VGain
                    End If
                Else
                    Frame12.Enabled = True
                    tbxVGain.Text = 0
                    Erase VGain
                End If
    
                If gFSO.FileExists(fields(7)) Then
                    If IsHPatFile(fields(7)) Then
                        Frame11.Enabled = False
                        LoadHPat fields(7), HRotate
                        UpdateLabels
                    ElseIf LCase(gFSO.GetExtensionName(fields(7))) = "gh" Then
                        Frame11.Enabled = False
                        LoadGenericHPat fields(7), HRotate
                    Else
                        Frame11.Enabled = True
                        tbxAzGain.Text = 0
                        Erase HGain
                    End If
                Else
                        Frame11.Enabled = True
                        tbxAzGain.Text = 0
                        Erase HGain
                End If
                peak = AutoAddStation(fields(0))
                
                'Debug.Print fields(0) & "Max = " & peak
            End If
        End If 'comment
    Loop While Not TS.AtEndOfStream
    
    Screen.MousePointer = vbNormal
    
    MsgBox "Complete"
    
End Sub

Private Sub cmdAddStation_Click()
    Dim i As Long, j As Long
    Dim OffSet As Long
    Dim MinPct As Double, MinExp As Double
    Dim MaxPct As Double, MaxExp As Double
    Dim MaxX As Long, MaxY As Long
    
    'boundaries
    Dim bNorth As Double, bSouth As Double, bEast As Double, bWest As Double

    OffSet = MultiiH / 2 'to center the source in the bitmap
    'meters/incr north, south, east and west of source
    bNorth = OffSet
    bSouth = -OffSet + 1
    bEast = OffSet
    bWest = -OffSet + 1
    
    'set the source location
    MPE.SH = CDbl(tbxZ1.Text)
    TP.CZ1 = MPE.SH
    MPE.SE = CDbl(tbxX1.Text)
    TP.CX1 = MPE.SE
    MPE.SN = CDbl(tbxY1.Text)
    TP.CY1 = MPE.SN
    
    'black dot at the source location
    On Error Resume Next
        Multiiexp((MultiIncr * MPE.SE) + OffSet, (MultiIncr * MPE.SN) + OffSet) = -999
        MultiiOcc((MultiIncr * MPE.SE) + OffSet, (MultiIncr * MPE.SN) + OffSet) = -999
    On Error GoTo 0
    
    Screen.MousePointer = vbHourglass
        For i = bSouth To bNorth 'south to north
            MPE.TN = CSng(i / MultiIncr)
            For j = bWest To bEast 'west to east
                MPE.TE = CSng(j / MultiIncr)
                'update calculations
                SilentUpdate
                'save percent to arrays
                If Multiiexp(j + OffSet, i + OffSet) <> -999 Then
                    Multiiexp(j + OffSet, i + OffSet) = Multiiexp(j + OffSet, i + OffSet) + (MPE.Pgp * 100)
                    MultiiOcc(j + OffSet, i + OffSet) = MultiiOcc(j + OffSet, i + OffSet) + (MPE.Pocc * 100)
                End If
                DoEvents
            Next j
        Next i
    Screen.MousePointer = vbNormal
    

End Sub

Private Sub cmdHPolarity_Click()
    If cmdHPolarity.Caption = "E" Then
        cmdHPolarity.Caption = "W"
    Else
        cmdHPolarity.Caption = "E"
    End If
End Sub

Private Sub cmdKey_Click()
    Dim outfile As String
    Dim Rng As Integer, i As Integer
    For i = 0 To 4
        If optRange(i).Value Then
            Rng = i
            Exit For
        End If
    Next i
    outfile = GetbmpFileForSave

    MakeBMPKey outfile, Rng
End Sub

Private Sub cmdMakeOccBMP_Click()
    Dim Rng As Integer, i As Integer
    Dim oFile As String, oPath As String

    'get save-to filename
    outfile = GetbmpFileForSave
    Screen.MousePointer = vbHourglass
        If Len(outfile) > 0 Then
            'make a bitmap from the mpe.pgp
            For i = 0 To 5
                If optRange(i).Value Then
                    Rng = i
                    Exit For
                End If
            Next i
            If Rng = 5 Then
                oFile = gFSO.GetBaseName(outfile)
                oPath = gFSO.GetParentFolderName(outfile)
                For i = 0 To 4
                    Select Case i
                    Case 0
                        outfile = gFSO.BuildPath(oPath, "1-5pct_" & oFile & ".bmp")
                    Case 1
                        outfile = gFSO.BuildPath(oPath, "2-12pct_" & oFile & ".bmp")
                    Case 2
                        outfile = gFSO.BuildPath(oPath, "5-25pct_" & oFile & ".bmp")
                    Case 3
                        outfile = gFSO.BuildPath(oPath, "10-50pct_" & oFile & ".bmp")
                    Case 4
                        outfile = gFSO.BuildPath(oPath, "20-100pct_" & oFile & ".bmp")
                    End Select
                    MakeMultiExposureBMP MultiiH, MultiiW, outfile, MultiiOcc(), i
                Next i
            Else
                MakeMultiExposureBMP MultiiH, MultiiW, outfile, MultiiOcc(), Rng
            End If
        End If
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdScanZ_Click()
    Dim TS As TextStream
    Dim i As Single
    Dim j As Single
    Dim Steps As Single
    Dim Incr As Single
    Dim fStart As Single, fEnd As Single
    
    'scan up the tower
    If IsNumeric(tbxScanZ.Text) And IsNumeric(tbxScanIncrZ.Text) Then
        Steps = CSng(tbxScanZ.Text)
        Incr = CSng(tbxScanIncrZ.Text)
        exportHeader
        If gFSO.FileExists(outfile) Then
            Set TS = gFSO.OpenTextFile(outfile, ForAppending, False)
                fStart = MPE.TH
                fEnd = MPE.TH + Steps
                For i = fStart To fEnd Step Incr
                    TbxZ2.Text = i
                    TS.Write MPE.ID & Chr(9)
                    TS.Write MPE.SE & Chr(9)
                    TS.Write MPE.SN & Chr(9)
                    TS.Write MPE.SH & Chr(9)
                    TS.Write MPE.TE & Chr(9)
                    TS.Write MPE.TN & Chr(9)
                    TS.Write MPE.TH & Chr(9)
                    TS.Write MPE.Freq & Chr(9)
                    TS.Write MPE.VGain & Chr(9)
                    TS.Write MPE.AGain & Chr(9)
                    TS.Write MPE.MERP & Chr(9)
                    TS.Write MPE.VAng & Chr(9)
                    TS.Write MPE.AAng & Chr(9)
                    TS.Write MPE.Dist & Chr(9)
                    TS.Write MPE.AERP & Chr(9)
                    TS.Write MPE.Pden & Chr(9)
                    TS.Write MPE.Pocc & Chr(9)
                    TS.writeline MPE.Pgp
                    DoEvents
                Next i
            TS.Close
        End If
    End If

End Sub

Private Sub cmdVerPat_Click()
    Dim VFileName As String
    
    VFileName = GetVPatFileForOpen
    If gFSO.FileExists(VFileName) Then
        If IsVPatFile(VFileName) Then
            Frame12.Enabled = False
            If LCase(gFSO.GetExtensionName(VFileName)) = "tsv" Then
                LoadVPat VFileName
            Else
                LoadERIVPat VFileName
            End If
            
        ElseIf LCase(gFSO.GetExtensionName(VFileName)) = "gv" Then
            Frame12.Enabled = False
            LoadGenericVPat VFileName, 0
        Else
            Frame12.Enabled = True
            tbxVGain.Text = 0
            Erase VGain
        End If
    Else
        Frame12.Enabled = True
        tbxVGain.Text = 0
        Erase VGain
    End If
    UpdateLabels

End Sub

Private Sub LoadHPat(FileName As String, Optional OffSet As Integer = 0)
    Dim TS As TextStream
    Dim res As Boolean
    Dim Txt As String
    Dim i As Integer, deg As Integer
    '0 to 359 degrees
    'antenna gain in dB
    'data follows "RFE Antenna Pattern H"
    res = False
    Set TS = gFSO.OpenTextFile(FileName, ForReading)
        Do While Not TS.AtEndOfStream
            Txt = TS.ReadLine
            res = InStr(Txt, "RFE Antenna Pattern H") > 0
            If res Then Exit Do
        Loop
        If res Then
            For i = 0 To 359
                If TS.AtEndOfStream Then
                    Erase HGain
                    Exit For
                End If
                Txt = TS.ReadLine
                deg = i + OffSet
                If deg > 359 Then deg = deg - 359
                If deg < 0 Then deg = deg + 359
                If IsNumeric(Txt) Then
                    HGain(i) = CSng(Txt)
                Else
                    HGain(i) = 0
                End If
            Next i
        End If
    TS.Close
End Sub

Private Sub LoadVPat(FileName As String, Optional tilt As Integer = 0)
    Dim TS As TextStream
    Dim res As Boolean
    Dim Txt As String
    Dim i As Integer, deg As Integer
    
    res = False
    Set TS = gFSO.OpenTextFile(FileName, ForReading)
        Do While Not TS.AtEndOfStream
            Txt = TS.ReadLine
            res = InStr(Txt, "RFE Antenna Pattern V") > 0
            If res Then Exit Do
        Loop
        If res Then
            For i = 90 To -90 Step -1
                If TS.AtEndOfStream Then
                    Erase VGain
                    Exit For
                End If
                Txt = TS.ReadLine
                deg = i + tilt 'factor in the tilt
                If i < 91 And i > -91 Then
                    If IsNumeric(Txt) Then
                        VGain(i) = CSng(Txt)
                    Else
                        VGain(i) = 0
                    End If
                End If
            Next i
        End If
    TS.Close
End Sub

Private Sub LoadERIVPat(FileName As String)
    Dim TS As TextStream
    Dim res As Boolean
    Dim Txt As String
    Dim i As Integer
    Dim Vdata() As String
    
    res = False
    Set TS = gFSO.OpenTextFile(FileName, ForReading)
        'scan until we find the start of data
        Do While Not TS.AtEndOfStream
            Txt = TS.ReadLine
            res = InStr(Txt, "Elevation       Field") > 0
            If res Then Exit Do
        Loop
        If res Then
            For i = 90 To -90 Step -1
                VGain(i) = -1
            Next i
            
            Do While Not TS.AtEndOfStream
                Txt = TS.ReadLine
                Vdata = Split(Txt, Chr(9))
                Txt = Vdata(0)
                Txt = Left$(Txt, Len(Txt) - 1)
                
                If IsNumeric(Txt) And IsNumeric(Vdata(2)) Then
                    'and convert the field values to dB
                    'If CInt(Txt) = 42 Then Stop
                    If val(Vdata(2)) = 0 Then
                        Vdata(2) = 0.001
                    End If
                    i = CInt(Txt)
                    VGain(i) = 20 * Log10(CSng(Vdata(2)))
                    If VGain(i) < -30 Then VGain(i) = -30
                End If
            
            Loop
            'fill in the back side of the pattern
            For i = 90 To -90 Step -1
                If VGain(i) = -1 Then
                    VGain(i) = VGain(-i)
                End If
            Next i
        
        End If
    TS.Close
End Sub

Public Sub LoadGenericVPat(FileName As String, Optional tilt As Integer = 0)
    Dim TS As TextStream
    Dim res As Boolean
    Dim Txt As String
    Dim i As Integer, j As Integer
    Dim Vdata() As String
    Dim VField(-90 To 90) As Single
    Dim Gflag As Boolean
    Dim deg As Integer
    
    'load a vertical pattern file where the data is comma delimited
    'elevation,Field <CRLF>
    '
    'The data starts with "DATA START"
    'The Data ends with an azimuth and field value of zero
    'or end of file.
    
    res = False
    Set TS = gFSO.OpenTextFile(FileName, ForReading)
        'scan until we find the start of data
        Do While Not TS.AtEndOfStream
            Txt = UCase(TS.ReadLine)
            res = InStr(Txt, "DATA START") > 0
            If res Then Exit Do
        Loop
        If res Then
            'initialize the data array to -1
            For i = 90 To -90 Step -1
                VGain(i) = -1
                VField(i) = -1
            Next i
            
            Gflag = True
            Do While (Not TS.AtEndOfStream) And Gflag
                'next line
                Txt = TS.ReadLine
                If Left$(Txt, 1) <> "#" Then 'skip lines that begin with "#"
                    Vdata = Split(Txt, ",") 'Chr(9))
                    If IsNumeric(Vdata(0)) And IsNumeric(Vdata(1)) Then
                        If val((Vdata(0)) = 0 And _
                           val(Vdata(1) = 0)) Then _
                           Gflag = False
                        If (val(CInt(Vdata(0))) < -90 Or _
                            val(CInt(Vdata(0))) > 90) Then _
                            Gflag = False
                        
                        If val(Vdata(1)) <= 0 Then
                            Vdata(1) = 0.001
                        ElseIf val(Vdata(1) > 1#) Then
                            Vdata(1) = 1#
                        End If
                        If CInt(Vdata(0)) = val(Vdata(0)) Then
                            'everything seems to be within valid ranges!
                            'apply the tilt
                            deg = CInt(Vdata(0)) + tilt
                            If deg > -91 And deg < 91 Then
                                VField(deg) = CSng(Vdata(1))
                            End If
                        End If
                    Else
                        Gflag = False
                    End If
                End If
            Loop
            
            'flesh out the pattern
            FillOutVertPat VField()
            
            For i = -90 To 90
                
                'and convert the field values to dB
                'and copy it to the VGain Array
                VGain(i) = 20 * Log10(CDbl(VField(i)))
                'limit gain to -30 dB
                If VGain(i) < -30 Then VGain(i) = -30
            Next i
        End If
    TS.Close
End Sub

Private Sub FillOutVertPat(VField() As Single)
    Dim A1 As Double, A2 As Double, A3 As Double
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i As Integer, j As Integer
    Dim fDone As Boolean
    Dim Steps As Integer, perStep As Double
    
    'first fill in the leading and trailing tails.
    A1 = -1
    For i = -90 To 90
        If VField(i) > 0 Then
            A1 = VField(i)
            i1 = i
            Exit For
        End If
    Next i
    If A1 > 0 Then
        'Backfill the first valid value to the -90-deg tail
        For i = i1 - 1 To -90 Step -1
            VField(i) = A1
        Next i
    End If
    A2 = -1
    For i = 90 To -90 Step -1
        If VField(i) > 0 Then
            A2 = VField(i)
            i2 = i
            Exit For
        End If
    Next i
    If A2 > 0 Then
        'Backfill the last valid value to the 90-deg tail
        For i = i2 + 1 To 90
            VField(i) = A2
        Next i
    End If
    'loop through the center until all invalid values have been filled in
    Do
        fDone = True
        For i = i1 To i2
            If VField(i) < 0 Then 'there are still fields to be set!
                fDone = False
                Exit For
            End If
        Next i
        If Not fDone Then
            i1 = i - 1 'new starting point
            For i = i1 + 1 To i2
                If VField(i) > 0 Then 'found ending point
                    Steps = i - i1 + 1
                    perStep = (VField(i) - VField(i1)) / Steps
                    For j = i1 + 1 To i - 1
                        VField(j) = VField(j - 1) + perStep
                    Next j
                    i1 = i 'set new starting point
                    Exit For
                End If
            Next i
        End If
    Loop Until fDone
    
End Sub

Public Sub LoadGenericHPat(FileName As String, Optional Hrot As Integer = 0)
    Dim TS As TextStream
    Dim res As Boolean
    Dim Txt As String
    Dim i As Integer, j As Integer
    Dim Hdata() As String
    Dim HField(0 To 359) As Single
    Dim Gflag As Boolean
    Dim deg As Integer
    'load a Horizontal pattern file where the data is comma delimited
    'azimuth,Field <CRLF>
    '
    'The data starts with "DATA START"
    'The Data ends with an azimuth and field value of zero
    'or end of file.
    'insure that rotation is between 0 and 359
    Do While Hrot > 359
        Hrot = Hrot - 360
    Loop
    Do While Hrot < 0
        Hrot = Hrot + 360
    Loop
    res = False
    Set TS = gFSO.OpenTextFile(FileName, ForReading)
        'scan until we find the start of data
        Do While Not TS.AtEndOfStream
            Txt = UCase(TS.ReadLine)
            res = InStr(Txt, "DATA START") > 0
            If res Then Exit Do
        Loop
        If res Then
            'initialize the data array to -1
            For i = 0 To 359 'Step -1
                HGain(i) = -1
                HField(i) = -1
            Next i
            
            Gflag = True
            Do While (Not TS.AtEndOfStream) And Gflag
                'next line
                Txt = TS.ReadLine
                If Left$(Txt, 1) <> "#" And Trim(Txt) <> "" Then 'skip lines that begin with "#"
                    Hdata = Split(Txt, ",") 'Chr(9))
                    If IsNumeric(Hdata(0)) And IsNumeric(Hdata(1)) Then
                        If val((Hdata(0)) = 0 And _
                           val(Hdata(1) = 0)) Then _
                           Gflag = False
                        If (val(CInt(Hdata(0))) < 0 Or _
                            val(CInt(Hdata(0))) > 359) Then _
                            Gflag = False
                        
                        If val(Hdata(1)) <= 0 Then
                            Hdata(1) = 0.001
                        ElseIf val(Hdata(1) > 1#) Then
                            Hdata(1) = 1#
                        End If
                        'everything seems to be within valid ranges!
                        'apply horizontal rotation
                        deg = CInt(Hdata(0)) + Hrot
                        If deg > 359 Then deg = deg - 360
                        If deg < 0 Then deg = deg + 360
                        If Gflag Then HField(deg) = CSng(Hdata(1))
                    Else
                        Gflag = False
                    End If
                End If
            Loop
            
            'flesh out the pattern
            FillOutHorizPat HField()
            
            For i = 0 To 359
                
                'and convert the field values to dB
                'and copy it to the hGain Array
                HGain(i) = 20 * Log10(CDbl(HField(i)))
                'limit gain to -30 dB
                If HGain(i) < -30 Then HGain(i) = -30
                
'                Debug.Print i, HField(i), HGain(i)
                
                
'                Stop
                
            Next i
        
'Stop
        
        End If
    TS.Close
End Sub

'''for i = 0 to 359 : if (i mod 10 = 0) then ? "" : ? i, hgain(i): next i


Private Sub FillOutHorizPat(HField() As Single)
    Dim A1 As Double, A2 As Double
    Dim i1 As Integer, i2 As Integer, itop As Integer
    Dim i As Integer, j As Integer
    Dim fDone As Boolean
    Dim Steps As Integer, perStep As Double
    
    'find the first non-negative value
    A1 = -1
    For i = 0 To 359
        If HField(i) > 0 Then
            i1 = i
            itop = i
            A1 = HField(i)
            Exit For
        End If
    Next i
    If A1 > 0 Then
        
        'find the last non-negative value
        For i = 359 To 0 Step -1
            If HField(i) > 0 Then
                i2 = i
                A2 = HField(i)
                Exit For
            End If
        Next i
            If A2 > 0 Then
            
            Do
                fDone = True
                For i = i1 To i2
                    If HField(i) < 0 Then 'there are still fields to be set!
                        fDone = False
                        Exit For
                    End If
                Next i
                If Not fDone Then
                    i1 = i - 1 'new starting point
                    For i = i1 + 1 To i2
                        If HField(i) > 0 Then 'found ending point
                            Steps = i - i1 + 1
                            perStep = (HField(i) - HField(i1)) / Steps
                            For j = i1 + 1 To i - 1
                                HField(j) = HField(j - 1) + perStep
                            Next j
                            i1 = i 'set new starting point
                            Exit For
                        End If
                    Next i
                End If
            Loop Until fDone
            
            'now need to fill in between i2 and around the top to the original i1
            Steps = itop + (359 - i2) + 1
            perStep = (HField(itop) - HField(i2)) / Steps
            'from i2 to 359-degrees
            For i = i2 + 1 To 359
                HField(i) = HField(i - 1) + perStep
            Next i
            'special case of no value for due north
            If HField(0) < 0 Then HField(0) = HField(359) + perStep
            'from 1-degree to iTop
            For i = 1 To itop
                HField(i) = HField(i - 1) + perStep
            Next i
        End If 'A2 value found
    End If 'a1 value found
    
End Sub

Private Sub cmdLoad_Click()

    LoadMPE
    
    With MPE
        NetGain = 0
        tbxSourceID.Text = .ID
        tbxX1.Text = .SE
        tbxX2.Text = .TE
        tbxY1.Text = .SN
        tbxY2.Text = .TN
        tbxZ1.Text = .SH
        TbxZ2.Text = .TH
        tbxFreq.Text = .Freq
        tbxERP.Text = .MERP
        tbxVGain.Text = .VGain
        tbxAzGain.Text = .AGain
    End With
    tbxFreq_Change
    tbxERP_Change

End Sub

Private Sub cmdSave_Click()
    saveMPE
End Sub

Private Sub cmdVerPat0_Click()
    Frame12.Enabled = True
    tbxVGain.Text = 0
    Erase VGain

End Sub


Private Sub cmdVPolarity_Click()
    If cmdVPolarity.Caption = "N" Then
        cmdVPolarity.Caption = "S"
    Else
        cmdVPolarity.Caption = "N"
    End If
End Sub

Private Sub Form_Load()
    Dim itop As Integer
    Dim ileft As Integer
    
    gxoFileTool.GetLastScreenPosn Me.Name, itop, ileft
    Me.Move ileft, itop
    With MPE
        .ID = ""
        .Locc = 1
        .Lgp = 1
        .SN = 0
        .TN = 0
        .SE = 0
        .TE = 0
        .SH = 100
        .TH = 2
        .Freq = 100
        .VGain = 0
        .AGain = 0
        .MERP = 250
        NetGain = 0
        tbxSourceID.Text = .ID
        tbxX1.Text = .SE
        tbxX2.Text = .TE
        tbxY1.Text = .SN
        tbxY2.Text = .TN
        tbxZ1.Text = .SH
        TbxZ2.Text = .TH
        tbxFreq.Text = .Freq
        tbxERP.Text = .MERP
        tbxVGain.Text = .VGain
        tbxAzGain.Text = .AGain
    End With
    tbxFreq_Change
    tbxERP_Change
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
    Case vbFormControlMenu 'The user chose the Close command from the Control menu on the form.
        cmdClose_Click
    Case vbFormCode 'The Unload statement is invoked from code.
    Case vbAppWindows 'The current Microsoft Windows operating environment session is ending.
        cmdClose_Click
    Case vbAppTaskManager 'The Microsoft Windows Task Manager is closing the application.
        cmdClose_Click
    Case vbFormMDIForm 'An MDI child form is closing because the MDI form is closing.
        cmdClose_Click
    Case vbFormOwner 'A form is closing because its owner is closing.
        cmdClose_Click
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    gxoFileTool.SaveLastScreenPosn Me.Name, Me.Top, Me.Left

End Sub


Private Sub tbxERP_Change()
    If IsNumeric(tbxERP.Text) Then
        MPE.MERP = CDbl(tbxERP.Text)
        UpdateLabels
    End If
End Sub

Private Sub tbxFreq_Change()
    Dim lFreq As Double
    Dim fq As Double
    
    If IsNumeric(tbxFreq.Text) Then
        lFreq = CDbl(tbxFreq.Text)
        Select Case lFreq
            Case 0.3 To 1.34
                MPE.Locc = 100
                MPE.Lgp = 100
            Case 1.34 To 3
                fq = (lFreq * lFreq)
                MPE.Locc = 100
                MPE.Lgp = 180 / fq
            Case 3 To 30
                fq = (lFreq * lFreq)
                MPE.Locc = 900 / fq
                MPE.Lgp = 180 / fq
            Case 30 To 300
                MPE.Locc = 1
                MPE.Lgp = 0.2
            Case 300 To 1500
                MPE.Locc = lFreq / 300
                MPE.Lgp = lFreq / 1500
            Case 1500 To 100000
                MPE.Locc = 5
                MPE.Lgp = 1
            Case Else
                Exit Sub
        End Select
        
        UpdateLabels
    End If
    
End Sub

Private Sub tbxScanIncrXY_Validate(Cancel As Boolean)
    Dim val As Single
    If IsNumeric(tbxScanIncrXY.Text) Then
        val = CSng(tbxScanIncrXY.Text)
        val = Abs(val)
        tbxScanIncrXY.Text = CStr(val)
    End If
End Sub

Private Sub tbxScanIncrX_Validate(Cancel As Boolean)
    Dim val As Single
    If IsNumeric(tbxScanIncrX.Text) Then
        val = CSng(tbxScanIncrX.Text)
        val = Abs(val)
        tbxScanIncrX.Text = CStr(val)
    End If
End Sub

Private Sub tbxScanIncrZ_Validate(Cancel As Boolean)
    Dim val As Single
    If IsNumeric(tbxScanIncrZ.Text) Then
        val = CSng(tbxScanIncrZ.Text)
        val = Abs(val)
        tbxScanIncrZ.Text = CStr(val)
    End If
End Sub

Private Sub tbxScanX_Validate(Cancel As Boolean)
    Dim val As Integer
    If IsNumeric(tbxScanX.Text) Then
        val = CInt(tbxScanX.Text)
        val = Abs(val)
        If val < 1 Then val = 1
        If val > 500 Then val = 500
        tbxScanX.Text = CStr(val)
    End If
End Sub

Private Sub tbxScanXY_Validate(Cancel As Boolean)
    Dim val As Integer
    If IsNumeric(tbxScanXY.Text) Then
        val = CInt(tbxScanXY.Text)
        val = Abs(val)
        If val < 1 Then val = 1
        If val > 250 Then val = 250
        tbxScanXY.Text = CStr(val)
    End If
End Sub

Private Sub tbxScanZ_Validate(Cancel As Boolean)
    Dim val As Integer
    If IsNumeric(tbxScanZ.Text) Then
        val = CInt(tbxScanZ.Text)
        val = Abs(val)
        If val < 1 Then val = 1
        If val > 800 Then val = 800
        tbxScanZ.Text = CStr(val)
    End If
End Sub

Private Sub tbxSourceID_Change()
    MPE.ID = tbxSourceID.Text

End Sub

Private Sub tbxVGain_Change()
    NetGain = 0
    If IsNumeric(tbxAzGain.Text) Then
        NetGain = CDbl(tbxAzGain.Text)
    End If
    If IsNumeric(tbxVGain.Text) Then
        MPE.VGain = CDbl(tbxVGain.Text)
        NetGain = NetGain + MPE.VGain
        UpdateLabels
    End If
End Sub

Private Sub tbxAzGain_Change()
    NetGain = 0
    If IsNumeric(tbxVGain.Text) Then
        NetGain = CDbl(tbxVGain.Text)
    End If
    If IsNumeric(tbxAzGain.Text) Then
        MPE.AGain = CDbl(tbxAzGain.Text)
        NetGain = NetGain + MPE.AGain
        UpdateLabels
    End If
End Sub

Private Sub tbxX1_Change()
    If IsNumeric(tbxX1.Text) Then
        MPE.SE = CDbl(tbxX1.Text)
        TP.CX1 = MPE.SE
        UpdateLabels
    End If
End Sub

Private Sub tbxX2_Change()
    If IsNumeric(tbxX2.Text) Then
        MPE.TE = CDbl(tbxX2.Text)
        TP.CX2 = MPE.TE
        UpdateLabels
    End If
End Sub

Private Sub tbxY1_Change()
    If IsNumeric(tbxY1.Text) Then
        MPE.SN = CDbl(tbxY1.Text)
        TP.CY1 = MPE.SN
        UpdateLabels
    End If
End Sub

Private Sub tbxY2_Change()
    If IsNumeric(tbxY2.Text) Then
        MPE.TN = CDbl(tbxY2.Text)
        TP.CY2 = MPE.TN
        UpdateLabels
    End If
End Sub

Private Sub tbxZ1_Change()
    If IsNumeric(tbxZ1.Text) Then
        MPE.SH = CDbl(tbxZ1.Text)
        TP.CZ1 = MPE.SH
        UpdateLabels
    End If
End Sub

Private Sub tbxZ2_Change()
    If IsNumeric(TbxZ2.Text) Then
        MPE.TH = CDbl(TbxZ2.Text)
        TP.CZ2 = MPE.TH
        UpdateLabels
    End If
End Sub

Private Sub UpdateLabels()
    Dim dTmp As Double
    Dim Ang As Double
    Dim PDens As Double
    Dim adjERP As Double
    Dim iAng As Integer
    
    With TP
        dTmp = .CZ2
            .CZ2 = .CZ1 'Make delta Z zero, so the azimuth angle is in the XY Plane
            Ang = CDbl(.Beta) / PiOvr180
            If Ang > 359.499999999 Then Ang = Ang - 360
            MPE.AAng = Ang
            lblAzimuth.Caption = Format(Ang, "##0.0")
        .CZ2 = dTmp 'reset elevation angle
        
        Ang = CDbl(.Gamma) / PiOvr180
        MPE.VAng = 90 - Ang
        lblGamma.Caption = Format(MPE.VAng, "##0.0")
        
        MPE.Dist = .Mag
        lblMag.Caption = Format(.Mag, "#,##0.00 meters")
        
        If Not Frame12.Enabled Then 'read elevation gain from array
            iAng = CInt(MPE.VAng)
            tbxVGain.Text = VGain(iAng)
        End If
        
        If Not Frame11.Enabled Then
            iAng = CInt(MPE.AAng)
            tbxAzGain.Text = HGain(iAng)
        End If
        
        If NetGain <> 0 Then
            MPE.AERP = (10 ^ (NetGain / 10)) * MPE.MERP
        Else
            MPE.AERP = MPE.MERP
        End If
        
        On Error Resume Next
            MPE.Pden = ((33.4 * MPE.AERP) / (.Mag ^ 2)) / 1000
        On Error GoTo 0
        lblAdjERP.Caption = Format(MPE.AERP, "###,###,##0.### Watts")
    End With
    MPE.Freq = CSng(tbxFreq.Text)
    lblPwrDens.Caption = Format(MPE.Pden * 1000, "#,##0.0#### uW/cm^2")
    MPE.Pocc = MPE.Pden / MPE.Locc
    lblPctOcc.Caption = Format(MPE.Pocc, "##0.00%") & " of " & Format(MPE.Locc, "##0.00# mW/cm^2")
    MPE.Pgp = MPE.Pden / MPE.Lgp
    lblPctGP.Caption = Format(MPE.Pgp, "##0.00%") & " of " & Format(MPE.Lgp * 1000, "#,##0.00 uW/cm^2")
    
End Sub

Private Sub cmdExport_Click()
    Export
End Sub

Private Sub cmdScanXY_Click()
    Dim TS As TextStream
    Dim i As Single
    Dim j As Single
    Dim Steps As Single
    Dim Incr As Single
    Dim bNorth As Double, bSouth As Double, bEast As Double, bWest As Double
    
    If IsNumeric(tbxScanXY.Text) And IsNumeric(tbxScanIncrXY.Text) Then
        Steps = CSng(tbxScanXY.Text)
        Incr = CSng(tbxScanIncrXY.Text)
        exportHeader
        If gFSO.FileExists(outfile) Then
            bNorth = MPE.TN
            bSouth = bNorth - Steps
            bEast = MPE.TE
            bWest = bEast - Steps
            Set TS = gFSO.OpenTextFile(outfile, ForAppending, False)
                For i = bNorth To bSouth Step -Incr
                    tbxY2.Text = i
                    For j = bWest To bEast Step Incr
                        tbxX2.Text = j
                        TS.Write MPE.ID & Chr(9)
                        TS.Write MPE.SE & Chr(9)
                        TS.Write MPE.SN & Chr(9)
                        TS.Write MPE.SH & Chr(9)
                        TS.Write MPE.TE & Chr(9)
                        TS.Write MPE.TN & Chr(9)
                        TS.Write MPE.TH & Chr(9)
                        TS.Write MPE.Freq & Chr(9)
                        TS.Write MPE.VGain & Chr(9)
                        TS.Write MPE.AGain & Chr(9)
                        TS.Write MPE.MERP & Chr(9)
                        TS.Write MPE.VAng & Chr(9)
                        TS.Write MPE.AAng & Chr(9)
                        TS.Write MPE.Dist & Chr(9)
                        TS.Write MPE.AERP & Chr(9)
                        TS.Write MPE.Pden & Chr(9)
                        TS.Write MPE.Pocc & Chr(9)
                        TS.writeline MPE.Pgp
                        DoEvents
                    Next j
                Next i
            TS.Close
        End If
    End If

End Sub

Private Sub cmdScanX_Click()
    Dim TS As TextStream
    Dim j As Single
    Dim Steps As Single, Incr As Single
    Dim bNorth As Double, bSouth As Double, bEast As Double, bWest As Double
    
    If IsNumeric(tbxScanX.Text) Then
        Steps = CSng(tbxScanX.Text)
        Incr = CSng(tbxScanIncrX.Text)
        exportHeader
        If gFSO.FileExists(outfile) Then
            bEast = MPE.TE
            bWest = bEast - Steps
            Set TS = gFSO.OpenTextFile(outfile, ForAppending, False)
                For j = bWest To bEast Step Incr
                    tbxX2.Text = j
                    TS.Write MPE.ID & Chr(9)
                    TS.Write MPE.SE & Chr(9)
                    TS.Write MPE.SN & Chr(9)
                    TS.Write MPE.SH & Chr(9)
                    TS.Write MPE.TE & Chr(9)
                    TS.Write MPE.TN & Chr(9)
                    TS.Write MPE.TH & Chr(9)
                    TS.Write MPE.Freq & Chr(9)
                    TS.Write MPE.VGain & Chr(9)
                    TS.Write MPE.AGain & Chr(9)
                    TS.Write MPE.MERP & Chr(9)
                    TS.Write MPE.VAng & Chr(9)
                    TS.Write MPE.AAng & Chr(9)
                    TS.Write MPE.Dist & Chr(9)
                    TS.Write MPE.AERP & Chr(9)
                    TS.Write MPE.Pden & Chr(9)
                    TS.Write MPE.Pocc & Chr(9)
                    TS.writeline MPE.Pgp
                    DoEvents
                Next j
            TS.Close
        End If
    End If

End Sub

Private Sub cmdRasterGP_Click()
    Dim TS As TextStream
    Dim i As Single
    Dim j As Single
    Dim Steps As Single
    Dim Incr As Single
    Dim bNorth As Double, bSouth As Double, bEast As Double, bWest As Double

    If IsNumeric(tbxScanXY.Text) And IsNumeric(tbxScanIncrXY.Text) Then
        Steps = CSng(tbxScanXY.Text)
        Incr = CSng(tbxScanIncrXY.Text)
        bNorth = MPE.TN
        bSouth = bNorth - Steps
        bEast = MPE.TE
        bWest = bEast - Steps
        outfile = GetTxtFileForRasterSave
        If Len(outfile) > 0 Then
            Set TS = gFSO.OpenTextFile(outfile, ForAppending, True)
                'top header  'west to east
                TS.Write "X" & Chr(9)
                For j = bWest To bEast - Incr Step Incr
                    TS.Write j & Chr(9)
                Next j
                TS.Write bEast 'the last heading
                TS.writeline
                
                For i = bNorth To bSouth Step -Incr 'north to south
                    tbxY2.Text = i
                    TS.Write MPE.TN & Chr(9)
                    For j = bWest To bEast Step Incr 'west to east
                        tbxX2.Text = j
                        TS.Write MPE.Pgp & Chr(9)
                        DoEvents
                    Next j
                    TS.writeline
                Next i
            TS.Close
        End If
    End If

End Sub

Private Sub cmdRasterOcc_Click()
    Dim TS As TextStream
    Dim i As Single
    Dim j As Single
    Dim Steps As Single
    Dim Incr As Single
    Dim bNorth As Double, bSouth As Double, bEast As Double, bWest As Double
    
    If IsNumeric(tbxScanXY.Text) And IsNumeric(tbxScanIncrXY.Text) Then
        Steps = CSng(tbxScanXY.Text)
        Incr = CSng(tbxScanIncrXY.Text)
        bNorth = MPE.TN
        bSouth = bNorth - Steps
        bEast = MPE.TE
        bWest = bEast - Steps
        outfile = GetTxtFileForRasterSave
        If Len(outfile) > 0 Then
            Set TS = gFSO.OpenTextFile(outfile, ForAppending, True)
                'top header  'west to east
                TS.Write "X" & Chr(9)
                For j = bWest To bEast - Incr Step Incr
                    TS.Write i & Chr(9)
                Next j
                TS.Write bEast 'the last heading
                TS.writeline
                
                For i = bNorth To bSouth Step -Incr 'north to south
                    tbxY2.Text = i
                    TS.Write MPE.TN & Chr(9)
                    For j = bWest To bEast Step Incr 'west to east
                        tbxX2.Text = j
                        TS.Write MPE.Pocc & Chr(9)
                        DoEvents
                    Next j
                    TS.writeline
                Next i
            TS.Close
        End If
    End If

End Sub

Private Function IsVPatFile(filespec As String) As Boolean
    Dim TS As TextStream
    Dim res As Boolean
    Dim Txt As String
    
    res = False
    Set TS = gFSO.OpenTextFile(filespec, ForReading)
        Do While Not TS.AtEndOfStream
            Txt = TS.ReadLine
            res = InStr(Txt, "RFE Antenna Pattern V") > 0
            If res Then Exit Do
        Loop
    TS.Close
    If Not res Then 'it is an ERI file?
        If gFSO.GetExtensionName(filespec) = "txt" Then
            Set TS = gFSO.OpenTextFile(filespec, ForReading)
                Do While Not TS.AtEndOfStream
                    Txt = TS.ReadLine
                    res = InStr(Txt, "Elevation       Field") > 0
                    If res Then Exit Do
                Loop
            TS.Close
        End If
    End If
    IsVPatFile = res
End Function

Private Function IsHPatFile(filespec As String) As Boolean
    Dim TS As TextStream
    Dim res As Boolean
    Dim Txt As String
    
    res = False
    Set TS = gFSO.OpenTextFile(filespec, ForReading)
        Do While Not TS.AtEndOfStream
            Txt = TS.ReadLine
            res = InStr(Txt, "RFE Antenna Pattern H") > 0
            If res Then Exit Do
        Loop
        IsHPatFile = res
    TS.Close
End Function

Private Sub cmdInit_Click()
    
    'area of analysis in pixels
    MultiiH = CLng(val(tbxPixels.Text))
    MultiiW = MultiiH
    ReDim Multiiexp(1 To MultiiW, 1 To MultiiH)
    ReDim MultiiOcc(1 To MultiiW, 1 To MultiiH)

    MultiIncr = CInt((1 / val(tbxCmpP.Text)) * 100)  '20 = 5cm steps, 10 = 10cm steps, 1 = 1 meter steps
    
End Sub

Private Function AutoAddStation(StaName As String) As Single
    'return value is the peak reading found.
    Dim i As Long, j As Long
    Dim OffSet As Long
    Dim MinPct As Double, MinExp As Double
    Dim MaxPct As Double, MaxExp As Double
    Dim MaxX As Long, MaxY As Long
    Dim iFile As Integer
    
    Dim peak As Single
    peak = -9999
    'boundaries
    Dim bNorth As Double, bSouth As Double, bEast As Double, bWest As Double

    OffSet = MultiiH / 2 'to center the source in the bitmap
    'meters/incr north, south, east and west of source
    bNorth = OffSet
    bSouth = -OffSet + 1
    bEast = OffSet
    bWest = -OffSet + 1

    'set the source location
    MPE.SH = CDbl(tbxZ1.Text)
    TP.CZ1 = MPE.SH
    MPE.SE = CDbl(tbxX1.Text)
    TP.CX1 = MPE.SE
    MPE.SN = CDbl(tbxY1.Text)
    TP.CY1 = MPE.SN

    'black dot at the source location
    On Error Resume Next
        Multiiexp((MultiIncr * MPE.SE) + OffSet, (MultiIncr * MPE.SN) + OffSet) = -999
        MultiiOcc((MultiIncr * MPE.SE) + OffSet, (MultiIncr * MPE.SN) + OffSet) = -999
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
        iFile = FreeFile
        Open StaName For Binary Access Write As iFile
        Put #iFile, , MPE
        For i = bSouth To bNorth 'south to north
            MPE.TN = CSng(i / MultiIncr)
            For j = bWest To bEast 'west to east
                MPE.TE = CSng(j / MultiIncr)
                'update calculations
                SilentUpdate
                
                If (j > -40 And j < 90) And (i > -400 And i < -315) Then
                    If MPE.Pden > peak Then peak = MPE.Pden
                End If
                
                Put #iFile, , MPE.Pden
                Put #iFile, , MPE.Pgp
                Put #iFile, , MPE.Pocc
                'save percent Public to array
                If Multiiexp(j + OffSet, i + OffSet) <> -999 Then
                    Multiiexp(j + OffSet, i + OffSet) = Multiiexp(j + OffSet, i + OffSet) + (MPE.Pgp * 100)
                    MultiiOcc(j + OffSet, i + OffSet) = MultiiOcc(j + OffSet, i + OffSet) + (MPE.Pocc * 100)
                End If
                DoEvents
            Next j
        Next i
        Close iFile
    Screen.MousePointer = vbNormal
    AutoAddStation = peak

End Function

Private Sub cmdMakeBMP_Click()
    Dim Rng As Integer, i As Integer
    Dim oFile As String, oPath As String
    'get save-to filename
    outfile = GetbmpFileForSave
    Screen.MousePointer = vbHourglass
        If Len(outfile) > 0 Then
            'make a bitmap from the mpe.pgp
            For i = 0 To 5
                If optRange(i).Value Then
                    Rng = i
                    Exit For
                End If
            Next i
            If Rng = 5 Then
                oFile = gFSO.GetBaseName(outfile)
                oPath = gFSO.GetParentFolderName(outfile)
                For i = 0 To 4
                    Select Case i
                    Case 0
                        outfile = gFSO.BuildPath(oPath, "1-5pct_" & oFile & ".bmp")
                    Case 1
                        outfile = gFSO.BuildPath(oPath, "2-12pct_" & oFile & ".bmp")
                    Case 2
                        outfile = gFSO.BuildPath(oPath, "5-25pct_" & oFile & ".bmp")
                    Case 3
                        outfile = gFSO.BuildPath(oPath, "10-50pct_" & oFile & ".bmp")
                    Case 4
                        outfile = gFSO.BuildPath(oPath, "20-100pct_" & oFile & ".bmp")
                    End Select
                    MakeMultiExposureBMP MultiiH, MultiiW, outfile, Multiiexp(), i
                Next i
            Else
                MakeMultiExposureBMP MultiiH, MultiiW, outfile, Multiiexp(), Rng
            End If
        End If
    Screen.MousePointer = vbNormal

End Sub

Public Function GetStationFileForOpen() As String
    Static LastFile As String

    With frmSplash!CommonDialog1
        .DialogTitle = "Read From pipe delimited File"
        .DefaultExt = ".txt"
        '.InitDir = gJobDIR
        .FileName = LastFile
        .Flags = cdlOFNFileMustExist
        .Filter = "Pipe Delimmited File (*.txt)|*.txt"
        .CancelError = True
        On Error Resume Next
            .ShowOpen
            If Err <> 0 Then
                GetStationFileForOpen = ""
            Else
                GetStationFileForOpen = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function

