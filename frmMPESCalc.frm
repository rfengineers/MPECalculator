VERSION 5.00
Begin VB.Form frmMPESCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPE Calculator:"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmMPESCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About..."
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load..."
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Beta"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   45
      Top             =   6120
      Width           =   1455
      Begin VB.Label lblBeta 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Max ERP (Watts)"
      Height          =   615
      Left            =   0
      TabIndex        =   44
      Top             =   3240
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
      Caption         =   "Frequency (MHz)"
      Height          =   615
      Index           =   3
      Left            =   0
      TabIndex        =   43
      Top             =   3840
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
   Begin VB.Frame Frame10 
      Height          =   4455
      Left            =   3240
      TabIndex        =   28
      Top             =   0
      Width           =   2655
      Begin VB.Frame Frame5 
         Caption         =   "Distance:"
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   2415
         Begin VB.Label lblMag 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Elevation:"
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1095
         Begin VB.Label lblGamma 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Azimuth:"
         Height          =   615
         Index           =   0
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   1095
         Begin VB.Label lblAzimuth 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Power Density:"
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Width           =   2415
         Begin VB.Label lblPwrDens 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "% General Public Exposure:"
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   2415
         Begin VB.Label lblPctGP 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "% Occupational Exposure:"
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   3720
         Width           =   2415
         Begin VB.Label lblPctOcc 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Adj ERP:"
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   2415
         Begin VB.Label lblAdjERP 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Azimuth Gain:"
      Height          =   615
      Left            =   1680
      TabIndex        =   26
      Top             =   3840
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
         TabIndex        =   27
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Elevation Gain:"
      Height          =   615
      Left            =   1680
      TabIndex        =   24
      Top             =   3240
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
         TabIndex        =   25
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Position in Meters:"
      Height          =   2415
      Left            =   0
      TabIndex        =   17
      Top             =   720
      Width           =   3135
      Begin VB.Frame Frame4 
         Caption         =   "Source Height"
         Height          =   615
         Left            =   120
         TabIndex        =   23
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
      Begin VB.Frame Frame3 
         Caption         =   "Source North"
         Height          =   615
         Left            =   120
         TabIndex        =   22
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
      Begin VB.Frame Frame2 
         Caption         =   "Source East"
         Height          =   615
         Left            =   120
         TabIndex        =   21
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
      Begin VB.Frame Frame6 
         Caption         =   "Target East"
         Height          =   615
         Left            =   1680
         TabIndex        =   20
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
      Begin VB.Frame Frame7 
         Caption         =   "Target North"
         Height          =   615
         Left            =   1680
         TabIndex        =   19
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
      Begin VB.Frame Frame8 
         Caption         =   "Target Height"
         Height          =   615
         Left            =   1680
         TabIndex        =   18
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
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export..."
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame14 
      Caption         =   "Source ID:"
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3135
      Begin VB.TextBox tbxSourceID 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmMPESCalc"
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

Private Sub cmdAbout_Click()
    frmSplash.Timer1.Interval = 0
    frmSplash.DisplayForm
End Sub

Private Sub cmdClose_Click()
    Dim f As Form
    
    For Each f In Forms
        Unload f
    Next f
End Sub

Private Sub cmdExport_Click()
    Export
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
            If Ang > 359 Then Ang = Ang - 360
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

