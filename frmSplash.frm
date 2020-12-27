VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0CCA
   ScaleHeight     =   6090
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   0
      Top             =   3360
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()

    Me.Top = -10000
    Me.Left = -10000
    
    
#If InHouse = 0 Then
    frmMPECalc.Show
#Else
    frmMPESCalc.Show
#End If
    
    Timer1.Enabled = False

End Sub

Private Sub Form_Load()
    
    Me.Height = 6090
    Me.Width = 3260
    
    DisplayForm

End Sub

Public Sub DisplayForm()
    
    Me.Top = (Screen.Height / 2) - 3050
    Me.Left = (Screen.Width / 2) - 1630

    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

    Form_Click

End Sub
