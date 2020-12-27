VERSION 5.00
Begin VB.Form frmSource 
   Caption         =   "Add New Source"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   Icon            =   "frmSource2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Source Decription:"
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3255
      Begin VB.TextBox tbxDescrip 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Center of Radiation (meters)"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   3255
      Begin VB.Frame Frame2 
         Caption         =   "X"
         Height          =   615
         Left            =   0
         TabIndex        =   13
         Top             =   240
         Width           =   1095
         Begin VB.TextBox tbxX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Y"
         Height          =   615
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   1095
         Begin VB.TextBox tbxY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Z (height)"
         Height          =   615
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1095
         Begin VB.TextBox tbxZ 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Load Antenna Pattern:"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   3255
      Begin VB.CommandButton tbxHorizPat 
         Caption         =   "Horizontal"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton tbxVertPat 
         Caption         =   "Vertical"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Max ERP (Watts)"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "250"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frequency (MHz)"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
      Begin VB.TextBox tbxFreq 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "100"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdAddSource 
      Caption         =   "Add Source"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

