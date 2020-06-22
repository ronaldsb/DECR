VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca del programa."
   ClientHeight    =   6435
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6270
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4439.481
   ScaleMode       =   0  'User
   ScaleWidth      =   5893.488
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   336.791
      ScaleMode       =   0  'User
      ScaleWidth      =   336.791
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4920
      TabIndex        =   0
      Top             =   5760
      Width           =   900
   End
   Begin VB.Label lblDescription 
      Caption         =   "Email: ronaldsb@me.com"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   15
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      Caption         =   "Ronald Steinvorth Berrocal"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      Caption         =   "Realizado por:"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Diseño de elementos en concreto reforzado."
      Height          =   252
      Left            =   960
      TabIndex        =   12
      Top             =   240
      Width           =   3252
   End
   Begin VB.Label Label7 
      Caption         =   "Ing. Rolando Aguilar A."
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Ing. Victor Rojas Q."
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Ing. Alfonso Bravo H."
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Ing. Alfonso Salvo S."
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Tutores externos:"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Tutor de tesis: Ing. Jorge A. Ruiz M."
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Educacional."
      Height          =   228
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   1008
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.794
      X2              =   5337.977
      Y1              =   1904.113
      Y2              =   1904.113
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":030A
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.794
      X2              =   5323.878
      Y1              =   1904.113
      Y2              =   1904.113
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión: 1.0"
      Height          =   228
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   912
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":03C7
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   480
      TabIndex        =   3
      Top             =   4440
      Width           =   5745
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    
    Private Sub cmdOK_Click()
         
         
    Unload Me
    
    
    End Sub

