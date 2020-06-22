VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FColumnas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diseño de elementos en concreto reforzado."
   ClientHeight    =   9840
   ClientLeft      =   48
   ClientTop       =   612
   ClientWidth     =   12912
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9840
   ScaleMode       =   0  'User
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabColumnas 
      Height          =   9672
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12732
      _ExtentX        =   22458
      _ExtentY        =   17060
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Flexo-compresión"
      TabPicture(0)   =   "ColumnasForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LHora"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label35"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label34"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label33"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "THora"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TFecha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CColumna"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CFlexion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text22"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CEje"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "THecho"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TProyecto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Diseño por capacidad y confinamiento"
      TabPicture(1)   =   "ColumnasForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label25"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label24"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label23"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label22"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label30"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label29"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label52"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label53"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label54"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label55"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label56"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label57"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label58"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label59"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label60"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label61"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label62"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label63"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label65"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label48"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label49"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label67"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label69"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label44"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label45"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label46"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label9"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label15"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label21"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label66"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label68"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label70"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label47"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Label50"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Label83"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Label71"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Label73"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label27"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label51"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Command1"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Talv"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "TAros"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "TFyh2"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Th2"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Tb2"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Td2"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Trl"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Trt"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "The"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "TMcpj"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "TMcpi"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "TFc2"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Tdp2"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "TAros2"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Tablacapa"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "TablaConfina"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "Frame3"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "TPmaxcomp"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "TPmaxten"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Tsanalisis"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Thx"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "TphiC"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).ControlCount=   63
      Begin VB.TextBox TphiC 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71160
         TabIndex        =   174
         Text            =   "0.85"
         Top             =   840
         Width           =   612
      End
      Begin VB.TextBox Thx 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -64080
         TabIndex        =   169
         Text            =   "18.5"
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox Tsanalisis 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -64080
         Locked          =   -1  'True
         TabIndex        =   166
         Top             =   840
         Width           =   732
      End
      Begin VB.TextBox TPmaxten 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "0"
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox TPmaxcomp 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   160
         Text            =   "0"
         Top             =   2880
         Width           =   732
      End
      Begin VB.TextBox TProyecto 
         BackColor       =   &H80000004&
         Height          =   288
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   960
         Width           =   2292
      End
      Begin VB.TextBox THecho 
         BackColor       =   &H80000004&
         Height          =   288
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   2292
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resultados Finales."
         Height          =   5535
         Left            =   -74760
         TabIndex        =   128
         Top             =   3840
         Width           =   4932
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   5175
            Left            =   1680
            Picture         =   "ColumnasForm.frx":0038
            ScaleHeight     =   5172
            ScaleWidth      =   1332
            TabIndex        =   135
            Top             =   240
            Width           =   1332
         End
         Begin VB.TextBox T1 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   134
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox T2 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   133
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox T5 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   132
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox T6 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   131
            Top             =   3960
            Width           =   735
         End
         Begin VB.TextBox T3 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   130
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox T4 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   129
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "Lo ="
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label28 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   146
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label31 
            Caption         =   "S ="
            Height          =   255
            Left            =   200
            TabIndex        =   145
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label32 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   144
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label Label36 
            Caption         =   "Lo ="
            Height          =   255
            Left            =   120
            TabIndex        =   143
            Top             =   3960
            Width           =   375
         End
         Begin VB.Label Label37 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   142
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label Label38 
            Caption         =   "S ="
            Height          =   255
            Left            =   200
            TabIndex        =   141
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label Label39 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   140
            Top             =   3960
            Width           =   255
         End
         Begin VB.Label Label40 
            Caption         =   "L ="
            Height          =   255
            Left            =   3360
            TabIndex        =   139
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label Label41 
            Caption         =   "cm"
            Height          =   255
            Left            =   4560
            TabIndex        =   138
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label42 
            Caption         =   "S ="
            Height          =   255
            Left            =   3360
            TabIndex        =   137
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label Label43 
            Caption         =   "cm"
            Height          =   255
            Left            =   4560
            TabIndex        =   136
            Top             =   3120
            Width           =   255
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid TablaConfina 
         Height          =   4620
         Left            =   -65520
         TabIndex        =   121
         Top             =   4800
         Width           =   2655
         _ExtentX        =   4678
         _ExtentY        =   8149
         _Version        =   393216
         Rows            =   19
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         Enabled         =   0   'False
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tablacapa 
         Height          =   1965
         Left            =   -69120
         TabIndex        =   120
         Top             =   4800
         Width           =   2535
         _ExtentX        =   4466
         _ExtentY        =   3471
         _Version        =   393216
         Rows            =   8
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         Enabled         =   0   'False
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.TextBox TAros2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -68040
         TabIndex        =   49
         Text            =   "2"
         Top             =   3600
         Width           =   732
      End
      Begin VB.TextBox Tdp2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   107
         Text            =   "5"
         Top             =   2250
         Width           =   615
      End
      Begin VB.TextBox TFc2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   39
         Text            =   "210"
         Top             =   795
         Width           =   615
      End
      Begin VB.TextBox TMcpi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -68040
         TabIndex        =   46
         Text            =   "80"
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox TMcpj 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -68040
         TabIndex        =   47
         Text            =   "80"
         Top             =   2148
         Width           =   732
      End
      Begin VB.TextBox The 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -68040
         TabIndex        =   48
         Text            =   "3"
         Top             =   2520
         Width           =   732
      End
      Begin VB.TextBox Trt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   45
         Text            =   "4"
         Top             =   3315
         Width           =   615
      End
      Begin VB.TextBox Trl 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   44
         Text            =   "8"
         Top             =   2955
         Width           =   615
      End
      Begin VB.TextBox Td2 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "55"
         Top             =   2595
         Width           =   615
      End
      Begin VB.TextBox Tb2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   41
         Text            =   "60"
         Top             =   1515
         Width           =   615
      End
      Begin VB.TextBox Th2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   42
         Text            =   "60"
         Top             =   1875
         Width           =   615
      End
      Begin VB.TextBox TFyh2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   40
         Text            =   "2800"
         Top             =   1155
         Width           =   615
      End
      Begin VB.TextBox TAros 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -64080
         TabIndex        =   50
         Text            =   "2"
         Top             =   2508
         Width           =   732
      End
      Begin VB.TextBox Talv 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -64080
         TabIndex        =   51
         Text            =   "3"
         Top             =   2148
         Width           =   732
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Dar resultados"
         Height          =   495
         Left            =   -65280
         TabIndex        =   52
         Top             =   3360
         Width           =   2052
      End
      Begin VB.ComboBox CEje 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         ItemData        =   "ColumnasForm.frx":1BB3A
         Left            =   1800
         List            =   "ColumnasForm.frx":1BB3C
         TabIndex        =   3
         Text            =   "M3"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   78
         Text            =   "EJE ="
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton CFlexion 
         Caption         =   "Actualizar cálculos."
         Height          =   495
         Left            =   3480
         TabIndex        =   35
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   68
         Text            =   "ELEMENTO #"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox CColumna 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         ItemData        =   "ColumnasForm.frx":1BB3E
         Left            =   1800
         List            =   "ColumnasForm.frx":1BB40
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TFecha 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   564
         Width           =   1455
      End
      Begin VB.TextBox THora 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   924
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   8280
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   12492
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tabla1 
            Height          =   3930
            Left            =   7920
            TabIndex        =   119
            Top             =   4080
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   6943
            _Version        =   393216
            Rows            =   16
            Cols            =   5
            FixedRows       =   2
            Enabled         =   0   'False
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.OptionButton OFlexion 
            Caption         =   "Flexión."
            Height          =   255
            Left            =   8880
            TabIndex        =   157
            Top             =   3480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OCortante 
            Caption         =   "Cortante."
            Height          =   255
            Left            =   10440
            TabIndex        =   156
            Top             =   3480
            Width           =   975
         End
         Begin VB.PictureBox Picture1 
            Height          =   3015
            Left            =   8040
            Picture         =   "ColumnasForm.frx":1BB42
            ScaleHeight     =   2964
            ScaleWidth      =   4140
            TabIndex        =   124
            Top             =   240
            Width           =   4185
         End
         Begin VB.CommandButton Diagrama 
            Caption         =   "Actualizar diagrama."
            Height          =   492
            Left            =   5760
            TabIndex        =   36
            Top             =   2640
            Width           =   1812
         End
         Begin VB.TextBox TOK 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   1920
            Width           =   735
         End
         Begin VB.Frame Frame2 
            Height          =   1880
            Left            =   4560
            TabIndex        =   73
            Top             =   360
            Width           =   3280
            Begin VB.TextBox Tgato 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   155
               Top             =   1560
               Width           =   255
            End
            Begin VB.TextBox Tgato 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   154
               Top             =   1320
               Width           =   255
            End
            Begin VB.TextBox Tgato 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   153
               Top             =   1080
               Width           =   255
            End
            Begin VB.TextBox Tgato 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   152
               Text            =   "#"
               Top             =   840
               Width           =   255
            End
            Begin VB.TextBox Tgato 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   151
               Text            =   "#"
               Top             =   600
               Width           =   255
            End
            Begin VB.TextBox Tgato 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   150
               Text            =   "#"
               Top             =   360
               Width           =   255
            End
            Begin VB.TextBox Tdx1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   2400
               TabIndex        =   34
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox Tdx1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   2400
               TabIndex        =   30
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox Tdx1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   2400
               TabIndex        =   26
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox Tdx1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   2400
               TabIndex        =   22
               Text            =   "55"
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox Tdx1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   2400
               TabIndex        =   18
               Text            =   "30"
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox Tdx1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   2400
               TabIndex        =   14
               Text            =   "5"
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox TAs1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   1560
               TabIndex        =   33
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox TAs1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   1560
               TabIndex        =   29
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox TAs1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1560
               TabIndex        =   25
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox TAs1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1560
               TabIndex        =   21
               Text            =   "15.2"
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox TAs1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1560
               TabIndex        =   17
               Text            =   "10.13"
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox TAs1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1560
               TabIndex        =   13
               Text            =   "15.2"
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox Tdiam 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   1200
               TabIndex        =   32
               Top             =   1560
               Width           =   375
            End
            Begin VB.TextBox Tdiam 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   1200
               TabIndex        =   28
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox Tdiam 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1200
               TabIndex        =   24
               Top             =   1080
               Width           =   375
            End
            Begin VB.TextBox Tdiam 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1200
               TabIndex        =   20
               Text            =   "8"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox Tdiam 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1200
               TabIndex        =   16
               Text            =   "8"
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox Tdiam 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   12
               Text            =   "8"
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox Tnum 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   600
               TabIndex        =   31
               Top             =   1560
               Width           =   375
            End
            Begin VB.TextBox Tnum 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   600
               TabIndex        =   27
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox Tnum 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   600
               TabIndex        =   23
               Top             =   1080
               Width           =   375
            End
            Begin VB.TextBox Tnum 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   600
               TabIndex        =   19
               Text            =   "3"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox Tnum 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   600
               TabIndex        =   15
               Text            =   "2"
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox Tnum 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   600
               TabIndex        =   11
               Text            =   "3"
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox Text24 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   116
               Text            =   "6"
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox Text11 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   114
               Text            =   "5"
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Text15 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   115
               Text            =   "4"
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   113
               Text            =   "3"
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox Text18 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   111
               Text            =   "2"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox Text28 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   112
               Text            =   "1"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   77
               Text            =   "Capa"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   75
               Text            =   "As (cm2)"
               Top             =   120
               Width           =   855
            End
            Begin VB.TextBox Text9 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   74
               Text            =   "x (cm)"
               Top             =   120
               Width           =   855
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   600
               Locked          =   -1  'True
               TabIndex        =   76
               Text            =   "Varillas"
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.TextBox TEs 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            TabIndex        =   10
            Text            =   "2.1E6"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox TEus 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0.002"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TEuc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            TabIndex        =   8
            Text            =   "0.003"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Tb 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Text            =   "60"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Th 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Text            =   "60"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TFc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Text            =   "210"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TFy 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Text            =   "4200"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox Tpa 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox TLong 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1920
            Width           =   615
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tabla2 
            Height          =   3930
            Left            =   7920
            TabIndex        =   158
            Top             =   4080
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   6943
            _Version        =   393216
            Rows            =   16
            Cols            =   5
            FixedRows       =   2
            Enabled         =   0   'False
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Label Label6 
            Caption         =   "Momentos en ton-m y carga axial en ton."
            Height          =   255
            Left            =   8520
            TabIndex        =   125
            Top             =   3840
            Width           =   3015
         End
         Begin VB.Label Label16 
            Caption         =   "Cortante y carga axial en ton."
            Height          =   255
            Left            =   9000
            TabIndex        =   159
            Top             =   3840
            Width           =   2175
         End
         Begin VB.OLE OLE1 
            AutoActivate    =   3  'Automatic
            BackColor       =   &H80000004&
            Class           =   "Excel.Chart.8"
            Enabled         =   0   'False
            Height          =   5532
            Left            =   120
            OleObjectBlob   =   "ColumnasForm.frx":CA9D0
            SizeMode        =   1  'Stretch
            TabIndex        =   123
            Top             =   2520
            Width           =   7644
         End
         Begin VB.Label Label2 
            Caption         =   "% Acero ="
            Height          =   252
            Left            =   2520
            TabIndex        =   118
            Top             =   1920
            Width           =   852
         End
         Begin VB.Label Label18 
            Caption         =   "su ="
            Height          =   252
            Left            =   2880
            TabIndex        =   81
            Top             =   840
            Width           =   372
         End
         Begin VB.Label Label8 
            Caption         =   "cu ="
            Height          =   252
            Left            =   2880
            TabIndex        =   80
            Top             =   480
            Width           =   372
         End
         Begin VB.Label Label5 
            Caption         =   "e"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   2760
            TabIndex        =   79
            Top             =   780
            Width           =   132
         End
         Begin VB.Label Label20 
            Caption         =   "Es ="
            Height          =   252
            Left            =   2880
            TabIndex        =   72
            Top             =   1200
            Width           =   372
         End
         Begin VB.Label Label17 
            Caption         =   "e"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   2760
            TabIndex        =   71
            Top             =   420
            Width           =   132
         End
         Begin VB.Label Label1 
            Caption         =   "b ="
            Height          =   252
            Left            =   480
            TabIndex        =   65
            Top             =   480
            Width           =   372
         End
         Begin VB.Label Label3 
            Caption         =   "cm"
            Height          =   252
            Left            =   1560
            TabIndex        =   64
            Top             =   480
            Width           =   372
         End
         Begin VB.Label Label4 
            Caption         =   "h ="
            Height          =   252
            Left            =   480
            TabIndex        =   63
            Top             =   840
            Width           =   372
         End
         Begin VB.Label Label7 
            Caption         =   "cm"
            Height          =   252
            Left            =   1560
            TabIndex        =   62
            Top             =   840
            Width           =   372
         End
         Begin VB.Label Label10 
            Caption         =   "f'c ="
            Height          =   252
            Left            =   360
            TabIndex        =   61
            Top             =   1200
            Width           =   372
         End
         Begin VB.Label Label11 
            Caption         =   "fy ="
            Height          =   252
            Left            =   360
            TabIndex        =   60
            Top             =   1560
            Width           =   372
         End
         Begin VB.Label Label12 
            Caption         =   "% Acero ="
            Height          =   252
            Left            =   2520
            TabIndex        =   59
            Top             =   1560
            Width           =   852
         End
         Begin VB.Label Label13 
            Caption         =   "kg/cm2"
            Height          =   252
            Left            =   1560
            TabIndex        =   58
            Top             =   1200
            Width           =   612
         End
         Begin VB.Label Label14 
            Caption         =   "kg/cm2"
            Height          =   252
            Left            =   1560
            TabIndex        =   57
            Top             =   1560
            Width           =   612
         End
         Begin VB.Label Label81 
            Caption         =   "Alt ="
            Height          =   252
            Left            =   360
            TabIndex        =   56
            Top             =   1920
            Width           =   372
         End
         Begin VB.Label Label82 
            Caption         =   "m"
            Height          =   252
            Left            =   1560
            TabIndex        =   55
            Top             =   1920
            Width           =   252
         End
      End
      Begin VB.Label Label51 
         Caption         =   "f ="
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.4
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   175
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "# Patas ="
         Height          =   252
         Left            =   -68880
         TabIndex        =   173
         Top             =   3600
         Width           =   732
      End
      Begin VB.Label Label73 
         Caption         =   "Aros #"
         Height          =   255
         Left            =   -73920
         TabIndex        =   172
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label71 
         Caption         =   "V.L #"
         Height          =   255
         Left            =   -73800
         TabIndex        =   171
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label83 
         Caption         =   "# Patas ="
         Height          =   252
         Left            =   -64920
         TabIndex        =   170
         Top             =   2520
         Width           =   732
      End
      Begin VB.Label Label50 
         Caption         =   "cm"
         Height          =   252
         Left            =   -63240
         TabIndex        =   168
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label47 
         Caption         =   "Separación de aros debida al cortante máximo del análisis: ="
         Height          =   252
         Left            =   -68520
         TabIndex        =   167
         Top             =   840
         Width           =   4452
      End
      Begin VB.Label Label70 
         Caption         =   "Pumax a tensión ="
         Height          =   252
         Left            =   -69480
         TabIndex        =   165
         Top             =   3240
         Width           =   1332
      End
      Begin VB.Label Label68 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -67200
         TabIndex        =   164
         Top             =   3240
         Width           =   372
      End
      Begin VB.Label Label66 
         Caption         =   "Pumax a compresión ="
         Height          =   252
         Left            =   -69780
         TabIndex        =   162
         Top             =   2880
         Width           =   1692
      End
      Begin VB.Label Label21 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -67200
         TabIndex        =   161
         Top             =   2880
         Width           =   372
      End
      Begin VB.Label Label33 
         Caption         =   "Hecho por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5880
         TabIndex        =   149
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label34 
         Caption         =   "Proyecto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   6000
         TabIndex        =   148
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label15 
         Caption         =   "Resultados por confinamiento."
         Height          =   252
         Left            =   -65280
         TabIndex        =   127
         Top             =   4440
         Width           =   2292
      End
      Begin VB.Label Label9 
         Caption         =   "Resultados por capacidad."
         Height          =   252
         Left            =   -68880
         TabIndex        =   126
         Top             =   4440
         Width           =   2052
      End
      Begin VB.Label Label46 
         Caption         =   "# de var."
         Height          =   255
         Left            =   -72480
         TabIndex        =   110
         Top             =   3330
         Width           =   735
      End
      Begin VB.Label Label45 
         Caption         =   "Aros por confinamiento:"
         Height          =   252
         Left            =   -65040
         TabIndex        =   109
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label Label44 
         Caption         =   "Diseño por capacidad:"
         Height          =   252
         Left            =   -69000
         TabIndex        =   108
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label Label69 
         Caption         =   "Mcpi ="
         Height          =   252
         Left            =   -68640
         TabIndex        =   106
         Top             =   1800
         Width           =   492
      End
      Begin VB.Label Label67 
         Caption         =   "Mcpj ="
         Height          =   252
         Left            =   -68640
         TabIndex        =   105
         Top             =   2148
         Width           =   492
      End
      Begin VB.Label Label49 
         Caption         =   "Luz libre (capac.) ="
         Height          =   252
         Left            =   -69480
         TabIndex        =   104
         Top             =   2508
         Width           =   1452
      End
      Begin VB.Label Label48 
         Caption         =   "m"
         Height          =   255
         Left            =   -67150
         TabIndex        =   103
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label65 
         Caption         =   "# de var."
         Height          =   255
         Left            =   -72480
         TabIndex        =   102
         Top             =   2955
         Width           =   735
      End
      Begin VB.Label Label63 
         Caption         =   "cm"
         Height          =   255
         Left            =   -72480
         TabIndex        =   101
         Top             =   2235
         Width           =   375
      End
      Begin VB.Label Label62 
         Caption         =   "d' ="
         Height          =   255
         Left            =   -73680
         TabIndex        =   100
         Top             =   2235
         Width           =   375
      End
      Begin VB.Label Label61 
         Caption         =   "cm"
         Height          =   255
         Left            =   -72480
         TabIndex        =   99
         Top             =   2595
         Width           =   375
      End
      Begin VB.Label Label60 
         Caption         =   "d ="
         Height          =   255
         Left            =   -73680
         TabIndex        =   98
         Top             =   2610
         Width           =   375
      End
      Begin VB.Label Label59 
         Caption         =   "kg/cm2"
         Height          =   255
         Left            =   -72480
         TabIndex        =   97
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label58 
         Caption         =   "fyh ="
         Height          =   255
         Left            =   -73800
         TabIndex        =   96
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label Label57 
         Caption         =   "b ="
         Height          =   255
         Left            =   -73680
         TabIndex        =   95
         Top             =   1530
         Width           =   375
      End
      Begin VB.Label Label56 
         Caption         =   "cm"
         Height          =   255
         Left            =   -72480
         TabIndex        =   94
         Top             =   1515
         Width           =   375
      End
      Begin VB.Label Label55 
         Caption         =   "cm"
         Height          =   255
         Left            =   -72480
         TabIndex        =   93
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label Label54 
         Caption         =   "h ="
         Height          =   255
         Left            =   -73680
         TabIndex        =   92
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label Label53 
         Caption         =   "f'c ="
         Height          =   255
         Left            =   -73800
         TabIndex        =   91
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label52 
         Caption         =   "kg/cm2"
         Height          =   255
         Left            =   -72480
         TabIndex        =   90
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "m"
         Height          =   255
         Left            =   -72720
         TabIndex        =   89
         Top             =   1275
         Width           =   255
      End
      Begin VB.Label Label29 
         Caption         =   "hx ="
         Height          =   252
         Left            =   -64560
         TabIndex        =   88
         Top             =   1788
         Width           =   372
      End
      Begin VB.Label Label30 
         Caption         =   "cm"
         Height          =   252
         Left            =   -63240
         TabIndex        =   87
         Top             =   1788
         Width           =   372
      End
      Begin VB.Label Label22 
         Caption         =   "m"
         Height          =   252
         Left            =   -63240
         TabIndex        =   86
         Top             =   2160
         Width           =   372
      End
      Begin VB.Label Label23 
         Caption         =   "Luz libre (confina.) ="
         Height          =   252
         Left            =   -65640
         TabIndex        =   85
         Top             =   2160
         Width           =   1452
      End
      Begin VB.Label Label24 
         Caption         =   "Ton*m"
         Height          =   252
         Left            =   -67200
         TabIndex        =   84
         Top             =   2160
         Width           =   492
      End
      Begin VB.Label Label25 
         Caption         =   "Ton*m"
         Height          =   252
         Left            =   -67200
         TabIndex        =   83
         Top             =   1788
         Width           =   492
      End
      Begin VB.Label Label35 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   9600
         TabIndex        =   70
         Top             =   576
         Width           =   612
      End
      Begin VB.Label LHora 
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   9720
         TabIndex        =   69
         Top             =   924
         Width           =   492
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   9480
      Left            =   120
      TabIndex        =   82
      Top             =   120
      Width           =   11628
      _ExtentX        =   20511
      _ExtentY        =   16722
      _Version        =   393216
      Rows            =   19
      Cols            =   9
      FixedCols       =   0
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "FRAME|LOAD|M2|M3|P|STATION|T|V2|V3"
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.mdb"
      InitDir         =   "C:\My Documents\Diseño\"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   9480
      Left            =   120
      TabIndex        =   122
      Top             =   120
      Width           =   11625
      _ExtentX        =   20511
      _ExtentY        =   16722
      _Version        =   393216
      Rows            =   19
      Cols            =   10
      FixedCols       =   0
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "Column|Load|Loc|M2|M3|P|Story|T|V2|V3"
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
   Begin VB.Menu MVigas 
      Caption         =   "    Vigas    "
   End
   Begin VB.Menu MColumnas 
      Caption         =   "    Columnas    "
   End
   Begin VB.Menu MMuros 
      Caption         =   "    Muros    "
   End
   Begin VB.Menu Msalida 
      Caption         =   "    Base de datos        "
   End
   Begin VB.Menu Msalida2 
      Caption         =   "    Salir de la base de datos    "
   End
   Begin VB.Menu MImprimir 
      Caption         =   "    Imprimir    "
   End
   Begin VB.Menu MGuardar 
      Caption         =   "    Guardar Diseño    "
      Visible         =   0   'False
   End
   Begin VB.Menu MAcerca 
      Caption         =   "    Acerca de...    "
   End
   Begin VB.Menu Salida 
      Caption         =   "    Salir    "
   End
End
Attribute VB_Name = "FColumnas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Const MARGIN_SIZE = 60
    Private datPrimaryRS As ADODB.Recordset


    'Cosas que el programa carga al inicio del módulo de columnas.
    Private Sub Form_Load()

    
    FColumnas.Visible = True

    
    FDialog.Visible = False


    FDialog.Visible = False


    Importar


    Acomodar


    MsgBox ("Este programa es para usos académicos únicamente!"), vbCritical
 
 
    THecho = Hechos
    
    
    TProyecto = Proyectos
 

    'Fin de las cosas que el programa carga al inicio del módulo de columnas.


End Sub
    
    
    Private Sub Importar()
    'Este procedimiento importa los datos del SAP2000 v7.40 al Visual Basic.
    
    
    On Error GoTo Sinsalida
Sinsalida:
    Resume Next

    
    Dim sConnect As String
    Dim sSQL As String
    Dim dfwConn As ADODB.Connection

       
    MSHFlexGrid1.Visible = True
    MSHFlexGrid2.Visible = False
    
    
    sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;User ID=Admin;Data Source=" & Archivo & ";Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Locale Identifier=1033;Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Global Partial Bulk Ops=2"
    sSQL = "select FRAME,LOAD,M2,M3,P,STATION,T,V2,V3 from FrameForces"


    Set dfwConn = New Connection
    dfwConn.Open sConnect

    
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, dfwConn, adOpenForwardOnly, adLockReadOnly

    
    Set MSHFlexGrid1.DataSource = datPrimaryRS

    
    With MSHFlexGrid1

        .Redraw = False
        .ColWidth(0) = -1
        .ColWidth(1) = -1
        .ColWidth(2) = -1
        .ColWidth(3) = -1
        .ColWidth(4) = -1
        .ColWidth(5) = -1
        .ColWidth(6) = -1
        .ColWidth(7) = -1
        .ColWidth(8) = -1

        
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        
        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

        
    End With

    
    ReDim Datos(MSHFlexGrid1.Rows, 8)
    
    
    'Pasa los datos de salida a una matriz.
    For c = 0 To 1
    For r = 0 To MSHFlexGrid1.Rows - 1
    Datos(r, c) = (MSHFlexGrid1.TextMatrix(r, c))
    Next r
    Next c
    'Fin de pasar los datos de salida a una matriz.
    

    'Convierte los datos que son números de kilogramos a toneladas.
    For c = 2 To MSHFlexGrid1.Cols - 1
    For r = 1 To MSHFlexGrid1.Rows - 1
    Datos(r, c) = (MSHFlexGrid1.TextMatrix(r, c)) / 1000
    Next r
    Next c
    'Fin de convertir los datos que son números de kilogramos a toneladas.
    
    
    ReDim Frame(MSHFlexGrid1.Rows)
    
    
    'Guarda los números de los elementos.
    j = 1
    For i = 1 To MSHFlexGrid1.Rows - 1
    If Datos(i, 0) <> Datos(i + 1, 0) Then
    Frame(j) = Datos(i, 0)
    j = j + 1
    End If
    Next i
    FrameU = j
    'Fin de guardar los números de los elementos.
    
    
    'Fin del Procedimiento de importación para el Sap2000 v7.40.
   
     
    'Pone la lista de elementos en el combobox de # de columna.
    CColumna.Clear
    For i = 1 To FrameU - 1
    CColumna.AddItem Frame(i)
    Next i
    'Fin de poner la lista de elementos en el combobox de # de columna.

    CColumna = CColumna.List(0)

    'Pone la lista de los ejes locales.
    CEje.AddItem "M3"
    CEje.AddItem "M2"
    'Fin de poner la lista de los ejes locales.
           
     
     End Sub


    Sub Acomodar()
    'Este procedimiento acomoda los datos importados y el tamaño de las tablas.


    'Centra todos los datos de la Tabla1.
    For i = 0 To 4
    For j = 0 To 15
    With Tabla1
        .Row = j
        .Col = i
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
    End With
    Next j
    Next i
    'Fin de centrar todos los datos de la Tabla1.


    'Le da el ancho deseado a las columnas de la Tabla 1.
    For i = 1 To 4
    With Tabla1
        .ColWidth(0) = 840
        .ColWidth(i) = 840
    End With
    Next i
    'Fin de darle el ancho deseado a las columnas de la Tabla 1.


    'Centra toda la Tabla2.
    For i = 0 To 4
    For j = 0 To 15
    With Tabla2
        .Row = j
        .Col = i
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
    End With
    Next j
    Next i
    'Fin de centrar todos los datos de la Tabla2.
      
    
    'Le da el ancho deseado a la columnas de la Tabla 2.
    For i = 1 To 4
    With Tabla2
        .ColWidth(0) = 840
        .ColWidth(i) = 840
    End With
    Next i
    'Fin de darle el ancho deseado a las columnas de la Tabla 2.


    'Centra toda la TablaCapa.
    For i = 0 To 2
    For j = 0 To 7
    With Tablacapa
           .Row = j
           .Col = i
           .CellAlignment = flexAlignCenterCenter
           .CellFontSize = 8
    End With
    Next j
    Next i
    'Fin de centrar toda la TablaCapa.


    'Le da el ancho deseado a la columnas de la TablaCapa.
    With Tablacapa
           .ColWidth(0) = 1100
           .ColWidth(1) = 800
           .ColWidth(2) = 650
    End With
    'Fin de darle el ancho deseado a la columnas de la TablaCapa.


    'Centra toda la TablaConfina.
    For i = 0 To 2
    For j = 0 To 18
    With TablaConfina
       .Row = j
       .Col = i
       .CellAlignment = flexAlignCenterCenter
       .CellFontSize = 8
    End With
    Next j
    Next i
    'Fin de centrar toda la TablaConfina.

    
    'Le da el ancho deseado a la columnas de la TablaConfina.
    With TablaConfina
        .ColWidth(0) = 1200
        .ColWidth(1) = 800
        .ColWidth(2) = 650
    End With
    'Fin de darle el ancho deseado a la columnas de la TablaConfina.


    'Pone los títulos de la tabla1.
    Tabla1.TextMatrix(1, 0) = "CARGAS"
    Tabla1.TextMatrix(2, 0) = "Muerta"
    Tabla1.TextMatrix(3, 0) = "Viva"
    Tabla1.TextMatrix(4, 0) = "Sismo X"
    Tabla1.TextMatrix(5, 0) = "Sismo Y"
    Tabla1.TextMatrix(6, 0) = "C1"
    Tabla1.TextMatrix(7, 0) = "C2 X "
    Tabla1.TextMatrix(8, 0) = "C2 Y"
    Tabla1.TextMatrix(9, 0) = "C3 X"
    Tabla1.TextMatrix(10, 0) = "C3 Y"
    Tabla1.TextMatrix(11, 0) = "C4 X "
    Tabla1.TextMatrix(12, 0) = "C4 Y"
    Tabla1.TextMatrix(13, 0) = "C5 X"
    Tabla1.TextMatrix(14, 0) = "C5 Y"
    Tabla1.TextMatrix(15, 0) = "MAX"
    

    Tabla1.TextMatrix(0, 1) = "END I"
    Tabla1.TextMatrix(0, 2) = "END I"
    Tabla1.TextMatrix(0, 3) = "END J"
    Tabla1.TextMatrix(0, 4) = "END J"
    Tabla1.TextMatrix(1, 1) = "MOM"
    Tabla1.TextMatrix(1, 2) = "AXIAL"
    Tabla1.TextMatrix(1, 3) = "MOM"
    Tabla1.TextMatrix(1, 4) = "AXIAL"

    
    Tabla2.TextMatrix(1, 0) = "CARGAS"
    Tabla2.TextMatrix(2, 0) = "Muerta"
    Tabla2.TextMatrix(3, 0) = "Viva"
    Tabla2.TextMatrix(4, 0) = "Sismo X"
    Tabla2.TextMatrix(5, 0) = "Sismo Y"
    Tabla2.TextMatrix(6, 0) = "C1"
    Tabla2.TextMatrix(7, 0) = "C2 X "
    Tabla2.TextMatrix(8, 0) = "C2 Y"
    Tabla2.TextMatrix(9, 0) = "C3 X"
    Tabla2.TextMatrix(10, 0) = "C3 Y"
    Tabla2.TextMatrix(11, 0) = "C4 X "
    Tabla2.TextMatrix(12, 0) = "C4 Y"
    Tabla2.TextMatrix(13, 0) = "C5 X"
    Tabla2.TextMatrix(14, 0) = "C5 Y"
    Tabla2.TextMatrix(15, 0) = "MAX"
    
    
    Tabla2.TextMatrix(0, 1) = "END I"
    Tabla2.TextMatrix(0, 2) = "END I"
    Tabla2.TextMatrix(0, 3) = "END J"
    Tabla2.TextMatrix(0, 4) = "END J"
    Tabla2.TextMatrix(1, 1) = "CORT"
    Tabla2.TextMatrix(1, 2) = "AXIAL"
    Tabla2.TextMatrix(1, 3) = "CORT"
    Tabla2.TextMatrix(1, 4) = "AXIAL"

    
    'Pone los títulos y unidades de la Tablacapa.
    Tablacapa.TextMatrix(0, 0) = "Vu="
    Tablacapa.TextMatrix(1, 0) = "Vc="
    Tablacapa.TextMatrix(2, 0) = "Vs="
    Tablacapa.TextMatrix(3, 0) = "Av="
    Tablacapa.TextMatrix(4, 0) = "S="

    
    Tablacapa.TextMatrix(6, 0) = "Smax 1="
    Tablacapa.TextMatrix(7, 0) = "Smax 2="


    Tablacapa.TextMatrix(0, 2) = "Ton"
    Tablacapa.TextMatrix(1, 2) = "Ton"
    Tablacapa.TextMatrix(2, 2) = "Ton"
    Tablacapa.TextMatrix(3, 2) = "cm2"
    Tablacapa.TextMatrix(4, 2) = "cm"

    
    Tablacapa.TextMatrix(6, 2) = "cm"
    Tablacapa.TextMatrix(7, 2) = "cm"
 
  
    'Pone los títulos y unidades de la TablaConfina.
    TablaConfina.TextMatrix(0, 0) = "Ag="
    TablaConfina.TextMatrix(1, 0) = "Ach="
    TablaConfina.TextMatrix(2, 0) = "Area del aro="
    TablaConfina.TextMatrix(3, 0) = "Ash="
    TablaConfina.TextMatrix(4, 0) = "hc="

    
    TablaConfina.TextMatrix(6, 0) = "Smax 1 (Ext)="
    TablaConfina.TextMatrix(7, 0) = "Smax 2 (Ext)="
    TablaConfina.TextMatrix(8, 0) = "sx (Ext)="


    TablaConfina.TextMatrix(10, 0) = "Smax 4 (Ext)="
    TablaConfina.TextMatrix(11, 0) = "Smax 5 (Ext)="

    
    TablaConfina.TextMatrix(13, 0) = "Smax 1 (Cen)="
    TablaConfina.TextMatrix(14, 0) = "Smax 2 (Cen)="

    
    TablaConfina.TextMatrix(16, 0) = "Lo 1="
    TablaConfina.TextMatrix(17, 0) = "Lo 2="
    TablaConfina.TextMatrix(18, 0) = "Lo 3="

    
    TablaConfina.TextMatrix(0, 2) = "cm2"
    TablaConfina.TextMatrix(1, 2) = "cm2"
    TablaConfina.TextMatrix(2, 2) = "cm2"
    TablaConfina.TextMatrix(3, 2) = "cm2"
    TablaConfina.TextMatrix(4, 2) = "cm2"

    
    TablaConfina.TextMatrix(6, 2) = "cm"
    TablaConfina.TextMatrix(7, 2) = "cm"
    TablaConfina.TextMatrix(8, 2) = "cm"


    TablaConfina.TextMatrix(10, 2) = "cm"
    TablaConfina.TextMatrix(11, 2) = "cm"


    TablaConfina.TextMatrix(13, 2) = "cm"
    TablaConfina.TextMatrix(14, 2) = "cm"


    TablaConfina.TextMatrix(16, 2) = "cm"
    TablaConfina.TextMatrix(17, 2) = "cm"
    TablaConfina.TextMatrix(18, 2) = "cm"
      

    'Fin del procedimiento que acomoda los datos importados y el tamaño de las tablas.


    End Sub


    Private Sub Command1_Click()
    'Este procedimiento hace los cálculos del diseño por capacidad y confinamiento.
    

    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 6
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 11
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 13
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    End Select
    'Fin del errorhandler de no poner nada!
            
    
    'Diseño de columnas por confinamiento
    Ag = Tb2 * Th2
    Aaro = ((Trt / 8) * 2.54) ^ 2 * 3.14159265359 / 4
    Ash = TAros * Aaro
    hc = (Tb2 - (2 * Tdp2) + ((Trl / 8) * 2.54) + ((Trt / 8) * 2.54))
    Ahc = (Tb2 - (2 * Tdp2) + ((Trl / 8) * 2.54) + (2 * ((Trt / 8) * 2.54))) * (Th2 - (2 * Tdp2) + ((Trl / 8) * 2.54) + (2 * ((Trt / 8) * 2.54)))


    If Tb2 < Th2 Then
    smaxext1 = Tb2 * 0.25
    Else:
    smaxext1 = Th2 * 0.25
    End If


    smaxext2 = 6 * ((Trl / 8) * 2.54)


    If ((10.16 + ((35.56 - Thx) / 3))) > 15 Then
    smaxext3 = 15
    End If

    
    If ((10.16 + ((35.56 - Thx) / 3))) < 10 Then
    smaxext3 = 10
    End If


    If 15 > ((10.16 + ((35.56 - Thx) / 3))) < 10 Then
    smaxext3 = ((10.16 + ((35.56 - Thx) / 3)))
    End If


    smaxext4 = ((Ash) / (0.3 * hc * (TFc2 / TFyh2) * ((Ag / Ahc) - 1)))
    smaxext5 = (Ash) / (0.09 * hc * (TFc2 / TFyh2))


    smaxcentro1 = Round(6 * Trl, 2)
    smaxcentro2 = Round(6 * 2.54, 2)

    
    If Tb2 > Th2 Then
    lo1 = Tb2
    Else:
    lo1 = Th2
    End If


    lo2 = Talv * 100 / 6
    lo3 = 45


    TablaConfina.TextMatrix(0, 1) = Round(Ag, 2)
    TablaConfina.TextMatrix(1, 1) = Round(Ahc, 2)
    TablaConfina.TextMatrix(2, 1) = Round(Aaro, 2)
    TablaConfina.TextMatrix(3, 1) = Round(Ash, 2)
    TablaConfina.TextMatrix(4, 1) = Round(hc, 2)

    
    TablaConfina.TextMatrix(6, 1) = Round(smaxext1, 2)
    TablaConfina.TextMatrix(7, 1) = Round(smaxext2, 2)
    TablaConfina.TextMatrix(8, 1) = Round(smaxext3, 2)

    
    TablaConfina.TextMatrix(10, 1) = Round(smaxext4, 2)
    TablaConfina.TextMatrix(11, 1) = Round(smaxext5, 2)

    
    TablaConfina.TextMatrix(13, 1) = Round(smaxcentro1, 2)
    TablaConfina.TextMatrix(14, 1) = Round(smaxcentro2, 2)

    
    TablaConfina.TextMatrix(16, 1) = Round(lo1, 2)
    TablaConfina.TextMatrix(17, 1) = Round(lo2, 2)
    TablaConfina.TextMatrix(18, 1) = Round(lo3, 2)
    'Fin del diseño de columnas por confinamiento.


    'Diseño de columnas por Capacidad.
    
    Fc2 = Val(TFc2)
    b = Val(Tb2)
    d = Val(Td2)
    
    Vu = ((Val(TMcpi) + Val(TMcpj)) / Val(The))
    Vc = 0.53 * ((Fc2) ^ 0.5) * b * d * (1 / 1000)

    Pmaxcomp = Val(TPmaxcomp)
    Pmaxten = Val(TPmaxten)
      
      
    
    Fi = 0.85
    
    ash2 = TAros2 * Aaro
    Av = ash2
    
    If (Pmaxcomp * 1000) < ((Ag * Fc2) / 20) Then
    Vc = 0
    Else
    End If
       
       
    If Pmaxten > 0 Then
    Vc = 0
    Else
    End If
    
    
    
    Vs = (Vu - (Fi * Vc)) / Fi
    s = (Av * TFyh2 * Td2) / (Vs * 1000)


    smax1 = 6 * 2.54
    smax2 = 6 * ((Trl / 8) * 2.54)


    Tablacapa.TextMatrix(0, 1) = Round(Vu, 2)
    Tablacapa.TextMatrix(1, 1) = Round(Vc, 2)
    Tablacapa.TextMatrix(2, 1) = Round(Vs, 2)
    Tablacapa.TextMatrix(3, 1) = Round(Av, 2)
    Tablacapa.TextMatrix(4, 1) = Round(s, 2)

    
    Tablacapa.TextMatrix(6, 1) = Round(smax1, 2)
    Tablacapa.TextMatrix(7, 1) = Round(smax2, 2)
    'Fin del diseño de columnas por capacidad.
      
        
    'Anota los resultados finales de los diseños por confinamiento y capacidad.
    lomax = Val(lo1)

    
    If lomax < Val(lo2) Then
    lomax = Val(lo2)
    End If


    If lomax < Val(lo3) Then
    lomax = Val(lo3)
    End If


    T1 = Round(lomax, 2)
    T3 = Round((Talv * 100) - (lomax * 2), 2)
    T6 = Round(lomax, 2)



    'Diseño por cortante por resistencia.
    
    
    'Obtiene el cortante de diseño por resistencia.
    
    Vnmayor = Abs(Val(Tabla2.TextMatrix(15, 1)))
    
    
    If Vnmayor < Abs(Val(Tabla2.TextMatrix(15, 3))) Then
    Vnmayor = Abs(Val(Tabla2.TextMatrix(15, 3)))
    Else
    End If
    'Fin de obtener el cortante de diseño por resistencia.
    
        
    Fi = Val(TphiC)
    Vs = (Vnmayor - (Fi * Vc)) / Fi
    sanalisis = (Av * Val(TFyh2) * Val(Td2)) / (Vs * 1000)
    
    If sanalisis < 0 Then
    
    sanalisis = 0
    
    End If
    
    
    Tsanalisis = Round(sanalisis, 2)

    
    'Fin del diseño por cortante de análisis.

    
    sminmax = smaxext1


    If sminmax > smaxext1 Then
    sminmax = smaxext1
    End If


    If sminmax > smaxext2 Then
    sminmax = smaxext2
    End If


    If sminmax > smaxext3 Then
    sminmax = smaxext3
    End If


    If sminmax > smaxext4 Then
    sminmax = smaxext4
    End If


    If sminmax > smaxext5 Then
    sminmax = smaxext5
    End If


    If sminmax > s Then
    sminmax = s
    End If

    If sanalisis > 0 And sminmax > sanalisis Then
    sminmax = sanalisis
    End If

    sminmax2 = smax1
    If sminmax2 > smax2 Then
    sminmax2 = smax2
    End If


    If sminmax2 > smaxcentro1 Then
    sminmax2 = smaxcentro1
    End If

    
    If sminmax2 > smaxcentro2 Then
    sminmax2 = smaxcentro2
    End If


    If sminmax2 > s Then
    sminmax2 = s
    End If



    T2 = Round(sminmax, 2)
    T4 = Round(sminmax2, 2)
    T5 = Round(sminmax, 2)
    'Fin de anotar los resultados finales.

    'Fin del procedimiento hace los cálculos del diseño por capacidad y confinamiento.


    End Sub


    Sub Importacion()
    'Este procedimiento pone todos los valores de cargas y sus combinaciones.
 
    
    'Guarda la posicion del número del elemento.
    For i = 1 To MSHFlexGrid1.Rows - 1
    If Datos(i, 0) = CColumna Then
    Posicion = i
    i = MSHFlexGrid1.Rows - 1
    End If
    Next i
    'Fin de guardar la posicion del número del elemento.

    
    'Lee la longitud del elemento.
    TLong = Round(Datos(Posicion + 6, 5) * 1000, 2)
    'Fin de leer la longitud del elemento.


    'Inserta los datos de CM,CV,CS en la Tabla1.
    
    If CEje = "M3" Then
    
    'Carga Muerta.
    Tabla1.TextMatrix(2, 1) = Round(Datos(Posicion + 0, 3), 3)
    Tabla1.TextMatrix(2, 2) = Round(Datos(Posicion + 0, 4), 3)
    Tabla1.TextMatrix(2, 3) = Round(Datos(Posicion + 6, 3), 3)
    Tabla1.TextMatrix(2, 4) = Round(Datos(Posicion + 6, 4), 3)


    'Carga Viva.
    Tabla1.TextMatrix(3, 1) = Round(Datos(Posicion + 7, 3), 3)
    Tabla1.TextMatrix(3, 2) = Round(Datos(Posicion + 7, 4), 3)
    Tabla1.TextMatrix(3, 3) = Round(Datos(Posicion + 13, 3), 3)
    Tabla1.TextMatrix(3, 4) = Round(Datos(Posicion + 13, 4), 3)


    'Sismo en X.
    Tabla1.TextMatrix(4, 1) = Round(Datos(Posicion + 14, 3), 3)
    Tabla1.TextMatrix(4, 2) = Round(Datos(Posicion + 14, 4), 3)
    Tabla1.TextMatrix(4, 3) = Round(Datos(Posicion + 20, 3), 3)
    Tabla1.TextMatrix(4, 4) = Round(Datos(Posicion + 20, 4), 3)
    
    
    'Sismo en Y.
    Tabla1.TextMatrix(5, 1) = Round(Datos(Posicion + 21, 3), 3)
    Tabla1.TextMatrix(5, 2) = Round(Datos(Posicion + 21, 4), 3)
    Tabla1.TextMatrix(5, 3) = Round(Datos(Posicion + 27, 3), 3)
    Tabla1.TextMatrix(5, 4) = Round(Datos(Posicion + 27, 4), 3)
    
    
    Else:
    
    'Carga Muerta.
    Tabla1.TextMatrix(2, 1) = Round(Datos(Posicion + 0, 2), 3)
    Tabla1.TextMatrix(2, 2) = Round(Datos(Posicion + 0, 4), 3)
    Tabla1.TextMatrix(2, 3) = Round(Datos(Posicion + 6, 2), 3)
    Tabla1.TextMatrix(2, 4) = Round(Datos(Posicion + 6, 4), 3)


    'Carga Viva.
    Tabla1.TextMatrix(3, 1) = Round(Datos(Posicion + 7, 2), 3)
    Tabla1.TextMatrix(3, 2) = Round(Datos(Posicion + 7, 4), 3)
    Tabla1.TextMatrix(3, 3) = Round(Datos(Posicion + 13, 2), 3)
    Tabla1.TextMatrix(3, 4) = Round(Datos(Posicion + 13, 4), 3)


    'Sismo en X.
    Tabla1.TextMatrix(4, 1) = Round(Datos(Posicion + 14, 2), 3)
    Tabla1.TextMatrix(4, 2) = Round(Datos(Posicion + 14, 4), 3)
    Tabla1.TextMatrix(4, 3) = Round(Datos(Posicion + 20, 2), 3)
    Tabla1.TextMatrix(4, 4) = Round(Datos(Posicion + 20, 4), 3)
    
    
    'Sismo en Y.
    Tabla1.TextMatrix(5, 1) = Round(Datos(Posicion + 21, 2), 3)
    Tabla1.TextMatrix(5, 2) = Round(Datos(Posicion + 21, 4), 3)
    Tabla1.TextMatrix(5, 3) = Round(Datos(Posicion + 27, 2), 3)
    Tabla1.TextMatrix(5, 4) = Round(Datos(Posicion + 27, 4), 3)
    
    End If
    'Fin de insertar los datos.

    
    'Calcula y anota las combinaciones en la Tabla1.
    'COMBO 1.
    Tabla1.TextMatrix(6, 1) = Round(1.4 * Val(Tabla1.TextMatrix(2, 1)) + (1.7 * Val(Tabla1.TextMatrix(3, 1))), 3)
    Tabla1.TextMatrix(6, 2) = Round(1.4 * Val(Tabla1.TextMatrix(2, 2)) + (1.7 * Val(Tabla1.TextMatrix(3, 2))), 3)
    Tabla1.TextMatrix(6, 3) = Round(1.4 * Val(Tabla1.TextMatrix(2, 3)) + (1.7 * Val(Tabla1.TextMatrix(3, 3))), 3)
    Tabla1.TextMatrix(6, 4) = Round(1.4 * Val(Tabla1.TextMatrix(2, 4)) + (1.7 * Val(Tabla1.TextMatrix(3, 4))), 3)


    'COMBO 2 en X.
    Tabla1.TextMatrix(7, 1) = Round(0.75 * Val(Tabla1.TextMatrix(6, 1)) + Val(Tabla1.TextMatrix(4, 1)), 3)
    Tabla1.TextMatrix(7, 2) = Round(0.75 * Val(Tabla1.TextMatrix(6, 2)) + Val(Tabla1.TextMatrix(4, 2)), 3)
    Tabla1.TextMatrix(7, 3) = Round(0.75 * Val(Tabla1.TextMatrix(6, 3)) + Val(Tabla1.TextMatrix(4, 3)), 3)
    Tabla1.TextMatrix(7, 4) = Round(0.75 * Val(Tabla1.TextMatrix(6, 4)) + Val(Tabla1.TextMatrix(4, 4)), 3)

    
    'COMBO 2 en Y.
    Tabla1.TextMatrix(8, 1) = Round(0.75 * Val(Tabla1.TextMatrix(6, 1)) + Val(Tabla1.TextMatrix(5, 1)), 3)
    Tabla1.TextMatrix(8, 2) = Round(0.75 * Val(Tabla1.TextMatrix(6, 2)) + Val(Tabla1.TextMatrix(5, 2)), 3)
    Tabla1.TextMatrix(8, 3) = Round(0.75 * Val(Tabla1.TextMatrix(6, 3)) + Val(Tabla1.TextMatrix(5, 3)), 3)
    Tabla1.TextMatrix(8, 4) = Round(0.75 * Val(Tabla1.TextMatrix(6, 4)) + Val(Tabla1.TextMatrix(5, 4)), 3)

    
    'COMBO 3 en X.
    Tabla1.TextMatrix(9, 1) = Round(0.75 * Val(Tabla1.TextMatrix(6, 1)) - Val(Tabla1.TextMatrix(4, 1)), 3)
    Tabla1.TextMatrix(9, 2) = Round(0.75 * Val(Tabla1.TextMatrix(6, 2)) - Val(Tabla1.TextMatrix(4, 2)), 3)
    Tabla1.TextMatrix(9, 3) = Round(0.75 * Val(Tabla1.TextMatrix(6, 3)) - Val(Tabla1.TextMatrix(4, 3)), 3)
    Tabla1.TextMatrix(9, 4) = Round(0.75 * Val(Tabla1.TextMatrix(6, 4)) - Val(Tabla1.TextMatrix(4, 4)), 3)

    
    'COMBO 3 en Y.
    Tabla1.TextMatrix(10, 1) = Round(0.75 * Val(Tabla1.TextMatrix(6, 1)) - Val(Tabla1.TextMatrix(5, 1)), 3)
    Tabla1.TextMatrix(10, 2) = Round(0.75 * Val(Tabla1.TextMatrix(6, 2)) - Val(Tabla1.TextMatrix(5, 2)), 3)
    Tabla1.TextMatrix(10, 3) = Round(0.75 * Val(Tabla1.TextMatrix(6, 3)) - Val(Tabla1.TextMatrix(5, 3)), 3)
    Tabla1.TextMatrix(10, 4) = Round(0.75 * Val(Tabla1.TextMatrix(6, 4)) - Val(Tabla1.TextMatrix(5, 4)), 3)

    
    'COMBO 4 en X.
    Tabla1.TextMatrix(11, 1) = Round(0.95 * Val(Tabla1.TextMatrix(2, 1)) + Val(Tabla1.TextMatrix(4, 1)), 3)
    Tabla1.TextMatrix(11, 2) = Round(0.95 * Val(Tabla1.TextMatrix(2, 2)) + Val(Tabla1.TextMatrix(4, 2)), 3)
    Tabla1.TextMatrix(11, 3) = Round(0.95 * Val(Tabla1.TextMatrix(2, 3)) + Val(Tabla1.TextMatrix(4, 3)), 3)
    Tabla1.TextMatrix(11, 4) = Round(0.95 * Val(Tabla1.TextMatrix(2, 4)) + Val(Tabla1.TextMatrix(4, 4)), 3)

    
    'COMBO 4 en Y.
    Tabla1.TextMatrix(12, 1) = Round(0.95 * Val(Tabla1.TextMatrix(2, 1)) + Val(Tabla1.TextMatrix(5, 1)), 3)
    Tabla1.TextMatrix(12, 2) = Round(0.95 * Val(Tabla1.TextMatrix(2, 2)) + Val(Tabla1.TextMatrix(5, 2)), 3)
    Tabla1.TextMatrix(12, 3) = Round(0.95 * Val(Tabla1.TextMatrix(2, 3)) + Val(Tabla1.TextMatrix(5, 3)), 3)
    Tabla1.TextMatrix(12, 4) = Round(0.95 * Val(Tabla1.TextMatrix(2, 4)) + Val(Tabla1.TextMatrix(5, 4)), 3)

    
    'COMBO 5 en X.
    Tabla1.TextMatrix(13, 1) = Round(0.95 * Val(Tabla1.TextMatrix(2, 1)) - Val(Tabla1.TextMatrix(4, 1)), 3)
    Tabla1.TextMatrix(13, 2) = Round(0.95 * Val(Tabla1.TextMatrix(2, 2)) - Val(Tabla1.TextMatrix(4, 2)), 3)
    Tabla1.TextMatrix(13, 3) = Round(0.95 * Val(Tabla1.TextMatrix(2, 3)) - Val(Tabla1.TextMatrix(4, 3)), 3)
    Tabla1.TextMatrix(13, 4) = Round(0.95 * Val(Tabla1.TextMatrix(2, 4)) - Val(Tabla1.TextMatrix(4, 4)), 3)

    
    'COMBO 5 en Y.
    Tabla1.TextMatrix(14, 1) = Round(0.95 * Val(Tabla1.TextMatrix(2, 1)) - Val(Tabla1.TextMatrix(5, 1)), 3)
    Tabla1.TextMatrix(14, 2) = Round(0.95 * Val(Tabla1.TextMatrix(2, 2)) - Val(Tabla1.TextMatrix(5, 2)), 3)
    Tabla1.TextMatrix(14, 3) = Round(0.95 * Val(Tabla1.TextMatrix(2, 3)) - Val(Tabla1.TextMatrix(5, 3)), 3)
    Tabla1.TextMatrix(14, 4) = Round(0.95 * Val(Tabla1.TextMatrix(2, 4)) - Val(Tabla1.TextMatrix(5, 4)), 3)
    'Fin de calcular y anotar las combinaciones en la Tabla1.

    
    'Máximos.
    For j = 1 To 4
    Valmax = Val(Tabla1.TextMatrix(6, j))
    For i = 6 To 14
    If Abs(Val(Tabla1.TextMatrix(i, j))) > Abs(Valmax) Then
    Valmax = Val(Tabla1.TextMatrix(i, j))
    End If
    Next i
    Tabla1.TextMatrix(15, j) = Round(Valmax, 3)
    Next j
    'Fin de sacar y poner los máximos.
     
    
    'Inserta los datos de CM,CV,CS en la tabla2.
    
    If CEje = "M3" Then
    
    'Carga Muerta.
    Tabla2.TextMatrix(2, 1) = Round(Datos(Posicion + 0, 7), 3)
    Tabla2.TextMatrix(2, 2) = Round(Datos(Posicion + 0, 4), 3)
    Tabla2.TextMatrix(2, 3) = Round(Datos(Posicion + 6, 7), 3)
    Tabla2.TextMatrix(2, 4) = Round(Datos(Posicion + 6, 4), 3)
    
    
    'Carga Viva.
    Tabla2.TextMatrix(3, 1) = Round(Datos(Posicion + 7, 7), 3)
    Tabla2.TextMatrix(3, 2) = Round(Datos(Posicion + 7, 4), 3)
    Tabla2.TextMatrix(3, 3) = Round(Datos(Posicion + 13, 7), 3)
    Tabla2.TextMatrix(3, 4) = Round(Datos(Posicion + 13, 4), 3)
    
    
    'Sismo en X.
    Tabla2.TextMatrix(4, 1) = Round(Datos(Posicion + 14, 7), 3)
    Tabla2.TextMatrix(4, 2) = Round(Datos(Posicion + 14, 4), 3)
    Tabla2.TextMatrix(4, 3) = Round(Datos(Posicion + 20, 7), 3)
    Tabla2.TextMatrix(4, 4) = Round(Datos(Posicion + 20, 4), 3)
    
    
    'Sismo en Y.
    Tabla2.TextMatrix(5, 1) = Round(Datos(Posicion + 21, 7), 3)
    Tabla2.TextMatrix(5, 2) = Round(Datos(Posicion + 21, 4), 3)
    Tabla2.TextMatrix(5, 3) = Round(Datos(Posicion + 27, 7), 3)
    Tabla2.TextMatrix(5, 4) = Round(Datos(Posicion + 27, 4), 3)
    
    
    Else:
    
    
    'Carga Muerta.
    Tabla2.TextMatrix(2, 1) = Round(Datos(Posicion + 0, 8), 3)
    Tabla2.TextMatrix(2, 2) = Round(Datos(Posicion + 0, 4), 3)
    Tabla2.TextMatrix(2, 3) = Round(Datos(Posicion + 6, 8), 3)
    Tabla2.TextMatrix(2, 4) = Round(Datos(Posicion + 6, 4), 3)
    
    
    'Carga Viva.
    Tabla2.TextMatrix(3, 1) = Round(Datos(Posicion + 7, 8), 3)
    Tabla2.TextMatrix(3, 2) = Round(Datos(Posicion + 7, 4), 3)
    Tabla2.TextMatrix(3, 3) = Round(Datos(Posicion + 13, 8), 3)
    Tabla2.TextMatrix(3, 4) = Round(Datos(Posicion + 13, 4), 3)
    
    
    'Sismo en X.
    Tabla2.TextMatrix(4, 1) = Round(Datos(Posicion + 14, 8), 3)
    Tabla2.TextMatrix(4, 2) = Round(Datos(Posicion + 14, 4), 3)
    Tabla2.TextMatrix(4, 3) = Round(Datos(Posicion + 20, 8), 3)
    Tabla2.TextMatrix(4, 4) = Round(Datos(Posicion + 20, 4), 3)
    
    
    'Sismo en Y.
    Tabla2.TextMatrix(5, 1) = Round(Datos(Posicion + 21, 8), 3)
    Tabla2.TextMatrix(5, 2) = Round(Datos(Posicion + 21, 4), 3)
    Tabla2.TextMatrix(5, 3) = Round(Datos(Posicion + 27, 8), 3)
    Tabla2.TextMatrix(5, 4) = Round(Datos(Posicion + 27, 4), 3)
    
 
    End If

    'Fin de insertar los datos.
  
      
    'Calcula y anota las combinaciones en la tabla2.
    'COMBO 1.
    Tabla2.TextMatrix(6, 1) = Round(1.4 * Val(Tabla2.TextMatrix(2, 1)) + (1.7 * Val(Tabla2.TextMatrix(3, 1))), 3)
    Tabla2.TextMatrix(6, 2) = Round(1.4 * Val(Tabla2.TextMatrix(2, 2)) + (1.7 * Val(Tabla2.TextMatrix(3, 2))), 3)
    Tabla2.TextMatrix(6, 3) = Round(1.4 * Val(Tabla2.TextMatrix(2, 3)) + (1.7 * Val(Tabla2.TextMatrix(3, 3))), 3)
    Tabla2.TextMatrix(6, 4) = Round(1.4 * Val(Tabla2.TextMatrix(2, 4)) + (1.7 * Val(Tabla2.TextMatrix(3, 4))), 3)
    
    
    'COMBO 2 en X.
    Tabla2.TextMatrix(7, 1) = Round(0.75 * Val(Tabla2.TextMatrix(6, 1)) + Val(Tabla2.TextMatrix(4, 1)), 3)
    Tabla2.TextMatrix(7, 2) = Round(0.75 * Val(Tabla2.TextMatrix(6, 2)) + Val(Tabla2.TextMatrix(4, 2)), 3)
    Tabla2.TextMatrix(7, 3) = Round(0.75 * Val(Tabla2.TextMatrix(6, 3)) + Val(Tabla2.TextMatrix(4, 3)), 3)
    Tabla2.TextMatrix(7, 4) = Round(0.75 * Val(Tabla2.TextMatrix(6, 4)) + Val(Tabla2.TextMatrix(4, 4)), 3)
    
    
    'COMBO 2 en Y.
    Tabla2.TextMatrix(8, 1) = Round(0.75 * Val(Tabla2.TextMatrix(6, 1)) + Val(Tabla2.TextMatrix(5, 1)), 3)
    Tabla2.TextMatrix(8, 2) = Round(0.75 * Val(Tabla2.TextMatrix(6, 2)) + Val(Tabla2.TextMatrix(5, 2)), 3)
    Tabla2.TextMatrix(8, 3) = Round(0.75 * Val(Tabla2.TextMatrix(6, 3)) + Val(Tabla2.TextMatrix(5, 3)), 3)
    Tabla2.TextMatrix(8, 4) = Round(0.75 * Val(Tabla2.TextMatrix(6, 4)) + Val(Tabla2.TextMatrix(5, 4)), 3)
    
    
    'COMBO 3 en X.
    Tabla2.TextMatrix(9, 1) = Round(0.75 * Val(Tabla2.TextMatrix(6, 1)) - Val(Tabla2.TextMatrix(4, 1)), 3)
    Tabla2.TextMatrix(9, 2) = Round(0.75 * Val(Tabla2.TextMatrix(6, 2)) - Val(Tabla2.TextMatrix(4, 2)), 3)
    Tabla2.TextMatrix(9, 3) = Round(0.75 * Val(Tabla2.TextMatrix(6, 3)) - Val(Tabla2.TextMatrix(4, 3)), 3)
    Tabla2.TextMatrix(9, 4) = Round(0.75 * Val(Tabla2.TextMatrix(6, 4)) - Val(Tabla2.TextMatrix(4, 4)), 3)
    
    
    'COMBO 3 en Y.
    Tabla2.TextMatrix(10, 1) = Round(0.75 * Val(Tabla2.TextMatrix(6, 1)) - Val(Tabla2.TextMatrix(5, 1)), 3)
    Tabla2.TextMatrix(10, 2) = Round(0.75 * Val(Tabla2.TextMatrix(6, 2)) - Val(Tabla2.TextMatrix(5, 2)), 3)
    Tabla2.TextMatrix(10, 3) = Round(0.75 * Val(Tabla2.TextMatrix(6, 3)) - Val(Tabla2.TextMatrix(5, 3)), 3)
    Tabla2.TextMatrix(10, 4) = Round(0.75 * Val(Tabla2.TextMatrix(6, 4)) - Val(Tabla2.TextMatrix(5, 4)), 3)
    
    
    'COMBO 4 en X.
    Tabla2.TextMatrix(11, 1) = Round(0.95 * Val(Tabla2.TextMatrix(2, 1)) + Val(Tabla2.TextMatrix(4, 1)), 3)
    Tabla2.TextMatrix(11, 2) = Round(0.95 * Val(Tabla2.TextMatrix(2, 2)) + Val(Tabla2.TextMatrix(4, 2)), 3)
    Tabla2.TextMatrix(11, 3) = Round(0.95 * Val(Tabla2.TextMatrix(2, 3)) + Val(Tabla2.TextMatrix(4, 3)), 3)
    Tabla2.TextMatrix(11, 4) = Round(0.95 * Val(Tabla2.TextMatrix(2, 4)) + Val(Tabla2.TextMatrix(4, 4)), 3)
    
    
    'COMBO 4 en Y.
    Tabla2.TextMatrix(12, 1) = Round(0.95 * Val(Tabla2.TextMatrix(2, 1)) + Val(Tabla2.TextMatrix(5, 1)), 3)
    Tabla2.TextMatrix(12, 2) = Round(0.95 * Val(Tabla2.TextMatrix(2, 2)) + Val(Tabla2.TextMatrix(5, 2)), 3)
    Tabla2.TextMatrix(12, 3) = Round(0.95 * Val(Tabla2.TextMatrix(2, 3)) + Val(Tabla2.TextMatrix(5, 3)), 3)
    Tabla2.TextMatrix(12, 4) = Round(0.95 * Val(Tabla2.TextMatrix(2, 4)) + Val(Tabla2.TextMatrix(5, 4)), 3)
    
    
    'COMBO 5 en X.
    Tabla2.TextMatrix(13, 1) = Round(0.95 * Val(Tabla2.TextMatrix(2, 1)) - Val(Tabla2.TextMatrix(4, 1)), 3)
    Tabla2.TextMatrix(13, 2) = Round(0.95 * Val(Tabla2.TextMatrix(2, 2)) - Val(Tabla2.TextMatrix(4, 2)), 3)
    Tabla2.TextMatrix(13, 3) = Round(0.95 * Val(Tabla2.TextMatrix(2, 3)) - Val(Tabla2.TextMatrix(4, 3)), 3)
    Tabla2.TextMatrix(13, 4) = Round(0.95 * Val(Tabla2.TextMatrix(2, 4)) - Val(Tabla2.TextMatrix(4, 4)), 3)
    
    
    'COMBO 5 en Y.
    Tabla2.TextMatrix(14, 1) = Round(0.95 * Val(Tabla2.TextMatrix(2, 1)) - Val(Tabla2.TextMatrix(5, 1)), 3)
    Tabla2.TextMatrix(14, 2) = Round(0.95 * Val(Tabla2.TextMatrix(2, 2)) - Val(Tabla2.TextMatrix(5, 2)), 3)
    Tabla2.TextMatrix(14, 3) = Round(0.95 * Val(Tabla2.TextMatrix(2, 3)) - Val(Tabla2.TextMatrix(5, 3)), 3)
    Tabla2.TextMatrix(14, 4) = Round(0.95 * Val(Tabla2.TextMatrix(2, 4)) - Val(Tabla2.TextMatrix(5, 4)), 3)
    'Fin de calcular y anotar las combinaciones en la tabla2.
    
    
    'Máximos.
    For j = 1 To 4
    Valmax = Val(Tabla2.TextMatrix(6, j))
    For i = 6 To 14
    If Abs(Val(Tabla2.TextMatrix(i, j))) > Abs(Valmax) Then
    Valmax = Val(Tabla2.TextMatrix(i, j))
    End If
    Next i
    Tabla2.TextMatrix(15, j) = Round(Valmax, 3)
    Next j
    'Fin de sacar y poner los máximos.
    
    
     
    'Carga axial máxima a compresion.
    
    Valmaxcomp = Val(Tabla1.TextMatrix(7, 2))
        
    For i = 7 To 14
    If Val(Tabla1.TextMatrix(i, 2)) < Val(Valmaxcomp) Then
    Valmaxcomp = Val(Tabla1.TextMatrix(i, 2))
    End If
    Next i
    
    
    For i = 7 To 14
    If Val(Tabla1.TextMatrix(i, 4)) < Val(Valmaxcomp) Then
    Valmaxcomp = Val(Tabla1.TextMatrix(i, 4))
    End If
    Next i
    
    If Valmaxcomp > 0 Then
    Valmaxcomp = 0
    End If
    
    TPmaxcomp = Round(Abs(Valmaxcomp), 3)
    
    'Fin de sacar la carga axial máxima a compresion.
    
    
    'Carga axial máxima a tension.
    
    Valmaxten = Val(Tabla1.TextMatrix(7, 2))
        
    For i = 7 To 14
    If Val(Tabla1.TextMatrix(i, 2)) > Val(Valmaxten) Then
    Valmaxten = Val(Tabla1.TextMatrix(i, 2))
    End If
    Next i
    
    
    For i = 7 To 14
    If Val(Tabla1.TextMatrix(i, 4)) > Val(Valmaxten) Then
    Valmaxten = Val(Tabla1.TextMatrix(i, 4))
    End If
    Next i
    
    If Valmaxten < 0 Then
    Valmaxten = 0
    End If
        
    TPmaxten = Round(Abs(Valmaxten), 3)
        
    'Fin de sacar la carga axial máxima a tension.
            

    'Fin del procedimiento que pone todos los valores de cargas y sus combinaciones.
    
    
    End Sub


    Private Sub Diagrama_Click()
    'Este procedimiento actualiza el diagrama de interacción.
    
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 6
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 11
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 13
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    End Select
    'Fin del errorhandler de no poner nada!

    
    CColumna_Click
    
    
    'Pasa todos los datos a un gráfico de EXCEL.
    For i = 0 To puntos
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 1).Value = fMn2(i) / 100000
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 2).Value = -fPn2(i) / 1000
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 3).Value = Mn2(i) / 100000
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 4).Value = -Pn2(i) / 1000
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 5).Value = Mnc2(i) / 100000
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 6).Value = -Pnc2(i) / 1000
    Next i

    
    For i = 0 To 1
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 7).Value = (Abs(Mnp2(i))) / 100000
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 8).Value = -Pnp2(i) / 1000
    Next i


    For i = 2 To 17
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 7).Value = Abs(Mnp2(i)) / 100000
        OLE1.object.Worksheets("Sheet1").Cells(i + 2, 8).Value = -Pnp2(i) / 1000
    Next i
 
  
        OLE1.object.Worksheets("Sheet1").Cells(2, 9).Value = 0
        OLE1.object.Worksheets("Sheet1").Cells(3, 9).Value = Mmax

        
        OLE1.object.Worksheets("Sheet1").Cells(2, 10).Value = -Pu / 1000
        OLE1.object.Worksheets("Sheet1").Cells(3, 10).Value = -Pu / 1000
  
  
        OLE1.object.Worksheets("Sheet1").Cells(1, 8).Value = "Puntos Pn"
        OLE1.object.Worksheets("Sheet1").Cells(1, 7).Value = "Puntos Mn"
        OLE1.object.Worksheets("Sheet1").Cells(1, 6).Value = "Pnc"
        OLE1.object.Worksheets("Sheet1").Cells(1, 5).Value = "Mnc"
        OLE1.object.Worksheets("Sheet1").Cells(1, 4).Value = "Pn"
        OLE1.object.Worksheets("Sheet1").Cells(1, 3).Value = "Mn"
        OLE1.object.Worksheets("Sheet1").Cells(1, 2).Value = "fPn"
        OLE1.object.Worksheets("Sheet1").Cells(1, 1).Value = "fMn"
    'Fin de pasar todos los datos a un gráfico de EXCEL.

    
    'Fin del procedimiento que actualiza el diagrama de interacción.
    
    
    End Sub


    Private Sub CColumna_Click()
    'Este procedimiento realiza todos los cálculos de flexo-compresion.
    
    
    TFecha = Format$(Now, "d / m / yyyy")
    
    
    THora = Format$(Now, "h:mm AM/PM")
    
    
    Importacion

    
    Dim Ass(5), EEs(5), Fs(5), dx(5), Fcapa(5), Pn(), Mn(), fPn(), fMn() As Variant
    Dim Assb(5), EEsb(5), Fsb(5), dxb(5), Fcapab(5) As Variant
    Dim Assc(5), EEsc(5), Fsc(5), dxc(5), Fcapac(5), Pnc(), Mnc() As Variant
    Dim Mnp(17), Pnp(17) As Variant
    Dim Tdiam(5), Tnum(5), Tgato(5) As Variant
    
    puntos = 15
    ReDim Pn(puntos), Mn(puntos), fPn(puntos), fMn(puntos), Pnc(puntos), Mnc(puntos)
 
    'Inicio del cálculo del Diagrama de Interacción.
    Es = Val(TEs)
    Euc = Val(TEuc)
    Eus = Val(TFy) / Es
    TEus = Round(Eus, 4)


    Fc = Val(TFc)
    Fy = Val(TFy)
    b = Val(Tb)


    'Cálculo de B1.
    If Fc <= 280 Then
    B1 = 0.85
    Else
    B1 = (0.85 - 0.05 * (Fc - 280) / (70))
    End If
        If Fc >= 560 Then
    B1 = 0.65
    Else
    End If
    'Fin del cálculo de B1.


    h = Val(Th)
    deltac = ((h) - 6) / (puntos)

    
    Limite = -0.1 * Fc * b * h / 0.7


    c = h / 10

    
    For j = 1 To puntos
    a = B1 * c

    
    For i = 0 To 5
    dx(i) = Val(Tdx1(i))

    
    Ecu = Val(TEuc)
    EEs(i) = Ecu * ((dx(i) / c) - 1)
    Fs(i) = EEs(i) * Es
    If Fs(i) > Fy Then
    Fs(i) = Fy
    End If
    If Fs(i) < -Fy Then
    Fs(i) = -Fy
    End If

    
    Ass(i) = Val(TAs1(i))
    Next i


    Ast = Ass(0) + Ass(1) + Ass(2) + Ass(3) + Ass(4) + Ass(5)
    Pa = 100 * Ast / (b * h)
    Tpa = Round(Pa, 2)
    If Pa > 1 And Pa < 6 Then
    TOK = "OK!"
    Else
    TOK = "NO OK!"
    End If

    
    Pconcreto = -(0.85 * Fc * a * b)
    
    
    For i = 0 To 5
    Fcapa(i) = (Ass(i) * Fs(i))
    Next i
    
    
    dconcreto = ((a / 2) - (h / 2))
    Pn(j) = Pconcreto
    Mn(j) = Pconcreto * dconcreto
    
    
    For i = 0 To 5
    Pn(j) = (Fcapa(i) + Pn(j))
    Mn(j) = ((Fcapa(i) * (dx(i) - (h / 2))) + Mn(j))
    Next i
    
    
    Select Case Pn(j)
        Case Is < Limite
            fPn(j) = Pn(j) * 0.7
            fMn(j) = Mn(j) * 0.7
        Case Is > 0
            fPn(j) = Pn(j) * 0.9
            fMn(j) = Mn(j) * 0.9
        Case Else
            fPn(j) = (0.9 - (0.2 * Pn(j)) / Limite) * Pn(j)
            fMn(j) = (0.9 - (0.2 * Pn(j)) / Limite) * Mn(j)
    End Select
        
    
    c = c + deltac
    Next j
    
    
    Pn(puntos) = -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fy))
    Mn(puntos) = 0
    
    
    Pn(0) = (Ast * Fy)
    Mn(0) = 0
    
    
    fPn(0) = (Ast * Fy) * 0.9
    fMn(0) = 0
    fPn(puntos) = 0.7 * -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fy))
    fMn(puntos) = 0
    
    
    Mmax = Mn(0)
    For i = 1 To puntos
    If Mmax < Mn(i) Then
    Mmax = Mn(i)
    End If
    Next i
    
    
    Pu = -0.8 * 0.7 * (0.85 * (b * h - Ast) * Fc + Ast * TFy)
     
     
    'Fin del diagrama de interacción.
    
    
    'Inico del cálculo del punto de falla balanceada.
    dmax = dx(0)
    For i = 1 To 5
        If dmax < dx(i) Then dmax = dx(i)
    Next i
    
    
    cb = (Euc * dmax) / (Euc + (Fy / Es))
    a = B1 * cb
    
    
    For i = 0 To 5
    dxb(i) = Val(Tdx1(i))
    
    
    EEsb(i) = Euc * ((dxb(i) / cb) - 1)
    Fsb(i) = EEsb(i) * Es
    If Fsb(i) >= Fy Then
    Fsb(i) = Fy
    End If
    If Fsb(i) < -Fy Then
    Fsb(i) = -Fy
    End If
    
    
    Next i
    
    
    Pconcreto = -(0.85 * Fc * a * b)
    
    
    For i = 0 To 5
    Fcapab(i) = (Ass(i) * Fsb(i))
    Next i
    
    
    dconcreto = ((a / 2) - (h / 2))
    Pnb = Pconcreto
    Mnb = Pconcreto * dconcreto
    
    
    For i = 0 To 5
    Pnb = (Fcapab(i) + Pnb)
    Mnb = ((Fcapab(i) * (dxb(i) - (h / 2))) + Mnb)
    Next i
    'Fin del cálculo del punto de falla balanceada.
        
        
    'Inicio del cálculo del Diagrama de Interacción utilizando 1.25 Fy.
    Es = Val(TEs)
    Euc = Val(TEuc)
    Eus = Val(TFy) / Es
    
    
    Fc = Val(TFc)
    Fyc = (Val(TFy) * 1.25)
    b = Val(Tb)
    
    
    c = h / 10
    For j = 1 To puntos
    a = B1 * c
    
    
    For i = 0 To 5
    dxc(i) = Val(Tdx1(i))
    
    
    Ecu = Val(TEuc)
    EEsc(i) = Ecu * ((dxc(i) / c) - 1)
    Fsc(i) = EEsc(i) * Es
    If Fsc(i) >= Fyc Then
    Fsc(i) = Fyc
    End If
    
    
    If Fsc(i) < -Fyc Then
    Fsc(i) = -Fyc
    End If
    
    
    Assc(i) = Val(TAs1(i))
    Next i
    
    
    Ast = Ass(0) + Ass(1) + Ass(2) + Ass(3) + Ass(4) + Ass(5)
    Pa = 100 * Ast / (b * h)
    Tpa = Round(Pa, 2)
    
    
    Pconcreto = -(0.85 * Fc * a * b)
    
    
    For i = 0 To 5
    Fcapac(i) = (Assc(i) * Fsc(i))
    Next i
    
    
    dconcreto = ((a / 2) - (h / 2))
    Pnc(j) = Pconcreto
    Mnc(j) = Pconcreto * dconcreto
    
    
    For i = 0 To 5
    Pnc(j) = (Fcapac(i) + Pnc(j))
    Mnc(j) = ((Fcapac(i) * (dxc(i) - (h / 2))) + Mnc(j))
    Next i
    
    
    c = c + deltac
    Next j
    
    
    Pnc(puntos) = -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fyc))
    Mnc(puntos) = 0
    
    
    Pnc(0) = (Ast * Fyc)
    Mnc(0) = 0

    
    'Fin del cálculo del Diagrama de Interacción utilizando 1.25 Fy.
    
    
    'Grafica los puntos del diagrama.
    
    Mnp(0) = Abs(Val(Tabla1.TextMatrix(6, 1)) * 100000)
    Pnp(0) = Val(Tabla1.TextMatrix(6, 2)) * 1000
    Mnp(1) = Abs(Val(Tabla1.TextMatrix(6, 3)) * 100000)
    Pnp(1) = Val(Tabla1.TextMatrix(6, 4)) * 1000
    Mnp(2) = Abs(Val(Tabla1.TextMatrix(7, 1)) * 100000)
    Pnp(2) = Val(Tabla1.TextMatrix(7, 2)) * 1000
    Mnp(3) = Abs(Val(Tabla1.TextMatrix(7, 3)) * 100000)
    Pnp(3) = Val(Tabla1.TextMatrix(7, 4)) * 1000
    Mnp(4) = Abs(Val(Tabla1.TextMatrix(8, 1)) * 100000)
    Pnp(4) = Val(Tabla1.TextMatrix(8, 2)) * 1000
    Mnp(5) = Abs(Val(Tabla1.TextMatrix(8, 3)) * 100000)
    Pnp(5) = Val(Tabla1.TextMatrix(8, 4)) * 1000
    Mnp(6) = Abs(Val(Tabla1.TextMatrix(9, 1)) * 100000)
    Pnp(6) = Val(Tabla1.TextMatrix(9, 2)) * 1000
    Mnp(7) = Abs(Val(Tabla1.TextMatrix(9, 3)) * 100000)
    Pnp(7) = Val(Tabla1.TextMatrix(9, 4)) * 1000
    Mnp(8) = Abs(Val(Tabla1.TextMatrix(10, 1)) * 100000)
    Pnp(8) = Val(Tabla1.TextMatrix(10, 2)) * 1000
    Mnp(9) = Abs(Val(Tabla1.TextMatrix(10, 3)) * 100000)
    Pnp(9) = Val(Tabla1.TextMatrix(10, 4)) * 1000
    Mnp(10) = Abs(Val(Tabla1.TextMatrix(11, 1)) * 100000)
    Pnp(10) = Val(Tabla1.TextMatrix(11, 2)) * 1000
    Mnp(11) = Abs(Val(Tabla1.TextMatrix(11, 3)) * 100000)
    Pnp(11) = Val(Tabla1.TextMatrix(11, 4)) * 1000
    Mnp(12) = Abs(Val(Tabla1.TextMatrix(12, 1)) * 100000)
    Pnp(12) = Val(Tabla1.TextMatrix(12, 2)) * 1000
    Mnp(13) = Abs(Val(Tabla1.TextMatrix(12, 3)) * 100000)
    Pnp(13) = Val(Tabla1.TextMatrix(12, 4)) * 1000
    Mnp(14) = Abs(Val(Tabla1.TextMatrix(13, 1)) * 100000)
    Pnp(14) = Val(Tabla1.TextMatrix(13, 2)) * 1000
    Mnp(15) = Abs(Val(Tabla1.TextMatrix(13, 3)) * 100000)
    Pnp(15) = Val(Tabla1.TextMatrix(13, 4)) * 1000
    Mnp(16) = Abs(Val(Tabla1.TextMatrix(14, 1)) * 100000)
    Pnp(16) = Val(Tabla1.TextMatrix(14, 2)) * 1000
    Mnp(17) = Abs(Val(Tabla1.TextMatrix(14, 3)) * 100000)
    Pnp(17) = Val(Tabla1.TextMatrix(14, 4)) * 1000
    
    
    For i = 0 To puntos
    Pn2(i) = Pn(i)
    Mn2(i) = Mn(i)
    
    
    fPn2(i) = fPn(i)
    fMn2(i) = fMn(i)
    
    
    Pnc2(i) = Pnc(i)
    Mnc2(i) = Mnc(i)
    
    
    Next i
    
    
    For i = 0 To 17
    Pnp2(i) = Pnp(i)
    Mnp2(i) = Mnp(i)
    Next i
    
    
    Mmax = fMn(0) / 100000
    For i = 1 To puntos
    If fMn(i) / 100000 > Mmax Then
    Mmax = fMn(i) / 100000
    End If
    Next i
    'Fin de graficar los puntos del diagrama.

    
    'Fin del procedimiento que realiza todos los cálculos de flexo-compresion.
    
    
    End Sub
    
    
    Private Sub Cflexion_Click()
    'Este procedimiento recalcula todos los datos del diseño.
    
   
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 6
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 9
    'Para poner en blanco las casillas cada corrida!
    For i = 2 To 15
    For j = 1 To 4
    Tabla1.TextMatrix(i, j) = ""
    Next j
    Next i
    'Fin de casillas en blanco.
    MsgBox ("Precaución!! No se importó ningún archivo!"), vbCritical
    Exit Sub
    Case 11
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 13
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    End Select
    'Fin del errorhandler de no poner nada!
    
    
    CColumna_Click
    
    
    'Fin del procedimiento que recalcula todos los datos del diseño.
 
 
    End Sub
        
        
    Private Sub CEje_Click()
    'Este procedimiento actualiza los datos al cambiar el número de elemento.
    
    
    CColumna_Click
    
    
    'Fin del procedimiento que actualiza los datos al cambiar el número de elemento.
    
    
    End Sub
    
    
    Private Sub CNivel_Click()
    'Este procedimiento actualiza los datos al cambiar el número del nivel.
   
    
    Cflexion_Click
    
    
    'Fin del procedimiento que actualiza los datos al cambiar el número del nivel.


    End Sub


    Private Sub MAcerca_Click()
    'Este procedimiento pone a la vista los datos acerca del programa.
  
  
    FAbout.Visible = True
    
    
    'Fin del procedimiento que pone a la vista los datos acerca del programa.
    
    
    End Sub
    
    
    Private Sub Mimprimir_Click()
   'Este procedimiento imprime la pantala actual.
    
    
    CommonDialog1.Orientation = cdlLandscape
    
    
    MsgPrompt = "Porfavor verifique si la impresora (" & Printer.DeviceName & ") está lista"
    i = MsgBox(MsgPrompt, vbOKCancel, "Confirmation")

    
    If i = vbCancel Then
    Exit Sub
    End If
    

    PrintForm
    
    
    'Fin del procedimiento que imprime la pantala actual.
    
    
    End Sub
    
    
    Private Sub MSalida_Click()
     
     
     'Este procedimiento enseña la base de datos importada.
    
    
    SSTabColumnas.Visible = False
    MSHFlexGrid1.Visible = True
    MSHFlexGrid2.Visible = False
    
     
     'Fin del procedimiento que enseña la base de datos importada.
   
   
    End Sub
    
    
    Private Sub MSalida2_Click()
    
    
    'Este procedimiento oculta la base de datos importada.


    SSTabColumnas.Visible = True
    
    
    MSHFlexGrid1.Visible = False
    
    
    MSHFlexGrid2.Visible = False
    
    
    'Fin del procedimiento que oculta la base de datos importada.
    
    
    End Sub
    
    
    Private Sub MVigas_Click()
    
    
    'Este procedimiento vuelve a la pantalla principal.
    
    
    Unload FColumnas
    
    
    FDialog.Visible = True
    
    
    FDialog.Option1.Value = True
    
    
    'Fin del procedimiento que vuelve a la pantalla principal.
 
    
    End Sub
    
    
    Private Sub MColumnas_Click()
    
    
    'Este procedimiento vuelve a la pantalla principal.
    
    
    Unload FColumnas
    
    
    FDialog.Visible = True
    
    
    FDialog.Option2.Value = True
     
    
    'Fin del procedimiento que vuelve a la pantalla principal.

    
    
    End Sub
    
    
    Private Sub MMuros_Click()
     
     
    'Este procedimiento vuelve a la pantalla principal.
   
    
    Unload FColumnas
    
    
    FDialog.Visible = True
    
    
    FDialog.Option3.Value = True
    
    
    'Fin del procedimiento que vuelve a la pantalla principal.
    
    
    End Sub
    
    
    Private Sub OCortante_Click()

 
    
    If OCortante.Value = True Then
    
    
    Label16.Visible = True
    
    
    Tabla2.Visible = True
     
    
    Label6.Visible = False
    
    
    Tabla1.Visible = False
    
    
    End If
    


    End Sub

    Private Sub Oflexion_Click()


    If OFlexion.Value = True Then
    
    
    Label6.Visible = True
    
    
    Tabla1.Visible = True
         
    
    Label16.Visible = False
    
    
    Tabla2.Visible = False
     
    
    End If
    

    End Sub

    Private Sub Salida_Click()
    
    
     'Este procedimiento hace que el programa se cierre completamente.
    
    
    
    End
    
    
     'Fin del procedimiento que hace que el programa se cierre completamente.
    
    
    End Sub


    Private Sub Tb_Change()


   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!


    Th = Tb
    
    
    
    Tb2 = Tb
    
    
    Th2 = Th
    
    
    Tdx1(0) = 5
    
    
    Tdx1(1) = Th / 2
    
    
    Tdx1(2) = Th - 5


    End Sub

    
    Private Sub Tb2_Change()
    
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    Tb = Tb2


    End Sub
 

    Private Sub Tdp2_Change()


    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    Td2 = Th2 - Tdp2


    End Sub

    
    Private Sub Tdx1_Change(Index As Integer)
    
    
    Tdp2 = Tdx1(0)


    End Sub

    
    Private Sub TFc_Change()

   
   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!


    TFc2 = TFc


    End Sub

    Private Sub TFc2_Change()


   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    TFc = TFc2


    End Sub


Private Sub Th_Change()
   
   
   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!


    Tdx1(0) = 5
    
    
    Tdx1(1) = Th / 2
    
    
    Tdx1(2) = Th - 5
    
    
    Th2 = Th
    
    
    End Sub

    
    Private Sub Th2_Change()


    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    Th = Th2
    
    
    Td2 = Th2 - Tdp2

    
    End Sub


    Private Sub Tnum_Change(Index As Integer)
        
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!
    

    For i = 0 To 5
    
    
    If Val(Tnum(i)) And Val(Tdiam(i)) > 0 Then
    
    
    TAs1(i) = Round(Val(Tnum(i)) * 3.14159265359 * ((Val(Tdiam(i)) / 8) * 2.54) ^ 2 / 4, 2)
    Tgato(i) = "#"
    
    Else:
    
    
    TAs1(i) = ""
    Tgato(i) = ""
    
    
    End If
    
    
    Next i
    
    
    End Sub


    Private Sub Tdiam_Change(Index As Integer)
    
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!
    

    For i = 0 To 5
    
    
    If Val(Tnum(i)) And Val(Tdiam(i)) > 0 Then
    
    
    TAs1(i) = Round(Val(Tnum(i)) * 3.14159265359 * ((Val(Tdiam(i)) / 8) * 2.54) ^ 2 / 4, 2)
    Tgato(i) = "#"
    
    Else
    
    
    TAs1(i) = ""
    Tgato(i) = ""
    
    
    End If
    
    
    Next i
    
    
    End Sub
