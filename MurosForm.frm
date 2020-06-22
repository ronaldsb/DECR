VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FMuros 
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
   ScaleHeight     =   9831.02
   ScaleMode       =   0  'User
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabMuros 
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      _ExtentX        =   22458
      _ExtentY        =   17166
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Flexo-compresión"
      TabPicture(0)   =   "MurosForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label35"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LHora"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label34"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label33"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CFlexion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CMuro"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TFecha"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "THora"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "THecho"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TProyecto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CEje"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text22"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Diseño por cortante y los elementos de borde"
      TabPicture(1)   =   "MurosForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label24"
      Tab(1).Control(3)=   "Label26"
      Tab(1).Control(4)=   "Label27"
      Tab(1).Control(5)=   "Label29"
      Tab(1).Control(6)=   "Label30"
      Tab(1).Control(7)=   "Label31"
      Tab(1).Control(8)=   "Label32"
      Tab(1).Control(9)=   "Label36"
      Tab(1).Control(10)=   "Label48"
      Tab(1).Control(11)=   "Label79"
      Tab(1).Control(12)=   "Label71"
      Tab(1).Control(13)=   "Label80"
      Tab(1).Control(14)=   "Label25"
      Tab(1).Control(15)=   "Label42"
      Tab(1).Control(16)=   "Label46"
      Tab(1).Control(17)=   "Label56"
      Tab(1).Control(18)=   "Label57"
      Tab(1).Control(19)=   "Label52"
      Tab(1).Control(20)=   "Label58"
      Tab(1).Control(21)=   "Label60"
      Tab(1).Control(22)=   "Label64"
      Tab(1).Control(23)=   "Label53"
      Tab(1).Control(24)=   "Label72"
      Tab(1).Control(25)=   "Label73"
      Tab(1).Control(26)=   "Label87"
      Tab(1).Control(27)=   "Label88"
      Tab(1).Control(28)=   "Label82"
      Tab(1).Control(29)=   "Label89"
      Tab(1).Control(30)=   "Label91"
      Tab(1).Control(31)=   "Label90"
      Tab(1).Control(32)=   "Label43"
      Tab(1).Control(33)=   "Label44"
      Tab(1).Control(34)=   "Label47"
      Tab(1).Control(35)=   "Label83"
      Tab(1).Control(36)=   "Label84"
      Tab(1).Control(37)=   "Label92"
      Tab(1).Control(38)=   "Label93"
      Tab(1).Control(39)=   "Label23"
      Tab(1).Control(40)=   "Label59"
      Tab(1).Control(41)=   "Label63"
      Tab(1).Control(42)=   "Label45"
      Tab(1).Control(43)=   "Label78"
      Tab(1).Control(44)=   "Label95"
      Tab(1).Control(45)=   "Label96"
      Tab(1).Control(46)=   "Label97"
      Tab(1).Control(47)=   "Label98"
      Tab(1).Control(48)=   "TGc"
      Tab(1).Control(49)=   "Thw"
      Tab(1).Control(50)=   "TFc2"
      Tab(1).Control(51)=   "TLw"
      Tab(1).Control(52)=   "ThwLw"
      Tab(1).Control(53)=   "TFyh2"
      Tab(1).Control(54)=   "Tmalla"
      Tab(1).Control(55)=   "Command1"
      Tab(1).Control(56)=   "TVar"
      Tab(1).Control(57)=   "Frame3"
      Tab(1).Control(58)=   "TFy2"
      Tab(1).Control(59)=   "Frame4"
      Tab(1).Control(60)=   "Tsreq"
      Tab(1).Control(61)=   "TVn"
      Tab(1).Control(62)=   "Tmalla2"
      Tab(1).Control(63)=   "Tsmax"
      Tab(1).Control(64)=   "TVmax"
      Tab(1).Control(65)=   "Tsmaxtemp"
      Tab(1).Control(66)=   "TR2"
      Tab(1).Control(67)=   "TR1"
      Tab(1).Control(68)=   "Tb3"
      Tab(1).Control(69)=   "TphiC2"
      Tab(1).Control(70)=   "TphiLw"
      Tab(1).Control(71)=   "TDebe"
      Tab(1).Control(72)=   "TVc"
      Tab(1).Control(73)=   "TVs"
      Tab(1).Control(74)=   "Text1"
      Tab(1).Control(75)=   "Text3"
      Tab(1).Control(76)=   "Text4"
      Tab(1).ControlCount=   77
      Begin VB.TextBox Text4 
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
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   233
         Text            =   "4"
         Top             =   8400
         Width           =   375
      End
      Begin VB.TextBox Text3 
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
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   230
         Text            =   "4"
         Top             =   8040
         Width           =   375
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
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   227
         Text            =   "4"
         Top             =   7680
         Width           =   375
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
         Left            =   240
         TabIndex        =   226
         Text            =   "EJE ="
         Top             =   960
         Width           =   1335
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
         Height          =   315
         ItemData        =   "MurosForm.frx":0038
         Left            =   1680
         List            =   "MurosForm.frx":003A
         TabIndex        =   225
         Text            =   "M3"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TVs 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   215
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox TVc 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   212
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox TDebe 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   210
         Text            =   "No"
         Top             =   9720
         Width           =   735
      End
      Begin VB.TextBox TphiLw 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   207
         Text            =   "0.80"
         Top             =   7080
         Width           =   735
      End
      Begin VB.TextBox TphiC2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   205
         Text            =   "0.85"
         Top             =   6720
         Width           =   732
      End
      Begin VB.TextBox Tb3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   200
         Text            =   "20"
         Top             =   1920
         Width           =   732
      End
      Begin VB.TextBox TProyecto 
         BackColor       =   &H80000004&
         Height          =   288
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   960
         Width           =   2292
      End
      Begin VB.TextBox THecho 
         BackColor       =   &H80000004&
         Height          =   288
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   600
         Width           =   2292
      End
      Begin VB.TextBox TR1 
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
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   175
         Text            =   "4"
         Top             =   8760
         Width           =   375
      End
      Begin VB.TextBox TR2 
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   174
         Top             =   8760
         Width           =   735
      End
      Begin VB.TextBox Tsmaxtemp 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   173
         Top             =   7680
         Width           =   735
      End
      Begin VB.TextBox TVmax 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   172
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox Tsmax 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   171
         Text            =   "45.00"
         Top             =   8040
         Width           =   735
      End
      Begin VB.TextBox Tmalla2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   65
         Text            =   "2"
         Top             =   6360
         Width           =   735
      End
      Begin VB.TextBox TVn 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   170
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox Tsreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   169
         Top             =   8400
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Height          =   4092
         Left            =   -69360
         TabIndex        =   132
         Top             =   720
         Width           =   6852
         Begin VB.TextBox Tc 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   141
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox Tc1 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   140
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox Tc2 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   139
            Top             =   2400
            Width           =   615
         End
         Begin VB.TextBox Tborde 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   138
            Text            =   "No"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox Tdu 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   66
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Tborde2 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   137
            Text            =   "No"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Tb1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5040
            TabIndex        =   67
            Text            =   "0"
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox Tb2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5040
            TabIndex        =   68
            Text            =   "0"
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox TAsbordemin 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   2400
            Width           =   615
         End
         Begin VB.TextBox TAsbordemax 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   135
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox TBordemin 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   134
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox TAsborde 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   133
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label Label94 
            Caption         =   "Nota:  Recordar los requerimientos de confinamiento en muros y sus elementos de borde!"
            Height          =   252
            Left            =   120
            TabIndex        =   224
            Top             =   3600
            Width           =   6492
         End
         Begin VB.Label Label38 
            Caption         =   "c ="
            Height          =   252
            Left            =   1680
            TabIndex        =   165
            Top             =   1680
            Width           =   252
         End
         Begin VB.Label Label39 
            Caption         =   "cm"
            Height          =   252
            Left            =   2760
            TabIndex        =   164
            Top             =   1680
            Width           =   252
         End
         Begin VB.Label Label28 
            Caption         =   "cm"
            Height          =   252
            Left            =   2760
            TabIndex        =   163
            Top             =   2040
            Width           =   252
         End
         Begin VB.Label Label37 
            Caption         =   "c - 0.1Lw ="
            Height          =   252
            Left            =   1152
            TabIndex        =   162
            Top             =   2040
            Width           =   852
         End
         Begin VB.Label Label40 
            Caption         =   "cm"
            Height          =   252
            Left            =   2760
            TabIndex        =   161
            Top             =   2400
            Width           =   252
         End
         Begin VB.Label Label41 
            Caption         =   "c/2 ="
            Height          =   252
            Left            =   1536
            TabIndex        =   160
            Top             =   2400
            Width           =   492
         End
         Begin VB.Label Label50 
            Caption         =   "Req. elem. borde ="
            Height          =   252
            Left            =   480
            TabIndex        =   159
            Top             =   1080
            Width           =   1452
         End
         Begin VB.Label Label67 
            Caption         =   "cm"
            Height          =   255
            Left            =   2760
            TabIndex        =   158
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label68 
            Caption         =   "d m ="
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   7.8
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   1560
            TabIndex        =   157
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label61 
            Caption         =   "Según Art. 21.6.6.2"
            Height          =   252
            Left            =   1320
            TabIndex        =   156
            Top             =   360
            Width           =   1452
         End
         Begin VB.Label Label62 
            Caption         =   "Req. elem. borde ="
            Height          =   252
            Left            =   3480
            TabIndex        =   155
            Top             =   720
            Width           =   1452
         End
         Begin VB.Label Label65 
            Caption         =   "Según Art. 21.6.6.3"
            Height          =   252
            Left            =   4200
            TabIndex        =   154
            Top             =   360
            Width           =   1452
         End
         Begin VB.Label Label66 
            Caption         =   "cm"
            Height          =   252
            Left            =   5760
            TabIndex        =   153
            Top             =   1680
            Width           =   252
         End
         Begin VB.Label Label69 
            Caption         =   "b' ="
            Height          =   252
            Left            =   4560
            TabIndex        =   152
            Top             =   1680
            Width           =   372
         End
         Begin VB.Label Label70 
            Caption         =   "cm"
            Height          =   252
            Left            =   5760
            TabIndex        =   151
            Top             =   2040
            Width           =   252
         End
         Begin VB.Label Label555 
            Caption         =   "L' ="
            Height          =   252
            Left            =   4680
            TabIndex        =   150
            Top             =   2040
            Width           =   372
         End
         Begin VB.Label Label74 
            Caption         =   "cm2"
            Height          =   252
            Left            =   5760
            TabIndex        =   149
            Top             =   2400
            Width           =   492
         End
         Begin VB.Label Label75 
            Caption         =   "As min ="
            Height          =   252
            Left            =   4200
            TabIndex        =   148
            Top             =   2400
            Width           =   732
         End
         Begin VB.Label Label76 
            Caption         =   "cm2"
            Height          =   252
            Left            =   5760
            TabIndex        =   147
            Top             =   2760
            Width           =   612
         End
         Begin VB.Label Label77 
            Caption         =   "As max ="
            Height          =   252
            Left            =   4200
            TabIndex        =   146
            Top             =   2760
            Width           =   732
         End
         Begin VB.Label Label54 
            Caption         =   "b' min ="
            Height          =   252
            Left            =   1320
            TabIndex        =   145
            Top             =   2760
            Width           =   612
         End
         Begin VB.Label Label55 
            Caption         =   "cm"
            Height          =   252
            Left            =   2760
            TabIndex        =   144
            Top             =   2760
            Width           =   252
         End
         Begin VB.Label Label49 
            Caption         =   "As req ="
            Height          =   252
            Left            =   4200
            TabIndex        =   143
            Top             =   3120
            Width           =   732
         End
         Begin VB.Label Label51 
            Caption         =   "cm2"
            Height          =   252
            Left            =   5760
            TabIndex        =   142
            Top             =   3120
            Width           =   612
         End
      End
      Begin VB.TextBox TFy2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   57
         Text            =   "4200"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   -69360
         TabIndex        =   128
         Top             =   4920
         Width           =   6852
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tabla2 
            Height          =   3930
            Left            =   1320
            TabIndex        =   129
            Top             =   480
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
         Begin VB.Label Label86 
            Caption         =   "Cortantes  y cargas axiales en ton."
            Height          =   252
            Left            =   2160
            TabIndex        =   168
            Top             =   240
            Width           =   2532
         End
      End
      Begin VB.TextBox TVar 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   59
         Text            =   "4"
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar cálculos."
         Height          =   495
         Left            =   -71280
         TabIndex        =   69
         Top             =   840
         Width           =   1692
      End
      Begin VB.TextBox Tmalla 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   125
         Text            =   "2"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox TFyh2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   58
         Text            =   "2800"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox ThwLw 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox TLw 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   61
         Text            =   "200"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TFc2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   56
         Text            =   "210"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Thw 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73080
         TabIndex        =   60
         Text            =   "350"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TGc 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   3360
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Height          =   8295
         Left            =   120
         TabIndex        =   73
         Top             =   1320
         Width           =   12492
         Begin VB.TextBox TphiC 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   203
            Text            =   "0.85"
            Top             =   1440
            Width           =   612
         End
         Begin VB.PictureBox Picture1 
            Height          =   1995
            Left            =   240
            Picture         =   "MurosForm.frx":003C
            ScaleHeight     =   1944
            ScaleWidth      =   4248
            TabIndex        =   166
            Top             =   1920
            Width           =   4290
         End
         Begin VB.TextBox TFyh 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Text            =   "2800"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox Thp 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   5
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox Tbp 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   4
            Text            =   "0"
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Actualizar diagrama."
            Height          =   492
            Left            =   10560
            TabIndex        =   53
            Top             =   3240
            Width           =   1695
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tabla1 
            Height          =   3930
            Left            =   240
            TabIndex        =   106
            Top             =   4320
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
         Begin VB.Frame Frame2 
            Height          =   2840
            Left            =   7800
            TabIndex        =   75
            Top             =   120
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
               Index           =   9
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   199
               Text            =   "#"
               Top             =   2520
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
               Index           =   8
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   198
               Text            =   "#"
               Top             =   2280
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
               Index           =   7
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   197
               Text            =   "#"
               Top             =   2040
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
               Index           =   6
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   196
               Top             =   1800
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
               Index           =   5
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   195
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
               TabIndex        =   194
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
               TabIndex        =   193
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
               TabIndex        =   192
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
               TabIndex        =   191
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
               TabIndex        =   190
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
               Index           =   9
               Left            =   2400
               TabIndex        =   51
               Text            =   "195"
               Top             =   2520
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
               Index           =   8
               Left            =   2400
               TabIndex        =   47
               Text            =   "180"
               Top             =   2280
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
               Index           =   7
               Left            =   2400
               TabIndex        =   43
               Text            =   "165"
               Top             =   2040
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
               Index           =   6
               Left            =   2400
               TabIndex        =   39
               Top             =   1800
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
               Index           =   5
               Left            =   2400
               TabIndex        =   35
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
               TabIndex        =   31
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
               TabIndex        =   27
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
               TabIndex        =   23
               Text            =   "35"
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
               TabIndex        =   19
               Text            =   "20"
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
               TabIndex        =   15
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
               Index           =   9
               Left            =   1560
               TabIndex        =   50
               Text            =   "15.2"
               Top             =   2520
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
               Index           =   8
               Left            =   1560
               TabIndex        =   46
               Text            =   "10.13"
               Top             =   2280
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
               Index           =   7
               Left            =   1560
               TabIndex        =   42
               Text            =   "15.2"
               Top             =   2040
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
               Index           =   6
               Left            =   1560
               TabIndex        =   38
               Top             =   1800
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
               TabIndex        =   34
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
               TabIndex        =   30
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
               TabIndex        =   26
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
               TabIndex        =   22
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
               TabIndex        =   18
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
               TabIndex        =   14
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
               Index           =   9
               Left            =   1200
               TabIndex        =   49
               Text            =   "8"
               Top             =   2520
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
               Index           =   8
               Left            =   1200
               TabIndex        =   45
               Text            =   "8"
               Top             =   2280
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
               Index           =   7
               Left            =   1200
               TabIndex        =   41
               Text            =   "8"
               Top             =   2040
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
               Index           =   6
               Left            =   1200
               TabIndex        =   37
               Top             =   1800
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
               Index           =   5
               Left            =   1200
               TabIndex        =   33
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
               TabIndex        =   29
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
               TabIndex        =   25
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
               TabIndex        =   21
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
               TabIndex        =   17
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
               TabIndex        =   13
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
               Index           =   9
               Left            =   600
               TabIndex        =   48
               Text            =   "3"
               Top             =   2520
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
               Index           =   8
               Left            =   600
               TabIndex        =   44
               Text            =   "2"
               Top             =   2280
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
               Index           =   7
               Left            =   600
               TabIndex        =   40
               Text            =   "3"
               Top             =   2040
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
               Index           =   6
               Left            =   600
               TabIndex        =   36
               Top             =   1800
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
               TabIndex        =   32
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
               TabIndex        =   28
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
               TabIndex        =   24
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
               TabIndex        =   20
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
               TabIndex        =   16
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
               TabIndex        =   12
               Text            =   "3"
               Top             =   360
               Width           =   375
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
               TabIndex        =   88
               Text            =   "Varillas"
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox Text47 
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
               TabIndex        =   76
               Text            =   "10"
               Top             =   2520
               Width           =   615
            End
            Begin VB.TextBox Text43 
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
               Text            =   "9"
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox Text39 
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
               TabIndex        =   78
               Text            =   "8"
               Top             =   2040
               Width           =   615
            End
            Begin VB.TextBox Text35 
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
               TabIndex        =   79
               Text            =   "7"
               Top             =   1800
               Width           =   615
            End
            Begin VB.TextBox Text27 
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
               TabIndex        =   80
               Text            =   "6"
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox Text23 
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
               TabIndex        =   81
               Text            =   "5"
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Text30 
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
               TabIndex        =   82
               Text            =   "4"
               Top             =   1080
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
               TabIndex        =   83
               Text            =   "3"
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox Text14 
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
               TabIndex        =   84
               Text            =   "2"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox Text10 
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
               TabIndex        =   85
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
               TabIndex        =   89
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
               TabIndex        =   87
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
               TabIndex        =   86
               Text            =   "x (cm)"
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.TextBox TEs 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5040
            TabIndex        =   11
            Text            =   "2.1E6"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TEus 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5040
            TabIndex        =   10
            Text            =   "0.002"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TEuc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5040
            TabIndex        =   9
            Text            =   "0.003"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Tb 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   2
            Text            =   "20"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Th 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   3
            Text            =   "200"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TFc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   6
            Text            =   "210"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox TFy 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   7
            Text            =   "4200"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Tpa 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label81 
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
            Left            =   2160
            TabIndex        =   204
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label85 
            Caption         =   "Momentos en ton-m y carga axial en ton."
            Height          =   255
            Left            =   1080
            TabIndex        =   167
            Top             =   4080
            Width           =   3015
         End
         Begin VB.Label Label22 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   112
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label21 
            Caption         =   "L ="
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label19 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   110
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "b' ="
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Top             =   1080
            Width           =   375
         End
         Begin VB.OLE OLE1 
            AutoActivate    =   3  'Automatic
            BackColor       =   &H80000004&
            Class           =   "Excel.Chart.8"
            Enabled         =   0   'False
            Height          =   5055
            Left            =   5280
            OleObjectBlob   =   "MurosForm.frx":A43EE
            SizeMode        =   1  'Stretch
            TabIndex        =   108
            Top             =   3120
            Width           =   7095
         End
         Begin VB.Label Label18 
            Caption         =   "su ="
            Height          =   255
            Left            =   4560
            TabIndex        =   105
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "cu ="
            Height          =   255
            Left            =   4560
            TabIndex        =   104
            Top             =   360
            Width           =   375
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
            Height          =   255
            Left            =   4395
            TabIndex        =   103
            Top             =   660
            Width           =   135
         End
         Begin VB.Label Label20 
            Caption         =   "Es ="
            Height          =   255
            Left            =   4560
            TabIndex        =   102
            Top             =   1080
            Width           =   375
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
            Height          =   255
            Left            =   4395
            TabIndex        =   101
            Top             =   300
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "b ="
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   99
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "L' ="
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "cm"
            Height          =   255
            Left            =   1320
            TabIndex        =   97
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "f'c ="
            Height          =   255
            Left            =   2040
            TabIndex        =   96
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "fy ="
            Height          =   255
            Left            =   2040
            TabIndex        =   95
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "% Acero ="
            Height          =   255
            Left            =   4200
            TabIndex        =   94
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "kg/cm2"
            Height          =   255
            Left            =   3240
            TabIndex        =   93
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "kg/cm2"
            Height          =   255
            Left            =   3240
            TabIndex        =   92
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "fyh ="
            Height          =   255
            Left            =   2040
            TabIndex        =   91
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label16 
            Caption         =   "kg/cm2"
            Height          =   255
            Left            =   3240
            TabIndex        =   90
            Top             =   1080
            Width           =   615
         End
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
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   936
         Width           =   1455
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
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   576
         Width           =   1455
      End
      Begin VB.ComboBox CMuro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "MurosForm.frx":A8806
         Left            =   1680
         List            =   "MurosForm.frx":A8808
         TabIndex        =   1
         Top             =   600
         Width           =   1215
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
         Left            =   240
         TabIndex        =   62
         Text            =   "ELEMENTO #"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CFlexion 
         Caption         =   "Actualizar cálculos."
         Height          =   495
         Left            =   3120
         TabIndex        =   52
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label98 
         Caption         =   "#"
         Height          =   252
         Left            =   -73680
         TabIndex        =   235
         Top             =   8412
         Width           =   132
      End
      Begin VB.Label Label97 
         Caption         =   "@"
         Height          =   252
         Left            =   -73128
         TabIndex        =   234
         Top             =   8412
         Width           =   252
      End
      Begin VB.Label Label96 
         Caption         =   "#"
         Height          =   252
         Left            =   -73680
         TabIndex        =   232
         Top             =   8052
         Width           =   132
      End
      Begin VB.Label Label95 
         Caption         =   "@"
         Height          =   252
         Left            =   -73128
         TabIndex        =   231
         Top             =   8052
         Width           =   252
      End
      Begin VB.Label Label78 
         Caption         =   "#"
         Height          =   252
         Left            =   -73680
         TabIndex        =   229
         Top             =   7692
         Width           =   132
      End
      Begin VB.Label Label45 
         Caption         =   "@"
         Height          =   252
         Left            =   -73128
         TabIndex        =   228
         Top             =   7692
         Width           =   252
      End
      Begin VB.Label Label63 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72000
         TabIndex        =   223
         Top             =   7680
         Width           =   372
      End
      Begin VB.Label Label59 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72000
         TabIndex        =   222
         Top             =   8400
         Width           =   372
      End
      Begin VB.Label Label23 
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   7.8
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -73560
         TabIndex        =   221
         Top             =   3360
         Width           =   132
      End
      Begin VB.Label Label93 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   7.8
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -73440
         TabIndex        =   220
         Top             =   9720
         Width           =   132
      End
      Begin VB.Label Label92 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   7.8
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -74700
         TabIndex        =   219
         Top             =   9720
         Width           =   132
      End
      Begin VB.Label Label84 
         Caption         =   "As Final ="
         Height          =   252
         Left            =   -74520
         TabIndex        =   218
         Top             =   8760
         Width           =   732
      End
      Begin VB.Label Label83 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -72240
         TabIndex        =   217
         Top             =   5280
         Width           =   372
      End
      Begin VB.Label Label47 
         Caption         =   "Vs ="
         Height          =   252
         Left            =   -73560
         TabIndex        =   216
         Top             =   5280
         Width           =   372
      End
      Begin VB.Label Label44 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -72240
         TabIndex        =   214
         Top             =   4920
         Width           =   372
      End
      Begin VB.Label Label43 
         Caption         =   "Vc ="
         Height          =   252
         Left            =   -73560
         TabIndex        =   213
         Top             =   4920
         Width           =   372
      End
      Begin VB.Label Label90 
         Caption         =   "  v  debe ser  >=      n ) ?"
         Height          =   252
         Left            =   -74640
         TabIndex        =   211
         Top             =   9720
         Width           =   1812
      End
      Begin VB.Label Label91 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.4
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -73680
         TabIndex        =   209
         Top             =   7080
         Width           =   132
      End
      Begin VB.Label Label89 
         Caption         =   "Lw ="
         Height          =   252
         Left            =   -73560
         TabIndex        =   208
         Top             =   7080
         Width           =   372
      End
      Begin VB.Label Label82 
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
         Height          =   252
         Left            =   -73440
         TabIndex        =   206
         Top             =   6720
         Width           =   372
      End
      Begin VB.Label Label88 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72240
         TabIndex        =   202
         Top             =   1920
         Width           =   372
      End
      Begin VB.Label Label87 
         Caption         =   "b ="
         Height          =   252
         Left            =   -73440
         TabIndex        =   201
         Top             =   1920
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
         Left            =   6600
         TabIndex        =   189
         Top             =   612
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
         Left            =   6720
         TabIndex        =   188
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label73 
         Caption         =   "@"
         Height          =   252
         Left            =   -73128
         TabIndex        =   187
         Top             =   8772
         Width           =   252
      End
      Begin VB.Label Label72 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72000
         TabIndex        =   186
         Top             =   8760
         Width           =   372
      End
      Begin VB.Label Label53 
         Caption         =   "#"
         Height          =   252
         Left            =   -73680
         TabIndex        =   185
         Top             =   8772
         Width           =   132
      End
      Begin VB.Label Label64 
         Caption         =   "s max temp ="
         Height          =   252
         Left            =   -74760
         TabIndex        =   184
         Top             =   7680
         Width           =   972
      End
      Begin VB.Label Label60 
         Caption         =   "s req ="
         Height          =   252
         Left            =   -74280
         TabIndex        =   183
         Top             =   8400
         Width           =   492
      End
      Begin VB.Label Label58 
         Caption         =   "Vu max ="
         Height          =   252
         Left            =   -73920
         TabIndex        =   182
         Top             =   4200
         Width           =   612
      End
      Begin VB.Label Label52 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -72240
         TabIndex        =   181
         Top             =   4200
         Width           =   372
      End
      Begin VB.Label Label57 
         Caption         =   "s max ="
         Height          =   252
         Left            =   -74364
         TabIndex        =   180
         Top             =   8040
         Width           =   612
      End
      Begin VB.Label Label56 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72000
         TabIndex        =   179
         Top             =   8040
         Width           =   372
      End
      Begin VB.Label Label46 
         Caption         =   "# de mallas ="
         Height          =   252
         Left            =   -74160
         TabIndex        =   178
         Top             =   6360
         Width           =   972
      End
      Begin VB.Label Label42 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -72240
         TabIndex        =   177
         Top             =   4560
         Width           =   372
      End
      Begin VB.Label Label25 
         Caption         =   "Vu mayor ="
         Height          =   252
         Left            =   -74040
         TabIndex        =   176
         Top             =   4560
         Width           =   852
      End
      Begin VB.Label Label80 
         Caption         =   "fy ="
         Height          =   252
         Left            =   -73680
         TabIndex        =   131
         Top             =   1200
         Width           =   492
      End
      Begin VB.Label Label71 
         Caption         =   "kg/cm2"
         Height          =   252
         Left            =   -72240
         TabIndex        =   130
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label79 
         Caption         =   "# Varilla ="
         Height          =   252
         Left            =   -73920
         TabIndex        =   127
         Top             =   6000
         Width           =   732
      End
      Begin VB.Label Label48 
         Caption         =   "# de mallas req.="
         Height          =   252
         Left            =   -74400
         TabIndex        =   126
         Top             =   3720
         Width           =   1332
      End
      Begin VB.Label Label36 
         Caption         =   "c ="
         Height          =   252
         Left            =   -73440
         TabIndex        =   123
         Top             =   3360
         Width           =   252
      End
      Begin VB.Label Label32 
         Caption         =   "kg/cm2"
         Height          =   252
         Left            =   -72240
         TabIndex        =   122
         Top             =   1560
         Width           =   612
      End
      Begin VB.Label Label31 
         Caption         =   "kg/cm2"
         Height          =   252
         Left            =   -72240
         TabIndex        =   121
         Top             =   840
         Width           =   612
      End
      Begin VB.Label Label30 
         Caption         =   "fyh ="
         Height          =   252
         Left            =   -73680
         TabIndex        =   120
         Top             =   1560
         Width           =   492
      End
      Begin VB.Label Label29 
         Caption         =   "f'c ="
         Height          =   252
         Left            =   -73680
         TabIndex        =   119
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label27 
         Caption         =   "hw/Lw ="
         Height          =   252
         Left            =   -73800
         TabIndex        =   118
         Top             =   3000
         Width           =   732
      End
      Begin VB.Label Label26 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72240
         TabIndex        =   117
         Top             =   2640
         Width           =   372
      End
      Begin VB.Label Label24 
         Caption         =   "Lw ="
         Height          =   252
         Left            =   -73560
         TabIndex        =   116
         Top             =   2280
         Width           =   372
      End
      Begin VB.Label Label6 
         Caption         =   "hw ="
         Height          =   252
         Left            =   -73560
         TabIndex        =   115
         Top             =   2640
         Width           =   372
      End
      Begin VB.Label Label2 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72240
         TabIndex        =   114
         Top             =   2280
         Width           =   372
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
         Left            =   10440
         TabIndex        =   71
         Top             =   948
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
         Left            =   10320
         TabIndex        =   70
         Top             =   600
         Width           =   612
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   9600
      Left            =   120
      TabIndex        =   72
      Top             =   120
      Width           =   12585
      _ExtentX        =   22204
      _ExtentY        =   16933
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
      Left            =   600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.mdb"
      InitDir         =   "C:\My Documents\Diseño\"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   9600
      Left            =   120
      TabIndex        =   107
      Top             =   120
      Width           =   12585
      _ExtentX        =   22204
      _ExtentY        =   16933
      _Version        =   393216
      Rows            =   19
      Cols            =   10
      FixedCols       =   0
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "Load|Loc|M2|M3|P|Pier|Story|T|V2|V3"
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
   Begin VB.Menu MSalida 
      Caption         =   "    Archivo de Salida    "
   End
   Begin VB.Menu MSalida2 
      Caption         =   "    Salir del archivo de Salida    "
   End
   Begin VB.Menu Mimprimir 
      Caption         =   "    Imprimir    "
   End
   Begin VB.Menu MGuardar 
      Caption         =   "    Guardar Diseño    "
      Visible         =   0   'False
   End
   Begin VB.Menu MAcerca 
      Caption         =   "    Acerca de    "
   End
   Begin VB.Menu Salida 
      Caption         =   "    Salir    "
   End
End
Attribute VB_Name = "FMuros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Const MARGIN_SIZE = 60
    Private datPrimaryRS As ADODB.Recordset
    
      
    Private Sub Form_Load()
    'Cosas que el programa carga al inicio del módulo de muros.
    
    
    FMuros.Visible = True
    
    
    FDialog.Visible = False
    
    
    FDialog.Visible = False
    
    
    Importar
    
    
    Acomodar
    
    
    MsgBox ("Este programa es para usos académicos únicamente!"), vbCritical
    
    
    THecho = Hechos
    
    
    TProyecto = Proyectos
     
     
    'Fin de las cosas que el programa carga al inicio del módulo de muros.
  
    
    End Sub
    
    
    Private Sub Importar()
    
      
    On Error GoTo Sinsalida
Sinsalida:
    Resume Next
    
    
    Dim sConnect As String
    Dim sSQL As String
    Dim dfwConn As ADODB.Connection
    
        
    MSHFlexGrid1.Visible = True
    MSHFlexGrid2.Visible = False
    
        
    ' set strings
    sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;User ID=Admin;Data Source=" & Archivo & ";Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Locale Identifier=1033;Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Global Partial Bulk Ops=2"
    sSQL = "select FRAME,LOAD,M2,M3,P,STATION,T,V2,V3 from FrameForces"
    
        
    ' open connection
    Set dfwConn = New Connection
     dfwConn.Open sConnect
    
        
    ' create a recordset using the provided collection
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, dfwConn, adOpenForwardOnly, adLockReadOnly
    
        
    Set MSHFlexGrid1.DataSource = datPrimaryRS
    
        
    With MSHFlexGrid1
    
            
    .Redraw = False
    ' set grid's column widths
    .ColWidth(0) = -1
    .ColWidth(1) = -1
    .ColWidth(2) = -1
    .ColWidth(3) = -1
    .ColWidth(4) = -1
    .ColWidth(5) = -1
    .ColWidth(6) = -1
    .ColWidth(7) = -1
    .ColWidth(8) = -1
    
            
    ' set grid's style
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
        
        
    'Convierte los datos que son números a toneladas.
    For c = 2 To MSHFlexGrid1.Cols - 1
    For r = 1 To MSHFlexGrid1.Rows - 1
    Datos(r, c) = (MSHFlexGrid1.TextMatrix(r, c)) / 1000
    Next r
    Next c
    'Fin de pasar los datos de salida a una matriz.
        
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
    
    'Pone la lista de los ejes locales.
    CEje.AddItem "M3"
    CEje.AddItem "M2"
    'Fin de poner la lista de los ejes locales.

    
    
    End Sub
    
    
    Private Sub Command1_Click()
    
    
    CMuro_Click
    
    
    End Sub
    
    
    Private Sub Command2_Click()


    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 6
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 9
    'Para poner en blanco las casillas cada corrida!
    For i = 1 To 21
    For j = 1 To 7
    Tabla3.TextMatrix(i, j) = ""
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
    
    
    CMuro_Click
    
    
    For i = 0 To 15
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
    
    
    End Sub
    
    
    Sub Acomodar()
        
        
    'Centra toda la Tabla1.
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
    
    
    'Le da el ancho deseado a la columnas de la tabla 1.
    For i = 1 To 4
    With Tabla1
        .ColWidth(0) = 840
        .ColWidth(i) = 840
    End With
    Next i
    
    
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
    
    
    'Le da el ancho deseado a la columnas de la tabla 2.
    For i = 1 To 4
    With Tabla2
        .ColWidth(0) = 840
        .ColWidth(i) = 840
    End With
    Next i
    
    
    'Pone los datos de la tabla iniciales.
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
    
    
    'Pone la lista de elementos en el combobox de # del Muro.
    CMuro.Clear
    For i = 1 To FrameU - 1
    CMuro.AddItem Frame(i)
    Next i
    
    CMuro = CMuro.List(0)
       
    End Sub
    
    
    Sub Importacion()
    
        'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 6
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 9
    'Para poner en blanco las casillas cada corrida!
    For i = 1 To 15
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

    
    'Pone la posicion del número del muro.
    For i = 1 To MSHFlexGrid1.Rows - 1
    If Datos(i, 0) = CMuro Then
    Posicion = i
    i = MSHFlexGrid1.Rows - 1
    End If
    Next i
    'Fin de poner la posicion del número del muro.
    
    
    'Lee la longitud del Muro!
    TLong = Round(Datos(Posicion + 6, 5) * 1000, 2)
    'Fin de leer la longitud del muro!
    
    
    'Inserta los datos de CM,CV,CS en la tabla1.
    
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

    
    
    'Calcula y anota las combinaciones en la tabla1.
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
    'Fin de calcular y anotar las combinaciones en la tabla1.
    
    
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
    
    
    End Sub
    
    
    Private Sub CMuro_Click()
    
    
    TFecha = Format$(Now, "d / m / yyyy")
    
    
    THora = Format$(Now, "h:mm AM/PM")
    
    
    Importacion
    
    
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
    
    
    'Diagrama de Interacción Nominal.
    Dim Ass(9), EEs(9), Fs(9), dx(9), Fcapa(9), Pn(), Mn(), fPn(), fMn() As Variant
    Dim Assb(9), EEsb(9), Fsb(9), dxb(9), Fcapab(9) As Variant
    Dim Assc(9), EEsc(9), Fsc(9), dxc(9), Fcapac(9), Pnc(), Mnc() As Variant
    Dim Mnp(18), Pnp(18), Pvert1(18), Pvert2(18) As Variant
    Dim Tnum(9), Tdiam(9), Tgato(9) As Variant
    
    puntos = 15
    ReDim Pn(puntos), Mn(puntos), fPn(puntos), fMn(puntos), Pnc(puntos), Mnc(puntos)
    
    
    Tborde = "No"
    Tborde2 = "No"
    Es = Val(TEs)
    Euc = Val(TEuc)
   
   
    Dim Eus As Double
   
    
    Eus = Val(TFy) / Es
    TEus = Round(Eus, 4)
    
    
    
    Fc = Val(TFc)
    Fy = Val(TFy)
    b = Val(Tb)
    bp = Val(Tbp)
    hp = Val(Thp)
    
    
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
    
    
    For i = 0 To 9
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
    
    
    Ast = Ass(0) + Ass(1) + Ass(2) + Ass(3) + Ass(4) + Ass(5) + Ass(6) + Ass(7) + Ass(8) + Ass(9)
    Pa = 100 * Ast / ((2 * (bp * hp)) + (b * (h - (2 * hp))))
    Tpa = Round(Pa, 2)
    If Pa > 1 And Pa < 6 Then
    TOK = "OK!"
    Else
    TOK = "NO OK!"
    End If
    
    Pconcreto = -(0.85 * Fc * a * b)
    dconcreto = ((a / 2) - (h / 2))
    
    If a < hp Then
    
    
    'Procedimiento 1.
    P1 = -(0.85 * Fc * a * bp)
    br1 = ((a / 2) - (h / 2))
    P2 = 0
    br2 = 0
    P3 = 0
    br3 = 0
    'Fin del procedimiento 1.
    
    
    Else
    
    
    If a < h - hp Then
    
    
    'Procedimiento 2.
    P1 = -(0.85 * Fc * hp * bp)
    br1 = ((hp / 2) - (h / 2))
    P2 = -(0.85 * Fc * (a - hp) * b)
    br2 = (hp + a - h) / 2
    P3 = 0
    br3 = 0
    'Fin del procedimiento 2.
    
    
    Else
    
    
    'Procedimiento 3.
    P1 = -(0.85 * Fc * hp * bp)
    br1 = ((hp / 2) - (h / 2))
    P2 = -(0.85 * Fc * (h - (2 * hp)) * b)
    br2 = 0
    P3 = -(0.85 * Fc * ((a - (h - hp)) * bp))
    br3 = ((a / 2) - (hp / 2))
    'Fin del procedimiento 3.
    End If
    End If
    
    
    For i = 0 To 9
    Fcapa(i) = (Ass(i) * Fs(i))
    Next i
    
    
    If hp = 0 Then
       Pn(j) = Pconcreto
       Mn(j) = Pconcreto * dconcreto
    End If

    If bp >= b Then
        If hp > 0 Then
           Pn(j) = P1 + P2 + P3
           Mn(j) = P1 * br1 + P2 * br2 + P3 * br3
        End If
    Else
           Pn(j) = Pconcreto
           Mn(j) = Pconcreto * dconcreto
    End If


    
    For i = 0 To 9
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
    
    
 '  Pn(puntos) = -((0.85 * Fc * (((b * (h - (2 * hp))) + (2 * bp * hp)) - (Ast))) + (Ast * Fy))
    Mn(puntos) = 0
    
    
    Pn(0) = (Ast * Fy)
    Mn(0) = 0
    
    
    fPn(0) = (Ast * Fy) * 0.9
    fMn(0) = 0
  ' fPn(puntos) = 0.7 * -((0.85 * Fc * (((b * (h - (2 * hp))) + (2 * bp * hp)) - (Ast))) + (Ast * Fy))
    fMn(puntos) = 0
    
    
    Mmax = Mn(0)
    For i = 1 To puntos
    If Mmax < Mn(i) Then
    Mmax = Mn(i)
    End If
    Next i
    
    
         If hp = 0 Then
            Pu = -0.8 * 0.7 * (0.85 * (b * h - Ast) * Fc + (Ast * Fy))
            Pn(puntos) = -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fy))
            fPn(puntos) = 0.7 * -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fy))
        End If

        If bp >= b Then
            If hp > 0 Then
                Pu = -0.8 * 0.7 * (0.85 * (b * h - Ast + (2 * (bp - b) * hp)) * Fc + (Ast * Fy))
                Pn(puntos) = -((0.85 * Fc * ((b * h) - (Ast) + (2 * (bp - b) * hp))) + (Ast * Fy))
                fPn(puntos) = 0.7 * -((0.85 * Fc * ((b * h) - (Ast) + (2 * (bp - b) * hp))) + (Ast * Fy))
            End If
        Else
            Pu = -0.8 * 0.7 * (0.85 * (b * h - Ast) * Fc + (Ast * Fy))
            Pn(puntos) = -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fy))
            fPn(puntos) = 0.7 * -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fy))
        End If
 
    
    
    'Fin del diagrama de interacción nominal.
    
    
        'Inicio del cálculo del Diagrama de Interacción utilizando 1.25 Fy.

        Fyc = (Val(TFy.Text) * 1.25)
        c = h / 10

        For j = 1 To puntos
            a = B1 * c


            For i = 0 To 9

                Ecu = Val(TEuc.Text)
                EEsc(i) = Ecu * ((dx(i) / c) - 1)
                Fsc(i) = EEsc(i) * Es
                If Fsc(i) >= Fyc Then
                    Fsc(i) = Fyc
                End If


                If Fsc(i) < -Fyc Then
                    Fsc(i) = -Fyc
                End If

            Next i


            Pconcreto = -(0.85 * Fc * a * b)
            dconcreto = ((a / 2) - (h / 2))


            For i = 0 To 9
                Fcapac(i) = (Ass(i) * Fsc(i))
            Next i


            If a < hp Then


                'Procedimiento 1.
                P1 = -(0.85 * Fc * a * bp)
                br1 = ((a / 2) - (h / 2))
                P2 = 0
                br2 = 0
                P3 = 0
                br3 = 0
                'Fin del procedimiento 1.


            Else


                If a < h - hp Then


                    'Procedimiento 2.
                    P1 = -(0.85 * Fc * hp * bp)
                    br1 = ((hp / 2) - (h / 2))
                    P2 = -(0.85 * Fc * (a - hp) * b)
                    br2 = (hp + a - h) / 2
                    P3 = 0
                    br3 = 0
                    'Fin del procedimiento 2.


                Else


                    'Procedimiento 3.
                    P1 = -(0.85 * Fc * hp * bp)
                    br1 = ((hp / 2) - (h / 2))
                    P2 = -(0.85 * Fc * (h - (2 * hp)) * b)
                    br2 = 0
                    P3 = -(0.85 * Fc * ((a - (h - hp)) * bp))
                    br3 = ((a / 2) - (hp / 2))
                    'Fin del procedimiento 3.
                End If
            End If



            If hp = 0 Then
                Pnc(j) = Pconcreto
                Mnc(j) = Pconcreto * dconcreto
            End If

            If bp >= b Then
                If hp > 0 Then
                    Pnc(j) = P1 + P2 + P3
                    Mnc(j) = P1 * br1 + P2 * br2 + P3 * br3
                End If
            Else
                Pnc(j) = Pconcreto
                Mnc(j) = Pconcreto * dconcreto
            End If


            For i = 0 To 9
                Pnc(j) = (Fcapac(i) + Pnc(j))
                Mnc(j) = ((Fcapac(i) * (dx(i) - (h / 2))) + Mnc(j))
            Next i


            c = c + deltac
        Next j



        'Pnc(puntos) = -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fyc))
        Mnc(puntos) = 0


        Pnc(0) = (Ast * Fyc)
        Mnc(0) = 0


        If hp = 0 Then
            Pnc(puntos) = -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fyc))
         End If

        If bp >= b Then
            If hp > 0 Then
                Pnc(puntos) = -((0.85 * Fc * ((b * h) - (Ast) + (2 * (bp - b) * hp))) + (Ast * Fyc))
            End If
        Else
            Pnc(puntos) = -((0.85 * Fc * ((b * h) - (Ast))) + (Ast * Fyc))
        End If


        'Fin del cálculo del Diagrama de Interacción utilizando 1.25 Fy.

    
    
    'Puntos del diagrama.
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
    
    
    Pnpmax = Pnp(0)
    For k = 1 To 17
    If Pnp(k) < Pnpmax Then
    Pnpmax = Pnp(k)
    End If
    Next k
    
    
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
    'Fin de los puntos del diagrama.
    
        
    'Cálculo del "c".
    
    puntos = Val(Th)
    Dim Pncc(), Mncc(), CC()
    ReDim Pncc(puntos), Mncc(puntos), CC(puntos)
    
    
    Es = Val(TEs)
    Euc = Val(TEuc)
    Eus = Val(TFy) / Es
    
    
    Fc = Val(TFc)
    Fy = Val(TFy)
    b = Val(Tb)
    bp = Val(Tbp)
    hp = Val(Thp)
    
    
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
    deltac = 1
    
    
    Limite = -0.1 * Fc * b * h / 0.7
    
    
    c = 1
    
    
    For j = 1 To puntos
    
    
    CC(j) = c
    a = B1 * c
    
    
    For i = 0 To 9
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
    
    
    Ast = Ass(0) + Ass(1) + Ass(2) + Ass(3) + Ass(4) + Ass(5) + Ass(6) + Ass(7) + Ass(8) + Ass(9)
    
    If a < hp Then
    
      
    'Procedimiento 1.
    P1 = -(0.85 * Fc * a * bp)
    br1 = ((a / 2) - (h / 2))
    P2 = 0
    br2 = 0
    P3 = 0
    br3 = 0
    'Fin del procedimiento 1.
    
    
    Else
    
    
    If a < h - hp Then
    
    
    'Procedimiento 2.
    P1 = -(0.85 * Fc * hp * bp)
    br1 = ((hp / 2) - (h / 2))
    P2 = -(0.85 * Fc * (a - hp) * b)
    br2 = (hp + a - h) / 2
    P3 = 0
    br3 = 0
    'Fin del procedimiento 2.
    
    
    Else
    
    
    'Procedimiento 3.
    P1 = -(0.85 * Fc * hp * bp)
    br1 = ((hp / 2) - (h / 2))
    P2 = -(0.85 * Fc * (h - (2 * hp)) * b)
    br2 = 0
    P3 = -(0.85 * Fc * ((a - (h - hp)) * bp))
    br3 = ((a / 2) - (hp / 2))
    'Fin del procedimiento 3.
    
    
    End If
    End If
    
    
    For i = 0 To 9
    Fcapa(i) = (Ass(i) * Fs(i))
    Next i
    
    
    Pncc(j) = P1 + P2 + P3
    Mncc(j) = P1 * br1 + P2 * br2 + P3 * br3
    
    
    For i = 0 To 9
    Pncc(j) = (Fcapa(i) + Pncc(j))
    Mncc(j) = ((Fcapa(i) * (dx(i) - (h / 2))) + Mncc(j))
    Next i
    
    
    c = c + deltac
    Next j
    
    
    Pncc(puntos) = -((0.85 * Fc * (((b * (h - (2 * hp))) + (2 * bp * hp)) - (Ast))) + (Ast * Fy))
    Mncc(puntos) = 0
    
    
    Pncc(0) = (Ast * Fy)
    Mncc(0) = 0
    
    
    Fyh2 = Val(TFyh2)
    Fy2 = Val(TFy2)
    Fc2 = Val(TFc2)
    Gc = Val(TGc)
    TLw = Val(Th)
    
    
    'Cálculo del "c" para los elementos de borde.
    tt = 0
    For v = 0 To puntos
    If Pncc(v) <= Pnpmax And tt = 0 Then
    ccc = CC(v)
    Tc = ccc
    tt = 1
    End If
    Next v
    'Fin del cálculo del "c".
    
    
    If TFyh2 <= 3500 Then
    Tsmaxtemp = (Val(Tmalla2) * ((((Val(TVar) / 8) * 2.54) ^ 2) * 3.14159265359 / 4)) / (0.002 * b)
    Else
    Tsmaxtemp = (Val(Tmalla2) * ((((Val(TVar) / 8) * 2.54) ^ 2) * 3.14159265359 / 4)) / (0.0018 * b)
    End If
    
    Tsmaxtemp = Round(Tsmaxtemp, 2)
      
    
    Acp = (b * h) - ((b - bp) * 2 * hp)
    Acv = b * h

   'Obtiene el cortante de diseño.
    Vnmayor = Abs(Val(Tabla2.TextMatrix(15, 1)))
    
    
    If Vnmayor <= Abs(Val(Tabla2.TextMatrix(15, 3))) Then
    Vnmayor = Abs(Val(Tabla2.TextMatrix(15, 3)))
    Else
    End If
    
    
    TVn = Round(Vnmayor, 3)
    'Fin de obtener el cortante de diseño.
 
    
    ThwLw = Round(Thw / TLw, 2)
    
    
    Select Case ThwLw
        Case Is <= 1.5
            TGc = 0.79
        Case Is >= 2
            TGc = 0.53
        Case Else
            TGc = Round(((((Thw / TLw) - 1.5) * (0.53 - 0.79)) / 0.5) + 0.79, 2)
    End Select
    
        
    'Art. 21.6.6.2
    Fact = (Val(Tdu) / Val(Thw))
    If Fact < 0.007 Then
    Fact = 0.007
    End If
    
    
    If Val(Tc) >= Val(Th) / (600 * (Fact)) Then
    Tborde = "Sí"
    Else
    Tborde = "No"
    End If
    
    If Val(Tdu) = 0 Then
    Tborde = "No"
    End If
    
    'Fin del art 21.6.6.2
    
    
    'Art. 21.6.6.3
    cent = Val(Th / 2)
    Inercia = ((bp * h ^ 3) / 12) - ((bp - b) * (h - (2 * bp)) ^ 3 / 12)
    Area = (b * h) - ((b - bp) * 2 * hp)
    
    
    Tbbb = 0
    For i = 0 To 17
    If ((-Pnp(i) / Area) + ((Mnp(i) * cent / Inercia))) >= (0.2 * Fc2) Or ((-Pnp(i) / Area) - ((Mnp(i) * cent / Inercia))) >= (0.2 * Fc2) Then
    Tbbb = Tbbb + 1
    End If
    Next i
    
    
    If Tbbb > 0 Then
    Tborde2 = "Sí"
    Else
    Tborde2 = "No"
    End If
    
    'Fin del art 21.6.6.3
      
      
    If (Vnmayor * 1000) >= (0.265 * Acv * ((TFc2) ^ 0.5)) Then
    smax1 = (Val(Tmalla2) * ((((Val(TVar) / 8) * 2.54) ^ 2) * 3.14159265359 / 4)) / 0.0025 * b
    Else:
    smax1 = 45
    End If
        
        
    smax2 = b * 3
        
    smax3 = h / 5
    
    smax4 = 45
    
    If Val(ThwLw) < 2 Then
    TDebe = "Sí"
    Else
    TDebe = "No"
    End If
                    
          
    smax = smax1
    If smax2 < smax1 Then
    smax = smax2
    End If
    
    If smax3 < smax Then
    smax = smax3
    End If

    If smax4 < smax Then
    smax = smax4
    End If

    TSmax = Round(smax, 2)
        
    
    'Obtiene el número de mallas requeridas.
    If (Vnmayor * 1000) >= (0.53 * Acv * ((Val(Fc2)) ^ 0.5)) Then
    Tmalla = 2
    Else
    Tmalla = 1
    End If
    
    
    If b >= 20 Then
    Tmalla = 2
    End If
    'Fin de obtener el número de mallas requeridas.
    
    
    'Obtiene el cortante máximo permitido.
    Vmax = (2.65 * Acp * 0.001 * ((Val(TFc2)) ^ 0.5))
    TVmax = Round(Vmax, 2)
    
    
    TVmax.FontBold = False
    TVmax.ForeColor = vbBlack
    
    
    If Vnmayor > Vmax Then
    TVmax.FontBold = True
    TVmax.ForeColor = vbRed
    
    Else
    End If
    'Fin de obtener el cortante máximo permitido.
    
    
    
    'Obtiene acero horizontal requerido.
    
    Vc = Val(TGc) * (Fc2) ^ 0.5 * b * Val(TphiLw) * Val(TLw)
    
    TVc = Round(Vc / 1000, 2)
    
    Vs = (Vnmayor / TphiC2) - (Vc / 1000)
    
    TVs = Round(Vs, 2)
    
    
    If Vs > 0 Then
    sreq = (Val(Tmalla2) * ((((Val(TVar) / 8) * 2.54) ^ 2) * 3.14159265359 / 4) * Fyh2 * Val(TphiLw) * Val(TLw)) / Vs / 1000
    Else
    sreq = 0
    End If
    
    
    Tsreq = Round(sreq, 2)
    
    
    sfinal = Val(Tsmaxtemp)

    If Val(TSmax) < sfinal Then
    sfinal = Val(TSmax)
    End If
    
    If Val(Tsreq) < sfinal And Val(Tsreq) <> 0 Then
    sfinal = Val(Tsreq)
    End If
        
    
    Text1 = Val(TVar)
    Text3 = Val(TVar)
    Text4 = Val(TVar)
    
    TR1 = Val(TVar)
    TR2 = Round(sfinal, 2)
    
    
    'Fin de obtener el acero horizontal requerido.
    
    
    'Obtiene el mínimo tamaño del elemento de borde.
    Tc1 = Round(Val(Tc) - (0.1 * h), 2)
    Tc2 = Round((Val(Tc) / 2), 2)
    
    
    If Tc1 > Tc2 Then
    TBordemin = Round(Tc1, 2)
    Else
    TBordemin = Round(Tc2, 2)
    End If
    'Fin de obtenter el mínimo tamaño del elemento de borde
       
       
    
    'Cargas de diseño para los elementos de borde!
    
    Pvert1max = Pvert1(0)
    Pvert2max = Pvert2(0)
    Pvert1min = Pvert1(0)
    Pvert2min = Pvert2(0)
    
    For i = 0 To 17
    Pvert1(i) = (-Pnp(i) / 2) + (Abs(Mnp(i) / h))
    Pvert2(i) = (-Pnp(i) / 2) - (Abs(Mnp(i) / h))
        
    If Pvert1(i) > Pvert1max Then
    Pvert1max = Pvert1(i)
    End If
    
    If Pvert2(i) > Pvert2max Then
    Pvert2max = Pvert2(i)
    End If
    
    If Pvert1(i) < Pvert1min Then
    Pvert1min = Pvert1(i)
    End If
    
    If Pvert2(i) < Pvert2min Then
    Pvert2min = Pvert2(i)
    End If
    
    Next i
    
    If Pvert1max > Pvert2max Then
    Pcomp = Pvert1max
    Else
    Pcomp = Pvert2max
    End If
    
    If Pvert1min < Pvert2min Then
    Pten = Abs(Pvert1min)
    Else
    Pten = Abs(Pvert2min)
    End If
    
    
    'Fin de cargas de diseño para los elementos de borde!
    
    
    Ag = (TB1 * Tb2)
    Asbordemin = Round(Ag * 0.01, 2)
    Asbordemax = Round(Ag * 0.06, 2)
    
    TAsbordemin = Round(Asbordemin, 2)
    TAsbordemax = Round(Asbordemax, 2)
    
    
    Ascompresion = ((Pcomp / (0.8 * 0.7) - (0.85 * Fc2 * Ag)) / (-0.85 * Fc2 + Fy2))
    If Ascompresion < 0 Then
    Ascompresion = 0
    End If
    
    
    Astension = (Pten) / (0.9 * Fy2)
    If Astension < 0 Then
    Astension = 0
    End If
    
    
    Asbordereq = Ascompresion
    
    If Astension > Ascompresion Then
    Asbordereq = Astension
    End If
    
    If Asbordemin > Asbordereq Then
    Asbordereq = Asbordemin
    End If
    
    
    TAsbordemax.FontBold = False
    TAsbordemax.ForeColor = vbBlack
    
    
    If Asbordereq > Asbordemax Then
    TAsbordemax.FontBold = True
    TAsbordemax.ForeColor = vbRed
    End If
    
    If TB1 = 0 And Tb2 = 0 Then
    TAsbordemax.FontBold = False
    TAsbordemax.ForeColor = vbBlack
    End If
    
    
    TAsborde = Round(Asbordereq, 2)
    
    
    End Sub
    
    
    Private Sub CEje_Click()
    
    
    CMuro_Click
    
    
    End Sub
    
    
    Private Sub Cflexion_Click()
    
  
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
    
    
    CMuro_Click
    
    
    End Sub


    Private Sub MAcerca_Click()
    
    
    FAbout.Visible = True
    
    
    End Sub
    
    
    Private Sub MColumnas_Click()
    
    Unload FMuros
    
    
    FDialog.Visible = True
    
    
    FDialog.Option2.Value = True
    
    
    End Sub
    
    
    Private Sub Mimprimir_Click()
    
    
    CommonDialog1.Orientation = cdlLandscape
    
    
    MsgPrompt = "Porfavor verifique si la impresora (" & Printer.DeviceName & ") está lista"
    i = MsgBox(MsgPrompt, vbOKCancel, "Confirmation")


    If i = vbCancel Then
        Exit Sub
    End If
    
    
    PrintForm
    
    
    End Sub
    
    
    Private Sub MMuros_Click()
    
    Unload FMuros
    
    
    FDialog.Visible = True
    
    
    FDialog.Option3.Value = True
     
    
    End Sub
    
    
    Private Sub MSalida_Click()
    
    
    SSTabMuros.Visible = False
    
    
    MSHFlexGrid1.Visible = True
    
    
    MSHFlexGrid2.Visible = False
    
    
    End Sub
    
    
    Private Sub MSalida2_Click()
    
    
    SSTabMuros.Visible = True
    
    
    MSHFlexGrid1.Visible = False
    
    
    MSHFlexGrid2.Visible = False
    
    
    End Sub
    
    
    Private Sub MVigas_Click()
    
    
    Unload FMuros
    
    
    FDialog.Visible = True
    
    
    FDialog.Option1.Value = True
    
    
    End Sub


    Private Sub Salida_Click()
      
    
    End
    
    
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
    
    
    Tb3 = Tb


End Sub


    Private Sub Tb1_Change()


    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!


    Tbp = TB1


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

    Thp = Tb2

    End Sub

    Private Sub Tb3_Change()
   
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!
    
    Tb = Tb3
    
    End Sub

    Private Sub Tbp_Change()
    
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    TB1 = Tbp

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


    Private Sub TFy_Change()


   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    TFy2 = TFy


    End Sub

    
    Private Sub TFy2_Change()
   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!


    TFy = TFy2


    End Sub

    
    Private Sub TFyh_Change()


   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!


    TFyh2 = TFyh

    
    End Sub

    
    Private Sub TFyh2_Change()


   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    TFyh = TFyh2


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
    
    
    TLw = Th
    
    
    Tdx1(0) = 5
    
    
    Tdx1(1) = 25
    
    
    Tdx1(2) = 45
    
    
    Tdx1(7) = Th - 45
    
    
    Tdx1(8) = Th - 25
    
    
    Tdx1(9) = Th - 5
    
    
    End Sub


    Private Sub Thp_Change()


    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    Tb2 = Thp


    End Sub

    
    Private Sub TLw_Change()


   'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!

    
    Th = TLw

    
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
    

    For i = 0 To 9
    
    
    If Val(Tnum(i)) And Val(Tdiam(i)) > 0 Then
    
    
    TAs1(i) = Round(Val(Tnum(i)) * 3.14159265359 * ((Val(Tdiam(i)) / 8) * 2.54) ^ 2 / 4, 2)
    Tgato(i) = "#"
    
    
    Else
    
    
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
        
    
    For i = 0 To 9
    
    
    If Val(Tnum(i)) And Val(Tdiam(i)) > 0 Then
    
    
    TAs1(i) = Round(Val(Tnum(i)) * 3.14159265359 * ((Val(Tdiam(i)) / 8) * 2.54) ^ 2 / 4, 2)
    Tgato(i) = "#"
    
    Else
    
    
    TAs1(i) = ""
    Tgato(i) = ""
    
    
    End If
    
    
    Next i
    
    
    End Sub
    
    
    Private Sub TphiC_Change()
    
    
    TphiC2 = TphiC
    
    
    End Sub
    
    
    Private Sub TphiC2_Change()
    
    
    TphiC = TphiC2
    
    
    End Sub
