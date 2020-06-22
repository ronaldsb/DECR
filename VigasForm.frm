VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form FVigas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diseño de elementos en concreto reforzado."
   ClientHeight    =   9840
   ClientLeft      =   48
   ClientTop       =   612
   ClientWidth     =   12912
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9831.02
   ScaleMode       =   0  'User
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabVigas 
      Height          =   9672
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   17060
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Flexión"
      TabPicture(0)   =   "VigasForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label34"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label35"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LHora"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CFlexion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CViga"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TFecha"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "THora"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TProyecto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "THecho"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Resultados"
      TabPicture(1)   =   "VigasForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label23"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label25"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label26"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label28"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label36"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label39"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label38"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label84"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label85"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label86"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label96"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label97"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label98"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label41"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label24"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label27"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label73"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label99"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label100"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label102"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label103"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Tpb"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Tpmax"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TAsmax"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "TAsmin"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "TVc"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TSmax2"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TVsmax"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Tpmin"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Tabla3"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "TSmax"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).ControlCount=   32
      TabCaption(2)   =   "Envolventes"
      TabPicture(2)   =   "VigasForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSChartFlexion"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "MSChartCortante"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Diseño por capacidad"
      TabPicture(3)   =   "VigasForm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label42"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label43"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label44"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label45"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label46"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label47"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label48"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label49"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label50"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label51"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label52"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label53"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label54"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label55"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label56"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label57"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Label74"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Label75"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Label76"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "Label77"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "Label58"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "Label59"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "Label60"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "Label61"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "Label62"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "Label63"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Label64"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "Label65"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "Label66"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "Label67"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "Label68"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "Label69"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Label40"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "Label78"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Label79"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "Label80"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "Label88"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "Label87"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Label89"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Label90"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Label91"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Label92"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Label93"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Label94"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "Label95"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "Label32"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).Control(46)=   "Label37"
      Tab(3).Control(46).Enabled=   0   'False
      Tab(3).Control(47)=   "TFyh2"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).Control(48)=   "TFy2"
      Tab(3).Control(48).Enabled=   0   'False
      Tab(3).Control(49)=   "TFc2"
      Tab(3).Control(49).Enabled=   0   'False
      Tab(3).Control(50)=   "Td2"
      Tab(3).Control(50).Enabled=   0   'False
      Tab(3).Control(51)=   "Thd2"
      Tab(3).Control(51).Enabled=   0   'False
      Tab(3).Control(52)=   "Th2"
      Tab(3).Control(52).Enabled=   0   'False
      Tab(3).Control(53)=   "Tb2"
      Tab(3).Control(53).Enabled=   0   'False
      Tab(3).Control(54)=   "TAsrai"
      Tab(3).Control(54).Enabled=   0   'False
      Tab(3).Control(55)=   "TAsrad"
      Tab(3).Control(55).Enabled=   0   'False
      Tab(3).Control(56)=   "TAsrabd"
      Tab(3).Control(56).Enabled=   0   'False
      Tab(3).Control(57)=   "TAsrabi"
      Tab(3).Control(57).Enabled=   0   'False
      Tab(3).Control(58)=   "TWcm"
      Tab(3).Control(58).Enabled=   0   'False
      Tab(3).Control(59)=   "TWcv"
      Tab(3).Control(59).Enabled=   0   'False
      Tab(3).Control(60)=   "TPcm1"
      Tab(3).Control(60).Enabled=   0   'False
      Tab(3).Control(61)=   "TPcv1"
      Tab(3).Control(61).Enabled=   0   'False
      Tab(3).Control(62)=   "TPcv2"
      Tab(3).Control(62).Enabled=   0   'False
      Tab(3).Control(63)=   "TR2"
      Tab(3).Control(63).Enabled=   0   'False
      Tab(3).Control(64)=   "Tdp2"
      Tab(3).Control(64).Enabled=   0   'False
      Tab(3).Control(65)=   "TR1"
      Tab(3).Control(65).Enabled=   0   'False
      Tab(3).Control(66)=   "TPcm2"
      Tab(3).Control(66).Enabled=   0   'False
      Tab(3).Control(67)=   "Tdist2"
      Tab(3).Control(67).Enabled=   0   'False
      Tab(3).Control(68)=   "Tdist1"
      Tab(3).Control(68).Enabled=   0   'False
      Tab(3).Control(69)=   "TLong2"
      Tab(3).Control(69).Enabled=   0   'False
      Tab(3).Control(70)=   "TRmax"
      Tab(3).Control(70).Enabled=   0   'False
      Tab(3).Control(71)=   "TPatas2"
      Tab(3).Control(71).Enabled=   0   'False
      Tab(3).Control(72)=   "TablaCapa1"
      Tab(3).Control(72).Enabled=   0   'False
      Tab(3).Control(73)=   "TablaCapa2"
      Tab(3).Control(73).Enabled=   0   'False
      Tab(3).Control(74)=   "Picture2"
      Tab(3).Control(74).Enabled=   0   'False
      Tab(3).Control(75)=   "Picture4"
      Tab(3).Control(75).Enabled=   0   'False
      Tab(3).Control(76)=   "Command1"
      Tab(3).Control(76).Enabled=   0   'False
      Tab(3).ControlCount=   77
      Begin VB.TextBox TSmax 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   202
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox THecho 
         BackColor       =   &H80000004&
         Height          =   288
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   200
         Top             =   480
         Width           =   2292
      End
      Begin VB.TextBox TProyecto 
         BackColor       =   &H80000004&
         Height          =   288
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   199
         Top             =   840
         Width           =   2292
      End
      Begin MSChart20Lib.MSChart MSChartCortante 
         Height          =   3960
         Left            =   -74520
         OleObjectBlob   =   "VigasForm.frx":0070
         TabIndex        =   148
         Top             =   5160
         Width           =   11712
      End
      Begin MSChart20Lib.MSChart MSChartFlexion 
         Height          =   3960
         Left            =   -74520
         OleObjectBlob   =   "VigasForm.frx":18D9
         TabIndex        =   147
         Top             =   840
         Width           =   11712
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Dar resultados"
         Height          =   495
         Left            =   -67080
         TabIndex        =   39
         Top             =   1920
         Width           =   2655
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   -68520
         Picture         =   "VigasForm.frx":3495
         ScaleHeight     =   900
         ScaleWidth      =   5532
         TabIndex        =   156
         Top             =   3120
         Width           =   5535
      End
      Begin VB.PictureBox Picture2 
         Height          =   4695
         Left            =   -74160
         Picture         =   "VigasForm.frx":B73EF
         ScaleHeight     =   4644
         ScaleWidth      =   4524
         TabIndex        =   155
         Top             =   4560
         Width           =   4575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid TablaCapa2 
         Height          =   2460
         Left            =   -68280
         TabIndex        =   151
         Top             =   6480
         Width           =   4815
         _ExtentX        =   8488
         _ExtentY        =   4339
         _Version        =   393216
         Rows            =   10
         Cols            =   5
         FixedCols       =   0
         Enabled         =   0   'False
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid TablaCapa1 
         Height          =   780
         Left            =   -67200
         TabIndex        =   150
         Top             =   5280
         Width           =   2772
         _ExtentX        =   4890
         _ExtentY        =   1376
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         Enabled         =   0   'False
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tabla3 
         Height          =   6570
         Left            =   -68520
         TabIndex        =   149
         Top             =   1320
         Width           =   5625
         _ExtentX        =   9927
         _ExtentY        =   11599
         _Version        =   393216
         Rows            =   27
         Cols            =   8
         Enabled         =   0   'False
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.TextBox Tpmin 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1320
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Height          =   8280
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   12492
         Begin VB.TextBox TSmax3 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   3708
            Locked          =   -1  'True
            TabIndex        =   205
            Top             =   7680
            Width           =   735
         End
         Begin VB.Frame Frame3 
            Height          =   1812
            Left            =   5400
            TabIndex        =   168
            Top             =   4320
            Width           =   6855
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   0
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   183
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   1
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   181
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   2
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   180
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   3
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   179
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   4
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   178
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   5
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   177
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   6
               Left            =   6000
               Locked          =   -1  'True
               TabIndex        =   176
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   7
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   175
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   8
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   174
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   9
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   173
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   10
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   172
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   11
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   171
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   12
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   170
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox Tas 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   13
               Left            =   6000
               Locked          =   -1  'True
               TabIndex        =   169
               Top             =   1320
               Width           =   615
            End
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               Height          =   900
               Left            =   960
               Picture         =   "VigasForm.frx":126941
               ScaleHeight     =   900
               ScaleWidth      =   5652
               TabIndex        =   182
               Top             =   480
               Width           =   5655
            End
            Begin VB.Label Label29 
               Caption         =   "As  [cm2]"
               Height          =   255
               Left            =   120
               TabIndex        =   185
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label30 
               Caption         =   "As  [cm2]"
               Height          =   255
               Left            =   120
               TabIndex        =   184
               Top             =   1320
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1692
            Left            =   5400
            TabIndex        =   158
            Top             =   6480
            Width           =   6855
            Begin VB.TextBox Tss 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   0
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   165
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox Tss 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   1
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   164
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox Tss 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   2
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   163
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox Tss 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   3
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   162
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox Tss 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   4
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   161
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox Tss 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   5
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   160
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox Tss 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               Height          =   285
               Index           =   6
               Left            =   6000
               Locked          =   -1  'True
               TabIndex        =   159
               Top             =   1200
               Width           =   615
            End
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   900
               Left            =   960
               Picture         =   "VigasForm.frx":1DA89B
               ScaleHeight     =   900
               ScaleWidth      =   5652
               TabIndex        =   167
               Top             =   240
               Width           =   5655
            End
            Begin VB.Label Label31 
               Caption         =   "s  [cm]"
               Height          =   252
               Left            =   120
               TabIndex        =   166
               Top             =   1200
               Width           =   732
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   4935
            Left            =   360
            Picture         =   "VigasForm.frx":28E7F5
            ScaleHeight     =   4884
            ScaleWidth      =   4524
            TabIndex        =   157
            Top             =   2520
            Width           =   4575
         End
         Begin VB.TextBox TPatas 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4680
            TabIndex        =   13
            Text            =   "2"
            Top             =   1800
            Width           =   492
         End
         Begin VB.TextBox Tdp 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Text            =   "5"
            Top             =   1800
            Width           =   495
         End
         Begin VB.OptionButton OCortante 
            Caption         =   "Cortante."
            Height          =   255
            Left            =   9480
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OFlexion 
            Caption         =   "Flexión."
            Height          =   255
            Left            =   7800
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox TLong 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   124
            Top             =   1800
            Width           =   492
         End
         Begin VB.TextBox Tvl 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4680
            TabIndex        =   11
            Text            =   "8"
            Top             =   1080
            Width           =   492
         End
         Begin VB.TextBox Tvt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4680
            TabIndex        =   12
            Text            =   "3"
            Top             =   1440
            Width           =   492
         End
         Begin VB.TextBox Tphi 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4680
            TabIndex        =   10
            Text            =   "0.85"
            Top             =   720
            Width           =   492
         End
         Begin VB.TextBox TFyh 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4680
            TabIndex        =   9
            Text            =   "2800"
            Top             =   360
            Width           =   492
         End
         Begin VB.TextBox TB1 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0.85"
            Top             =   1440
            Width           =   492
         End
         Begin VB.TextBox Tppb 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Text            =   "0.75"
            Top             =   1080
            Width           =   492
         End
         Begin VB.TextBox TFy 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   7
            Text            =   "4200"
            Top             =   720
            Width           =   492
         End
         Begin VB.TextBox TFc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   6
            Text            =   "210"
            Top             =   360
            Width           =   492
         End
         Begin VB.TextBox Td 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "75"
            Top             =   1440
            Width           =   492
         End
         Begin VB.TextBox Thd 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Text            =   "5"
            Top             =   1080
            Width           =   492
         End
         Begin VB.TextBox Th 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   3
            Text            =   "80"
            Top             =   720
            Width           =   492
         End
         Begin VB.TextBox Tb 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            TabIndex        =   2
            Text            =   "30"
            Top             =   360
            Width           =   492
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tabla 
            Height          =   2955
            Left            =   6360
            TabIndex        =   153
            Top             =   1080
            Width           =   5850
            _ExtentX        =   10329
            _ExtentY        =   5207
            _Version        =   393216
            Rows            =   12
            Cols            =   8
            Enabled         =   0   'False
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   8
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Tabla1 
            Height          =   2700
            Left            =   6360
            TabIndex        =   154
            Top             =   1080
            Width           =   5850
            _ExtentX        =   10329
            _ExtentY        =   4763
            _Version        =   393216
            Rows            =   11
            Cols            =   8
            Enabled         =   0   'False
            ScrollBars      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   8
         End
         Begin VB.Label Label105 
            Caption         =   "Smax (2d) ="
            Height          =   252
            Left            =   2760
            TabIndex        =   207
            Top             =   7680
            Width           =   960
         End
         Begin VB.Label Label104 
            Caption         =   "cm"
            Height          =   252
            Left            =   4560
            TabIndex        =   206
            Top             =   7680
            Width           =   372
         End
         Begin VB.Label Label101 
            Caption         =   " ="
            Height          =   252
            Left            =   2208
            TabIndex        =   198
            Top             =   1440
            Width           =   132
         End
         Begin VB.Label Label72 
            Caption         =   "b ="
            Height          =   252
            Left            =   2112
            TabIndex        =   192
            Top             =   1080
            Width           =   252
         End
         Begin VB.Label Label71 
            Caption         =   "Espaciamiento requerido de los aros."
            Height          =   252
            Left            =   7560
            TabIndex        =   191
            Top             =   6240
            Width           =   2892
         End
         Begin VB.Label Label70 
            Caption         =   "Acero requerido en la sección."
            Height          =   252
            Left            =   7680
            TabIndex        =   190
            Top             =   4080
            Width           =   2292
         End
         Begin VB.Label Label83 
            Caption         =   "# Patas ="
            Height          =   252
            Left            =   3840
            TabIndex        =   129
            Top             =   1800
            Width           =   732
         End
         Begin VB.Label Label82 
            Caption         =   "m"
            Height          =   252
            Left            =   3120
            TabIndex        =   126
            Top             =   1800
            Width           =   252
         End
         Begin VB.Label Label81 
            Caption         =   "Long ="
            Height          =   252
            Left            =   1920
            TabIndex        =   125
            Top             =   1800
            Width           =   492
         End
         Begin VB.Label Label20 
            Caption         =   "V.L # ="
            Height          =   252
            Left            =   4080
            TabIndex        =   63
            Top             =   1080
            Width           =   492
         End
         Begin VB.Label Label19 
            Caption         =   "Aros #"
            Height          =   252
            Left            =   4080
            TabIndex        =   62
            Top             =   1440
            Width           =   492
         End
         Begin VB.Label Label17 
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
            Left            =   4320
            TabIndex        =   61
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label16 
            Caption         =   "kg/cm2"
            Height          =   252
            Left            =   5280
            TabIndex        =   60
            Top             =   360
            Width           =   612
         End
         Begin VB.Label Label15 
            Caption         =   "fyh ="
            Height          =   252
            Left            =   4200
            TabIndex        =   59
            Top             =   360
            Width           =   492
         End
         Begin VB.Label Label21 
            Caption         =   "d' ="
            Height          =   252
            Left            =   480
            TabIndex        =   58
            Top             =   1800
            Width           =   252
         End
         Begin VB.Label Label2 
            Caption         =   "cm"
            Height          =   252
            Left            =   1440
            TabIndex        =   57
            Top             =   1800
            Width           =   372
         End
         Begin VB.Label Label18 
            Caption         =   "b1"
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
            Left            =   2040
            TabIndex        =   56
            Top             =   1440
            Width           =   252
         End
         Begin VB.Label Label14 
            Caption         =   "kg/cm2"
            Height          =   252
            Left            =   3120
            TabIndex        =   55
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label13 
            Caption         =   "kg/cm2"
            Height          =   252
            Left            =   3120
            TabIndex        =   54
            Top             =   360
            Width           =   612
         End
         Begin VB.Label Label12 
            Caption         =   "%r"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   7.8
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1890
            TabIndex        =   53
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "fy ="
            Height          =   252
            Left            =   2040
            TabIndex        =   52
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label10 
            Caption         =   "f'c ="
            Height          =   252
            Left            =   2040
            TabIndex        =   51
            Top             =   360
            Width           =   372
         End
         Begin VB.Label Label9 
            Caption         =   "cm"
            Height          =   252
            Left            =   1440
            TabIndex        =   50
            Top             =   1440
            Width           =   372
         End
         Begin VB.Label Label8 
            Caption         =   "cm"
            Height          =   252
            Left            =   1440
            TabIndex        =   49
            Top             =   1080
            Width           =   372
         End
         Begin VB.Label Label7 
            Caption         =   "cm"
            Height          =   252
            Left            =   1440
            TabIndex        =   48
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label6 
            Caption         =   "d ="
            Height          =   252
            Left            =   480
            TabIndex        =   47
            Top             =   1440
            Width           =   252
         End
         Begin VB.Label Label5 
            Caption         =   "h-d ="
            Height          =   252
            Left            =   360
            TabIndex        =   46
            Top             =   1080
            Width           =   372
         End
         Begin VB.Label Label4 
            Caption         =   "h ="
            Height          =   252
            Left            =   480
            TabIndex        =   45
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label3 
            Caption         =   "cm"
            Height          =   252
            Left            =   1440
            TabIndex        =   44
            Top             =   360
            Width           =   372
         End
         Begin VB.Label Label1 
            Caption         =   "b ="
            Height          =   252
            Left            =   480
            TabIndex        =   43
            Top             =   360
            Width           =   372
         End
         Begin VB.Label LMomentos 
            Caption         =   "Momentos en toneladas por metro."
            Height          =   255
            Left            =   7920
            TabIndex        =   85
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label LCortantes 
            Caption         =   "Cortantes en toneladas."
            Height          =   255
            Left            =   8400
            TabIndex        =   186
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label LTorsion 
            Caption         =   "Torsión en toneladas por metro."
            Height          =   255
            Left            =   7920
            TabIndex        =   201
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.TextBox TPatas2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   25
         Text            =   "2"
         Top             =   3720
         Width           =   492
      End
      Begin VB.TextBox TRmax 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -68760
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TLong2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   26
         Text            =   "6"
         Top             =   4080
         Width           =   492
      End
      Begin VB.TextBox Tdist1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   31
         Text            =   "0"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Tdist2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   34
         Text            =   "0"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox TPcm2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   32
         Text            =   "0"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TR1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -68760
         Locked          =   -1  'True
         TabIndex        =   133
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TVsmax 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73668
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Tdp2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   21
         Text            =   "5"
         Top             =   2280
         Width           =   492
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
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TR2 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -68760
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TPcv2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   33
         Text            =   "0"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox TPcv1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   30
         Text            =   "0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TPcm1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   29
         Text            =   "0"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TWcv 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   28
         Text            =   "3000"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TWcm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71280
         TabIndex        =   27
         Text            =   "2000"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TAsrabi 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -68280
         TabIndex        =   36
         Text            =   "10.14"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox TAsrabd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -64080
         TabIndex        =   38
         Text            =   "10.14"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox TAsrad 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -64080
         TabIndex        =   37
         Text            =   "15.21"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox TAsrai 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -68280
         TabIndex        =   35
         Text            =   "15.21"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Tb2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   17
         Text            =   "30"
         Top             =   840
         Width           =   492
      End
      Begin VB.TextBox Th2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   18
         Text            =   "80"
         Top             =   1200
         Width           =   492
      End
      Begin VB.TextBox Thd2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   19
         Text            =   "5"
         Top             =   1560
         Width           =   492
      End
      Begin VB.TextBox Td2 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "75"
         Top             =   1920
         Width           =   492
      End
      Begin VB.TextBox TFc2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   22
         Text            =   "210"
         Top             =   2640
         Width           =   492
      End
      Begin VB.TextBox TFy2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   23
         Text            =   "4200"
         Top             =   3000
         Width           =   492
      End
      Begin VB.TextBox TFyh2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   24
         Text            =   "2800"
         Top             =   3360
         Width           =   492
      End
      Begin VB.TextBox TSmax2 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -71052
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox TVc 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   2160
         Width           =   735
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
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox CViga 
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
         ItemData        =   "VigasForm.frx":2FDD47
         Left            =   1680
         List            =   "VigasForm.frx":2FDD49
         TabIndex        =   1
         Top             =   765
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
         TabIndex        =   69
         Text            =   "ELEMENTO #"
         Top             =   765
         Width           =   1335
      End
      Begin VB.CommandButton CFlexion 
         Caption         =   "Actualizar cálculos."
         Height          =   495
         Left            =   3240
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TAsmin 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TAsmax 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Tpmax 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Tpb 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label103 
         Caption         =   "Smax ="
         Height          =   252
         Left            =   -71640
         TabIndex        =   204
         Top             =   2160
         Width           =   612
      End
      Begin VB.Label Label102 
         Caption         =   "cm"
         Height          =   252
         Left            =   -70188
         TabIndex        =   203
         Top             =   2160
         Width           =   372
      End
      Begin VB.Label Label100 
         Caption         =   " r"
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
         Left            =   -74400
         TabIndex        =   197
         Top             =   1680
         Width           =   132
      End
      Begin VB.Label Label99 
         Caption         =   "max ="
         Height          =   252
         Left            =   -74280
         TabIndex        =   196
         Top             =   1680
         Width           =   492
      End
      Begin VB.Label Label73 
         Caption         =   " r"
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
         Left            =   -74400
         TabIndex        =   195
         Top             =   1320
         Width           =   132
      End
      Begin VB.Label Label27 
         Caption         =   "min ="
         Height          =   252
         Left            =   -74280
         TabIndex        =   194
         Top             =   1320
         Width           =   492
      End
      Begin VB.Label Label24 
         Caption         =   " r"
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
         Left            =   -74280
         TabIndex        =   193
         Top             =   960
         Width           =   132
      End
      Begin VB.Label Label41 
         Caption         =   "Resultados generales de todos los diseños."
         Height          =   252
         Left            =   -67440
         TabIndex        =   189
         Top             =   960
         Width           =   3252
      End
      Begin VB.Label Label37 
         Caption         =   "Resultados del diseño por capacidad."
         Height          =   252
         Left            =   -66960
         TabIndex        =   188
         Top             =   6240
         Width           =   2772
      End
      Begin VB.Label Label32 
         Caption         =   "Espaciamiento requerido de los aros."
         Height          =   252
         Left            =   -67080
         TabIndex        =   187
         Top             =   4920
         Width           =   2772
      End
      Begin VB.Label Label98 
         Caption         =   "2. Mn (+) o Mn (-)   >=   0.25 [ Mn (+) , Mn (-) ] max"
         Height          =   252
         Left            =   -74280
         TabIndex        =   146
         Top             =   7200
         Width           =   3612
      End
      Begin VB.Label Label97 
         Caption         =   "1. Mn (+ ) , i, d   >=  (0.5*Mn (-)  i, d"
         Height          =   252
         Left            =   -74280
         TabIndex        =   145
         Top             =   6840
         Width           =   2772
      End
      Begin VB.Label Label96 
         Caption         =   "** Nota: A la hora de detallar, recordar que:"
         Height          =   252
         Left            =   -74640
         TabIndex        =   144
         Top             =   6240
         Width           =   3252
      End
      Begin VB.Label Label95 
         Caption         =   "# Patas ="
         Height          =   252
         Left            =   -74280
         TabIndex        =   143
         Top             =   3720
         Width           =   732
      End
      Begin VB.Label Label94 
         Caption         =   "kg"
         Height          =   252
         Left            =   -67800
         TabIndex        =   142
         Top             =   1560
         Width           =   372
      End
      Begin VB.Label Label93 
         Caption         =   "Rmax ="
         Height          =   252
         Left            =   -69600
         TabIndex        =   141
         Top             =   1560
         Width           =   612
      End
      Begin VB.Label Label92 
         Caption         =   "Long ="
         Height          =   252
         Left            =   -74040
         TabIndex        =   139
         Top             =   4080
         Width           =   492
      End
      Begin VB.Label Label91 
         Caption         =   "m"
         Height          =   252
         Left            =   -72840
         TabIndex        =   138
         Top             =   4080
         Width           =   252
      End
      Begin VB.Label Label90 
         Caption         =   "m"
         Height          =   252
         Left            =   -70320
         TabIndex        =   137
         Top             =   3360
         Width           =   372
      End
      Begin VB.Label Label89 
         Caption         =   "d2 ="
         Height          =   252
         Left            =   -71760
         TabIndex        =   136
         Top             =   3360
         Width           =   372
      End
      Begin VB.Label Label87 
         Caption         =   "m"
         Height          =   252
         Left            =   -70320
         TabIndex        =   135
         Top             =   2280
         Width           =   252
      End
      Begin VB.Label Label88 
         Caption         =   "d1 ="
         Height          =   252
         Left            =   -71760
         TabIndex        =   134
         Top             =   2280
         Width           =   372
      End
      Begin VB.Label Label86 
         Caption         =   "Vsmax ="
         Height          =   252
         Left            =   -74520
         TabIndex        =   83
         Top             =   2520
         Width           =   612
      End
      Begin VB.Label Label85 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -72828
         TabIndex        =   84
         Top             =   2520
         Width           =   372
      End
      Begin VB.Label Label84 
         Caption         =   "Ton"
         Height          =   252
         Left            =   -72840
         TabIndex        =   131
         Top             =   2160
         Width           =   372
      End
      Begin VB.Label Label38 
         Caption         =   "cm"
         Height          =   252
         Left            =   -70200
         TabIndex        =   130
         Top             =   2520
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
         Left            =   9960
         TabIndex        =   128
         Top             =   840
         Width           =   492
      End
      Begin VB.Label Label80 
         Caption         =   "R2 ="
         Height          =   252
         Left            =   -69360
         TabIndex        =   123
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label Label79 
         Caption         =   "kg"
         Height          =   252
         Left            =   -67800
         TabIndex        =   122
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label Label78 
         Caption         =   "Kg"
         Height          =   252
         Left            =   -67800
         TabIndex        =   121
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label40 
         Caption         =   "R1 ="
         Height          =   252
         Left            =   -69360
         TabIndex        =   120
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label69 
         Caption         =   "Pcm 2 ="
         Height          =   252
         Left            =   -72000
         TabIndex        =   118
         Top             =   2640
         Width           =   612
      End
      Begin VB.Label Label68 
         Caption         =   "Kg"
         Height          =   252
         Left            =   -70320
         TabIndex        =   117
         Top             =   2640
         Width           =   372
      End
      Begin VB.Label Label67 
         Caption         =   "kg"
         Height          =   252
         Left            =   -70320
         TabIndex        =   116
         Top             =   3000
         Width           =   372
      End
      Begin VB.Label Label66 
         Caption         =   "Pcv 2 ="
         Height          =   252
         Left            =   -72000
         TabIndex        =   115
         Top             =   3000
         Width           =   612
      End
      Begin VB.Label Label65 
         Caption         =   "Kg"
         Height          =   252
         Left            =   -70320
         TabIndex        =   114
         Top             =   1920
         Width           =   372
      End
      Begin VB.Label Label64 
         Caption         =   "Kg"
         Height          =   252
         Left            =   -70320
         TabIndex        =   113
         Top             =   1560
         Width           =   372
      End
      Begin VB.Label Label63 
         Caption         =   "Kg/m"
         Height          =   252
         Left            =   -70320
         TabIndex        =   112
         Top             =   1200
         Width           =   492
      End
      Begin VB.Label Label62 
         Caption         =   "Pcv 1 ="
         Height          =   252
         Left            =   -72000
         TabIndex        =   111
         Top             =   1920
         Width           =   612
      End
      Begin VB.Label Label61 
         Caption         =   "Pcm 1 ="
         Height          =   252
         Left            =   -72000
         TabIndex        =   110
         Top             =   1560
         Width           =   612
      End
      Begin VB.Label Label60 
         Caption         =   "Wcv ="
         Height          =   252
         Left            =   -71880
         TabIndex        =   109
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label59 
         Caption         =   "Kg/m"
         Height          =   252
         Left            =   -70320
         TabIndex        =   108
         Top             =   840
         Width           =   492
      End
      Begin VB.Label Label58 
         Caption         =   "Wcm ="
         Height          =   252
         Left            =   -71880
         TabIndex        =   107
         Top             =   840
         Width           =   612
      End
      Begin VB.Label Label77 
         Caption         =   "cm2"
         Height          =   252
         Left            =   -63480
         TabIndex        =   106
         Top             =   4080
         Width           =   372
      End
      Begin VB.Label Label76 
         Caption         =   "cm2"
         Height          =   252
         Left            =   -63360
         TabIndex        =   105
         Top             =   2760
         Width           =   372
      End
      Begin VB.Label Label75 
         Caption         =   "cm2"
         Height          =   252
         Left            =   -67560
         TabIndex        =   104
         Top             =   4080
         Width           =   372
      End
      Begin VB.Label Label74 
         Caption         =   "cm2"
         Height          =   252
         Left            =   -67560
         TabIndex        =   103
         Top             =   2760
         Width           =   372
      End
      Begin VB.Label Label57 
         Caption         =   "b ="
         Height          =   252
         Left            =   -73800
         TabIndex        =   102
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label56 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72840
         TabIndex        =   101
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label55 
         Caption         =   "h ="
         Height          =   252
         Left            =   -73800
         TabIndex        =   100
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label Label54 
         Caption         =   "h-d ="
         Height          =   252
         Left            =   -73920
         TabIndex        =   99
         Top             =   1560
         Width           =   372
      End
      Begin VB.Label Label53 
         Caption         =   "d ="
         Height          =   252
         Left            =   -73800
         TabIndex        =   98
         Top             =   1920
         Width           =   252
      End
      Begin VB.Label Label52 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72840
         TabIndex        =   97
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label Label51 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72840
         TabIndex        =   96
         Top             =   1560
         Width           =   372
      End
      Begin VB.Label Label50 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72840
         TabIndex        =   95
         Top             =   1920
         Width           =   372
      End
      Begin VB.Label Label49 
         Caption         =   "f'c ="
         Height          =   252
         Left            =   -73920
         TabIndex        =   94
         Top             =   2640
         Width           =   372
      End
      Begin VB.Label Label48 
         Caption         =   "fy ="
         Height          =   252
         Left            =   -73920
         TabIndex        =   93
         Top             =   3000
         Width           =   372
      End
      Begin VB.Label Label47 
         Caption         =   "kg/cm2"
         Height          =   252
         Left            =   -72840
         TabIndex        =   92
         Top             =   2640
         Width           =   612
      End
      Begin VB.Label Label46 
         Caption         =   "kg/cm2"
         Height          =   252
         Left            =   -72840
         TabIndex        =   91
         Top             =   3000
         Width           =   612
      End
      Begin VB.Label Label45 
         Caption         =   "cm"
         Height          =   252
         Left            =   -72840
         TabIndex        =   90
         Top             =   2280
         Width           =   372
      End
      Begin VB.Label Label44 
         Caption         =   "d' ="
         Height          =   252
         Left            =   -73800
         TabIndex        =   89
         Top             =   2280
         Width           =   252
      End
      Begin VB.Label Label43 
         Caption         =   "fyh ="
         Height          =   252
         Left            =   -73920
         TabIndex        =   88
         Top             =   3360
         Width           =   492
      End
      Begin VB.Label Label42 
         Caption         =   "kg/cm2"
         Height          =   252
         Left            =   -72840
         TabIndex        =   87
         Top             =   3360
         Width           =   612
      End
      Begin VB.Label Label39 
         Caption         =   "Smax (2d) ="
         Height          =   252
         Left            =   -72000
         TabIndex        =   82
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label36 
         Caption         =   "Vc ="
         Height          =   252
         Left            =   -74172
         TabIndex        =   81
         Top             =   2160
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
         Left            =   9840
         TabIndex        =   78
         Top             =   492
         Width           =   612
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
         Left            =   6360
         TabIndex        =   76
         Top             =   840
         Width           =   852
      End
      Begin VB.Label Label28 
         Caption         =   "b ="
         Height          =   252
         Left            =   -74160
         TabIndex        =   75
         Top             =   960
         Width           =   252
      End
      Begin VB.Label Label26 
         Caption         =   "cm2"
         Height          =   252
         Left            =   -70200
         TabIndex        =   74
         Top             =   960
         Width           =   372
      End
      Begin VB.Label Label22 
         Caption         =   "Asmin ="
         Height          =   252
         Left            =   -71760
         TabIndex        =   73
         Top             =   960
         Width           =   612
      End
      Begin VB.Label Label25 
         Caption         =   "Asmax ="
         Height          =   252
         Left            =   -71760
         TabIndex        =   72
         Top             =   1320
         Width           =   612
      End
      Begin VB.Label Label23 
         Caption         =   "cm2"
         Height          =   252
         Left            =   -70200
         TabIndex        =   71
         Top             =   1320
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
         Left            =   6240
         TabIndex        =   70
         Top             =   492
         Width           =   972
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   9600
      Left            =   120
      TabIndex        =   86
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
      Left            =   3240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.mdb"
      InitDir         =   "C:\My Documents\Diseño\"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   9600
      Left            =   120
      TabIndex        =   152
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
      FormatString    =   "Beam|Load|Loc|M2|M3|P|Story|T|V2|V3"
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
   Begin VB.Menu MImprimir 
      Caption         =   "    Imprimir    "
   End
   Begin VB.Menu MAbout 
      Caption         =   "    Acerca de     "
   End
   Begin VB.Menu Salida 
      Caption         =   "    Salir    "
   End
End
Attribute VB_Name = "FVigas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Private Const MARGIN_SIZE = 60      ' in Twips
    Private datPrimaryRS As ADODB.Recordset
    
    
    Private Sub Form_Load()
    'Cosas que el programa carga al inicio del módulo de vigas.
    
    
    FVigas.Visible = True
    
    
    FDialog.Visible = False
    
    
    FDialog.Visible = False
    
    
    OFlexion.Value = True
    
    
    SSTabVigas.Visible = True
    
    
    Importar
    
    
    Acomodar
    
    
    MsgBox ("Este programa es para usos académicos únicamente!"), vbCritical
    
    
    TProyecto = Proyectos
    
    
    THecho = Hechos
    
    
    'Fin de las cosas que el programa carga al inicio del módulo de vigas.
    
    
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
    
    
    'Convierte los datos que son números a toneladas
    For c = 2 To MSHFlexGrid1.Cols - 1
    For r = 1 To MSHFlexGrid1.Rows - 1
    Datos(r, c) = (MSHFlexGrid1.TextMatrix(r, c)) / 1000
    Next r
    Next c
    'Fin de pasar los datos de salida a una matriz
        
    ReDim Frame(MSHFlexGrid1.Rows)
    
    'Guarda los números de elementos
    j = 1
    For i = 1 To MSHFlexGrid1.Rows - 1
    If Datos(i, 0) <> Datos(i + 1, 0) Then
    Frame(j) = Datos(i, 0)
    j = j + 1
    End If
    Next i
    FrameU = j
    'Fin de guardar los números de frames
    
    
    End Sub
    
    
    Private Sub Command1_Click()
    
    
    CViga_Click
    
    
    End Sub
    
    
    Sub Acomodar()
    
    
    'Centra toda la Tabla.
    For i = 0 To 7
    For j = 0 To 11
    With Tabla
        .Row = j
        .Col = i
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
    End With
    Next j
    Next i
    
    
    'Le da el ancho deseado a la columnas de la tabla.
    For i = 1 To 11
    With Tabla
        .ColWidth(0) = 840
        .ColWidth(i) = 710
    End With
    Next i
    
    
    'Centra toda la Tabla1.
    For i = 0 To 7
    For j = 0 To 10
    With Tabla1
        .Row = j
        .Col = i
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
    End With
    Next j
    Next i
    
    
    'Le da el ancho deseado a la columnas de la tabla 1.
    For i = 1 To 10
    With Tabla1
        .ColWidth(0) = 840
        .ColWidth(i) = 710
    End With
    Next i
    
    
    'Centra toda la Tabla3.
    For i = 0 To 7
    For j = 0 To 26
    With Tabla3
        .Row = j
        .Col = i
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
    End With
    Next j
    Next i
    
    
    'Le da el ancho deseado a la columnas de la tabla 3.
    For i = 1 To 7
    With Tabla3
        .ColWidth(0) = 820
        .ColWidth(i) = 670
            
    End With
    Next i
    'Fin de Ancho.
    
    
    'Centra toda la TablaCapa1.
    For i = 0 To 2
    For j = 0 To 2
    With TablaCapa1
        .Row = j
        .Col = i
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
    End With
    Next j
    Next i
    
    
    'Centra toda la TablaCapa2.
    For i = 0 To 4
    For j = 0 To 9
    With TablaCapa2
        .Row = j
        .Col = i
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
    End With
    Next j
    Next i
    
    
    'Pone los datos de la tabla iniciales.
    Tabla.TextMatrix(0, 0) = "CARGAS"
    Tabla.TextMatrix(1, 0) = "Muerta"
    Tabla.TextMatrix(2, 0) = "Viva"
    Tabla.TextMatrix(3, 0) = "Sismo X"
    Tabla.TextMatrix(4, 0) = "Sismo Y"
    Tabla.TextMatrix(5, 0) = "C1"
    Tabla.TextMatrix(6, 0) = "C2"
    Tabla.TextMatrix(7, 0) = "C3"
    Tabla.TextMatrix(8, 0) = "C4"
    Tabla.TextMatrix(9, 0) = "C5"
    Tabla.TextMatrix(10, 0) = "MIN"
    Tabla.TextMatrix(11, 0) = "MAX"
    
    
    Tabla.TextMatrix(0, 1) = "END-I"
    Tabla.TextMatrix(0, 2) = "1/6-PT"
    Tabla.TextMatrix(0, 3) = "2/6-PT"
    Tabla.TextMatrix(0, 4) = "1/2-PT"
    Tabla.TextMatrix(0, 5) = "4/6-PT"
    Tabla.TextMatrix(0, 6) = "5/6-PT"
    Tabla.TextMatrix(0, 7) = "END-J"
    
    
    Tabla1.TextMatrix(0, 0) = "CARGAS"
    Tabla1.TextMatrix(1, 0) = "Muerta"
    Tabla1.TextMatrix(2, 0) = "Viva"
    Tabla1.TextMatrix(3, 0) = "Sismo X"
    Tabla1.TextMatrix(4, 0) = "Sismo Y"
    Tabla1.TextMatrix(5, 0) = "C1"
    Tabla1.TextMatrix(6, 0) = "C2"
    Tabla1.TextMatrix(7, 0) = "C3"
    Tabla1.TextMatrix(8, 0) = "C4"
    Tabla1.TextMatrix(9, 0) = "C5"
    Tabla1.TextMatrix(10, 0) = "MAX"
    
    
    Tabla1.TextMatrix(0, 1) = "END-I"
    Tabla1.TextMatrix(0, 2) = "1/6-PT"
    Tabla1.TextMatrix(0, 3) = "2/6-PT"
    Tabla1.TextMatrix(0, 4) = "1/2-PT"
    Tabla1.TextMatrix(0, 5) = "4/6-PT"
    Tabla1.TextMatrix(0, 6) = "5/6-PT"
    Tabla1.TextMatrix(0, 7) = "END-J"
      
    
    Tabla3.TextMatrix(1, 0) = "As"
    Tabla3.TextMatrix(2, 0) = "p/pb"
    Tabla3.TextMatrix(3, 0) = "a"
    Tabla3.TextMatrix(4, 0) = "1.33As"
    Tabla3.TextMatrix(5, 0) = "Refuerzo"
    Tabla3.TextMatrix(6, 0) = "As'"
    Tabla3.TextMatrix(7, 0) = "Amax+As'"
    Tabla3.TextMatrix(8, 0) = "c"
    Tabla3.TextMatrix(9, 0) = "Fs'"
    Tabla3.TextMatrix(10, 0) = "As Final"
    
    
    Tabla3.TextMatrix(12, 0) = "As"
    Tabla3.TextMatrix(13, 0) = "p/pb"
    Tabla3.TextMatrix(14, 0) = "a"
    Tabla3.TextMatrix(15, 0) = "1.33As"
    Tabla3.TextMatrix(16, 0) = "Refuerzo"
    Tabla3.TextMatrix(17, 0) = "As'"
    Tabla3.TextMatrix(18, 0) = "Amax+As'"
    Tabla3.TextMatrix(19, 0) = "c"
    Tabla3.TextMatrix(20, 0) = "Fs'"
    Tabla3.TextMatrix(21, 0) = "As Final"
    
    
    Tabla3.TextMatrix(23, 0) = "Vu"
    Tabla3.TextMatrix(24, 0) = "Vs"
    Tabla3.TextMatrix(26, 0) = "s final"
     
    
    Tabla3.TextMatrix(0, 1) = "END-I"
    Tabla3.TextMatrix(0, 2) = "1/6-PT"
    Tabla3.TextMatrix(0, 3) = "2/6-PT"
    Tabla3.TextMatrix(0, 4) = "1/2-PT"
    Tabla3.TextMatrix(0, 5) = "4/6-PT"
    Tabla3.TextMatrix(0, 6) = "5/6-PT"
    Tabla3.TextMatrix(0, 7) = "END-J"
       
    
    TablaCapa1.TextMatrix(0, 1) = "s # 3"
    TablaCapa1.TextMatrix(1, 1) = "s # 4"
    TablaCapa1.TextMatrix(2, 1) = "s # 5"
    
    
    TablaCapa2.TextMatrix(0, 0) = "Cálculos"
    TablaCapa2.TextMatrix(0, 1) = "END I"
    
    
    TablaCapa2.TextMatrix(1, 0) = "a Arriba"
    TablaCapa2.TextMatrix(2, 0) = "a Abajo"
    TablaCapa2.TextMatrix(3, 0) = "Mncpi Ar. "
    TablaCapa2.TextMatrix(4, 0) = "Mncpi Ab."
    TablaCapa2.TextMatrix(5, 0) = "Vu"
    TablaCapa2.TextMatrix(6, 0) = "Vu/2"
    TablaCapa2.TextMatrix(7, 0) = "Vleft"
    TablaCapa2.TextMatrix(8, 0) = "Vc"
    TablaCapa2.TextMatrix(9, 0) = "Vs"
    
    
    TablaCapa2.TextMatrix(0, 3) = "Cálculos"
    TablaCapa2.TextMatrix(0, 4) = "END J"
    
    
    TablaCapa2.TextMatrix(1, 3) = "a Arriba"
    TablaCapa2.TextMatrix(2, 3) = "a Abajo"
    TablaCapa2.TextMatrix(3, 3) = "Mncpj Ar."
    TablaCapa2.TextMatrix(4, 3) = "Mncpj Ab."
    TablaCapa2.TextMatrix(5, 3) = "Vu"
    TablaCapa2.TextMatrix(6, 3) = "Vu/2"
    TablaCapa2.TextMatrix(7, 3) = "Vleft"
    TablaCapa2.TextMatrix(8, 3) = "Vc"
    TablaCapa2.TextMatrix(9, 3) = "Vs"
    
    
    'Pone la lista de elementos en el combobox de # de Viga.
    
    
    CViga.Clear
    
    
    For i = 1 To FrameU - 1
    CViga.AddItem Frame(i)
    Next i
    'Fin de poner la lista de elementos en el combobox de # de Viga.
       
    CViga = CViga.List(0)
    
    End Sub
    
    
    Private Sub CViga_Click()
    
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 6
    MsgBox ("Precaución!! Algún dato de entrada es igual a 0 ó a un número NO VALIDO!"), vbCritical
    Exit Sub
    Case 9
    'Para poner en blanco las casillas cada corrida!
    For i = 1 To 11
    For j = 1 To 7
    Tabla.TextMatrix(i, j) = ""
    Next j
    Next i
    For i = 1 To 10
    For j = 1 To 7
    Tabla.TextMatrix(i, j) = ""
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
    
    
    
    TFecha = Format$(Now, "d / m / yyyy")
    THora = Format$(Now, "h:mm AM/PM")
    Importacion
    
    
    'Declaración de variables.
    Dim b, h, hd, d, dp, Fc, Fy, ppb, B1, Fyh, phic, vl, vt, Av, Patas As Single
    Dim Asmin, Asmax, amax, pb, pmin, pmax, AA, BB, CC, Es, Esp, Fs, Fsp, c As Single
    Dim Mu(14), Ass(14), Assn(14), pp(14), p(14), Mn, M1(14), Assc(14), Asfinal(14), Vs(14), smax(5), s(14), sfinal(14) As Single
    Dim ComboMin, ComboMax As Single
    
    
    'Declaración y colocación de los datos iniciales.
    Fc = TFc
    Fy = TFy
    Fyh = TFyh
    ppb = Tppb
    b = Tb
    h = Th
    hd = Thd
    d = h - hd
    Td = d
    dp = Tdp
    phic = Tphi
    
    
    Vc = 0.53 * ((Fc) ^ 0.5) * b * d * (1 / 1000)
    TVc = Round(Vc, 2)
    
    Vsmax = 4 * Vc
    TVsmax = Round(Vsmax, 2)
    
    
    If Fc <= 280 Then
    B1 = 0.85
    Else
    B1 = (0.85 - 0.05 * (Fc - 280) / (70))
    End If
    If Fc >= 560 Then
    B1 = 0.65
    Else
    End If
    
    
    TB1 = B1
    
    
    pb = 0.85 * B1 * (Fc / Fy) * ((6300) / (6300 + Fy))
    Tpb = Round(pb, 4)
    'Fin de la declaración y colocación de los datos iniciales.
    
    
    'Otros cálculos necesarios.
    Asmin = (14 / Fy) * b * d
    Asmax = ppb * pb * b * d
    TAsmin = Round(Asmin, 2)
    TAsmax = Round(Asmax, 2)
    
    
    pmin = Asmin / (b * d)
    pmax = Asmax / (b * d)
    Tpmin = Round(pmin, 4)
    Tpmax = Round(pmax, 4)
    'Fin de otros cálculos necesarios.
    
    
    'Error Handler para la formula del acero!
    On Error GoTo ErrorAss
ErrorAss:
    Select Case Err.Number
    Case 5
    MsgBox "Aumentar el tamaño del elemento y actualizar datos. (La sección no es capaz de resistir esa flexion)"
    Exit Sub
    End Select
    'Fin del errorhandler del acero.
    
    
    'Calcula el acero requerido a tension arriba.
    For i = 1 To 7
    Mu(i) = Abs(Tabla.TextMatrix(10, i))
    Ass(i) = ((0.9 * Fy * d) - (((-0.9 * Fy * d) ^ 2) - (4 * ((Fy ^ 2 * 0.9) / (1.7 * Fc * b)) * Abs(Mu(i)) * 100000)) ^ 0.5) / (2 * ((Fy ^ 2 * 0.9) / (1.7 * Fc * b)))
    
    
    'Aceros finales posibles.
    Asfinal(i) = Ass(i)
    If Ass(i) < Asmin Then Asfinal(i) = Asmin
    
    
    If Ass(i) * 1.33 < Asmin Then
    Asfinal(i) = 1.33 * Ass(i)
    End If
    
    
    If Tabla.TextMatrix(10, i) > 0 Then
    Ass(i) = 0
    End If
    
    
    If Ass(i) = 0 Then Asfinal(i) = Asmin
    'Fin de aceros finales posibles.
    
    
    Tabla3.TextMatrix(5, i) = "Simple"
    Tabla3.TextMatrix(10, i) = Round(Asfinal(i), 2)
    Tabla3.TextMatrix(1, i) = Round(Ass(i), 2)
    
    
    a = (Ass(i) * Fy) / (0.85 * Fc * b)
    Tabla3.TextMatrix(3, i) = Round(a, 2)
    Tabla3.TextMatrix(4, i) = Round(Ass(i) * 1.33, 2)
    
    
    p(i) = Ass(i) / (b * d)
    Tabla3.TextMatrix(2, i) = Round(p(i) / pb, 2)
    
    
    Next i
    'Fin de calcular el acero requerido a tension arriba.
    
    
    'Calcula el acero requerido a tension abajo.
    For j = 1 To 7
    Mu(j + 7) = Abs(Tabla.TextMatrix(11, j))
    Ass(j + 7) = ((0.9 * Fy * d) - (((-0.9 * Fy * d) ^ 2) - (4 * ((Fy ^ 2 * 0.9) / (1.7 * Fc * b)) * Abs(Mu(j + 7)) * 100000)) ^ 0.5) / (2 * ((Fy ^ 2 * 0.9) / (1.7 * Fc * b)))
    
    
    'Aceros finales posibles.
    Asfinal(j + 7) = Ass(j + 7)
    If Ass(j + 7) < Asmin Then Asfinal(j + 7) = Asmin
    
    If Ass(j + 7) * 1.33 < Asmin Then
    Asfinal(j + 7) = 1.33 * Ass(j + 7)
    End If
    
    
    If Tabla.TextMatrix(11, j) < 0 Then
    Ass(j + 7) = 0
    End If
    
    
    If Ass(j + 7) = 0 Then Asfinal(j + 7) = Asmin
    'Fin de aceros finales posibles.
    
    
    Tabla3.TextMatrix(16, j) = "Simple"
    Tabla3.TextMatrix(21, j) = Round(Asfinal(j + 7), 2)
    Tabla3.TextMatrix(12, j) = Round(Ass(j + 7), 2)
    
    
    a = (Ass(j + 7) * Fy) / (0.85 * Fc * b)
    Tabla3.TextMatrix(14, j) = Round(a, 2)
    Tabla3.TextMatrix(15, j) = Round(Ass(j + 7) * 1.33, 2)
    
    
    p(j + 7) = Ass(j + 7) / (b * d)
    Tabla3.TextMatrix(13, j) = Round(p(j + 7) / pb, 2)
    
    
    Next j
    'Fin de calcular el acero requerido a tension abajo.
    
    
    'Calcula el acero requerido a compresion arriba (Doble refuerzo).
    For i = 1 To 7
    
    
    If Ass(i) > Asmax Then
    Tabla3.TextMatrix(5, i) = "Doble"
    amax = (Asmax * Fy) / (0.85 * Fc * b)
    c = amax / B1
    Esp = (0.003 * (c - dp)) / c
    Es = (0.003 * (d - c)) / c
    
    
    Fs = Es * 2038901.781
    Fsp = Esp * 2038901.781
    
    
    'Usar como Máximo Fsp=Fy.
    If Fsp > Val(Fy) Then
    Fsp = Val(Fy)
    Else: End If
    Tabla3.TextMatrix(9, i) = Round(Fsp, 1)
    'Fin de Máximo de Fsp.
    
    
    Ey = Fy / 2038901.781
    
    
    Mn = Asmax * Fy * (d - (amax / 2))
    M2 = Mn / 100000
    
    
    M1(i) = (Abs(Tabla.TextMatrix(10, i)) / (0.9)) - M2
    Assc(i) = M1(i) * 100000 / (Fy * (d - dp))
    
    
    'Si NO FLUYE.
    If Fsp < Val(Fy) Then
    
    AA = 0.85 * Fc * B1 * b
    BB = 0.003 * 2038901.781 * Assc(i) - ((Asmax + Assc(i)) * Fy)
    CC = -dp * 0.003 * 2038901.781 * Assc(i)
    
    
    c = (-BB + ((BB ^ 2) - (4 * AA * CC)) ^ (0.5)) / (2 * AA)
    
    
    Fsp = 0.003 * 2038901.781 * ((c - dp) / (c))
    Tabla3.TextMatrix(9, i) = Round(Fsp, 1)
    
    
    Assc(i) = Assc(i) * (Fy / Fsp)
    pp(i) = (Assc(i) * (Fy / Fsp)) / (b * d)
    Assn(i) = (Assc(i) * (Fsp / Fy)) + Asmax
    
    
    Else:
    
    
    pp(i) = Assc(i) / (b * d)
    Assn(i) = Asmax + Assc(i)
    End If
    
    
    Tabla3.TextMatrix(8, i) = Round(c, 2)
    Tabla3.TextMatrix(6, i) = Round(Assc(i), 2)
    Tabla3.TextMatrix(7, i) = Round(Assn(i), 2)
    Asfinal(i) = Assn(i)
    Tabla3.TextMatrix(10, i) = Round(Asfinal(i), 2)
    
    
    Else: End If
    
    
    Next i
    'Calcula el acero requerido a compresion abajo (Doble refuerzo).
    
    
    For j = 1 To 7
    
    If Ass(j + 7) > Asmax Then
    Tabla3.TextMatrix(16, j) = "Doble"
    amax = (Asmax * Fy) / (0.85 * Fc * b)
    c = amax / B1
    Esp = (0.003 * (c - dp)) / c
    Es = (0.003 * (d - c)) / c
    
    
    Fs = Es * 2038901.781
    Fsp = Esp * 2038901.781
    
    
    'Usar como Máximo Fsp=Fy.
    If Fsp > Val(Fy) Then
    Fsp = Val(Fy)
    Else: End If
    Tabla3.TextMatrix(20, j) = Round(Fsp, 1)
    'Fin de usar Máximo de Fsp.
    
    
    Ey = Fy / 2038901.781
    
    
    Mn = Asmax * Fy * (d - (amax / 2))
    M2 = Mn / 100000
    
    
    M1(j + 7) = (Abs(Tabla.TextMatrix(11, j)) / (0.9)) - M2
    Assc(j + 7) = M1(j + 7) * 100000 / (Fy * (d - dp))
    
    
    'Si NO FLUYE.
    If Fsp < Val(Fy) Then
    
    
    AA = 0.85 * Fc * B1 * b
    BB = 0.003 * 2038901.781 * Assc(j + 7) - ((Asmax + Assc(j + 7)) * Fy)
    CC = -dp * 0.003 * 2038901.781 * Assc(j + 7)
    
    
    c = (-BB + ((BB ^ 2) - (4 * AA * CC)) ^ (0.5)) / (2 * AA)
    
    
    Fsp = 0.003 * 2038901.781 * ((c - dp) / (c))
    Tabla3.TextMatrix(20, j) = Round(Fsp, 1)
    
    
    Assc(j + 7) = Assc(j + 7) * (Fy / Fsp)
    pp(j + 7) = (Assc(j + 7) * (Fy / Fsp)) / (b * d)
    Assn(j + 7) = (Assc(j + 7) * (Fsp / Fy)) + Asmax
    Else:
    pp(j + 7) = Assc(j + 7) / (b * d)
    Assn(j + 7) = Asmax + Assc(j + 7)
    End If
    
    
    Tabla3.TextMatrix(19, j) = Round(c, 2)
    Tabla3.TextMatrix(17, j) = Round(Assc(j + 7), 2)
    Tabla3.TextMatrix(18, j) = Round(Assn(j + 7), 2)
    Asfinal(j + 7) = Assn(j + 7)
    Tabla3.TextMatrix(21, j) = Round(Asfinal(j + 7), 2)
    Else: End If
    Next j
    
    
    'Pone el valor respectivo de acero a compresion!
    For i = 1 To 7
    If Ass(i) > Asmax Then
    If Assc(i) > Asfinal(i + 7) Then
    If Assc(i) > Asmin Then
    Tabla3.TextMatrix(21, i) = Round(Assc(i), 2)
    Else
    Tabla3.TextMatrix(21, i) = Round(Asmin, 2)
    End If
    Else: End If
    Else: End If
    Next i
    
    
    For j = 1 To 7
    If Ass(j + 7) > Asmax Then
    If Assc(j + 7) > Asfinal(j) Then
    If Assc(j + 7) > Asmin Then
    Tabla3.TextMatrix(10, j) = Round(Assc(j + 7), 2)
    Else
    Tabla3.TextMatrix(10, j) = Round(Asmin, 2)
    End If
    Else: End If
    Else: End If
    Next j
    'Fin de poner el valor respectivo!!
    
    
    'Diseño por cortante'
    Patas = TPatas
    vt = Tvt
    vl = Tvl
    Av = ((Patas * (((vt * 2.54) / 8)) ^ 2 * 3.14159265359) / 4)
    
    
    'Mínimo de los máximos.
    smax(1) = d / 4
    smax(2) = 8 * vl
    smax(3) = 24 * vt
    smax(4) = 30
    smax(5) = smax(1)
    If smax(2) < smax(5) Then smax(5) = smax(2)
    If smax(3) < smax(5) Then smax(5) = smax(3)
    If smax(4) < smax(5) Then smax(5) = smax(4)
    TSmax2 = smax(5)
    TSmax3 = smax(5)
    
    
    TSmax = 30
    
    'Fin de mínimo de los máximos.
    
    
    'Cálculo del Vs y de s.
    For i = 1 To 7
    Tabla3.TextMatrix(23, i) = Abs(Val(Tabla1.TextMatrix(10, i)))
    Next i
    
    
    For i = 1 To 7
    Vs(i) = (Val(Tabla3.TextMatrix(23, i)) - (phic * Vc)) / phic
    Tabla3.TextMatrix(24, i) = Round(Vs(i), 2)
    If Vs(i) < 0 Then Tabla3.TextMatrix(24, i) = 0
    s(i) = (Av * Fyh * d) / (Vs(i) * 1000)
    If s(i) > smax(5) Then s(i) = TSmax
    If s(i) > 0 Then
    Tabla3.TextMatrix(26, i) = Round(s(i), 2)
    Else
    Tabla3.TextMatrix(26, i) = TSmax
    End If
    If Vs(i) > Vsmax Then
    MsgBox "Aumentar el tamaño del elemento y actualizar datos. (Vs > Vsmax)"
    Else
    End If
    Next i
    'Fin del cálculo del Vs y de s.
    
    
    'Fin de diseño a cortante
    
    
    'Gráficos de momentos.
    Dim X(1 To 3, 0 To 7) As Variant
    
    
    'Pone los títulos en X.
    X(1, 1) = Tabla.TextMatrix(0, 1)
    X(1, 2) = Tabla.TextMatrix(0, 2)
    X(1, 3) = Tabla.TextMatrix(0, 3)
    X(1, 4) = Tabla.TextMatrix(0, 4)
    X(1, 5) = Tabla.TextMatrix(0, 5)
    X(1, 6) = Tabla.TextMatrix(0, 6)
    X(1, 7) = Tabla.TextMatrix(0, 7)
    
    
    'Leyenda.
    X(2, 0) = "Mínimo"
    X(3, 0) = "Máximo"
    
    
    'Dibuja datos.
    For i = 2 To 3
    For j = 1 To 7
    X(i, j) = -Val(Tabla.TextMatrix(i - 1 + 9, j))
    Next j
    Next i
    MSChartFlexion.ChartData = X
    'Fin gráficos de momentos.
    
    
    'Gráficos de Cortantes.
    Dim xx(1 To 3, 0 To 7) As Variant
    
    
    xx(1, 1) = Tabla1.TextMatrix(0, 1)
    xx(1, 2) = Tabla1.TextMatrix(0, 2)
    xx(1, 3) = Tabla1.TextMatrix(0, 3)
    xx(1, 4) = Tabla1.TextMatrix(0, 4)
    xx(1, 5) = Tabla1.TextMatrix(0, 5)
    xx(1, 6) = Tabla1.TextMatrix(0, 6)
    xx(1, 7) = Tabla1.TextMatrix(0, 7)
    
    
    xx(2, 0) = "Mínimos"
    xx(3, 0) = "Máximos"
    
    
    For j = 1 To 7
    xx(2, j) = -Val(Tabla1.TextMatrix(10, j))
    Next j
    
    
    MSChartCortante.ChartData = xx
    'Fin gráficos de Cortantes.
    
    
    'Pasa los resultados finales de acero.
    For i = 0 To 6
    Tas(i) = Round(Tabla3.TextMatrix(10, i + 1), 2)
    Tas(i + 7) = Round(Tabla3.TextMatrix(21, i + 1), 2)
    Tss(i) = Round(Tabla3.TextMatrix(26, i + 1), 2)
    Next i
    'Pasa los resultados finales de acero.
        
    
    'Diseño por capacidad.
    TablaCapa2.TextMatrix(1, 1) = Round((1.25 * Val(TAsrai) * TFy2) / (0.85 * TFc2 * Tb2), 2)
    aai = Val(TablaCapa2.TextMatrix(1, 1))
    
    
    TablaCapa2.TextMatrix(2, 1) = Round((1.25 * Val(TAsrabi) * TFy2) / (0.85 * TFc2 * Tb2), 2)
    aabi = Val(TablaCapa2.TextMatrix(2, 1))
    
    
    TablaCapa2.TextMatrix(1, 4) = Round((1.25 * Val(TAsrad) * TFy2) / (0.85 * TFc2 * Tb2), 2)
    aad = Val(TablaCapa2.TextMatrix(1, 4))
    
    
    TablaCapa2.TextMatrix(2, 4) = Round((1.25 * Val(TAsrabd) * TFy2) / (0.85 * TFc2 * Tb2), 2)
    aabd = Val(TablaCapa2.TextMatrix(2, 4))
    
    
    TablaCapa2.TextMatrix(3, 1) = Round((1.25 * Val(TAsrai) * TFy2) * (Td2 - (aai / 2)) / (100000), 2)
    Mnai = Val(TablaCapa2.TextMatrix(3, 1))
    
    
    TablaCapa2.TextMatrix(4, 1) = Round((1.25 * Val(TAsrabi) * TFy2) * (Td2 - (aabi / 2)) / (100000), 2)
    Mnabi = Val(TablaCapa2.TextMatrix(4, 1))
    
    
    TablaCapa2.TextMatrix(3, 4) = Round((1.25 * Val(TAsrad) * TFy2) * (Td2 - (aad / 2)) / (100000), 2)
    Mnad = Val(TablaCapa2.TextMatrix(3, 4))
    
    
    TablaCapa2.TextMatrix(4, 4) = Round((1.25 * Val(TAsrabd) * TFy2) * (Td2 - (aabd / 2)) / (100000), 2)
    Mnabd = Val(TablaCapa2.TextMatrix(4, 4))
    
    
    Ra = ((1.4 * (TWcm + (Tb2 * Th2 * 0.0001 * 2400)) + 1.7 * TWcv) * (TLong2) * 0.5) + ((1.4 * TPcm1 + 1.7 * TPcv1) * (TLong2 - Tdist1) / TLong2) + ((1.4 * TPcm2 + 1.7 * TPcv2) * (TLong2 - Tdist2) / TLong2)
    Rb = ((1.4 * (TWcm + (Tb2 * Th2 * 0.0001 * 2400)) + 1.7 * TWcv) * (TLong2) * 0.5) + ((1.4 * TPcm1 + 1.7 * TPcv1) * (Tdist1) / TLong2) + ((1.4 * TPcm2 + 1.7 * TPcv2) * (Tdist2) / TLong2)
    
    
    If Ra > Rb Then
    Rmax = Ra
    Else
    Rmax = Rb
    End If
    
    
    TR1 = Ra
    TR2 = Rb
    TRmax = Rmax
    
    
    Vui = (((Mnai + Mnabd) / ((TLong2)) + (0.75 * (Rmax / 1000))))
    Vud = (((Mnabi + Mnad) / ((TLong2)) + (0.75 * (Rmax / 1000))))
    Vci = 0.53 * ((TFc2) ^ 0.5) * Tb2 * Td2 * (1 / 1000)
    Vcd = 0.53 * ((TFc2) ^ 0.5) * Tb2 * Td2 * (1 / 1000)
    TablaCapa2.TextMatrix(5, 1) = Round(Vui, 2)
    TablaCapa2.TextMatrix(5, 4) = Round(Vud, 2)
    TablaCapa2.TextMatrix(6, 1) = Round(Vui * 0.5, 2)
    TablaCapa2.TextMatrix(6, 4) = Round(Vud * 0.5, 2)
    TablaCapa2.TextMatrix(7, 1) = Round((Mnai + Mnabd) / TLong2, 2)
    TablaCapa2.TextMatrix(7, 4) = Round((Mnabi + Mnad) / TLong2, 2)
    Vlefti = TablaCapa2.TextMatrix(7, 1)
    Vleftd = TablaCapa2.TextMatrix(7, 4)
    
    
    If Val(Vlefti) > (Vui * 0.5) Then
    Vci = 0
    End If
    If Val(Vleftd) > (Vud * 0.5) Then
    Vcd = 0
    End If
    
    
    TablaCapa2.TextMatrix(8, 1) = Round(Vci, 2)
    TablaCapa2.TextMatrix(8, 4) = Round(Vcd, 2)
    TablaCapa2.TextMatrix(9, 1) = Round((Vui - phic * Vci) / phic, 2)
    TablaCapa2.TextMatrix(9, 4) = Round((Vud - phic * Vcd) / phic, 2)
    
    
    Scap3i = TPatas2 * 0.71 * TFyh2 * Td2 / (TablaCapa2.TextMatrix(9, 1) * 1000)
    Scap4i = TPatas2 * 1.27 * TFyh2 * Td2 / (TablaCapa2.TextMatrix(9, 1) * 1000)
    Scap5i = TPatas2 * 1.98 * TFyh2 * Td2 / (TablaCapa2.TextMatrix(9, 1) * 1000)
    Scap3d = TPatas2 * 0.71 * TFyh2 * Td2 / (TablaCapa2.TextMatrix(9, 4) * 1000)
    Scap4d = TPatas2 * 1.27 * TFyh2 * Td2 / (TablaCapa2.TextMatrix(9, 4) * 1000)
    Scap5d = TPatas2 * 1.98 * TFyh2 * Td2 / (TablaCapa2.TextMatrix(9, 4) * 1000)
    
    
    TablaCapa1.TextMatrix(0, 0) = Round(Scap3i, 2)
    TablaCapa1.TextMatrix(1, 0) = Round(Scap4i, 2)
    TablaCapa1.TextMatrix(2, 0) = Round(Scap5i, 2)
    TablaCapa1.TextMatrix(0, 2) = Round(Scap3d, 2)
    TablaCapa1.TextMatrix(1, 2) = Round(Scap4d, 2)
    TablaCapa1.TextMatrix(2, 2) = Round(Scap5d, 2)
    'Fin de capacidad.
    
    
    'Pone los colores si es doble refuerzo.
    For i = 1 To 7
    If Tabla3.TextMatrix(5, i) = "Doble" Then
    Tas(i - 1).FontBold = True
    Tas(i - 1).ForeColor = vbRed
    
    
    Else
    
    
    Tas(i - 1).FontBold = False
    Tas(i - 1).ForeColor = vbBlack
    End If
    Next i
    
    
    For i = 1 To 7
    If Tabla3.TextMatrix(16, i) = "Doble" Then
    Tas(i + 6).FontBold = True
    Tas(i + 6).ForeColor = vbRed
    
    
    Else
    
    
    Tas(i + 6).FontBold = False
    Tas(i + 6).ForeColor = vbBlack
    End If
    Next i
    'Fin de poner los colores si es doble refuerzo.
    
    
    For i = 1 To 14
    If (p(i) - pp(i)) > 0.025 Then
    MsgBox "La seccion no puede soportar esa flexion! ( p-p' > 0.025 )"
    Exit Sub
    Else
    End If
    Next i
    
    
    End Sub
    
    
    Sub Importacion()
  
      
    'Para poner en blanco las casillas cada corrida!
    For i = 1 To 21
    For j = 1 To 7
    Tabla3.TextMatrix(i, j) = ""
    Next j
    Next i
    'Fin de casillas en blanco.
    
    
    'Pone la posicion del número de viga.
    For i = 1 To MSHFlexGrid1.Rows - 1
    If Datos(i, 0) = CViga Then
    Posicion = i
    i = MSHFlexGrid1.Rows - 1
    End If
    Next i
    'Fin de posicion.
    
    
    'Lee la longitud de la viga!
    TLong = Round(Datos(Posicion + 6, 5) * 1000, 2)
    'Fin de leer la longitud de la viga!
    
    
    'Para flexion.
    'Inserta los datos de CM,CV,CS en la tabla.
    
    
    'Carga Muerta.
    For w = 0 To 6
    Tabla.TextMatrix(1, w + 1) = Round(Datos(Posicion + w, 3), 3)
    Next w
    
    
    'Carga Viva.
    For w = 0 To 6
    Tabla.TextMatrix(2, w + 1) = Round(Datos(Posicion + w + 7, 3), 3)
    Next w
    
    
    'Sismo en X.
    For w = 0 To 6
    Tabla.TextMatrix(3, w + 1) = Round(Datos(Posicion + w + 14, 3), 3)
    Next w
    
    
    'Sismo en Y.
    For w = 0 To 6
    Tabla.TextMatrix(4, w + 1) = Round(Datos(Posicion + w + 21, 3), 3)
    Next w
    'Fin de insertar los datos.
    
    
    'Saca el valor mayor de los sismos en cada tramo.
    For i = 0 To 6
    If Abs(Datos(Posicion + i + 14, 3)) > Abs(Datos(Posicion + i + 21, 3)) Then
    SismoM(i + 1) = Abs(Datos(Posicion + i + 14, 3))
    Else:
    SismoM(i + 1) = Abs(Datos(Posicion + i + 21, 3))
    End If
    Next i
    'Fin de Sacar el valor mayor de los sismos en cada tramo.
    
    
    'Calcula y anota las combinaciones en la tabla.
    'COMBO1.
    For v = 0 To 6
    Tabla.TextMatrix(5, v + 1) = Round(1.4 * Datos(Posicion + v, 3) + 1.7 * Datos(Posicion + v + 7, 3), 3)
    
    
    'COMBO2.
    Tabla.TextMatrix(6, v + 1) = Round((0.75 * (1.4 * Datos(Posicion + v, 3) + 1.7 * Datos(Posicion + v + 7, 3)) + SismoM(v + 1)), 3)
    
    
    'COMBO3.
    Tabla.TextMatrix(7, v + 1) = Round((0.75 * (1.4 * Datos(Posicion + v, 3) + 1.7 * Datos(Posicion + v + 7, 3)) - SismoM(v + 1)), 3)
    
    
    'COMBO4.
    Tabla.TextMatrix(8, v + 1) = Round((0.95 * (Datos(Posicion + v, 3)) + SismoM(v + 1)), 3)
    
    
    'COMBO5.
    Tabla.TextMatrix(9, v + 1) = Round((0.95 * (Datos(Posicion + v, 3)) - SismoM(v + 1)), 3)
    Next v
    'Fin de calcular y anotar las combinaciones en la tabla.
    
    
    'Pasa las combinaciones a una matriz.
    For c = 1 To 7
    For r = 5 To 9
    CombosM(r - 4, c) = Val(Tabla.TextMatrix(r, c))
    Next r
    Next c
    'Fin de pasar las combinaciones a una matriz.
    
    
    'Saca los valoresa Máximos y Mínimos y los pone en la tabla.
    'Mínimos
    For v = 1 To 7
    ComboMin = CombosM(1, v)
    For i = 2 To 5
    If ComboMin > CombosM(i, v) Then
    ComboMin = CombosM(i, v)
    End If
    Tabla.TextMatrix(10, v) = Round(ComboMin, 3)
    Next i
    Next v
    
    
    'Máximos.
    For v = 1 To 7
    ComboMax = CombosM(1, v)
    For i = 2 To 5
    If ComboMax < CombosM(i, v) Then
    ComboMax = CombosM(i, v)
    End If
    Tabla.TextMatrix(11, v) = Round(ComboMax, 3)
    Next i
    Next v
    'Fin de sacar y poner los máximos y mínimos.
    
    
    'Para cortante:
    'Inserta los datos de CM,CV,CS en la tabla1.
    
    
    'Carga Muerta.
    For w = 0 To 6
    Tabla1.TextMatrix(1, w + 1) = Round(Datos(Posicion + w, 7), 3)
    Next w
    
    
    'Carga Viva.
    For w = 0 To 6
    Tabla1.TextMatrix(2, w + 1) = Round(Datos(Posicion + w + 7, 7), 3)
    Next w
    
    
    'Sismo en X.
    For w = 0 To 6
    Tabla1.TextMatrix(3, w + 1) = Round(Datos(Posicion + w + 14, 7), 3)
    Next w
    
    
    'Sismo en Y.
    For w = 0 To 6
    Tabla1.TextMatrix(4, w + 1) = Round(Datos(Posicion + w + 21, 7), 3)
    Next w
    'Fin de insertar los datos.
    
    
    'Saca el valor mayor de los sismos en cada tramo.
    For i = 0 To 6
    If Abs(Datos(Posicion + i + 14, 7)) > Abs(Datos(Posicion + i + 21, 7)) Then
    SismoC(i + 1) = Abs(Datos(Posicion + i + 14, 7))
    Else:
    SismoC(i + 1) = Abs(Datos(Posicion + i + 21, 7))
    End If
    Next i
    'Fin de Sacar el valor mayor de los sismos en cada tramo.
    
    
    'Calcula y anota las combinaciones en la tabla1.
    
    
    'COMBO1.
    For v = 0 To 6
    Tabla1.TextMatrix(5, v + 1) = Round(1.4 * Datos(Posicion + v, 7) + 1.7 * Datos(Posicion + v + 7, 7), 3)
    
    
    'COMBO2.
    Tabla1.TextMatrix(6, v + 1) = Round((0.75 * (1.4 * Datos(Posicion + v, 7) + 1.7 * Datos(Posicion + v + 7, 7)) + SismoC(v + 1)), 3)
    
    
    'COMBO3.
    Tabla1.TextMatrix(7, v + 1) = Round((0.75 * (1.4 * Datos(Posicion + v, 7) + 1.7 * Datos(Posicion + v + 7, 7)) - SismoC(v + 1)), 3)
    
    
    'COMBO4.
    Tabla1.TextMatrix(8, v + 1) = Round((0.95 * (Datos(Posicion + v, 7)) + SismoC(v + 1)), 3)
    
    
    'COMBO5.
    Tabla1.TextMatrix(9, v + 1) = Round((0.95 * (Datos(Posicion + v, 7)) - SismoC(v + 1)), 3)
    Next v
    'Fin de calcular y anotar las combinaciones en la tabla1.
    
    
    'Pasa las combinaciones a una matriz.
    For c = 1 To 7
    For r = 5 To 9
    CombosC(r - 4, c) = Val(Tabla1.TextMatrix(r, c))
    Next r
    Next c
    'Fin de pasar las combinaciones a una matriz.
    
      
    'Saca los valores Máximos y los pone en la tabla1.
    For v = 1 To 7
    ComboMax = CombosC(1, v)
    For i = 2 To 5
    If Abs(ComboMax) < Abs(CombosC(i, v)) Then
    ComboMax = CombosC(i, v)
    End If
    Tabla1.TextMatrix(10, v) = Round(ComboMax, 3)
    Next i
    Next v
    'Fin de sacar y poner los máximos y mínimos.
    
    
    End Sub
    
    
    Private Sub Cflexion_Click()
    
    
    CViga_Click
    
    
    End Sub
    
    
    Private Sub Imprimir_Click()
    
    
    PrintForm
    
    
    End Sub
    
    
    Private Sub MColumnas_Click()
    
    
    Unload FVigas
    
    
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
    
    
    Unload FVigas
    
    
    FDialog.Visible = True
        
        
    FDialog.Option3.Value = True
        
        
    End Sub
    
    
    Private Sub MSalida_Click()
    
    
    SSTabVigas.Visible = False
    
    
    MSHFlexGrid1.Visible = True
    
    
    MSHFlexGrid2.Visible = False
    
    
    End Sub
    
    
    Private Sub MAbout_Click()
    
    
    FAbout.Visible = True
    
    
    End Sub
    
    
    Private Sub MSalida2_Click()
    
    
    SSTabVigas.Visible = True
    
    
    MSHFlexGrid1.Visible = False
    
    
    MSHFlexGrid2.Visible = False
    
    
    End Sub
    
    
    Private Sub MVigas_Click()
    
    
    Unload FVigas
    
    
    FDialog.Visible = True
    
    
    FDialog.Option1.Value = True
    
    
    End Sub
    
    
    Private Sub OCortante_Click()
    
    
    Tabla.Visible = False
    
    
    Tabla1.Visible = True
    
    
    LCortantes.Visible = True
    
    
    LMomentos.Visible = False
    
    
    LTorsion.Visible = False
    
    
    End Sub
    
    
    Private Sub Oflexion_Click()
    
    
    Tabla1.Visible = False
    
    
    Tabla.Visible = True
    
    
    LCortantes.Visible = False
    
    
    LMomentos.Visible = True
    
    
    LTorsion.Visible = False
    
    
    End Sub
    
    
    Private Sub Otorsion_Click()
    
    
    Tabla.Visible = False
    
    
    Tabla1.Visible = False
    
    
    LCortantes.Visible = False
    
    
    LMomentos.Visible = False
    
    
    LTorsion.Visible = True
    
    
    End Sub
    

    Private Sub Picture4_Click()
    
    
    CViga_Click
    
    
    End Sub


    Private Sub Tb_Change()
    
    
    Tb2 = Tb
    
    
    End Sub
    
    
    Private Sub Td_Change()
    
    
    Td2 = Td
    
    
    End Sub
    
    
    Private Sub Tdp_Change()
    
    
    Tdp2 = Tdp
    
    
    End Sub


    Private Sub TFc_Change()
    
    
    TFc2 = TFc
    
    
    End Sub
    
    
    Private Sub TFy_Change()
    
    
    TFy2 = TFy
    
    
    End Sub
    
    
    Private Sub TFyh_Change()
    
    
    TFyh2 = TFyh
    
    
    End Sub
    
    
    Private Sub Th_Change()
    
    
    Th2 = Th
    
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!
    Td = Th - Thd
    
    
    End Sub
    
    
    Private Sub Thd_Change()
    
            
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!
    
    
    Tdp = Thd
    
    
    Tdp2 = Tdp
    
    
    Thd2 = Thd
        
    
    Td = Th - Thd
    
    
    End Sub
    
    
    Private Sub TPatas_Change()
    
    
    TPatas2 = TPatas
    
    
    End Sub
    
    
    Private Sub Tb2_Change()
    
    
    Tb = Tb2
    
    
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
    
    
    Td2 = Th2 - Thd2
    
    
    End Sub
    
    
    Private Sub Thd2_Change()
    
    
    'Error Handler para los errores de no poner nada!
    On Error GoTo ErrorNada
ErrorNada:
    Select Case Err.Number
    Case 13
    Resume Next
    End Select
    'Fin del errorhandler de no poner nada!
    
    
    Tdp2 = Thd2
    
    
    Thd = Thd2
    
    
    Td2 = Th2 - Thd2
    
    
    End Sub
    
    
    Private Sub Td2_Change()
    
    
    Td = Td2
    
    
    End Sub
    
    
    Private Sub Tdp2_Change()
    
    
    Tdp = Tdp2
    
    
    End Sub
    
    
    Private Sub TFc2_Change()
    
    
    TFc = TFc2
    
    
    End Sub
    
    
    Private Sub TFy2_Change()
    
    
    TFy = TFy2
    
    
    End Sub
    
    
    Private Sub TFyh2_Change()
       
       
    TFyh = TFyh2
    
    
    End Sub
    
    
    Private Sub TPatas2_Change()
    
    
    TPatas = TPatas2
    
    
    End Sub
    
    
    Private Sub Salida_Click()
 
    
    End
    
    
    End Sub
