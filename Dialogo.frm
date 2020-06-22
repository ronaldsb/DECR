VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diseño de elementos en concreto reforzado."
   ClientHeight    =   4224
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   8736
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4225.021
   ScaleMode       =   0  'User
   ScaleWidth      =   8724
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CProyecto 
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
      Left            =   5400
      TabIndex        =   5
      Top             =   1440
      Width           =   3012
   End
   Begin VB.ComboBox CHecho 
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
      Left            =   5400
      TabIndex        =   4
      Top             =   960
      Width           =   3012
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1332
      Left            =   720
      TabIndex        =   10
      Top             =   960
      Width           =   3132
      Begin VB.OptionButton Option2 
         Caption         =   "Diseño de columnas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Diseño de muros de corte."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3132
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Diseño de vigas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Access Database  (*.mdb) |*.mdb|"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUSCAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   3972
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Left            =   4440
      TabIndex        =   12
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "Diseñador:"
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
      Left            =   4320
      TabIndex        =   11
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "Escoja el archivo de salida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Escoja el tipo de diseño que desea realizar:"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3852
   End
End
Attribute VB_Name = "FDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      
      
    Private Sub Form_Load()
     
    
    'Cosas que el programa carga al inicio del progama.
    Dim dates As Date
    dates = #1/1/2003# 'Enero 1 del 2003.
    dias = 300 'Días permitidos antes del vencimiento.
    
    
    If Protegido = "S" Then
    
    
    On Error GoTo ErrorArchivo
ErrorArchivo:
    Select Case Err.Number
    Case 53
    Exit Sub
    End Select
    
    
    'Proteccción de copia. Borra los archivos necesarios.
    If Date >= dates + dias Then
    Dim FileSystemObject As Object
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    FileSystemObject.DeleteFile "C:\Windows\System\display.cpl"
    FileSystemObject.DeleteFile "C:\Windows\System\win.cpl"
    MsgBox ("El tiempo válido del demostrativo ha expirado!. Actualizar el programa a una nueva versión!"), vbCritical
    End
    Else
    End If
    
    
    If Date < dates - dias Then
    Dim FileSystemObject1 As Object
    Set FileSystemObject1 = CreateObject("Scripting.FileSystemObject")
    FileSystemObject1.DeleteFile "C:\Windows\System\display.cpl"
    FileSystemObject1.DeleteFile "C:\Windows\System\win.cpl"
    MsgBox ("La fecha de la computadora ha sido alterada! El progama quedará desactivado completamente"), vbCritical
    Else
    End If
    'Fin de la proteccción de copia. Borra los archivos necesarios.
    
    
    Else:
    
    
    'Lo que debe de hacer si no esta protegido.
    'If Date >= dates + dias Then
    'MsgBox ("El tiempo válido del demostrativo ha expirado!. Actualizar el programa a una nueva versión!"), vbCritical
    'End
    'Else
    'End If
    
    'If Date < dates - dias Then
    'MsgBox ("La fecha de la computadora ha sido alterada! El progama quedará desactivado completamente"), vbCritical
    'End
    'Else
    'End If
    'Fin de lo que debe de hacer si no esta protegido.
    
    
    End If
    
    CHecho.AddItem "Grupo #1."
    CHecho.AddItem "Grupo #2."
    CHecho.AddItem "Grupo #3."
    CHecho.AddItem "Grupo #4."
    CHecho.AddItem "Grupo #5."
    CHecho.AddItem "Grupo #6."
    
    
    CProyecto.AddItem "Tarea #1."
    CProyecto.AddItem "Tarea #2."
    CProyecto.AddItem "Tarea #3."
    CProyecto.AddItem "Tarea #4."
    CProyecto.AddItem "Tarea #5."
    CProyecto.AddItem "Tarea #6."
    CProyecto.AddItem "Trabajo Final."
    
    
    'Fin de las cosas que el programa carga al inicio del progama.
    
    
    End Sub
   

    Private Sub OKButton_Click()
    'Cosas que el programa carga al poner OK.
     
    
    Proyectos = CProyecto.Text
    Hechos = CHecho.Text
     
    
    If Protegido = "S" Then
    'Lo que debe de hacer si esta protegido.
    
    
    'Protección de copia. Clave de Instalación del programa.
    If TClave = "123" Then
    TClave = ""
    'Fin de la Protección de copia. La clave de instalación del programa.
    
    
    'Protección de copia. Copia los archivos de la seguridad.
    Dim FileSystemObject2 As Object
    Set FileSystemObject2 = CreateObject("Scripting.FileSystemObject")
    FileSystemObject2.CopyFile "C:\Windows\display.txt", "C:\Windows\System\display.cpl"
    FileSystemObject2.CopyFile "C:\Windows\win.ini", "C:\Windows\System\win.cpl"
    Else:
    End If
    'Fin de la Protección de copia. Copia los archivos de la seguridad.
     
    
    'Protección de copia. Borra todos los archivos de la seguridad.
    If TClave = "borrar" Then
    TClave = ""
    
    
    On Error GoTo ErrorArchivo
ErrorArchivo:
    Select Case Err.Number
    Case 53
    MsgBox ("El archivo deseado ya no existe!"), vbCritical
    Exit Sub
    End Select
    
    
    Dim FileSystemObject3 As Object
    Set FileSystemObject3 = CreateObject("Scripting.FileSystemObject")
    FileSystemObject3.DeleteFile "C:\Windows\System\display.cpl"
    FileSystemObject3.DeleteFile "C:\Windows\System\win.cpl"
    MsgBox ("El programa desactivado exitosamente!"), vbInformation
    Exit Sub
    End If
    'Fin de la protección de copia. Borra los archivos de la seguridad.
    
    
    'Protección de copia. Busca un archivo y si no esta, no funciona el programa.
    On Error GoTo ErrorArchivo2
ErrorArchivo2:
    Select Case Err.Number
    Case 53
    MsgBox ("Programa ha sido copiado o distribuido ilegalmente!"), vbCritical
    End
    End Select
    
    
    Open "C:\Windows\System\display.cpl" For Input As #1
    Close #1
    Open "C:\Windows\System\win.cpl" For Input As #1
    Close #1
    ' Fin de la Protección de copia. Busca el archivo.
    
    
    'Pender las formas
    If Option1.Value = True Then
    FVigas.Visible = True
    FDialog.Visible = False
    End If
    
    
    If Option2.Value = True Then
    FColumnas.Visible = True
    FDialog.Visible = False
    End If
    
    
    If Option3.Value = True Then
    FMuros.Visible = True
    FDialog.Visible = False
    End If
    'Prender las formas
    
    
    'Fin de lo que debe de hacer si está protegido.
    
    
    Else:
    
    
    'Lo que debe de hacer si no esta protegido.
    
    
    'Prender las formas.
    If Option1.Value = True Then
    FVigas.Visible = True
    FDialog.Visible = False
    End If
    
    
    If Option2.Value = True Then
    FColumnas.Visible = True
    FDialog.Visible = False
    End If
    
    
    If Option3.Value = True Then
    FMuros.Visible = True
    FDialog.Visible = False
    End If
    'Fin de prender las formas.
    
    
    End If
        
    
    'Fin de las cosas que el programa carga al poner OK.
       
       
    End Sub
    
    
    Private Sub CancelButton_Click()
    'Este procedimiento termina el programa.
    
    
    End
    
    
    'Fin del procedimiento que termina el programa.
   
   
    End Sub
  
    
    Private Sub Command1_Click()
    'Este procedimiento pide la ubicación del archivo y el nombre del directorio de trabajo.
    
    
    CommonDialog1.ShowOpen
    
    
    Archivo = CommonDialog1.FileName
    
    
    If Archivo = "" Then
    Else
    Command1.Caption = Archivo
    End If
    
    Hecho = CHecho.Text
   
    
    'Fin del procedimiento que pide la ubicación del archivo y el nombre del directorio de trabajo.
    
    
    End Sub
