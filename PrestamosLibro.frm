VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PRESTAMOS"
   ClientHeight    =   9975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   LinkTopic       =   "Form3"
   ScaleHeight     =   9975
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   855
      Left            =   1200
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "AGREGAR"
      Height          =   855
      Left            =   8640
      Picture         =   "PrestamosLibro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "REGRESAR"
      Height          =   975
      Left            =   6240
      Picture         =   "PrestamosLibro.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   2280
      TabIndex        =   19
      Top             =   6600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos socio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   8655
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   24
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BUSCAR SOCIO"
         Height          =   975
         Left            =   7440
         Picture         =   "PrestamosLibro.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cedula"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Libro"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      TabIndex        =   2
      Top             =   4920
      Width           =   8655
      Begin VB.CommandButton Command5 
         Caption         =   "AGREGAR"
         Height          =   855
         Left            =   7200
         Picture         =   "PrestamosLibro.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Título"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton Guardar 
      Caption         =   "GUARDAR"
      Height          =   975
      Left            =   3600
      Picture         =   "PrestamosLibro.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DESHACER"
      Height          =   975
      Left            =   5040
      Picture         =   "PrestamosLibro.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1215
      Left            =   1080
      TabIndex        =   21
      Top             =   3600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2143
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   3000
      Picture         =   "PrestamosLibro.frx":34BC
      Top             =   0
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Préstamo de Libros"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   3450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cont%, i%, codigo%, m%
'el tercer boton de comando es solo un agregado para pruebas
'puede quitarlo si desea
Private Sub Guardar_Click()
Dim RS As New ADODB.Recordset 'crea ua variable recordset
Set RS = conexion.Execute("Select * From LibrosPrestados") 'extrae los datos de la tabla librosprestados en visual
                            'select from muestra la tabla en postgreSQL
Set RS = conexion.Execute("Insert into LibrosPrestados values ('" + Text8.Text + "','" + Text6.Text + "','" + Text1.Text + "' )")
'guarda los  datos de la tabla en la base de datos postgres
Set RS = conexion.Execute("Insert into GestionPrestamo values ('" + MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) + "','" + MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) + "','" + MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 3) + "','" + MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 4) + "' )")
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Set RS = conexion.Execute("Select * From GestionPrestamo")
RS.MoveLast 'mueve al sigiente registro
codigo = Val(RS!codigo_prestamo) + 1 'este y
Text1.Text = codigo ' este generan automaticamente el codigo del prestamo
Call formato
Call formato2
End Sub

Private Sub Command2_Click()
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Call formato
Call formato2
End Sub

Private Sub Command3_Click()
Dim RS As New ADODB.Recordset 'creamos una variable .recordset para crear la conexion con la BD
Set RS = conexion.Execute("Select * From SOCIOS where CedulaSocio = '" + Text3.Text + "'") 'hacer la busqueda con el select from
If RS.RecordCount = 0 Then
m = MsgBox("EL SOCIO NO EXISTE", vbInformation)
Else
Text4.Text = RS!nombre & "  " & RS!apellido
MSFlexGrid2.TextMatrix(1, 3) = RS!codigo_socio
End If
End Sub


Private Sub Command4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Command5_Click()

i = i + 1
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From LIBROS where nombrelibro = '" + Text5.Text + "'")
If RS.RecordCount = 0 Then
m = MsgBox("ESE LIBRO NO ESTA DISPONIBLE", vbInformation)
Else
Text6.Text = RS!codigo_libro
Set RS = conexion.Execute("Select * From librosprestados")
RS.MoveLast
cont = Val(RS!codlibroprestado) + 1 'es para generar el codigo automaticamente
Text8.Text = cont
MSFlexGrid1.TextMatrix(i, 1) = Text8.Text
MSFlexGrid1.TextMatrix(i, 2) = Text6.Text
MSFlexGrid1.TextMatrix(i, 3) = Text1.Text
End If
End Sub

Private Sub Command6_Click()
'select * from LIBROS inner join librosprestados on (LIBROS.codigo_libro = librosprestados.codigo_libro);
DataReport1.Show
'DataEnvironment1.Show
End Sub

Private Sub Command7_Click()

MSFlexGrid2.TextMatrix(1, 4) = Text7.Text
MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) = Text1.Text
MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = Text2.Text

End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From GestionPrestamo")
RS.MoveLast
codigo = Val(RS!codigo_prestamo) + 1
Text1.Text = codigo
Call formato
Call formato2
Text2.Text = Date
End Sub

Function formato()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 20
MSFlexGrid1.FormatString = "#|<CODIGO  |< Codigo Libro |< Codigo Prestamo"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.ColWidth(3) = 2000

End Function
 
Function formato2()
MSFlexGrid2.Clear
MSFlexGrid2.Rows = 20
MSFlexGrid2.FormatString = "#|<Codigo Prestamo|<   Fecha Prestamo|<    Codigo Socio|<   Observaciones"
MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(1) = 1500
MSFlexGrid2.ColWidth(2) = 2500
MSFlexGrid2.ColWidth(3) = 2000
MSFlexGrid2.ColWidth(4) = 2000
End Function

