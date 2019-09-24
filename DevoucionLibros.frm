VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Devoluciones"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12045
   LinkTopic       =   "Form6"
   ScaleHeight     =   7785
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   10080
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "REGRESAR"
      Height          =   975
      Left            =   10560
      Picture         =   "DevoucionLibros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CONFIRMAR DEVOLUCION"
      Height          =   975
      Left            =   10560
      Picture         =   "DevoucionLibros.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   960
      TabIndex        =   9
      Top             =   3360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos socio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   8535
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
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEDULA SOCIO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   2190
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUSCAR"
      Height          =   975
      Left            =   9120
      Picture         =   "DevoucionLibros.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DESHACER"
      Height          =   975
      Left            =   9120
      Picture         =   "DevoucionLibros.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   7680
      Picture         =   "DevoucionLibros.frx":2328
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestiòn de  Devoluciones"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   4395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   885
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m%
Private Sub Command1_Click()
Dim RS As New ADODB.Recordset
Dim c%, a%
Dim codigo%
Set RS = conexion.Execute("select gestionprestamo.codigo_prestamo,LIBROS.codigo_libro,LIBROS.nombrelibro,gestionprestamo.fechaprestamo,SOCIOS.cedulasocio,SOCIOS.nombre,SOCIOS.apellido from LIBROS,LibrosPrestados,GestionPrestamo,SOCIOS where LIBROS.codigo_libro = Librosprestados.codigo_libro and librosprestados.codigo_prestamo = GestionPrestamo.codigo_prestamo and gestionprestamo.codigosocio = SOCIOS.codigo_socio and cedulasocio =  '" + Text3.Text + "'")
If RS.RecordCount = 0 Then
MsgBox ("EL PRESTAMO QUE BUSCA NO EXISTE")
Else
MSFlexGrid1.Clear
Call formato
Call buscar
End If
End Sub

Private Sub Command2_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "

End Sub
Private Sub Command3_Click()
Dim RS As New ADODB.Recordset
Dim RSA As New ADODB.Recordset
Set RS = conexion.Execute("Select * From LIBROS")
Set RSA = conexion.Execute("SELECT codigo_socio from socios where cedulasocio =  '" + Text3.Text + "'")
codigo = RSA!codigo_socio
Set RS = conexion.Execute("delete From librosseleccionados where codigo_sc = '" + CStr(codigo) + "'")
Set RS = conexion.Execute("delete from librosprestados where codigo_prestamo =  '" + Text1.Text + "'")
Set RS = conexion.Execute("delete from gestionprestamo where codigo_prestamo =  '" + Text1.Text + "'")
m = MsgBox("¡DEVOLUCION DE LIBRO EXITOSA!")
Call formato
Call datos
End Sub

Private Sub Command6_Click()
Unload Me
Form2.Show
End Sub


Private Sub Form_Load()
Call formato
Call datos
End Sub

Private Sub MSFlexGrid1_Click()
Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
Text3.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
End Sub

Private Sub Option1_Click()
Text3.Enabled = True
Text4.Enabled = False
End Sub

Private Sub Option2_Click()
Text3.Enabled = False
Text4.Enabled = True
End Sub
Function formato()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 20
MSFlexGrid1.FormatString = "#|<CodigoPrestamo  |< Codigo Libro|<  Libro|<   Fecha Prestamo|<   Cedula|<   Nombre |< Apellido"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 800
MSFlexGrid1.ColWidth(2) = 800
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1300
MSFlexGrid1.ColWidth(6) = 2000
MSFlexGrid1.ColWidth(7) = 2000
End Function
 'codigo_prestamo | codigo_libro |     nombrelibro     | fechaprestamo | cedulasocio |     nombre
Function datos()
Dim RS As New ADODB.Recordset
Dim c%
Set RS = conexion.Execute("select gestionprestamo.codigo_prestamo,LIBROS.codigo_libro,LIBROS.nombrelibro,gestionprestamo.fechaprestamo,SOCIOS.cedulasocio,SOCIOS.nombre,SOCIOS.apellido from LIBROS,LibrosPrestados,GestionPrestamo,SOCIOS where LIBROS.codigo_libro = Librosprestados.codigo_libro and librosprestados.codigo_prestamo = GestionPrestamo.codigo_prestamo and gestionprestamo.codigosocio = SOCIOS.codigo_socio;")
RS.MoveFirst
For c = 0 To RS.RecordCount - 1
MSFlexGrid1.TextMatrix(c + 1, 1) = RS!codigo_prestamo
MSFlexGrid1.TextMatrix(c + 1, 2) = RS!codigo_libro
MSFlexGrid1.TextMatrix(c + 1, 3) = RS!nombrelibro
MSFlexGrid1.TextMatrix(c + 1, 4) = RS!fechaprestamo
MSFlexGrid1.TextMatrix(c + 1, 5) = RS!cedulasocio
MSFlexGrid1.TextMatrix(c + 1, 6) = RS!nombre
MSFlexGrid1.TextMatrix(c + 1, 7) = RS!apellido
RS.MoveNext
Next c
End Function
Function buscar()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("select gestionprestamo.codigo_prestamo,LIBROS.codigo_libro,LIBROS.nombrelibro,gestionprestamo.fechaprestamo,SOCIOS.cedulasocio,SOCIOS.nombre,SOCIOS.apellido from LIBROS,LibrosPrestados,GestionPrestamo,SOCIOS where LIBROS.codigo_libro = Librosprestados.codigo_libro and librosprestados.codigo_prestamo = GestionPrestamo.codigo_prestamo and gestionprestamo.codigosocio = SOCIOS.codigo_socio;")
c = 1
RS.MoveFirst
a = RS.RecordCount
While (c <> a)
If RS!cedulasocio = Text3.Text Then
c = RS.RecordCount
MSFlexGrid1.TextMatrix(1, 1) = RS!codigo_prestamo
MSFlexGrid1.TextMatrix(1, 2) = RS!codigo_libro
MSFlexGrid1.TextMatrix(1, 3) = RS!nombrelibro
MSFlexGrid1.TextMatrix(1, 4) = RS!fechaprestamo
MSFlexGrid1.TextMatrix(1, 5) = RS!cedulasocio
MSFlexGrid1.TextMatrix(1, 6) = RS!nombre
MSFlexGrid1.TextMatrix(1, 7) = RS!apellido
Else
RS.MoveNext
c = c + 1
End If
Wend
End Function
