VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Busqueda Socio"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11385
   LinkTopic       =   "Form8"
   ScaleHeight     =   7425
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "REPORTE DE SOCIOS"
      Height          =   975
      Left            =   9600
      Picture         =   "BusquedaSocio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
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
      Left            =   3840
      TabIndex        =   14
      Top             =   840
      Width           =   2655
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
      Height          =   525
      Left            =   3840
      TabIndex        =   13
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text3 
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
      Left            =   3840
      TabIndex        =   12
      Top             =   2040
      Width           =   2655
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cedula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2160
      Picture         =   "BusquedaSocio.frx":08CA
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apellido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2160
      Picture         =   "BusquedaSocio.frx":1194
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2160
      Picture         =   "BusquedaSocio.frx":1A5E
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editorial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2160
      Picture         =   "BusquedaSocio.frx":2328
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUSCAR"
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
      Left            =   6600
      Picture         =   "BusquedaSocio.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DESHACER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Picture         =   "BusquedaSocio.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "REGRESAR"
      Height          =   975
      Left            =   8280
      Picture         =   "BusquedaSocio.frx":3D86
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BUSCAR"
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
      Left            =   6600
      Picture         =   "BusquedaSocio.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BUSCAR"
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
      Left            =   6600
      Picture         =   "BusquedaSocio.frx":4F1A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   360
      Picture         =   "BusquedaSocio.frx":57E4
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUSQUEDA DE SOCIOS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   3240
      TabIndex        =   15
      Top             =   0
      Width           =   4845
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c%, a%
Private Sub Command1_Click()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From SOCIOS where cedulasocio = '" + Text1.Text + "'")
If RS.RecordCount = 0 Then
MsgBox ("NO EXISTE")
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
Set RS = conexion.Execute("Select * From SOCIOS where Nombre = '" + Text2.Text + "'")
If RS.RecordCount = 0 Then
MsgBox (" NO EXISTE")
Else
MSFlexGrid1.Clear
Call formato
Call buscar2
End If
End Sub


Private Sub Command4_Click()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From SOCIOS where Apellido = '" + Text3.Text + "'")
If RS.RecordCount = 0 Then
MsgBox (" NO EXISTE")
Else
MSFlexGrid1.Clear
Call formato
Call buscar3
End If
End Sub


Private Sub Command5_Click()
DataReport1.Show
End Sub

Private Sub Command6_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Load()
Call formato2
Call datosocio
End Sub

Private Sub Option1_Click()
Option3.Visible = True
Option4.Visible = True
Option5.Visible = True
Option6.Visible = False
Option7.Visible = False
Option8.Visible = False
Call formato2
Call datosocio
End Sub

Private Sub Option2_Click()
Option6.Visible = True
Option7.Visible = True
Option8.Visible = True
Option3.Visible = False
Option4.Visible = False
Option5.Visible = False
Call formato
Call datoslibro
End Sub

Private Sub Option3_Click()
Text1.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Option4_Click()
Text2.Enabled = True
Text1.Enabled = False
Text3.Enabled = False
Command1.Enabled = False
Command3.Enabled = True
Command4.Enabled = False
End Sub

Private Sub Option5_Click()
Text3.Enabled = True
Text2.Enabled = False
Text1.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Option6_Click()
Text1.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Option7_Click()
Text2.Enabled = True
Text1.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Option8_Click()
Text3.Enabled = True
Text2.Enabled = False
Text1.Enabled = False
End Sub
Function formato()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 20
MSFlexGrid1.FormatString = "#|<Codigo Libro  |<    Nombre Libro|<     Autor |<   Editorial|<     Tomo|<   Edicion"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 800
MSFlexGrid1.ColWidth(2) = 2700
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(5) = 800

End Function
Function formato2()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 20
MSFlexGrid1.FormatString = "#|<COD |<CEDULA   |<    APELLIDO |<     NOMBRE |<   EDAD |<     TELEFONO |<   CUIDAD"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 600
MSFlexGrid1.ColWidth(2) = 1100
MSFlexGrid1.ColWidth(3) = 2500
MSFlexGrid1.ColWidth(4) = 2500
MSFlexGrid1.ColWidth(5) = 800
MSFlexGrid1.ColWidth(6) = 1500
MSFlexGrid1.ColWidth(7) = 1500
End Function
Function datosocio()
Dim RS As New ADODB.Recordset
Dim c%
Set RS = conexion.Execute("Select * From SOCIOS")
RS.MoveFirst
For c = 0 To RS.RecordCount - 1
MSFlexGrid1.TextMatrix(c + 1, 1) = RS!codigo_socio
MSFlexGrid1.TextMatrix(c + 1, 2) = RS!cedulasocio
MSFlexGrid1.TextMatrix(c + 1, 3) = RS!apellido
MSFlexGrid1.TextMatrix(c + 1, 4) = RS!nombre
MSFlexGrid1.TextMatrix(c + 1, 5) = RS!Edad
MSFlexGrid1.TextMatrix(c + 1, 6) = RS!telefono
MSFlexGrid1.TextMatrix(c + 1, 7) = RS!cuidad
RS.MoveNext
Next c
End Function

Function buscar()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("select * from SOCIOS")
c = 1
RS.MoveFirst
a = RS.RecordCount
While (c <> a)
If RS!cedulasocio = Text1.Text Then
c = RS.RecordCount
MSFlexGrid1.TextMatrix(1, 1) = RS!codigo_socio
MSFlexGrid1.TextMatrix(1, 2) = RS!cedulasocio
MSFlexGrid1.TextMatrix(1, 3) = RS!apellido
MSFlexGrid1.TextMatrix(1, 4) = RS!nombre
MSFlexGrid1.TextMatrix(1, 5) = RS!Edad
MSFlexGrid1.TextMatrix(1, 6) = RS!telefono
MSFlexGrid1.TextMatrix(1, 7) = RS!cuidad
Else
RS.MoveNext
c = c + 1
End If
Wend
End Function
Function buscar2()

Dim RS As New ADODB.Recordset
Set RS = conexion.Execute(" select * from SOCIOS")
c = 1
RS.MoveFirst
a = RS.RecordCount
While (c <> a)
If RS!nombre = Text2.Text Then
c = RS.RecordCount
MSFlexGrid1.TextMatrix(1, 1) = RS!codigo_socio
MSFlexGrid1.TextMatrix(1, 2) = RS!cedulasocio
MSFlexGrid1.TextMatrix(1, 3) = RS!apellido
MSFlexGrid1.TextMatrix(1, 4) = RS!nombre
MSFlexGrid1.TextMatrix(1, 5) = RS!Edad
MSFlexGrid1.TextMatrix(1, 6) = RS!telefono
MSFlexGrid1.TextMatrix(1, 7) = RS!cuidad
Else
RS.MoveNext
c = c + 1
End If
Wend
End Function
Function buscar3()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("select * from SOCIOS")
c = 1
RS.MoveFirst
a = RS.RecordCount
While (c <> a)
If RS!apellido = Text3.Text Then
c = RS.RecordCount
MSFlexGrid1.TextMatrix(1, 1) = RS!codigo_socio
MSFlexGrid1.TextMatrix(1, 2) = RS!cedulasocio
MSFlexGrid1.TextMatrix(1, 3) = RS!apellido
MSFlexGrid1.TextMatrix(1, 4) = RS!nombre
MSFlexGrid1.TextMatrix(1, 5) = RS!Edad
MSFlexGrid1.TextMatrix(1, 6) = RS!telefono
MSFlexGrid1.TextMatrix(1, 7) = RS!cuidad
Else
RS.MoveNext
c = c + 1
End If
Wend
End Function

