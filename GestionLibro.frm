VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestión Libro"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   LinkTopic       =   "Form5"
   ScaleHeight     =   8970
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "REGRESAR"
      Height          =   855
      Left            =   7800
      Picture         =   "GestionLibro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text4 
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
      Left            =   1440
      TabIndex        =   17
      Top             =   2760
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7223
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
      Left            =   1440
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   2175
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
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Text5 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AÑADIR"
      Height          =   855
      Left            =   7800
      Picture         =   "GestionLibro.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GUARDAR"
      Height          =   855
      Left            =   7800
      Picture         =   "GestionLibro.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DESHACER"
      Height          =   855
      Left            =   7800
      Picture         =   "GestionLibro.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   855
      Left            =   7800
      Picture         =   "GestionLibro.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   2640
      Picture         =   "GestionLibro.frx":2BF2
      Top             =   0
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de Libro"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   240
      Width           =   2970
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
      Left            =   360
      TabIndex        =   13
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
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
      Left            =   480
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editorial"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autor"
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
      TabIndex        =   10
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edición"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   930
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomo"
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
      Left            =   480
      TabIndex        =   8
      Top             =   4080
      Width           =   720
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From LIBROS")
RS.MoveLast
cod = Val(RS!codigo_libro) + 1
For c = 0 To RS.RecordCount - 1
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = Text1.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Text2.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Text3.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Text4.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = Text6.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = Text5.Text
'RS.MoveNext
Next c
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
End Sub



Private Sub Command3_Click()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From LIBROS")
Set RS = conexion.Execute("Insert into LIBROS values ('" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) + "' )")
Call formato
Call datos
End Sub

Private Sub Command4_Click()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From LIBROS")
Set RS = conexion.Execute("delete from LIBROS where codigo_libro =  '" + Text1.Text + "'")
Call formato
Call datos
End Sub

Private Sub Command5_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
End Sub

Private Sub Command6_Click()
Unload Me
Form2.Show
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
MSFlexGrid1.ColWidth(6) = 800
End Function
'codigo_libro |       nombrelibro        |        autor        |      editorial       | tomo | edicion


Private Sub Form_Load()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From LIBROS")
RS.MoveLast
codigo = Val(RS!codigo_libro) + 1
Text1.Text = codigo
Call formato
Call datos
End Sub

Private Sub MSFlexGrid1_Click()
Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
Text3.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
Text4.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
Text6.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
Text5.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
End Sub
Function datos()
Dim RS As New ADODB.Recordset
Dim c%
Set RS = conexion.Execute("Select * From LIBROS")
RS.MoveFirst
For c = 0 To RS.RecordCount - 1
MSFlexGrid1.TextMatrix(c + 1, 1) = RS!codigo_libro
MSFlexGrid1.TextMatrix(c + 1, 2) = RS!nombrelibro
MSFlexGrid1.TextMatrix(c + 1, 3) = RS!autor
MSFlexGrid1.TextMatrix(c + 1, 4) = RS!Editorial
MSFlexGrid1.TextMatrix(c + 1, 5) = RS!Tomo
MSFlexGrid1.TextMatrix(c + 1, 6) = RS!Edicion
RS.MoveNext
Next c
End Function
