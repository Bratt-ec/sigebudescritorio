VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestion Socio"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   LinkTopic       =   "Form4"
   ScaleHeight     =   7905
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "REGRESAR"
      Height          =   855
      Left            =   10200
      Picture         =   "GestionSocio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2160
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4575
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8070
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
      Left            =   6240
      TabIndex        =   15
      Top             =   1080
      Width           =   2535
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
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
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
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
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
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   2535
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
      Left            =   6240
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "GestionSocio.frx":08CA
      Left            =   6240
      List            =   "GestionSocio.frx":08D7
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Anadir 
      Caption         =   "AÑADIR"
      Height          =   855
      Left            =   8880
      Picture         =   "GestionSocio.frx":08F8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GUARDAR"
      Height          =   855
      Left            =   8880
      Picture         =   "GestionSocio.frx":11C2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DESHACER"
      Height          =   855
      Left            =   10200
      Picture         =   "GestionSocio.frx":1A8C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   855
      Left            =   10200
      Picture         =   "GestionSocio.frx":2356
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   0
      Picture         =   "GestionSocio.frx":2C20
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "EDAD"
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
      Left            =   5280
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de Socio"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   0
      Width           =   2910
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
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
      Left            =   5160
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod%
Private Sub Anadir_Click() 'este comando añade resgistros
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From SOCIOS")
RS.MoveLast
cod = Val(RS!codigo_socio) + 1
For c = 0 To RS.RecordCount - 1
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cod
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Text1.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Text3.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Text2.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = Text4.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = Text5.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Combo1.Text
'RS.MoveNext
Next c
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
End Sub

Private Sub Command3_Click()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From SOCIOS")
Set RS = conexion.Execute("Insert into SOCIOS values ('" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) + "','" + MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) + "' )")
Call formato
Call datos
End Sub

Private Sub Command4_Click()
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("Select * From SOCIOS")
Set RS = conexion.Execute("delete from SOCIOS where cedulasocio=  '" + Text1.Text + "'")
Call formato
Call datos
End Sub

Private Sub Command5_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
End Sub

Private Sub Command6_Click()
Unload Me
Form2.Show
End Sub

Function formato()
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

Private Sub Form_Load()
Call formato
Call datos

End Sub

Private Sub MSFlexGrid1_Click()
Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
Text3.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
Text4.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
Text5.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
Combo1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
'muestra los datos seleccionados en los text
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("select * from SOCIOS")
c = 1
RS.MoveFirst
a = RS.RecordCount
While (c <> a)
If RS!cedulasocio = Text1.Text Then
c = RS.RecordCount
modificar.Visible = True
Command3.Visible = False
Else
RS.MoveNext
c = c + 1
End If
Wend
End Sub
Function datos()
'esta funcion busca los datos
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
'codigo_socio | cedulasocio |    apellido     | nombre  | edad |  telefono  |   cuidad
