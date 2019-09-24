VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INICIO DE SESION"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      Picture         =   "InicioSesion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      Picture         =   "InicioSesion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1830
      Left            =   1320
      Picture         =   "InicioSesion.frx":1194
      Top             =   120
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m%, c%, ac%
Private Sub Command1_Click()
Dim bandera As Integer
Dim RS As New ADODB.Recordset
Set RS = conexion.Execute("select * from usuarios") 'conexion con la tabla Uusarios

If Text1.Text = "Secretaria" And Text2.Text = "secretaria" Then
Form2.Show
Unload Me
c = RS.RecordCount
Else
Form1.Show
Unload Me
m = MsgBox("CONTRASEÑA O USUARIO INCORRECTOS", vbCritical)
ac = ac + 1
    If ac = 5 Then
    End
    Else
    m = MsgBox("ESTE ES SU INTENTO: " & ac & " DE 5 INTENTOS DISPONIBLES, POR FAVOR INGRESE UN USUARIO Y UNA CONTRASEÑA VALIDOS", vbInformation)
        If ac = 4 Then
        m = MsgBox("SOLO LE QUEDA UN INTENTO", vbCritical)
        End If
    End If
End If

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub

