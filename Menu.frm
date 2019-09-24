VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MENU"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11865
   LinkTopic       =   "Form2"
   ScaleHeight     =   4650
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "BUSQUEDA SOCIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9960
      Picture         =   "Menu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "BUSQUEDA LIBROS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8160
      Picture         =   "Menu.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CERRAR SESION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10680
      Picture         =   "Menu.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DEVOLUCIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      Picture         =   "Menu.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GESTION SOCIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      Picture         =   "Menu.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PRESTAMO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Picture         =   "Menu.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GESTION LIBROS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6240
      Picture         =   "Menu.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lbUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido, Francisco Tocto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   3675
      Width           =   3135
   End
   Begin VB.Label lbTipoUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Presidente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema Informático para la Gestión de la Biblioteca de la asociacion de obreros ""Unión y Disciplina"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Left            =   1440
      TabIndex        =   4
      Top             =   0
      Width           =   9525
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_Click()

End Sub

Private Sub Command1_Click()
Unload Me
Form6.Show
End Sub

Private Sub Command2_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command4_Click()
Unload Me
Form5.Show
End Sub

Private Sub Command5_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command6_Click()
Unload Me
Form7.Show
End Sub

Private Sub Command7_Click()
Form8.Show
Unload Me
End Sub

