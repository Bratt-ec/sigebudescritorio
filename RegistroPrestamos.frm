VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REGISTRO DE PRESTAMOS"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   14370
   LinkTopic       =   "Form8"
   ScaleHeight     =   8220
   ScaleWidth      =   14370
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6975
      Left            =   8400
      TabIndex        =   0
      Top             =   1200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12303
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   6975
      Left            =   -120
      TabIndex        =   2
      Top             =   1200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12303
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro Prestamo de Libros"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   5040
   End
   Begin VB.Menu Regresar 
      Caption         =   "Regresar"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call formato
Call formato2
End Sub
Function formato()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 20
MSFlexGrid1.FormatString = "#|< CODIGO|<   Codigo Libro|<   Codigo Prestamo"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 2000
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
