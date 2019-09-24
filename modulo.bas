Attribute VB_Name = "Module1"
Public conexion As New ADODB.Connection
Sub main()
'declara variabla cadena de conexion
Dim cadena As String
cadena = "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=BIBLIOTECA;Initial Catalog=proyectobiblioteca"
conexion.Open cadena
conexion.CursorLocation = adUseClient
Form1.Show
End Sub


