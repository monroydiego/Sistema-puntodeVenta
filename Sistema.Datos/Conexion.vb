Imports System.Data.SqlClient ' Se usa para poder conectarnos a SQL SERVER 

Public Class Conexion
#Region "Campos"
    Private _Base As String ' Variable para almacenar el nombre de la base de datos 
    Private _Servidor As String ' Variable para almacenar el nombre de nuestro servidor 
    Private _Usuario As String ' se almacena el usuario que se especifica desde la Base de datos previamente creada
    Private _Clave As String ' se almacena la clave con el usuario pueda acceder a la base de datos
    Private _Seguridad As Boolean = True
    Public conn As SqlConnection ' Variable para almacenar la conexion a la base de datps 
#End Region

#Region "Propiedades"
    Public Property Base As String
        Get
            Return _Base
        End Get
        Set(value As String)
            _Base = value
        End Set
    End Property

    Public Property Servidor As String
        Get
            Return _Servidor
        End Get
        Set(value As String)
            _Servidor = value
        End Set
    End Property

    Public Property Usuario As String
        Get
            Return _Usuario
        End Get
        Set(value As String)
            _Usuario = value
        End Set
    End Property

    Public Property Clave As String
        Get
            Return _Clave
        End Get
        Set(value As String)
            _Clave = value
        End Set
    End Property

    Public Property Seguridad As Boolean
        Get
            Return _Seguridad
        End Get
        Set(value As Boolean)
            _Seguridad = value
        End Set
    End Property

#End Region
    ' Creamos un constructor de la clase Conexion 
    Public Sub New()
        Me.Base = "dbsistema"
        Me.Servidor = "DESKTOP-CPRU6SB\SQLEXPRESS"
        Me.Usuario = "Monroy"
        Me.Clave = "Yeyo.monry0043"
        Me.conn = New SqlConnection(CrearCadena) ' Enviamos la cadena de conexion al objeto SqlConnection
    End Sub

    ' Declaramos una funcion 
    Public Function CrearCadena() As String
        Dim Cadena As String

        Cadena = "Server=" & Me.Servidor & ";Database=" & Me.Base & ";"
        ' Verificamos si la seguridad es de Windows o SQL 
        If Me.Seguridad Then
            Cadena &= "Integrated Security = SSPI" ' Si se trabaja con seguridad de Windows
        Else
            Cadena = Cadena & "User Id=" & Me.Usuario & ";Password=" & Me.Clave ' Si se trabaja con seguridad de SQL Server 
        End If
        Return Cadena
    End Function
End Class
