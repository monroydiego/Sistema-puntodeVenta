Imports System.Data.SqlClient

Public Class DRol

    Inherits Conexion

    Public Function Listar() As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("rol_listar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure ' indicamos que es un procedimiento almacenado 
            MyBase.conn.Open() ' Abrimos la conexion
            Resultado = Comando.ExecuteReader() ' Ejecutamos el comando y almacenamos el resultado en Resultado
            Tabla.Load(Resultado) ' Cargamos el resultado en la tabla
            MyBase.conn.Close() ' Cerramos la conexion
            Return Tabla ' Retornamos la tabla con los datos
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
