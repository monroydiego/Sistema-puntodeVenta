Imports System.Data.SqlClient
Imports Sistema.Entidades

Public Class DUsuario
    Inherits Conexion ' Hereda de la clase Conexion para poder usar sus propiedades y metodos

    ' Una tabla en memoria para almacenar los datos de la consulta
    Public Function Listar() As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("usuario_listar", MyBase.conn)
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

    Public Function Buscar(Valor As String) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("usuario_buscar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure ' indicamos que es un procedimiento almacenado 
            Comando.Parameters.Add("@Valor", SqlDbType.VarChar).Value = Valor ' Agregamos el parametro de busqueda del procedimiento almacenado
            MyBase.conn.Open() ' Abrimos la conexion
            Resultado = Comando.ExecuteReader() ' Ejecutamos el comando y almacenamos el resultado en Resultado
            Tabla.Load(Resultado) ' Cargamos el resultado en la tabla
            MyBase.conn.Close() ' Cerramos la conexion
            Return Tabla ' Retornamos la tabla con los datos
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function Login(Email As String, Clave As String) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("usuario_login", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure ' indicamos que es un procedimiento almacenado 
            Comando.Parameters.Add("@email", SqlDbType.VarChar).Value = Email ' Agregamos el parametro de busqueda del procedimiento almacenado
            Comando.Parameters.Add("@clave", SqlDbType.VarChar).Value = Clave ' Agregamos el parametro de busqueda del procedimiento almacenado
            MyBase.conn.Open() ' Abrimos la conexion
            Resultado = Comando.ExecuteReader() ' Ejecutamos el comando y almacenamos el resultado en Resultado
            Tabla.Load(Resultado) ' Cargamos el resultado en la tabla
            MyBase.conn.Close() ' Cerramos la conexion
            Return Tabla ' Retornamos la tabla con los datos
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ' Metodo para insertar una categoria
    Public Sub Insertar(Obj As Usuario)
        Try
            ' Obj hace una instancia en la clase Usuario
            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("usuario_insertar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure ' indicamos que es un procedimiento almacenado 
            Comando.Parameters.Add("@idrol", SqlDbType.Int).Value = Obj.IdRol
            Comando.Parameters.Add("@nombre", SqlDbType.VarChar).Value = Obj.Nombre
            Comando.Parameters.Add("@tipo_documento", SqlDbType.VarChar).Value = Obj.TipoDocumento
            Comando.Parameters.Add("@num_documento", SqlDbType.VarChar).Value = Obj.NumDocumento
            Comando.Parameters.Add("@direccion", SqlDbType.VarChar).Value = Obj.Direccion
            Comando.Parameters.Add("@telefono", SqlDbType.VarChar).Value = Obj.Telefono
            Comando.Parameters.Add("@email", SqlDbType.VarChar).Value = Obj.Email
            Comando.Parameters.Add("@clave", SqlDbType.VarChar).Value = Obj.Clave
            MyBase.conn.Open() ' Abrimos la conexion
            Comando.ExecuteNonQuery() ' Ejecutamos el comando y almacenamos el resultado en Resultado
            MyBase.conn.Close() ' Cerramos la conexion
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ' Metodo para eliminar un usuario 
    Public Sub Actualizar(Obj As Usuario)
        Try
            Dim Comando As New SqlCommand("usuario_actualizar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idrol", SqlDbType.Int).Value = Obj.IdRol
            Comando.Parameters.Add("@idusuario", SqlDbType.Int).Value = Obj.IdUsuario
            Comando.Parameters.Add("@nombre", SqlDbType.VarChar).Value = Obj.Nombre
            Comando.Parameters.Add("@tipo_documento", SqlDbType.VarChar).Value = Obj.TipoDocumento
            Comando.Parameters.Add("@num_documento", SqlDbType.VarChar).Value = Obj.NumDocumento
            Comando.Parameters.Add("@direccion", SqlDbType.VarChar).Value = Obj.Direccion
            Comando.Parameters.Add("@telefono", SqlDbType.VarChar).Value = Obj.Telefono
            Comando.Parameters.Add("@email", SqlDbType.VarChar).Value = Obj.Email
            Comando.Parameters.Add("@clave", SqlDbType.VarChar).Value = Obj.Clave
            MyBase.conn.Open() ' Abrimos la conexion
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Metodo para eliminar una categoria
    Public Sub Eliminar(Id As Integer)
        Try
            Dim Comando As New SqlCommand("usuario_eliminar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idusuario", SqlDbType.Int).Value = Id
            MyBase.conn.Open() ' Abrimos la conexion 
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Metodo para desactivar una categoria 

    Public Sub Desactivar(Id As Integer)
        Try
            Dim Comando As New SqlCommand("usuario_desactivar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idusuario", SqlDbType.Int).Value = Id ' se le envia el id de la categoria que se desea desactivar 
            MyBase.conn.Open() ' Abrimos la conexion 
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Metodo para activar una categoria 
    Public Sub Activar(Id As Integer)
        Try
            Dim Comando As New SqlCommand("usuario_activar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idusuario", SqlDbType.Int).Value = Id ' se envia id  de la categoria que se desea activar
            MyBase.conn.Open() ' Abrimos la conexion 
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
