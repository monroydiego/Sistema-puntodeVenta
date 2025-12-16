Imports System.Data.SqlClient
Imports Sistema.Entidades

Public Class DPersona
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
            Dim Comando As New SqlCommand("persona_listar", MyBase.conn)
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
    Public Function ListarProveedores() As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("persona_listar_proveedores", MyBase.conn)
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
    Public Function ListarClientes() As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("persona_listar_clientes", MyBase.conn)
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
            Dim Comando As New SqlCommand("persona_buscar", MyBase.conn)
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
    Public Function BuscarProveedores(Valor As String) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("persona_buscar_proveedores", MyBase.conn)
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
    Public Function BuscarClientes(Valor As String) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("persona_buscar_clientes", MyBase.conn)
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


    ' Metodo para insertar una persona
    Public Sub Insertar(Obj As Persona)
        Try
            ' Obj hace una instancia en la clase Usuario
            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("persona_insertar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure ' indicamos que es un procedimiento almacenado 
            Comando.Parameters.Add("@tipo_persona", SqlDbType.VarChar).Value = Obj.TipoPersona
            Comando.Parameters.Add("@nombre", SqlDbType.VarChar).Value = Obj.Nombre
            Comando.Parameters.Add("@tipo_documento", SqlDbType.VarChar).Value = Obj.TipoDocumento
            Comando.Parameters.Add("@num_documento", SqlDbType.VarChar).Value = Obj.NumDocumento
            Comando.Parameters.Add("@direccion", SqlDbType.VarChar).Value = Obj.Direccion
            Comando.Parameters.Add("@telefono", SqlDbType.VarChar).Value = Obj.Telefono
            Comando.Parameters.Add("@email", SqlDbType.VarChar).Value = Obj.Email

            MyBase.conn.Open() ' Abrimos la conexion
            Comando.ExecuteNonQuery() ' Ejecutamos el comando y almacenamos el resultado en Resultado
            MyBase.conn.Close() ' Cerramos la conexion
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ' Metodo para Actualizar  un usuario 
    Public Sub Actualizar(Obj As Persona)
        Try
            Dim Comando As New SqlCommand("persona_actualizar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idpersona", SqlDbType.Int).Value = Obj.IdPersona
            Comando.Parameters.Add("@tipo_persona", SqlDbType.VarChar).Value = Obj.TipoPersona
            Comando.Parameters.Add("@nombre", SqlDbType.VarChar).Value = Obj.Nombre
            Comando.Parameters.Add("@tipo_documento", SqlDbType.VarChar).Value = Obj.TipoDocumento
            Comando.Parameters.Add("@num_documento", SqlDbType.VarChar).Value = Obj.NumDocumento
            Comando.Parameters.Add("@direccion", SqlDbType.VarChar).Value = Obj.Direccion
            Comando.Parameters.Add("@telefono", SqlDbType.VarChar).Value = Obj.Telefono
            Comando.Parameters.Add("@email", SqlDbType.VarChar).Value = Obj.Email

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
            Dim Comando As New SqlCommand("persona_eliminar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idpersona", SqlDbType.Int).Value = Id
            MyBase.conn.Open() ' Abrimos la conexion 
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
