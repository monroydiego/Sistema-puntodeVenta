Imports System.Data.SqlClient
Imports Sistema.Entidades

Public Class DIngreso
    Inherits Conexion
    Public Function Listar() As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable ' Tabla hace una instancia en la clase DataTable 

            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("ingreso_listar", MyBase.conn)
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
            Dim Comando As New SqlCommand("ingreso_buscar", MyBase.conn)
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
    Public Sub Anular(Id As Integer)
        Try
            Dim Comando As New SqlCommand("ingreso_anular", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idingreso", SqlDbType.Int).Value = Id ' se le envia el id de la categoria que se desea desactivar 
            MyBase.conn.Open() ' Abrimos la conexion 
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ' Consulta vw_ComprasDetalladas por rango de fechas (una fila por línea de detalle)
    Public Function ConsultarDetallado(FechaInicio As Date, FechaFin As Date) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable
            Dim Comando As New SqlCommand("sp_ConsultaComprasDetalladas", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@fechaInicio", SqlDbType.DateTime).Value = FechaInicio
            Comando.Parameters.Add("@fechaFin", SqlDbType.DateTime).Value = FechaFin
            MyBase.conn.Open()
            Resultado = Comando.ExecuteReader()
            Tabla.Load(Resultado)
            MyBase.conn.Close()
            Return Tabla
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub Insertar(Obj As Ingreso, Det As DataTable)
        Try
            ' Obj hace una instancia en la clase Ingreso
            ' Creamos un comando SQL para ejecutar el procedimiento almacenado
            ' El primer parametro hace referencia  al procedimiento almacenado de la base de datos
            ' El segundo parametro hace refeencia a la cadena de conexion para la base de datos
            ' Como heredamos la clase de la clase conexion, podemos usar MyBase para acceder a sus propiedades y metodos
            Dim Comando As New SqlCommand("ingreso_insertar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure ' indicamos que es un procedimiento almacenado 
            Comando.Parameters.Add("@idingreso", SqlDbType.Int).Value = Obj.IdIngreso ' Agrega el parámetro @Nombre al comando SQL, asignando el valor de la propiedad Nombre del objeto Categoria. Esto permite enviar el nombre de la nueva categoría al procedimiento almacenado en la base de datos.
            Comando.Parameters("@idingreso").Direction = ParameterDirection.Output
            Comando.Parameters.Add("@idusuario", SqlDbType.Int).Value = Obj.IdUsuario
            Comando.Parameters.Add("@idproveedor", SqlDbType.Int).Value = Obj.IdProveedor
            Comando.Parameters.Add("@tipo_comprobante", SqlDbType.VarChar).Value = Obj.TipoComprobante
            Comando.Parameters.Add("@serie_comprobante", SqlDbType.VarChar).Value = Obj.SerieComprobante
            Comando.Parameters.Add("@num_comprobante", SqlDbType.VarChar).Value = Obj.NumComprobante
            Comando.Parameters.Add("@impuesto", SqlDbType.Decimal).Value = Obj.Impuesto
            Comando.Parameters.Add("@total", SqlDbType.Decimal).Value = Obj.Total
            Comando.Parameters.Add("@detalle", SqlDbType.Structured).Value = Det
            MyBase.conn.Open() ' Abrimos la conexion
            Comando.ExecuteNonQuery() ' Ejecutamos el comando y almacenamos el resultado en Resultado
            MyBase.conn.Close() ' Cerramos la conexion
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
