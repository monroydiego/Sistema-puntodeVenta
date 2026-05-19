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

    ' NUEVO-BUG-01 FIX: reemplaza SqlDbType.Structured (TVP) por transaccion
    ' explicita fila a fila, identico al patron de DVenta.InsertarConDetalle.
    ' TRG01 (trg_Ingreso_ActualizarStock) se activa automaticamente al insertar
    ' cada linea en detalle_ingreso e incrementa el stock del articulo.
    ' Det debe tener columnas: idarticulo, cantidad, precio
    Public Sub Insertar(Obj As Ingreso, Det As DataTable)
        Dim Trx As SqlTransaction = Nothing
        Try
            MyBase.conn.Open()
            Trx = MyBase.conn.BeginTransaction()

            ' 1. Insertar cabecera — recuperar idIngreso generado via OUTPUT
            Dim CmdCab As New SqlCommand("sp_IngresoInsertar", MyBase.conn, Trx)
            CmdCab.CommandType = CommandType.StoredProcedure
            CmdCab.Parameters.Add("@idProveedor", SqlDbType.Int).Value = Obj.IdProveedor
            CmdCab.Parameters.Add("@idUsuario", SqlDbType.Int).Value = Obj.IdUsuario
            CmdCab.Parameters.Add("@tipo_comprobante", SqlDbType.VarChar).Value = Obj.TipoComprobante
            CmdCab.Parameters.Add("@serie_comprobante", SqlDbType.VarChar).Value = If(Obj.SerieComprobante IsNot Nothing, Obj.SerieComprobante, "")
            CmdCab.Parameters.Add("@num_comprobante", SqlDbType.VarChar).Value = Obj.NumComprobante
            CmdCab.Parameters.Add("@impuesto", SqlDbType.Decimal).Value = Obj.Impuesto
            CmdCab.Parameters.Add("@total", SqlDbType.Decimal).Value = Obj.Total
            Dim ParamId As New SqlParameter("@idIngreso", SqlDbType.Int)
            ParamId.Direction = ParameterDirection.Output
            CmdCab.Parameters.Add(ParamId)
            CmdCab.ExecuteNonQuery()
            Dim NuevoId As Integer = Convert.ToInt32(ParamId.Value)

            ' 2. Insertar cada linea de detalle (TRG01 incrementa stock por linea)
            For Each Fila As DataRow In Det.Rows
                Dim CmdDet As New SqlCommand("sp_DetalleIngresoInsertar", MyBase.conn, Trx)
                CmdDet.CommandType = CommandType.StoredProcedure
                CmdDet.Parameters.Add("@idIngreso", SqlDbType.Int).Value = NuevoId
                CmdDet.Parameters.Add("@idArticulo", SqlDbType.Int).Value = Convert.ToInt32(Fila("idarticulo"))
                CmdDet.Parameters.Add("@cantidad", SqlDbType.Int).Value = Convert.ToInt32(Fila("cantidad"))
                CmdDet.Parameters.Add("@precio", SqlDbType.Decimal).Value = Convert.ToDecimal(Fila("precio"))
                CmdDet.ExecuteNonQuery()
            Next

            Trx.Commit()
            MyBase.conn.Close()
        Catch ex As Exception
            If Trx IsNot Nothing Then Trx.Rollback()
            MyBase.conn.Close()
            Throw ex
        End Try
    End Sub
End Class
