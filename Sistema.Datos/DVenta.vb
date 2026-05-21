Imports System.Data.SqlClient
Imports Sistema.Entidades

Public Class DVenta
    Inherits Conexion

    ' Lista todas las ventas en estado Activo (cabecera con JOIN a Cliente y TipoComprobante)
    Public Function Listar() As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable
            Dim Comando As New SqlCommand("sp_VentaListar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            MyBase.conn.Open()
            Resultado = Comando.ExecuteReader()
            Tabla.Load(Resultado)
            MyBase.conn.Close()
            Return Tabla
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ' Busca ventas activas por nombre de cliente, num documento o num comprobante
    Public Function Buscar(Valor As String) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable
            Dim Comando As New SqlCommand("sp_VentaBuscar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@valor", SqlDbType.VarChar).Value = Valor
            MyBase.conn.Open()
            Resultado = Comando.ExecuteReader()
            Tabla.Load(Resultado)
            MyBase.conn.Close()
            Return Tabla
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ' Busca ventas por rango de fechas usando sp_VentaBuscarPorFechas
    Public Function BuscarPorFechas(FechaInicio As Date, FechaFin As Date) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable
            Dim Comando As New SqlCommand("sp_VentaBuscarPorFechas", MyBase.conn)
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

    ' Retorna un DataSet con 2 tablas: Tabla(0)=cabecera, Tabla(1)=detalle
    Public Function ObtenerConDetalle(IdVenta As Integer) As DataSet
        Try
            Dim Ds As New DataSet
            Dim Comando As New SqlCommand("sp_ObtenerVentaConDetalle", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idVenta", SqlDbType.Int).Value = IdVenta
            Dim Da As New SqlDataAdapter(Comando)
            MyBase.conn.Open()
            Da.Fill(Ds)
            MyBase.conn.Close()
            Return Ds
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ' Anula una venta (TRG03 restaura el stock automáticamente)
    Public Sub Anular(IdVenta As Integer)
        Try
            Dim Comando As New SqlCommand("sp_VentaAnular", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idVenta", SqlDbType.Int).Value = IdVenta
            MyBase.conn.Open()
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Actualiza cabecera de venta (sin detalles)
    Public Sub Actualizar(Obj As Venta)
        Try
            Dim Comando As New SqlCommand("sp_VentaActualizar", MyBase.conn)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.Add("@idVenta", SqlDbType.Int).Value = Obj.IdVenta
            Comando.Parameters.Add("@idCliente", SqlDbType.Int).Value = Obj.IdCliente
            Comando.Parameters.Add("@idTipoComprobante", SqlDbType.Int).Value = Obj.IdTipoComprobante
            Comando.Parameters.Add("@numComprobante", SqlDbType.VarChar).Value = Obj.NumComprobante
            Comando.Parameters.Add("@impuesto", SqlDbType.Decimal).Value = Obj.Impuesto
            Comando.Parameters.Add("@totalVenta", SqlDbType.Decimal).Value = Obj.TotalVenta
            MyBase.conn.Open()
            Comando.ExecuteNonQuery()
            MyBase.conn.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Consulta vw_VentasDetalladas por rango de fechas (una fila por línea de detalle)
    Public Function ConsultarDetallado(FechaInicio As Date, FechaFin As Date) As DataTable
        Try
            Dim Resultado As SqlDataReader
            Dim Tabla As New DataTable
            Dim Comando As New SqlCommand("sp_ConsultaVentasDetalladas", MyBase.conn)
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

    ' Inserta cabecera + detalle en una sola transacción explícita.
    ' Cada fila de Det requiere columnas: idArticulo, cantidad, precio, descuento, subtotal
    ' TRG02 descuenta stock al insertar cada línea de DetalleVenta.
    Public Sub InsertarConDetalle(Obj As Venta, Det As DataTable)
        Dim Trx As SqlTransaction = Nothing
        Try
            MyBase.conn.Open()
            Trx = MyBase.conn.BeginTransaction()

            ' 1. Insertar cabecera — recuperar idVenta generado via OUTPUT
            Dim CmdCab As New SqlCommand("sp_VentaInsertar", MyBase.conn, Trx)
            CmdCab.CommandType = CommandType.StoredProcedure
            CmdCab.Parameters.Add("@idCliente", SqlDbType.Int).Value = Obj.IdCliente
            CmdCab.Parameters.Add("@idUsuario", SqlDbType.Int).Value = Obj.IdUsuario
            CmdCab.Parameters.Add("@idTipoComprobante", SqlDbType.Int).Value = Obj.IdTipoComprobante
            CmdCab.Parameters.Add("@numComprobante", SqlDbType.VarChar).Value = Obj.NumComprobante
            CmdCab.Parameters.Add("@fechaVenta", SqlDbType.DateTime).Value = Obj.FechaVenta
            CmdCab.Parameters.Add("@impuesto", SqlDbType.Decimal).Value = Obj.Impuesto
            CmdCab.Parameters.Add("@totalVenta", SqlDbType.Decimal).Value = Obj.TotalVenta
            Dim ParamId As New SqlParameter("@idVenta", SqlDbType.Int)
            ParamId.Direction = ParameterDirection.Output
            CmdCab.Parameters.Add(ParamId)
            CmdCab.ExecuteNonQuery()
            Dim NuevoId As Integer = Convert.ToInt32(ParamId.Value)

            ' 2. Insertar cada línea de detalle (TRG02 descuenta stock por línea)
            ' Bug #2 fix: columna se llama "idarticulo" (minúsculas) en DtDetalle
            For Each Fila As DataRow In Det.Rows
                Dim CmdDet As New SqlCommand("sp_DetalleVentaInsertar", MyBase.conn, Trx)
                CmdDet.CommandType = CommandType.StoredProcedure
                CmdDet.Parameters.Add("@idVenta", SqlDbType.Int).Value = NuevoId
                CmdDet.Parameters.Add("@idArticulo", SqlDbType.Int).Value = Convert.ToInt32(Fila("idarticulo"))
                CmdDet.Parameters.Add("@cantidad", SqlDbType.Int).Value = Convert.ToInt32(Fila("cantidad"))
                CmdDet.Parameters.Add("@precio", SqlDbType.Decimal).Value = Convert.ToDecimal(Fila("precio"))
                CmdDet.Parameters.Add("@descuento", SqlDbType.Decimal).Value = Convert.ToDecimal(Fila("descuento"))
                CmdDet.Parameters.Add("@subtotal", SqlDbType.Decimal).Value = Convert.ToDecimal(Fila("subtotal"))
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
