Imports Sistema.Datos
Imports Sistema.Entidades

Public Class NVenta

    ' Calcula el subtotal de una línea: (precio - descuento) * cantidad
    Public Function CalcularSubtotal(Precio As Decimal, Descuento As Decimal, Cantidad As Integer) As Decimal
        Return (Precio - Descuento) * Cantidad
    End Function

    ' Calcula el total de la venta: suma de subtotales * (1 + impuesto/100)
    Public Function CalcularTotal(Det As DataTable, Impuesto As Decimal) As Decimal
        Dim Subtotal As Decimal = 0
        For Each Fila As DataRow In Det.Rows
            Subtotal += Convert.ToDecimal(Fila("subtotal"))
        Next
        Return Subtotal * (1 + Impuesto / 100)
    End Function

    ' Lista todas las ventas activas
    Public Function Listar() As DataTable
        Try
            Dim Datos As New DVenta
            Return Datos.Listar()
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    ' Busca ventas en un rango de fechas usando vw_VentasDetalladas
    Public Function BuscarPorFechas(FechaInicio As Date, FechaFin As Date) As DataTable
        Try
            If FechaInicio > FechaFin Then
                Throw New Exception("La fecha de inicio no puede ser mayor a la fecha fin.")
            End If
            Dim Datos As New DVenta
            Return Datos.BuscarPorFechas(FechaInicio, FechaFin)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    ' Retorna DataSet con 2 tablas: cabecera (0) y detalle (1)
    Public Function ObtenerConDetalle(IdVenta As Integer) As DataSet
        Try
            Dim Datos As New DVenta
            Return Datos.ObtenerConDetalle(IdVenta)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    ' Anula una venta (TRG03 restaura stock automáticamente)
    Public Function Anular(IdVenta As Integer) As Boolean
        Try
            Dim Datos As New DVenta
            Datos.Anular(IdVenta)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    ' Inserta cabecera + detalle con validaciones de negocio
    Public Function Insertar(Obj As EVenta, Det As DataTable) As Boolean
        Try
            ' Validar cliente seleccionado
            If Obj.IdCliente <= 0 Then
                Throw New Exception("Debe seleccionar un cliente.")
            End If

            ' Validar que el detalle no esté vacío
            If Det Is Nothing OrElse Det.Rows.Count = 0 Then
                Throw New Exception("Debe agregar al menos un artículo a la venta.")
            End If

            ' Validar stock suficiente para cada artículo ANTES de insertar
            Dim DatosArt As New DArticulo
            For Each Fila As DataRow In Det.Rows
                Dim IdArt As Integer = Convert.ToInt32(Fila("idArticulo"))
                Dim Cant As Integer = Convert.ToInt32(Fila("cantidad"))
                Dim TablaArt As DataTable = DatosArt.Listar()
                Dim Filas() As DataRow = TablaArt.Select("idArticulo = " & IdArt)
                If Filas.Length > 0 Then
                    Dim StockDisponible As Integer = Convert.ToInt32(Filas(0)("stock"))
                    If Cant > StockDisponible Then
                        Throw New Exception("Stock insuficiente para el artículo: " &
                                            Filas(0)("nombre").ToString() &
                                            ". Disponible: " & StockDisponible)
                    End If
                End If
            Next

            Dim Datos As New DVenta
            Datos.InsertarConDetalle(Obj, Det)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    ' Actualiza datos de la cabecera
    Public Function Actualizar(Obj As EVenta) As Boolean
        Try
            If Obj.IdVenta <= 0 Then
                Throw New Exception("ID de venta inválido.")
            End If
            Dim Datos As New DVenta
            Datos.Actualizar(Obj)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
End Class
