Imports Sistema.Datos
Imports Sistema.Entidades

Public Class NVenta

    ' Calcula el subtotal de una línea: (precio - descuento) * cantidad
    Public Function CalcularSubtotal(Precio As Decimal, Descuento As Decimal, Cantidad As Integer) As Decimal
        Return (Precio - Descuento) * Cantidad
    End Function

    ' Calcula el total de la venta: suma de subtotales * (1 + impuesto)
    ' Bug #4 fix: impuesto ya es decimal (ej: 0.16), no dividir por 100
    Public Function CalcularTotal(Det As DataTable, Impuesto As Decimal) As Decimal
        Dim Subtotal As Decimal = 0
        For Each Fila As DataRow In Det.Rows
            Subtotal += Convert.ToDecimal(Fila("subtotal"))
        Next
        Return Subtotal * (1 + Impuesto)
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

    ' Consulta vw_VentasDetalladas por rango de fechas (detalle por línea de artículo)
    Public Function ConsultarDetallado(FechaInicio As Date, FechaFin As Date) As DataTable
        Try
            If FechaInicio > FechaFin Then
                Throw New Exception("La fecha de inicio no puede ser mayor a la fecha fin.")
            End If
            Dim Datos As New DVenta
            Return Datos.ConsultarDetallado(FechaInicio, FechaFin)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    ' Busca ventas activas por nombre de cliente, num documento o num comprobante
    Public Function Buscar(Valor As String) As DataTable
        Try
            Dim Datos As New DVenta
            Return Datos.Buscar(Valor)
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
    Public Function Insertar(Obj As Venta, Det As DataTable) As Boolean
        Try
            ' Validar cliente seleccionado
            If Obj.IdCliente <= 0 Then
                Throw New Exception("Debe seleccionar un cliente.")
            End If

            ' Validar que el detalle no esté vacío
            If Det Is Nothing OrElse Det.Rows.Count = 0 Then
                Throw New Exception("Debe agregar al menos un artículo a la venta.")
            End If

            ' La validación de stock la realiza TRG02 en la BD (trg_Venta_DescontarStock).
            ' Si el stock es insuficiente, el trigger hace ROLLBACK y lanza el error
            ' que sube por DVenta.InsertarConDetalle hasta este catch → MsgBox.

            Dim Datos As New DVenta
            Datos.InsertarConDetalle(Obj, Det)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    ' Actualiza datos de la cabecera
    Public Function Actualizar(Obj As Venta) As Boolean
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

    ' Reporte de ventas por período con clasificación y acumulado (usa SP con CURSOR)
    Public Function ReporteVentasPorPeriodo(FechaInicio As Date, FechaFin As Date) As DataSet
        Try
            If FechaInicio > FechaFin Then
                Throw New Exception("La fecha de inicio no puede ser mayor a la fecha fin.")
            End If
            Dim Datos As New DVenta
            Return Datos.ReporteVentasPorPeriodo(FechaInicio, FechaFin)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
End Class
