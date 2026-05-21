Public Class FrmReporteVentasPeriodo
    Private Sub BtnGenerar_Click(sender As Object, e As EventArgs) Handles BtnGenerar.Click
        Try
            Dim FechaInicio As Date = DtpFechaInicio.Value.Date
            Dim FechaFin As Date = DtpFechaFin.Value.Date

            Dim Neg As New Negocio.NVenta
            Dim Ds As DataSet = Neg.ReporteVentasPorPeriodo(FechaInicio, FechaFin)

            If Ds Is Nothing OrElse Ds.Tables.Count < 2 Then
                MsgBox("No hay datos para el período seleccionado.", vbOKOnly + vbInformation, "Sin datos")
                Return
            End If

            ' Tabla 0: Detalle con clasificación
            DgvDetalle.DataSource = Ds.Tables(0)
            FormatoDetalle()

            ' Tabla 1: Resumen con totales por clasificación
            DgvResumen.DataSource = Ds.Tables(1)
            FormatoResumen()

            ' Calcular totales
            CalcularTotales(Ds.Tables(0))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub FormatoDetalle()
        DgvDetalle.AllowUserToAddRows = False
        DgvDetalle.ReadOnly = True
        For Each col As DataGridViewColumn In DgvDetalle.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Sub FormatoResumen()
        DgvResumen.AllowUserToAddRows = False
        DgvResumen.ReadOnly = True
        For Each col As DataGridViewColumn In DgvResumen.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Sub CalcularTotales(Tabla As DataTable)
        If Tabla Is Nothing OrElse Tabla.Rows.Count = 0 Then
            LblTotalVentas.Text = "Total Ventas: 0"
            LblTotalMontoAlta.Text = "Monto Alta: S/. 0.00"
            LblTotalMontoMedia.Text = "Monto Media: S/. 0.00"
            LblTotalMontoSBaja.Text = "Monto Baja: S/. 0.00"
            Return
        End If

        Dim TotalVentas As Integer = Tabla.Rows.Count
        Dim MontoAlta As Decimal = 0
        Dim MontoMedia As Decimal = 0
        Dim MontoBaja As Decimal = 0

        For Each Fila As DataRow In Tabla.Rows
            Dim Clasificacion As String = Fila("clasificacion").ToString()
            Dim Total As Decimal = CDec(Fila("total"))

            Select Case Clasificacion
                Case "Alta"
                    MontoAlta += Total
                Case "Media"
                    MontoMedia += Total
                Case "Baja"
                    MontoBaja += Total
            End Select
        Next

        LblTotalVentas.Text = "Total Ventas: " & TotalVentas
        LblTotalMontoAlta.Text = "Monto Alta: S/. " & MontoAlta.ToString("0.00")
        LblTotalMontoMedia.Text = "Monto Media: S/. " & MontoMedia.ToString("0.00")
        LblTotalMontoSBaja.Text = "Monto Baja: S/. " & MontoBaja.ToString("0.00")
    End Sub

    Private Sub FrmReporteVentasPeriodo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DtpFechaInicio.Value = DateTime.Now.AddMonths(-1)
        DtpFechaFin.Value = DateTime.Now
    End Sub
End Class
