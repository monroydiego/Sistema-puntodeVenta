Public Class FrmConsultaVentas

    ' vw_VentasDetalladas columnas (22):
    ' 0=idventa, 1=idcliente, 2=cliente, 3=doc_cliente, 4=tipoComprobante,
    ' 5=serieComprobante, 6=numComprobante, 7=fechaVenta, 8=impuesto,
    ' 9=totalVenta, 10=estado, 11=iddetalle_venta, 12=codigoArticulo,
    ' 13=nombreArticulo, 14=categoria, 15=cantidad, 16=precio,
    ' 17=descuento, 18=precioNeto, 19=subtotal, 20=totalLineas, 21=sumaSubtotales
    Private Sub Formato()
        DgvResultado.Columns(0).Visible = False   ' idventa
        DgvResultado.Columns(1).Visible = False   ' idcliente
        DgvResultado.Columns(3).Visible = False   ' doc_cliente
        DgvResultado.Columns(5).Visible = False   ' serieComprobante
        DgvResultado.Columns(8).Visible = False   ' impuesto
        DgvResultado.Columns(11).Visible = False  ' iddetalle_venta
        DgvResultado.Columns(20).Visible = False  ' totalLineas
        DgvResultado.Columns(21).Visible = False  ' sumaSubtotales

        DgvResultado.Columns(2).HeaderText = "CLIENTE"
        DgvResultado.Columns(2).Width = 150
        DgvResultado.Columns(4).HeaderText = "COMPROBANTE"
        DgvResultado.Columns(4).Width = 90
        DgvResultado.Columns(6).HeaderText = "N° COMPROBANTE"
        DgvResultado.Columns(6).Width = 110
        DgvResultado.Columns(7).HeaderText = "FECHA"
        DgvResultado.Columns(7).Width = 130
        DgvResultado.Columns(9).HeaderText = "TOTAL VENTA"
        DgvResultado.Columns(9).Width = 90
        DgvResultado.Columns(10).HeaderText = "ESTADO"
        DgvResultado.Columns(10).Width = 70
        DgvResultado.Columns(12).HeaderText = "COD. ARTÍCULO"
        DgvResultado.Columns(12).Width = 100
        DgvResultado.Columns(13).HeaderText = "ARTÍCULO"
        DgvResultado.Columns(13).Width = 170
        DgvResultado.Columns(14).HeaderText = "CATEGORÍA"
        DgvResultado.Columns(14).Width = 110
        DgvResultado.Columns(15).HeaderText = "CANT."
        DgvResultado.Columns(15).Width = 60
        DgvResultado.Columns(16).HeaderText = "PRECIO"
        DgvResultado.Columns(16).Width = 80
        DgvResultado.Columns(17).HeaderText = "DESCUENTO"
        DgvResultado.Columns(17).Width = 80
        DgvResultado.Columns(18).HeaderText = "P. NETO"
        DgvResultado.Columns(18).Width = 80
        DgvResultado.Columns(19).HeaderText = "SUBTOTAL"
        DgvResultado.Columns(19).Width = 90
    End Sub

    Private Sub FrmConsultaVentas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DtpFechaInicio.Value = Date.Today.AddMonths(-1)
        DtpFechaFin.Value = Date.Today
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Try
            If DtpFechaInicio.Value.Date > DtpFechaFin.Value.Date Then
                MsgBox("La fecha de inicio no puede ser mayor a la fecha fin.",
                       vbOKOnly + vbExclamation, "Rango inválido")
                Return
            End If

            Dim Neg As New Negocio.NVenta
            Dim Tabla As DataTable = Neg.ConsultarDetallado(
                DtpFechaInicio.Value.Date,
                DtpFechaFin.Value.Date.AddDays(1).AddSeconds(-1))

            If Tabla IsNot Nothing Then
                DgvResultado.DataSource = Tabla
                LblTotal.Text = "Total Registros: " & Tabla.Rows.Count.ToString()
                Me.Formato()

                Dim SumaSubtotales As Decimal = 0
                For Each Fila As DataRow In Tabla.Rows
                    SumaSubtotales += Convert.ToDecimal(Fila("subtotal"))
                Next
                LblSumaTotal.Text = "Suma subtotales del período: S/. " & SumaSubtotales.ToString("F2")
            Else
                DgvResultado.DataSource = Nothing
                LblTotal.Text = "Total Registros: 0"
                LblSumaTotal.Text = "Suma subtotales del período: S/. 0.00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs) Handles BtnLimpiar.Click
        DtpFechaInicio.Value = Date.Today.AddMonths(-1)
        DtpFechaFin.Value = Date.Today
        DgvResultado.DataSource = Nothing
        LblTotal.Text = "Total Registros: 0"
        LblSumaTotal.Text = "Suma subtotales del período: S/. 0.00"
    End Sub

End Class
