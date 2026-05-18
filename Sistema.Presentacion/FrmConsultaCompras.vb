Public Class FrmConsultaCompras

    ' vw_ComprasDetalladas columnas (20):
    ' 0=idingreso, 1=idproveedor, 2=proveedor, 3=doc_proveedor, 4=tipoComprobante,
    ' 5=serieComprobante, 6=numComprobante, 7=fecha, 8=impuesto, 9=totalIngreso,
    ' 10=estado, 11=iddetalle_ingreso, 12=codigoArticulo, 13=nombreArticulo,
    ' 14=categoria, 15=cantidad, 16=precio, 17=importe, 18=totalLineas, 19=sumaImportes
    Private Sub Formato()
        DgvResultado.Columns(0).Visible = False   ' idingreso
        DgvResultado.Columns(1).Visible = False   ' idproveedor
        DgvResultado.Columns(3).Visible = False   ' doc_proveedor
        DgvResultado.Columns(5).Visible = False   ' serieComprobante
        DgvResultado.Columns(8).Visible = False   ' impuesto
        DgvResultado.Columns(11).Visible = False  ' iddetalle_ingreso
        DgvResultado.Columns(18).Visible = False  ' totalLineas
        DgvResultado.Columns(19).Visible = False  ' sumaImportes

        DgvResultado.Columns(2).HeaderText = "PROVEEDOR"
        DgvResultado.Columns(2).Width = 150
        DgvResultado.Columns(4).HeaderText = "COMPROBANTE"
        DgvResultado.Columns(4).Width = 90
        DgvResultado.Columns(6).HeaderText = "N° COMPROBANTE"
        DgvResultado.Columns(6).Width = 110
        DgvResultado.Columns(7).HeaderText = "FECHA"
        DgvResultado.Columns(7).Width = 130
        DgvResultado.Columns(9).HeaderText = "TOTAL INGRESO"
        DgvResultado.Columns(9).Width = 100
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
        DgvResultado.Columns(17).HeaderText = "IMPORTE"
        DgvResultado.Columns(17).Width = 90
    End Sub

    Private Sub FrmConsultaCompras_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

            Dim Neg As New Negocio.NIngreso
            Dim Tabla As DataTable = Neg.ConsultarDetallado(
                DtpFechaInicio.Value.Date,
                DtpFechaFin.Value.Date.AddDays(1).AddSeconds(-1))

            If Tabla IsNot Nothing Then
                DgvResultado.DataSource = Tabla
                LblTotal.Text = "Total Registros: " & Tabla.Rows.Count.ToString()
                Me.Formato()

                Dim SumaImportes As Decimal = 0
                For Each Fila As DataRow In Tabla.Rows
                    SumaImportes += Convert.ToDecimal(Fila("importe"))
                Next
                LblSumaTotal.Text = "Suma importes del período: S/. " & SumaImportes.ToString("F2")
            Else
                DgvResultado.DataSource = Nothing
                LblTotal.Text = "Total Registros: 0"
                LblSumaTotal.Text = "Suma importes del período: S/. 0.00"
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
        LblSumaTotal.Text = "Suma importes del período: S/. 0.00"
    End Sub

End Class
