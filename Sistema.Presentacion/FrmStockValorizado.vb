Public Class FrmStockValorizado

    ' vw_StockValorizado columnas (13):
    ' 0=idcategoria, 1=categoria, 2=idarticulo, 3=codigoArticulo, 4=nombreArticulo,
    ' 5=descripcion, 6=stockActual, 7=precioVenta, 8=valorTotal, 9=estadoStock,
    ' 10=valorCategoria, 11=articulosEnCategoria, 12=precioPromedioCategoria
    Private Sub Formato()
        DgvResultado.Columns(0).Visible = False   ' idcategoria
        DgvResultado.Columns(2).Visible = False   ' idarticulo

        DgvResultado.Columns(1).HeaderText = "CATEGORÍA"
        DgvResultado.Columns(1).Width = 130
        DgvResultado.Columns(3).HeaderText = "CÓDIGO"
        DgvResultado.Columns(3).Width = 90
        DgvResultado.Columns(4).HeaderText = "ARTÍCULO"
        DgvResultado.Columns(4).Width = 180
        DgvResultado.Columns(5).HeaderText = "DESCRIPCIÓN"
        DgvResultado.Columns(5).Width = 200
        DgvResultado.Columns(6).HeaderText = "STOCK"
        DgvResultado.Columns(6).Width = 70
        DgvResultado.Columns(7).HeaderText = "PRECIO VENTA"
        DgvResultado.Columns(7).Width = 100
        DgvResultado.Columns(8).HeaderText = "VALOR TOTAL"
        DgvResultado.Columns(8).Width = 100
        DgvResultado.Columns(9).HeaderText = "ESTADO STOCK"
        DgvResultado.Columns(9).Width = 100
        DgvResultado.Columns(10).HeaderText = "VALOR CATEGORÍA"
        DgvResultado.Columns(10).Width = 120
        DgvResultado.Columns(11).HeaderText = "ARTS. CATEGORÍA"
        DgvResultado.Columns(11).Width = 120
        DgvResultado.Columns(12).HeaderText = "PRECIO PROM. CAT."
        DgvResultado.Columns(12).Width = 130
    End Sub

    Private Sub FrmStockValorizado_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CargarStock()
    End Sub

    Private Sub CargarStock()
        Try
            Dim Neg As New Negocio.NArticulo
            Dim Tabla As DataTable = Neg.ConsultarStockValorizado()

            If Tabla IsNot Nothing Then
                DgvResultado.DataSource = Tabla
                LblTotal.Text = "Total Artículos: " & Tabla.Rows.Count.ToString()
                Me.Formato()

                Dim ValorInventario As Decimal = 0
                For Each Fila As DataRow In Tabla.Rows
                    ValorInventario += Convert.ToDecimal(Fila("valorTotal"))
                Next
                LblValorTotal.Text = "Valor total inventario: S/. " & ValorInventario.ToString("F2")
            Else
                DgvResultado.DataSource = Nothing
                LblTotal.Text = "Total Artículos: 0"
                LblValorTotal.Text = "Valor total inventario: S/. 0.00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BtnActualizar_Click(sender As Object, e As EventArgs) Handles BtnActualizar.Click
        Me.CargarStock()
    End Sub

End Class
