Public Class FrmReporteInventario
    Private Sub BtnGenerar_Click(sender As Object, e As EventArgs) Handles BtnGenerar.Click
        Try
            Dim StockMinimo As Integer = CInt(NumStockMinimo.Value)

            Dim Neg As New Negocio.NArticulo
            Dim Ds As DataSet = Neg.ReporteInventarioValorizado(StockMinimo)

            If Ds Is Nothing OrElse Ds.Tables.Count < 1 Then
                MsgBox("No hay datos de inventario.", vbOKOnly + vbInformation, "Sin datos")
                Return
            End If

            DgvInventario.DataSource = Ds.Tables(0)
            Formato()

            ' Calcular totales
            CalcularTotales(Ds.Tables(0), StockMinimo)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Formato()
        DgvInventario.AllowUserToAddRows = False
        DgvInventario.ReadOnly = True
        For Each col As DataGridViewColumn In DgvInventario.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Sub CalcularTotales(Tabla As DataTable, StockMinimo As Integer)
        If Tabla Is Nothing OrElse Tabla.Rows.Count = 0 Then
            LblTotalArticulos.Text = "Total Artículos: 0"
            LblValorInventario.Text = "Valor Total Inventario: S/. 0.00"
            LblArticulosCriticos.Text = "Artículos Críticos: 0"
            Return
        End If

        Dim TotalArticulos As Integer = Tabla.Rows.Count
        Dim ValorInventario As Decimal = 0
        Dim ArticulosCriticos As Integer = 0

        For Each Fila As DataRow In Tabla.Rows
            Dim Valor As Decimal = CDec(Fila("valor_economico"))
            Dim Nivel As String = Fila("nivel_stock").ToString()

            ValorInventario += Valor

            If Nivel = "Crítico" OrElse Nivel = "Agotado" Then
                ArticulosCriticos += 1
            End If
        Next

        LblTotalArticulos.Text = "Total Artículos: " & TotalArticulos
        LblValorInventario.Text = "Valor Total Inventario: S/. " & ValorInventario.ToString("0.00")
        LblArticulosCriticos.Text = "Artículos Críticos: " & ArticulosCriticos
    End Sub

    Private Sub FrmReporteInventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NumStockMinimo.Value = 5
    End Sub
End Class
