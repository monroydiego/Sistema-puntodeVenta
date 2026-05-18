
Public Class FrmVenta
    Private DtDetalle As New DataTable

    ' ── Formato del listado (TabPage1) ──────────────────────────────────────
    Private Sub Formato()
        DgvListado.Columns(0).Visible = False   ' Seleccionar (manual)
        DgvListado.Columns(1).Visible = False   ' idventa
        DgvListado.Columns(2).Visible = False   ' idcliente
        DgvListado.Columns(3).Width = 200       ' cliente
        DgvListado.Columns(4).Width = 100       ' doc_cliente
        DgvListado.Columns(5).Width = 110       ' tipo_comprobante
        DgvListado.Columns(6).Width = 70        ' serie_comprobante
        DgvListado.Columns(7).Width = 110       ' num_comprobante
        DgvListado.Columns(8).Width = 130       ' fecha
        DgvListado.Columns(9).Width = 75        ' impuesto
        DgvListado.Columns(10).Width = 100      ' total
        DgvListado.Columns(11).Width = 80       ' estado
        DgvListado.Columns(12).Width = 90       ' clasificacion
        DgvListado.Columns(13).Width = 80       ' num_articulos
        DgvListado.Columns.Item("Seleccionar").Visible = False
        BtnAnular.Visible = False
        ChkSeleccionar.CheckState = CheckState.Unchecked
    End Sub

    ' ── Formato del panel de búsqueda de artículos ───────────────────────────
    Private Sub FormatoArticulos()
        DgvArticulos.Columns(0).Visible = False  ' idarticulo
        DgvArticulos.Columns(1).Visible = False  ' idcategoria
        DgvArticulos.Columns(2).Width = 100      ' categoria
        DgvArticulos.Columns(3).Width = 100      ' codigo
        DgvArticulos.Columns(4).Width = 180      ' nombre
        DgvArticulos.Columns(5).Width = 90       ' precio_venta
        DgvArticulos.Columns(6).Width = 70       ' stock
        DgvArticulos.Columns(7).Width = 180      ' descripcion
        DgvArticulos.Columns(8).Width = 90       ' imagen
        DgvArticulos.Columns(9).Width = 70       ' estado
    End Sub

    ' ── Carga el listado de ventas activas ───────────────────────────────────
    Private Sub Listar()
        Try
            Dim Neg As New Negocio.NVenta
            DgvListado.DataSource = Neg.Listar()
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString()
            Me.Formato()
            Me.Limpiar()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ' ── Inicializa la tabla en memoria para el detalle ───────────────────────
    Private Sub CrearTablaDetalle()
        Me.DtDetalle = New DataTable("Detalle")
        Me.DtDetalle.Columns.Add("idarticulo", GetType(Integer))
        Me.DtDetalle.Columns.Add("codigo", GetType(String))
        Me.DtDetalle.Columns.Add("articulo", GetType(String))
        Me.DtDetalle.Columns.Add("cantidad", GetType(Integer))
        Me.DtDetalle.Columns.Add("precio", GetType(Decimal))
        Me.DtDetalle.Columns.Add("descuento", GetType(Decimal))
        Me.DtDetalle.Columns.Add("subtotal", GetType(Decimal))

        DgvDetalle.DataSource = Me.DtDetalle
        DgvDetalle.Columns(0).Visible = False

        DgvDetalle.Columns(1).HeaderText = "CÓDIGO"
        DgvDetalle.Columns(1).Width = 100
        DgvDetalle.Columns(2).HeaderText = "ARTÍCULO"
        DgvDetalle.Columns(2).Width = 220
        DgvDetalle.Columns(3).HeaderText = "CANTIDAD"
        DgvDetalle.Columns(3).Width = 90
        DgvDetalle.Columns(4).HeaderText = "PRECIO"
        DgvDetalle.Columns(4).Width = 100
        DgvDetalle.Columns(5).HeaderText = "DESCUENTO"
        DgvDetalle.Columns(5).Width = 100
        DgvDetalle.Columns(6).HeaderText = "SUBTOTAL"
        DgvDetalle.Columns(6).Width = 110

        DgvDetalle.Columns(1).ReadOnly = True   ' codigo — no editable
        DgvDetalle.Columns(2).ReadOnly = True   ' articulo — no editable
        DgvDetalle.Columns(4).ReadOnly = True   ' precio — no editable
        DgvDetalle.Columns(6).ReadOnly = True   ' subtotal — calculado
    End Sub


    ' ── Limpia los controles de la pestaña Nueva Venta ──────────────────────
    Private Sub Limpiar()
        BtnInsertar.Visible = True
        TxtValor.Text = ""
        TxtId.Text = ""
        TxtIdCliente.Text = ""
        TxtNombreCliente.Text = ""
        TxtSerieComprobante.Text = "S001"
        TxtNumComprobante.Text = ""
        ' Bug #1 fix: solo asignar si el ComboBox ya tiene items cargados
        If CboTipoComprobante.Items.Count > 0 Then
            CboTipoComprobante.SelectedIndex = 0
        End If
        DtDetalle.Clear()
        DgvDetalle.Refresh()
        TxtSubTotal.Text = "0.00"
        TxtTotalImpuesto.Text = "0.00"
        TxtTotal.Text = "0.00"
        TxtCodigo.Clear()
        Variables.IdCliente = ""
        Variables.NombreCliente = ""
    End Sub

    ' ── Agrega un artículo al detalle (evita duplicados) ────────────────────
    Private Sub AgregarDetalle(Obj As Entidades.Articulo)
        ' VALIDACIÓN: Verificar que hay cliente seleccionado
        If TxtIdCliente.Text = "" OrElse TxtIdCliente.Text = "0" Then
            MsgBox("Debe seleccionar un cliente antes de agregar artículos.", vbOKOnly + vbCritical, "Sin cliente")
            Return
        End If

        ' VALIDACIÓN: Evitar duplicados en el detalle
        For Each FilaTemp As DataGridViewRow In DgvDetalle.Rows
            If Convert.ToInt32(FilaTemp.Cells("idarticulo").Value) = Convert.ToInt32(Obj.IdArticulo) Then
                MsgBox("El artículo ya fue agregado al detalle." & vbCrLf &
                       "Edite la cantidad si desea aumentarla.", vbOKOnly + vbCritical, "Artículo duplicado")
                Return
            End If
        Next

        Dim Row As DataRow = Me.DtDetalle.NewRow()
        Row("idarticulo") = Obj.IdArticulo
        Row("codigo") = Obj.Codigo
        Row("articulo") = Obj.Nombre
        Row("cantidad") = 1
        Row("precio") = Obj.PrecioVenta
        Row("descuento") = 0D
        Row("subtotal") = Obj.PrecioVenta

        Me.DtDetalle.Rows.Add(Row)
        DgvDetalle.Refresh()
        Me.CalcularTotales()
    End Sub

    ' ── Recalcula SubTotal, IGV y Total desde las filas del detalle ──────────
    Private Sub CalcularTotales()
        Try
            Dim Total As Decimal = 0D
            Dim FilaCount As Integer = DgvDetalle.Rows.Count

            ' Iterar sobre todas las filas de detalle y sumar subtotales
            For i As Integer = 0 To FilaCount - 1
                Dim LineSubtotal As Decimal = CDec(DgvDetalle.Rows(i).Cells("subtotal").Value)
                Total += LineSubtotal
            Next

            ' Cálculo de subtotal e impuesto
            Dim Impuesto As Decimal = CDec(TxtImpuesto.Text)
            Dim SubTotalCalculado As Decimal = Math.Round(Total / (1 + Impuesto), 2)
            Dim TotalImpuesto As Decimal = Math.Round(Total - SubTotalCalculado, 2)

            ' Asignar valores calculados a los TextBox
            TxtSubTotal.Text = SubTotalCalculado.ToString("F2")
            TxtTotalImpuesto.Text = TotalImpuesto.ToString("F2")
            TxtTotal.Text = Total.ToString("F2")
        Catch ex As Exception
            MsgBox("Error al calcular totales: " & ex.Message)
        End Try
    End Sub

    ' ── Carga del formulario ─────────────────────────────────────────────────
    Private Sub FrmVenta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Bug #1 fix: inicializar ComboBox ANTES de llamar a Listar/Limpiar
        TxtImpuesto.Text = "0.16"
        If CboTipoComprobante.Items.Count = 0 Then
            CboTipoComprobante.Items.AddRange(New String() {"Factura", "Boleta", "Ticket"})
        End If
        CboTipoComprobante.SelectedIndex = 0
        Me.CrearTablaDetalle()
        Me.Listar()
    End Sub

    ' ── Buscar en listado ────────────────────────────────────────────────────
    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Try
            Dim Neg As New Negocio.NVenta
            DgvListado.DataSource = Neg.Listar()
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString()
            Me.Formato()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ' ── Popup selección de cliente ───────────────────────────────────────────
    Private Sub BtnBuscarCliente_Click(sender As Object, e As EventArgs) Handles BtnBuscarCliente.Click
        FrmCliente_Venta.ShowDialog()

        ' CORRECCIÓN: Asignar valores y verificar que se asignó correctamente
        Try
            If Variables.IdCliente <> "" AndAlso Variables.IdCliente <> "0" Then
                TxtIdCliente.Text = Convert.ToString(Variables.IdCliente)
                TxtNombreCliente.Text = Convert.ToString(Variables.NombreCliente)
                ' Limpiar el detalle cuando se cambia de cliente
                DtDetalle.Clear()
                DgvDetalle.Refresh()
                Me.CalcularTotales()
            End If
        Catch ex As Exception
            MsgBox("Error al seleccionar cliente: " & ex.Message)
        End Try
    End Sub

    ' ── Buscar artículo por código (Enter) ───────────────────────────────────
    Private Sub TxtCodigo_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtCodigo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Try
                ' VALIDACIÓN: Verificar cliente antes de agregar
                If TxtIdCliente.Text = "" OrElse TxtIdCliente.Text = "0" Then
                    MsgBox("Debe seleccionar un cliente primero.", vbOKOnly + vbCritical, "Sin cliente")
                    e.Handled = True
                    Return
                End If

                Dim Neg As New Negocio.NArticulo
                Dim Obj As Entidades.Articulo = Neg.BuscarCodigo(TxtCodigo.Text.Trim())
                If Obj Is Nothing Then
                    MsgBox("No existe artículo con ese código.", vbOKOnly + vbCritical, "No encontrado")
                Else
                    Me.AgregarDetalle(Obj)
                    TxtCodigo.Clear()
                End If
                e.Handled = True
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    ' ── Panel flotante de búsqueda de artículos ──────────────────────────────
    Private Sub BtnBuscarArticulos_Click(sender As Object, e As EventArgs) Handles BtnBuscarArticulos.Click
        ' VALIDACIÓN: Verificar cliente antes de abrir panel
        If TxtIdCliente.Text = "" OrElse TxtIdCliente.Text = "0" Then
            MsgBox("Debe seleccionar un cliente primero.", vbOKOnly + vbCritical, "Sin cliente")
            Return
        End If
        PanelArticulos.Visible = True
    End Sub

    Private Sub BtnCerrar_Click(sender As Object, e As EventArgs) Handles BtnCerrar.Click
        PanelArticulos.Visible = False
    End Sub

    Private Sub BtnBuscarArticulosDetalle_Click(sender As Object, e As EventArgs) Handles BtnBuscarArticulosDetalle.Click
        Try
            Dim Neg As New Negocio.NArticulo
            DgvArticulos.DataSource = Neg.Buscar(TxtBuscarArticulos.Text)
            LblTotalArticulos.Text = "Total Artículos: " & DgvArticulos.DataSource.Rows.Count
            Me.FormatoArticulos()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DgvArticulos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvArticulos.CellDoubleClick
        Try
            If e.RowIndex < 0 Then Return

            ' Acceso por fila + índice de columna (no SelectedCells — independiente del SelectionMode)
            ' Col 0=idarticulo(oculta), 1=idcategoria(oculta), 2=categoria, 3=codigo, 4=nombre, 5=precio_venta
            Dim Fila As DataGridViewRow = DgvArticulos.Rows(e.RowIndex)
            Dim Obj As New Entidades.Articulo
            Obj.IdArticulo = Fila.Cells(0).Value
            Obj.Codigo = Fila.Cells(3).Value
            Obj.Nombre = Fila.Cells(4).Value
            Obj.PrecioVenta = Fila.Cells(5).Value
            Me.AgregarDetalle(Obj)
            PanelArticulos.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ' ── Recalcula subtotal al editar cantidad o descuento ───────────────────
    Private Sub DgvDetalle_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DgvDetalle.CellEndEdit
        Try
            If e.RowIndex < 0 Then Return

            Dim Fila As DataGridViewRow = DgvDetalle.Rows(e.RowIndex)
            Dim Precio As Decimal = CDec(Fila.Cells("precio").Value)
            Dim Cantidad As Integer = 0
            Dim Descuento As Decimal = 0D

            ' VALIDACIÓN: Cantidad debe ser >= 1
            If Not Integer.TryParse(Fila.Cells("cantidad").Value.ToString(), Cantidad) OrElse Cantidad < 1 Then
                MsgBox("La cantidad debe ser un número mayor a 0.", vbOKOnly + vbCritical, "Cantidad inválida")
                Fila.Cells("cantidad").Value = 1
                Cantidad = 1
            End If

            ' VALIDACIÓN: Descuento no debe ser negativo ni mayor al precio
            If Not Decimal.TryParse(Fila.Cells("descuento").Value.ToString(), Descuento) Then
                Descuento = 0D
            End If

            If Descuento < 0 Or Descuento > Precio Then
                MsgBox("El descuento debe estar entre 0 y " & Precio.ToString("F2"), vbOKOnly + vbCritical, "Descuento inválido")
                Fila.Cells("descuento").Value = 0D
                Descuento = 0D
            End If

            ' Cálculo correcto del subtotal: (precio - descuento) * cantidad
            Fila.Cells("subtotal").Value = Math.Round((Precio - Descuento) * Cantidad, 2)
            Me.CalcularTotales()
        Catch ex As Exception
            MsgBox("Error al editar detalle: " & ex.Message)
        End Try
    End Sub

    Private Sub DgvDetalle_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DgvDetalle.RowsRemoved
        Me.CalcularTotales()
    End Sub

    ' ── ChkSeleccionar: muestra/oculta columna y botón Anular ───────────────
    Private Sub ChkSeleccionar_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSeleccionar.CheckedChanged
        DgvListado.Columns.Item("Seleccionar").Visible = ChkSeleccionar.Checked
        BtnAnular.Visible = ChkSeleccionar.Checked
    End Sub

    ' ── Anular la venta seleccionada en el listado ───────────────────────────
    Private Sub BtnAnular_Click(sender As Object, e As EventArgs) Handles BtnAnular.Click
        If DgvListado.CurrentRow Is Nothing Then
            MsgBox("Seleccione una venta de la lista.", vbOKOnly + vbExclamation, "Sin selección")
            Return
        End If
        Dim IdVenta As Integer = Convert.ToInt32(DgvListado.CurrentRow.Cells("idventa").Value)
        If MsgBox("¿Anular la venta seleccionada?" & vbCrLf &
                  "El stock se restaurará automáticamente (TRG03).",
                  vbYesNo + vbQuestion, "Confirmar anulación") = MsgBoxResult.Yes Then
            Try
                Dim Neg As New Negocio.NVenta
                If Neg.Anular(IdVenta) Then
                    MsgBox("Venta anulada correctamente.", vbOKOnly + vbInformation, "Anulado")
                    Me.Listar()
                End If
            Catch ex As Exception
                MsgBox("Error al anular venta: " & ex.Message)
            End Try
        End If
    End Sub

    ' ── Guardar nueva venta ──────────────────────────────────────────────────
    Private Sub BtnInsertar_Click(sender As Object, e As EventArgs) Handles BtnInsertar.Click
        Try
            ' VALIDACIONES completas antes de guardar
            If TxtIdCliente.Text = "" OrElse TxtIdCliente.Text = "0" Then
                MsgBox("Debe seleccionar un cliente.", vbOKOnly + vbCritical, "Falta dato")
                Return
            End If
            If TxtNumComprobante.Text.Trim() = "" Then
                MsgBox("Ingrese el número de comprobante.", vbOKOnly + vbCritical, "Falta dato")
                Return
            End If
            If DtDetalle.Rows.Count = 0 Then
                MsgBox("Agregue al menos un artículo al detalle.", vbOKOnly + vbCritical, "Sin detalle")
                Return
            End If

            ' Crear objeto Venta con datos del formulario
            Dim Obj As New Entidades.Venta
            Obj.IdCliente = Convert.ToInt32(TxtIdCliente.Text)
            Obj.IdUsuario = Convert.ToInt32(Variables.IdUsuario)
            Obj.IdTipoComprobante = CboTipoComprobante.SelectedIndex + 1
            Obj.NumComprobante = TxtNumComprobante.Text.Trim()
            Obj.FechaVenta = Date.Now
            Obj.Impuesto = CDec(TxtImpuesto.Text)
            Obj.TotalVenta = CDec(TxtTotal.Text)

            ' Insertar venta con detalle (transacción en BD)
            Dim Neg As New Negocio.NVenta
            If Neg.Insertar(Obj, DtDetalle) Then
                MsgBox("Venta registrada correctamente." & vbCrLf &
                       "Stock actualizado automáticamente.", vbOKOnly + vbInformation, "Registro correcto")
                Me.Listar()
            End If
        Catch ex As Exception
            MsgBox("Error al registrar venta: " & ex.Message)
        End Try
    End Sub

    ' ── Cancelar / limpiar formulario ────────────────────────────────────────
    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Limpiar()
    End Sub

End Class
