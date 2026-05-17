Public Class FrmIngreso
    Private DtDetalle As New DataTable ' Nos permite almacenar todas las filas de los detalles que vamos agregando a nuestro ingreso
    Private Sub Formato()
        ' se agregan 10 columnas propias del Stored Procedure 
        ' Agregaremos 1 columna del Listado (seleccionar) 
        ' el parametro de columns es para saber cuantas columnas vamos agregar 
        DgvListado.Columns(0).Visible = False ' Ocultamos la columna 0 que es el ID de la categoria, se inicia con una columna 0 
        DgvListado.Columns(2).Visible = False ' idcategoria (que no sea visible)
        DgvListado.Columns(0).Width = 100 ' id
        DgvListado.Columns(1).Width = 60 ' id categoria
        DgvListado.Columns(3).Width = 150 ' Categoria  
        DgvListado.Columns(4).Width = 150 ' codigo
        DgvListado.Columns(5).Width = 100 ' Nombre
        DgvListado.Columns(6).Width = 70 ' Precio Venta
        DgvListado.Columns(7).Width = 70 ' stock 
        DgvListado.Columns(8).Width = 60 ' Descripcion 
        DgvListado.Columns(9).Width = 100 ' Imagen 
        DgvListado.Columns(10).Width = 100 ' Estado
        DgvListado.Columns(11).Width = 100 ' Estado




        DgvListado.Columns.Item("Seleccionar").Visible = False ' Ocultamos la columna seleccionar 
        BtnAnular.Visible = False ' Ocultamos el el botón de desactivar
        ChkSeleccionar.CheckState = False
    End Sub

    Private Sub FormatoArticulos()
        DgvArticulos.Columns(0).Visible = False
        DgvArticulos.Columns(1).Visible = False
        DgvArticulos.Columns(2).Width = 100
        DgvArticulos.Columns(3).Width = 100
        DgvArticulos.Columns(4).Width = 150
        DgvArticulos.Columns(5).Width = 100
        DgvArticulos.Columns(6).Width = 100
        DgvArticulos.Columns(7).Width = 200
        DgvArticulos.Columns(8).Width = 100
        DgvArticulos.Columns(9).Width = 100

    End Sub

    ' Metodo buscar
    Private Sub Listar()
        Try
            'Creamos una instancia que haga referencia a la capa negocio
            Dim Neg As New Negocio.NIngreso ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
            DgvListado.DataSource = Neg.Listar() ' obtener el listado de categorias desde la capa negocio 
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString() ' Muestra el total de registros en el Label
            Me.Formato() ' Llamamos al metodo Formato para establecer el formato de las columnas
            Me.Limpiar()
        Catch ex As Exception
            MsgBox(ex.Message) ' Si ocurre un error, muestra un mensaje con la descripción del error
        End Try
    End Sub

    ' Metodo Buscar
    Private Sub Buscar()
        Try
            'Creamos una instancia que haga referencia a la capa negocio
            Dim Neg As New Negocio.NIngreso ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
            Dim Valor As String ' Declara una variable para almacenar el valor de búsqueda
            Valor = TxtValor.Text ' se almacenará el valor de la caja de texto a la variable valor
            DgvListado.DataSource = Neg.Buscar(Valor) ' obtener el listado de categorias desde la capa negocio, utilizando el valor de búsqueda
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString() ' Muestra el total de registros encontrados
            Me.Formato() ' Llamamos al metodo Formato para establecer el formato de las columnas
        Catch ex As Exception
            MsgBox(ex.Message) ' Si ocurre un error, muestra un mensaje con la descripción del error
        End Try
    End Sub

    Private Sub CrearTablaDetalle()
        ' Creamos una nueva tabla en memoria llamada Detalle
        Me.DtDetalle = New DataTable("Detalle")

        ' Agregamos las columnas con sus tipos de dato
        Me.DtDetalle.Columns.Add("idarticulo", GetType(Integer))      ' ID único del artículo
        Me.DtDetalle.Columns.Add("codigo", GetType(String))           ' Código del producto
        Me.DtDetalle.Columns.Add("articulo", GetType(String))         ' Nombre del producto
        Me.DtDetalle.Columns.Add("cantidad", GetType(Integer))        ' Cantidad a ingresar
        Me.DtDetalle.Columns.Add("precio", GetType(Decimal))          ' Precio unitario
        Me.DtDetalle.Columns.Add("importe", GetType(Decimal))         ' Total (cantidad × precio)


        DgvDetalle.DataSource = Me.DtDetalle
        DgvDetalle.Columns(0).Visible = False ' que la columna del articulo no sea visible

        DgvDetalle.Columns(1).HeaderText = "CODIGO"
        DgvDetalle.Columns(1).Width = 100
        DgvDetalle.Columns(2).HeaderText = "ARTICULO"
        DgvDetalle.Columns(2).Width = 200
        DgvDetalle.Columns(3).HeaderText = "CANTIDAD"
        DgvDetalle.Columns(3).Width = 100
        DgvDetalle.Columns(4).HeaderText = "PRECIO"
        DgvDetalle.Columns(4).Width = 100
        DgvDetalle.Columns(5).HeaderText = "IMPORTE"
        DgvDetalle.Columns(5).Width = 100

        DgvDetalle.Columns(1).ReadOnly = True
        DgvDetalle.Columns(2).ReadOnly = True
        DgvDetalle.Columns(5).ReadOnly = True
    End Sub

    ' Metodo Limpiar 
    Private Sub Limpiar()
        BtnInsertar.Visible = True
        TxtValor.Text = ""
        TxtId.Text = ""
        TxtCodigo.Text = ""
        TxtIdProveedor.Text = ""
        TxtNombreProveedor.Text = ""
        TxtSerieComprobante.Text = ""
        TxtNumComprobante.Text = ""
        DtDetalle.Clear()
        TxtSubTotal.Text = 0
        TxtTotalImpuesto.Text = 0
        TxtTotal.Text = 0

    End Sub

    Private Sub AgregarDetalle(Fila As Entidades.Articulo)
        Dim Agregar As Boolean = True


        For Each FilaTemp As DataGridViewRow In DgvDetalle.Rows
            If (Convert.ToInt32(FilaTemp.Cells("idarticulo").Value) = Convert.ToInt32(Fila.IdArticulo)) Then
                Agregar = False
                MsgBox("El artículo ya ha sido agregado", vbOKOnly + vbCritical, "Error agregando detalles")
            End If
        Next

        If (Agregar) Then
            Dim Row As DataRow
            Row = Me.DtDetalle.NewRow()
            Row("idarticulo") = Convert.ToString(Fila.IdArticulo)
            Row("codigo") = Convert.ToString(Fila.Codigo)
            Row("articulo") = Convert.ToString(Fila.Nombre)
            Row("cantidad") = Convert.ToString(1)
            Row("precio") = Convert.ToString(Fila.PrecioVenta)
            Row("importe") = Convert.ToString(Fila.PrecioVenta)

            Me.DtDetalle.Rows.Add(Row)
            Me.CalcularTotales()
        End If

    End Sub

    Private Sub CalcularTotales()
        Dim Total As Decimal = 0
        Dim SubTotal As Decimal = 0
        Dim TotalImpuesto As Decimal = 0
        For Each FilaTemp As DataGridViewRow In DgvDetalle.Rows
            Total = Total + CDec(FilaTemp.Cells("importe").Value)
        Next
        SubTotal = Math.Round((Total / (1 + TxtImpuesto.Text)), 2)
        TxtTotal.Text = Total
        TxtSubTotal.Text = SubTotal
        TxtTotalImpuesto.Text = CStr(Total - SubTotal)
    End Sub

    Private Sub FrmIngreso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Listar()
        Me.CrearTablaDetalle()
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Me.Buscar()
    End Sub

    Private Sub BtnBuscarProveedor_Click(sender As Object, e As EventArgs) Handles BtnBuscarProveedor.Click
        FrmProveedor_Ingreso.ShowDialog()
        TxtIdProveedor.Text = Variables.IdProveedor
        TxtNombreProveedor.Text = Variables.NombreProveedor
    End Sub

    Private Sub TxtCodigo_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtCodigo.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            Try
                Dim Neg As New Negocio.NArticulo
                Dim Obj As New Entidades.Articulo

                Obj = Neg.BuscarCodigo(TxtCodigo.Text)

                If (Obj Is Nothing) Then
                    MsgBox("No existe artículo con ese código de barra", vbOKOnly + vbCritical, "NO EXISTE ARTICULO")

                Else
                    Me.AgregarDetalle(Obj)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub BtnBuscarArticulos_Click(sender As Object, e As EventArgs) Handles BtnBuscarArticulos.Click
        PanelArticulos.Visible = True
    End Sub

    Private Sub BtnCerrar_Click(sender As Object, e As EventArgs) Handles BtnCerrar.Click
        PanelArticulos.Visible = False
    End Sub

    Private Sub BtnBuscarArticulosDetalle_Click(sender As Object, e As EventArgs) Handles BtnBuscarArticulosDetalle.Click
        Try
            Dim Neg As New Negocio.NArticulo
            Dim Valor As String

            Valor = TxtBuscarArticulos.Text
            DgvArticulos.DataSource = Neg.Buscar(Valor)
            LblTotalArticulos.Text = "Total Artículos: " & DgvArticulos.DataSource.Rows.Count
            Me.FormatoArticulos()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DgvArticulos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvArticulos.CellDoubleClick
        Try
            Dim Obj As New Entidades.Articulo
            Obj.IdArticulo = DgvArticulos.SelectedCells.Item(0).Value
            Obj.Codigo = DgvArticulos.SelectedCells.Item(3).Value
            Obj.Nombre = DgvArticulos.SelectedCells.Item(4).Value
            Obj.PrecioVenta = DgvArticulos.SelectedCells.Item(5).Value
            Me.AgregarDetalle(Obj)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DgvDetalle_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DgvDetalle.CellEndEdit
        Dim Fila As DataGridViewRow = CType(DgvDetalle.Rows(e.RowIndex), DataGridViewRow)
        Dim Precio As Decimal = Fila.Cells("precio").Value
        Dim Cantidad As Integer = Fila.Cells("cantidad").Value

        Fila.Cells("importe").Value = Precio * Cantidad
        Me.CalcularTotales()

    End Sub

    Private Sub DgvDetalle_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DgvDetalle.RowsRemoved
        Me.CalcularTotales()
    End Sub

    Private Sub BtnInsertar_Click(sender As Object, e As EventArgs) Handles BtnInsertar.Click
        Try
            If (TxtIdProveedor.Text <> "" And CboTipoComprobante.Text <> "" And TxtNumComprobante.Text <> "" And DtDetalle.Rows.Count > 0) Then
                Dim Obj As New Entidades.Ingreso
                Dim Neg As New Negocio.NIngreso

                Obj.IdUsuario = Variables.IdUsuario
                Obj.IdProveedor = TxtIdProveedor.Text
                Obj.TipoComprobante = CboTipoComprobante.Text
                Obj.SerieComprobante = TxtSerieComprobante.Text
                Obj.NumComprobante = TxtNumComprobante.Text
                Obj.Impuesto = TxtImpuesto.Text
                Obj.Total = TxtTotal.Text

                If (Neg.Insertar(Obj, DtDetalle)) Then
                    MsgBox("Se ha registrado corectamente", vbOKOnly + vbInformation, "Registro correcto")
                    Me.Listar()
                End If
                MsgBox("No se ha podido registrar el ingreso", vbOKOnly + vbCritical, "Registro incorrecto")
            Else
                MsgBox("Ingrese los datos Obligatorios, al menos agregue un detalle", vbOKOnly + vbCritical, "Faltan datos")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class