Public Class FrmProveedor_Ingreso
    Private Sub Formato()
        ' se agregan 10 columnas propias del Stored Procedure 
        ' Agregaremos 1 columna del Listado (seleccionar) 
        ' el parametro de columns es para saber cuantas columnas vamos agregar 
        DgvListado.Columns(0).Visible = False ' Ocultamos la columna 0 que es el ID de la categoria, se inicia con una columna 0 
        DgvListado.Columns(0).Width = 100 ' id
        DgvListado.Columns(1).Width = 80 ' id categoria
        DgvListado.Columns(2).Width = 120 ' Categoria  
        DgvListado.Columns(3).Width = 120 ' codigo
        DgvListado.Columns(4).Width = 100 ' Nombre
        DgvListado.Columns(5).Width = 100 ' Precio Venta
        DgvListado.Columns(6).Width = 120 ' stock 
        DgvListado.Columns(7).Width = 100 ' Descripcion 
        DgvListado.Columns(8).Width = 120 ' Imagen 




        DgvListado.Columns.Item("Seleccionar").Visible = False
    End Sub

    ' Metodo buscar
    Private Sub Listar()
        Try
            'Creamos una instancia que haga referencia a la capa negocio
            Dim Neg As New Negocio.NPersona ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
            DgvListado.DataSource = Neg.ListarProveedores() ' obtener el listado de categorias desde la capa negocio 
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString() ' Muestra el total de registros en el Label
            Me.Formato() ' Llamamos al metodo Formato para establecer el formato de las columnas
        Catch ex As Exception
            MsgBox(ex.Message) ' Si ocurre un error, muestra un mensaje con la descripción del error
        End Try
    End Sub

    ' Metodo Buscar
    Private Sub Buscar()
        Try
            'Creamos una instancia que haga referencia a la capa negocio
            Dim Neg As New Negocio.NPersona ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
            Dim Valor As String ' Declara una variable para almacenar el valor de búsqueda
            Valor = TxtValor.Text ' se almacenará el valor de la caja de texto a la variable valor
            DgvListado.DataSource = Neg.BuscarProveedores(Valor) ' obtener el listado de categorias desde la capa negocio, utilizando el valor de búsqueda
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString() ' Muestra el total de registros encontrados
            Me.Formato() ' Llamamos al metodo Formato para establecer el formato de las columnas
        Catch ex As Exception
            MsgBox(ex.Message) ' Si ocurre un error, muestra un mensaje con la descripción del error
        End Try
    End Sub
    Private Sub FrmProveedor_Ingreso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Listar()
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Me.Buscar()
    End Sub

    Private Sub DgvListado_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvListado.CellDoubleClick
        Variables.IdProveedor = DgvListado.SelectedCells.Item(1).Value
        Variables.NombreProveedor = DgvListado.SelectedCells.Item(3).Value
        Me.Close()
    End Sub
End Class