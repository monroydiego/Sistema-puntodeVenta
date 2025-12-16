Public Class FrmProveedores
    Private Sub Formato()
        ' se agregan 10 columnas propias del Stored Procedure 
        ' Agregaremos 1 columna del Listado (seleccionar) 
        ' el parametro de columns es para saber cuantas columnas vamos agregar 
        DgvListado.Columns(0).Visible = False ' Ocultamos la columna 0 que es el ID de la categoria, se inicia con una columna 0 
        DgvListado.Columns(0).Width = 100 ' id
        DgvListado.Columns(1).Width = 80 ' id categoria
        DgvListado.Columns(3).Width = 120 ' Categoria  
        DgvListado.Columns(4).Width = 120 ' codigo
        DgvListado.Columns(5).Width = 100 ' Nombre
        DgvListado.Columns(6).Width = 100 ' Precio Venta
        DgvListado.Columns(7).Width = 120 ' stock 
        DgvListado.Columns(8).Width = 100 ' Descripcion 
        DgvListado.Columns(9).Width = 120 ' Imagen 





        DgvListado.Columns.Item("Seleccionar").Visible = False ' Ocultamos la columna seleccionar 
        BtnEliminar.Visible = False ' Ocultamos el  botón de eliminar
        ChkSeleccionar.CheckState = False
    End Sub

    ' Metodo buscar
    Private Sub Listar()
        Try
            'Creamos una instancia que haga referencia a la capa negocio
            Dim Neg As New Negocio.NPersona ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
            DgvListado.DataSource = Neg.ListarProveedores() ' obtener el listado de categorias desde la capa negocio 
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

    ' Metodo Limpiar 
    Private Sub Limpiar()
        BtnInsertar.Visible = True 'indicamos que el botón insertar éste visible
        BtnActualizar.Visible = False ' Indicamos que al momento de limpiar no sea visible el botón de actualizar
        TxtValor.Text = "" ' Limpia el campo de texto de búsqueda
        TxtId.Text = "" ' Limpia el campo de texto del ID
        TxtNumDocumento.Text = ""
        TxtTelefono.Text = ""
        TxtEmail.Text = ""
        TxtClave.Text = ""
        TxtDireccion.Text = "" ' Limpia el campo texto de precio de venta
        TxtNumDocumento.Text = "" ' Limpia el campo de texto de Stock
        TxtNombre.Text = "" ' Limpia el campo de texto del nombre
        TxtTelefono.Text = "" ' Limpia el campo de texto de la descripción
    End Sub
    Private Sub FrmProveedores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Listar()
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Me.Buscar()
    End Sub
End Class