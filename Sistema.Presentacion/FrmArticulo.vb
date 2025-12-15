Imports System.Windows.Forms.VisualStyles
Imports System.IO ' Estapacio de nombres

Public Class FrmArticulo
    ' Variables globales
    Private RutaOrigen As String ' de donde estoy cargando la imagen
    Private RutaDestino As String ' a donde voy a cargar la imagen
    Private Directorio As String = "C:\sistema\"

    Private Sub Formato()
        ' se agregan 10 columnas propias del Stored Procedure 
        ' Agregaremos 1 columna del Listado (seleccionar) 
        ' el parametro de columns es para saber cuantas columnas vamos agregar 
        DgvListado.Columns(0).Visible = False ' Ocultamos la columna 0 que es el ID de la categoria, se inicia con una columna 0 
        DgvListado.Columns(2).Visible = False ' idcategoria (que no sea visible)
        DgvListado.Columns(0).Width = 100 ' id
        DgvListado.Columns(1).Width = 100 ' id categoria
        DgvListado.Columns(3).Width = 100 ' Categoria  
        DgvListado.Columns(4).Width = 100 ' codigo
        DgvListado.Columns(5).Width = 150 ' Nombre
        DgvListado.Columns(6).Width = 100 ' Precio Venta
        DgvListado.Columns(7).Width = 100 ' stock 
        DgvListado.Columns(8).Width = 200 ' Descripcion 
        DgvListado.Columns(9).Width = 100 ' Imagen 
        DgvListado.Columns(10).Width = 100 ' Estado




        DgvListado.Columns.Item("Seleccionar").Visible = False ' Ocultamos la columna seleccionar 
        BtnEliminar.Visible = False ' Ocultamos el  botón de eliminar
        BtnActivar.Visible = False ' Ocultamos el botón de activar
        BtnDesactivar.Visible = False ' Ocultamos el el botón de desactivar
        ChkSeleccionar.CheckState = False
    End Sub

    ' Metodo buscar
    Private Sub Listar()
        Try
            'Creamos una instancia que haga referencia a la capa negocio
            Dim Neg As New Negocio.NArticulo ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
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
            Dim Neg As New Negocio.NArticulo ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
            Dim Valor As String ' Declara una variable para almacenar el valor de búsqueda
            Valor = TxtValor.Text ' se almacenará el valor de la caja de texto a la variable valor
            DgvListado.DataSource = Neg.Buscar(Valor) ' obtener el listado de categorias desde la capa negocio, utilizando el valor de búsqueda
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
        TxtCodigo.Text = "" ' Limpia el campo de texto de Codigo
        TxtPrecioVenta.Text = "" ' Limpia el campo texto de precio de venta
        TxtStock.Text = "" ' Limpia el campo de texto de Stock
        txtImagen.Text = "" ' Limpia el campo de texto de imagen
        PicImagen.Image = Nothing ' Limpia la caja donde almacena la imagen
        TxtNombre.Text = "" ' Limpia el campo de texto del nombre
        TxtDescripcion.Text = "" ' Limpia el campo de texto de la descripción
        RutaOrigen = ""
    End Sub

    Private Sub CargarCategoria()
        Try
            Dim Neg As New Negocio.NCategoria
            CboCategoria.DataSource = Neg.Seleccionar
            CboCategoria.ValueMember = "idcategoria"
            CboCategoria.DisplayMember = "nombre"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub FrmArticulo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Listar()
        Me.CargarCategoria()


    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Me.Buscar()
    End Sub

    Private Sub BtnCargarImagen_Click(sender As Object, e As EventArgs) Handles BtnCargarImagen.Click
        Dim File As New OpenFileDialog() ' Instancia a OpenFileDialog, esto nos permite que se nos abra un bloque para cargar un archivo
        ' filtro
        File.Filter = "Image File (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png" ' Mostramos las extensiones que se pueden cargar e indicamos que solo acepte esas extensiones


        If (File.ShowDialog() = DialogResult.OK) Then
            ' mostramos la imagen en el picture box
            PicImagen.Image = Image.FromFile(File.FileName) ' Obtener una imagen de un archivo y obtener toda la ruta de la imagen

            RutaOrigen = File.FileName

            txtImagen.Text = File.FileName.Substring(File.FileName.LastIndexOf("\") + 1)
        End If
    End Sub

    Private Sub BtnInsertar_Click(sender As Object, e As EventArgs) Handles BtnInsertar.Click
        Try


            If Me.ValidateChildren = True And CboCategoria.Text <> "" And TxtCodigo.Text <> "" And TxtStock.Text <> "" And TxtPrecioVenta.Text <> "" And TxtNombre.Text <> "" Then ' Verifica si la validación de los controles hijos es exitosa y si el campo de nombre no está vacío
                Dim Obj As New Entidades.Articulo 'Declaramos un objeto en donde se haga la referencia a entidades, entidad hace referencia a Articulo
                Dim Neg As New Negocio.NArticulo ' se hace referencia a la capa negocio de NArticulo


                Obj.IdCategoria = CboCategoria.SelectedValue ' el idcategoria, será el valor seleccionado en el combo box
                Obj.Codigo = TxtCodigo.Text
                Obj.Nombre = TxtNombre.Text ' estos parametros se esperan de la capa de Ncategoria y Dcategoria en la funcion agregar
                Obj.Stock = TxtStock.Text
                Obj.PrecioVenta = TxtPrecioVenta.Text
                Obj.Descripcion = TxtDescripcion.Text
                Obj.Imagen = txtImagen.Text

                If (Neg.Insertar(Obj)) Then ' Intenta insertar la nueva categoría utilizando la capa de negocio
                    MsgBox("Se ha registrado correctamente", vbOKOnly + vbInformation, "Registro correcto") ' Muestra un mensaje de éxito

                    If (txtImagen.Text <> "") Then
                        RutaDestino = Directorio & txtImagen.Text ' Concatenamos la ruta de destino -- > c:\sistema\ ..
                        File.Copy(RutaOrigen, RutaDestino) ' que copie la ruta de la imagen  a la ruta de nuestro directoria
                    End If

                    Me.Listar() ' Actualiza la lista de categorías

                Else
                    MsgBox("No se ha podido registrar", vbOKOnly + vbCritical, "Registro Incorrecto") ' Muestra un mensaje de error si no se pudo insertar
                End If
            Else
                MsgBox("Rellene todos los campos obligatorios (*)", vbOKOnly + vbCritical, "Falta Ingresar datos") ' Muestra un mensaje si faltan campos obligatorios
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Limpiar()

        TabGeneral.SelectedIndex = 0

    End Sub

    Private Sub DgvListado_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvListado.CellDoubleClick
        Try
            TxtId.Text = DgvListado.SelectedCells.Item(1).Value
            CboCategoria.SelectedValue = DgvListado.SelectedCells.Item(2).Value
            TxtCodigo.Text = DgvListado.SelectedCells.Item(4).Value
            TxtNombre.Text = DgvListado.SelectedCells.Item(5).Value
            TxtPrecioVenta.Text = DgvListado.SelectedCells.Item(6).Value
            TxtStock.Text = DgvListado.SelectedCells.Item(7).Value
            TxtDescripcion.Text = DgvListado.SelectedCells.Item(8).Value

            Dim Imagen As String
            Imagen = DgvListado.SelectedCells.Item(9).Value
            ' al momento de actualizar mostraremos la imagen en el picturebox
            If (Imagen <> "") Then
                PicImagen.Image = Image.FromFile(Directorio & Imagen)
                txtImagen.Text = Imagen

            Else
                PicImagen.Image = Nothing
                txtImagen.Text = ""
            End If

            BtnInsertar.Visible = False
            BtnActualizar.Visible = True
            TabGeneral.SelectedIndex = 1

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BtnActualizar_Click(sender As Object, e As EventArgs) Handles BtnActualizar.Click
        Try

            ' Verifica si la validación de los controles hijos es exitosa y si el campo de nombre no está vacío
            If Me.ValidateChildren = True And CboCategoria.Text <> "" And TxtCodigo.Text <> "" And TxtStock.Text <> "" And TxtPrecioVenta.Text <> "" And TxtNombre.Text <> "" And TxtId.Text <> "" Then

                Dim Obj As New Entidades.Articulo 'Declaramos un objeto en donde se haga la referencia a entidades, entidad hace referencia a Articulo
                Dim Neg As New Negocio.NArticulo ' se hace referencia a la capa negocio de NArticulo

                Obj.IdArticulo = TxtId.Text
                Obj.IdCategoria = CboCategoria.SelectedValue ' el idcategoria, será el valor seleccionado en el combo box
                Obj.Codigo = TxtCodigo.Text
                Obj.Nombre = TxtNombre.Text ' estos parametros se esperan de la capa de Ncategoria y Dcategoria en la funcion agregar
                Obj.Stock = TxtStock.Text
                Obj.PrecioVenta = TxtPrecioVenta.Text
                Obj.Descripcion = TxtDescripcion.Text
                Obj.Imagen = txtImagen.Text

                If (Neg.Actualizar(Obj)) Then ' Intenta insertar la nueva categoría utilizando la capa de negocio
                    MsgBox("Se ha actualizado correctamente", vbOKOnly + vbInformation, "Actualización correcta") ' Muestra un mensaje de éxito
                    'Verificamos si la ruta de origen sea nueva, si es nueva se actualiza en el directorio
                    If (txtImagen.Text <> "" And RutaOrigen <> "") Then
                        RutaDestino = Directorio & txtImagen.Text ' Concatenamos la ruta de destino -- > c:\sistema\ ..
                        File.Copy(RutaOrigen, RutaDestino) ' que copie la ruta de la imagen  a la ruta de nuestro directoria
                    End If

                    Me.Listar() ' Actualiza la lista de categorías

                Else
                    MsgBox("No se ha podido Actualizar", vbOKOnly + vbCritical, "Actualización Incorrecta") ' Muestra un mensaje de error si no se pudo insertar
                End If
            Else
                MsgBox("Rellene todos los campos obligatorios (*)", vbOKOnly + vbCritical, "Falta Ingresar datos") ' Muestra un mensaje si faltan campos obligatorios
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ' Evento para poder Seleccionar un articulo
    Private Sub DgvListado_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvListado.CellContentClick
        If e.ColumnIndex = DgvListado.Columns.Item("Seleccionar").Index Then ' si el indice de la columna es igual a seleccionar
            ' entonces

            Dim chkcell As DataGridViewCheckBoxCell = DgvListado.Rows(e.RowIndex).Cells("Seleccionar") ' le damos el valor a chk a la columna que seleccionamos
            chkcell.Value = Not chkcell.Value ' capturamos el valor de esa columna que estamos seleccionando

        End If
    End Sub

    Private Sub ChkSeleccionar_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSeleccionar.CheckedChanged
        If ChkSeleccionar.CheckState = CheckState.Checked Then ' Si es estado del Checkbox selecionar es "Checked" Entonces
            DgvListado.Columns.Item("Seleccionar").Visible = True ' Vamos a mostrar la columna seleccionar
            BtnEliminar.Visible = True ' Mostramos el botón de eliminar
            BtnActivar.Visible = True ' Mostramos el botón de activar
            BtnDesactivar.Visible = True ' Mostramos el botón de desactivar
        Else
            DgvListado.Columns.Item("Seleccionar").Visible = False ' Ocultamos la columna seleccionar 
            BtnEliminar.Visible = False ' Ocultamos el  botón de eliminar
            BtnActivar.Visible = False ' Ocultamos el botón de activar
            BtnDesactivar.Visible = False ' Ocultamos el el botón de desactivar

        End If
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        If (MsgBox("¿Estas seguro de que quieres eliminar los registros seleccionados?", vbYesNo + vbQuestion, "Eliminar Registros") = vbYes) Then
            Try
                Dim Neg As New Negocio.NArticulo ' creamos una instancia que haga referencia a la capa Negocio y a la clase NArticulo
                ' para cada fila de DataGridViewRow recorrer todas las filas de mi dgvListado
                For Each row As DataGridViewRow In DgvListado.Rows

                    Dim marcado As Boolean = Convert.ToBoolean(row.Cells("Seleccionar").Value)
                    If marcado Then ' si  marcado es true, entonces
                        Dim OneKey As Integer = Convert.ToInt32(row.Cells("ID").Value) ' Creamos una variable para convertir toda la fila a un int 
                        Dim Imagen As String = Convert.ToString(row.Cells("Imagen").Value) ' Creamos una variable que convierta la fila de Imagen en String
                        File.Delete(Directorio & Imagen)
                        Neg.Eliminar(OneKey)
                    End If
                Next
                Me.Listar()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub BtnActivar_Click(sender As Object, e As EventArgs) Handles BtnActivar.Click
        If (MsgBox("¿Estas seguro de que quieres activar los registros seleccionados?", vbYesNo + vbQuestion, "Activar registros") = vbYes) Then
            Try
                Dim Neg As New Negocio.NArticulo ' creamos una instancia que haga referencia a la capa Negocio y a la clase NCategoria
                ' para cada fila de DataGridViewRow recorrer todas las filas de mi dgvListado
                For Each row As DataGridViewRow In DgvListado.Rows

                    Dim marcado As Boolean = Convert.ToBoolean(row.Cells("Seleccionar").Value)
                    If marcado Then ' si  marcado es true, entonces
                        Dim OneKey As Integer = Convert.ToInt32(row.Cells("ID").Value) ' Creamos una variable para convertir toda la fila a un int 
                        Neg.Activar(OneKey)
                    End If
                Next
                Me.Listar()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub BtnDesactivar_Click(sender As Object, e As EventArgs) Handles BtnDesactivar.Click
        If (MsgBox("¿Estas seguro de que quieres desactivar los registros seleccionados?", vbYesNo + vbQuestion, "Desactivar registros") = vbYes) Then
            Try
                Dim Neg As New Negocio.NArticulo ' creamos una instancia que haga referencia a la capa Negocio y a la clase NCategoria
                ' para cada fila de DataGridViewRow recorrer todas las filas de mi dgvListado
                For Each row As DataGridViewRow In DgvListado.Rows

                    Dim marcado As Boolean = Convert.ToBoolean(row.Cells("Seleccionar").Value)
                    If marcado Then ' si  marcado es true, entonces
                        Dim OneKey As Integer = Convert.ToInt32(row.Cells("ID").Value) ' Creamos una variable para convertir toda la fila a un int 
                        Neg.Desactivar(OneKey)
                    End If
                Next
                Me.Listar()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
End Class