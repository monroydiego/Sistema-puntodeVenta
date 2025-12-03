Public Class FrmCategoria
    ' de la capa presentacion se va hacer referencia de la capa negocios, funcion listar y funcion listar hace referencia a la capa datos
    ' Este metodo nos permite establecer el formato de las columnas que van aparecer en el listado 
    Private Sub Formato()
        ' se agregan 4 columnas propias del Stored Procedure 
        ' Agregaremos 1 columna del Listado (seleccionar) 
        ' el parametro de columns es para saber cuantas columnas vamos agregar 
        DgvListado.Columns(0).Visible = False ' Ocultamos la columna 0 que es el ID de la categoria, se inicia con una columna 0 
        DgvListado.Columns(0).Width = 100 ' establecemos el ancho  de la columna 0 
        DgvListado.Columns(1).Width = 100 ' id
        DgvListado.Columns(2).Width = 150 ' nombre 
        DgvListado.Columns(3).Width = 400 ' descripcion
        DgvListado.Columns(4).Width = 100 ' estado


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
            Dim Neg As New Negocio.NCategoria ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
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
            Dim Neg As New Negocio.NCategoria ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
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
        TxtNombre.Text = "" ' Limpia el campo de texto del nombre
        TxtDescripcion.Text = "" ' Limpia el campo de texto de la descripción
    End Sub
    Private Sub FrmCategoria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Listar() ' Llamamos al metodo listar cuando se carga el formulario 
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Me.Buscar() ' llamamos el metodo de buscar 
    End Sub

    Private Sub BtnInsertar_Click(sender As Object, e As EventArgs) Handles BtnInsertar.Click
        If Me.ValidateChildren = True And TxtNombre.Text <> "" Then ' Verifica si la validación de los controles hijos es exitosa y si el campo de nombre no está vacío
            Dim Obj As New Entidades.Categoria 'Declaramos un objeto en donde se haga la referencia a entidades, entidad hace referencia a categoria 
            Dim Neg As New Negocio.NCategoria ' se hace referencia a la capa negocio de ncategoria

            Obj.Nombre = TxtNombre.Text ' estos parametros se esperan de la capa de Ncategoria y Dcategoria en la funcion agregar
            Obj.Descripcion = TxtDescripcion.Text

            If (Neg.Insertar(Obj)) Then ' Intenta insertar la nueva categoría utilizando la capa de negocio
                MsgBox("Se ha registrado correctamente", vbOKOnly + vbInformation, "Registro correcto") ' Muestra un mensaje de éxito
                Me.Listar() ' Actualiza la lista de categorías
            Else
                MsgBox("No se ha podido registrar", vbOKOnly + vbCritical, "Registro Incorrecto") ' Muestra un mensaje de error si no se pudo insertar
            End If
        Else
            MsgBox("Rellene todos los campos obligatorios (*)", vbOKOnly + vbCritical, "Falta Ingresar datos") ' Muestra un mensaje si faltan campos obligatorios
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Limpiar() ' Llamamos al metodo limpiar para que limpie todo el formulario
        TabGeneral.SelectedIndex = 0 ' regresamos al listado
    End Sub

    Private Sub TxtNombre_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TxtNombre.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorIcono.SetError(sender, "")
        Else
            Me.ErrorIcono.SetError(sender, "Ingrese el nombre de la categría, este dato es obligatorio")
        End If
    End Sub

    ' Evento para poder actualizar a darle doble click al DgvListado
    Private Sub DgvListado_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvListado.CellDoubleClick
        TxtId.Text = DgvListado.SelectedCells.Item(1).Value ' Celdas que se muestran en el apartado de listar
        TxtNombre.Text = DgvListado.SelectedCells.Item(2).Value
        TxtDescripcion.Text = DgvListado.SelectedCells.Item(3).Value
        BtnInsertar.Visible = False ' al momento que le demos doble click,  el botón de insertar se ocultará y se mostrará el de actualizar
        BtnActualizar.Visible = True
        TabGeneral.SelectedIndex = 1
    End Sub

    Private Sub BtnActualizar_Click(sender As Object, e As EventArgs) Handles BtnActualizar.Click
        If Me.ValidateChildren = True And TxtNombre.Text <> "" And TxtId.Text <> "" Then ' Verifica si la validación de los controles hijos es exitosa y si el campo de nombre no está vacío
            Dim Obj As New Entidades.Categoria 'Declaramos un objeto en donde se haga la referencia a entidades, entidad hace referencia a categoria 
            Dim Neg As New Negocio.NCategoria ' se hace referencia a la capa negocio de ncategoria

            Obj.IdCategoria = TxtId.Text
            Obj.Nombre = TxtNombre.Text ' estos parametros se esperan de la capa de Ncategoria y Dcategoria en la funcion agregar
            Obj.Descripcion = TxtDescripcion.Text

            If (Neg.Actualizar(Obj)) Then ' Intenta insertar la nueva categoría utilizando la capa de negocio
                MsgBox("Se ha actualizado correctamente", vbOKOnly + vbInformation, "Actualización correcta") ' Muestra un mensaje de éxito

                Me.Listar() ' Actualiza la lista de categorías
                TabGeneral.SelectedIndex = 0
            Else
                MsgBox("No se ha podido actualizar", vbOKOnly + vbCritical, "Registro Incorrecto") ' Muestra un mensaje de error si no se pudo insertar
            End If
        Else
            MsgBox("Rellene todos los campos obligatorios (*)", vbOKOnly + vbCritical, "Falta Ingresar datos") ' Muestra un mensaje si faltan campos obligatorios
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

    Private Sub DgvListado_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvListado.CellContentClick
        If e.ColumnIndex = DgvListado.Columns.Item("Seleccionar").Index Then ' si el indice de la columna es igual a seleccionar
            ' entonces

            Dim chkcell As DataGridViewCheckBoxCell = DgvListado.Rows(e.RowIndex).Cells("Seleccionar") ' le damos el valor a chk a la columna que seleccionamos
            chkcell.Value = Not chkcell.Value ' capturamos el valor de esa columna que estamos seleccionando

        End If
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        If (MsgBox("¿Estas seguro de que quieres eliminar los registros seleccionados?", vbYesNo + vbQuestion, "Eliminar Registros") = vbYes) Then
            Try
                Dim Neg As New Negocio.NCategoria ' creamos una instancia que haga referencia a la capa Negocio y a la clase NCategoria
                ' para cada fila de DataGridViewRow recorrer todas las filas de mi dgvListado
                For Each row As DataGridViewRow In DgvListado.Rows

                    Dim marcado As Boolean = Convert.ToBoolean(row.Cells("Seleccionar").Value)
                    If marcado Then ' si  marcado es true, entonces
                        Dim OneKey As Integer = Convert.ToInt32(row.Cells("ID").Value) ' Creamos una variable para convertir toda la fila a un int 
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
                Dim Neg As New Negocio.NCategoria ' creamos una instancia que haga referencia a la capa Negocio y a la clase NCategoria
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
                Dim Neg As New Negocio.NCategoria ' creamos una instancia que haga referencia a la capa Negocio y a la clase NCategoria
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