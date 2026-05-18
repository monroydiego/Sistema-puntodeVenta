Public Class FrmCliente_Venta

    Private Sub Formato()
        DgvListado.Columns(0).Visible = False
        DgvListado.Columns(1).Width = 80
        DgvListado.Columns(2).Width = 120
        DgvListado.Columns(3).Width = 180
        DgvListado.Columns(4).Width = 100
        DgvListado.Columns(5).Width = 120
        DgvListado.Columns(6).Width = 100
        DgvListado.Columns(7).Width = 180
        DgvListado.Columns(8).Width = 120
    End Sub

    Private Sub Listar()
        Try
            Dim Neg As New Negocio.NPersona
            DgvListado.DataSource = Neg.ListarClientes()
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString()
            Me.Formato()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Buscar()
        Try
            Dim Neg As New Negocio.NPersona
            DgvListado.DataSource = Neg.BuscarClientes(TxtValor.Text)
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString()
            Me.Formato()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub FrmCliente_Venta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Listar()
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Me.Buscar()
    End Sub

    Private Sub DgvListado_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvListado.CellDoubleClick
        Try
            ' Validar que se seleccionó una fila válida
            If e.RowIndex < 0 Then
                MsgBox("Seleccione un cliente válido.", vbOKOnly + vbExclamation)
                Return
            End If

            ' CORRECCIÓN: Índice 1 = idpersona, Índice 3 = nombre
            ' (Índice 0 es la columna "Seleccionar" checkbox que está oculta)
            Dim IdSeleccionado = DgvListado.Rows(e.RowIndex).Cells(1).Value
            Dim NombreSeleccionado = DgvListado.Rows(e.RowIndex).Cells(3).Value

            ' Validar que los valores no sean NULL o DBNull
            If IdSeleccionado Is Nothing OrElse IsDBNull(IdSeleccionado) Then
                MsgBox("Error: ID de cliente no válido.", vbOKOnly + vbCritical, "ID Inválido")
                Return
            End If

            If NombreSeleccionado Is Nothing OrElse IsDBNull(NombreSeleccionado) Then
                MsgBox("Error: Nombre de cliente no válido.", vbOKOnly + vbCritical, "Nombre Inválido")
                Return
            End If

            ' Asignar a variables globales (convertir a string)
            Variables.IdCliente = Convert.ToString(IdSeleccionado)
            Variables.NombreCliente = Convert.ToString(NombreSeleccionado)

            ' Confirmación de asignación exitosa
            ' MsgBox("Cliente asignado: " & Variables.NombreCliente, vbOKOnly + vbInformation)

            Me.Close()
        Catch ex As Exception
            MsgBox("Error al seleccionar cliente: " & ex.Message, vbOKOnly + vbCritical, "Error")
        End Try
    End Sub

End Class
