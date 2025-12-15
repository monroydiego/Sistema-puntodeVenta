Public Class FrmRol
    Private Sub Formato()
        ' se agregan 4 columnas propias del Stored Procedure 
        ' Agregaremos 1 columna del Listado (seleccionar) 
        ' el parametro de columns es para saber cuantas columnas vamos agregar 
        DgvListado.Columns(0).Width = 100 ' establecemos el ancho  de la columna 0 
        DgvListado.Columns(1).Width = 150 ' id
        DgvListado.Columns(0).HeaderText = "ID"
        DgvListado.Columns(1).HeaderText = "Nombre"

    End Sub

    ' Metodo buscar
    Private Sub Listar()
        Try
            'Creamos una instancia que haga referencia a la capa negocio
            Dim Neg As New Negocio.NRol ' Objeto que hace una instania a la clase Ncategoria de la capa negocio 
            DgvListado.DataSource = Neg.Listar() ' obtener el listado de categorias desde la capa negocio 
            LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString() ' Muestra el total de registros en el Label
            Me.Formato() ' Llamamos al metodo Formato para establecer el formato de las columnas
        Catch ex As Exception
            MsgBox(ex.Message) ' Si ocurre un error, muestra un mensaje con la descripción del error
        End Try
    End Sub
    Private Sub FrmRol_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Listar()
    End Sub
End Class