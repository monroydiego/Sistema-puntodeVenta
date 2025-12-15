Imports Sistema.Datos

Public Class NRol
    Public Function Listar() As DataTable
        Try
            Dim Datos As New DRol ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.Listar() ' Llamamos el metodo listar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try

    End Function
End Class
