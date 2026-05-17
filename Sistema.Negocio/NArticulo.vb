Imports Sistema.Datos
Imports Sistema.Entidades

Public Class NArticulo
    Public Function Listar() As DataTable
        Try
            Dim Datos As New DArticulo ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.Listar() ' Llamamos el metodo listar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try

    End Function

    Public Function Buscar(Valor As String) As DataTable
        Try
            Dim Datos As New DArticulo ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.Buscar(Valor) ' Llamamos el metodo Buscar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function BuscarCodigo(Valor As String) As Articulo
        Try
            Dim Datos As New DArticulo ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Dim Art As New Articulo
            Tabla = Datos.BuscarCodigo(Valor) ' Llamamos el metodo Buscar de la clase de la capa datos
            If (Tabla.Rows.Count > 0) Then
                Art.IdArticulo = Tabla.Rows(0).Item(0).ToString
                Art.Codigo = Tabla.Rows(0).Item(1).ToString
                Art.Nombre = Tabla.Rows(0).Item(2).ToString
                Art.PrecioVenta = Tabla.Rows(0).Item(3).ToString
                Art.Stock = Tabla.Rows(0).Item(4).ToString

                Return Art
            Else
                Return Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function Insertar(Obj As Articulo) As Boolean ' Si logramos insertar devolvemos un True o sino un False 
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DArticulo
            Datos.Insertar(Obj) ' Llamamos el metodo Insertar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function Actualizar(Obj As Articulo) As Boolean ' Si logramos actualizar devolvemos un True o sino un False 
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DArticulo
            Datos.Actualizar(Obj) ' Llamamos el metodo Actualizar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function Eliminar(Id As Integer) As Boolean
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DArticulo
            Datos.Eliminar(Id) ' Llamamos el metodo Eliminar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function Desactivar(Id As Integer) As Boolean
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DArticulo
            Datos.Desactivar(Id) ' Llamamos el metodo Desativar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function Activar(Id As Integer) As Boolean
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DArticulo
            Datos.Activar(Id) ' Llamamos el metodo Activar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
End Class
