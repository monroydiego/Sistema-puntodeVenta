Imports Sistema.Datos
Imports Sistema.Entidades

Public Class NPersona
    Public Function Listar() As DataTable
        Try
            Dim Datos As New DPersona ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.Listar() ' Llamamos el metodo listar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try

    End Function
    Public Function ListarProveedores() As DataTable
        Try
            Dim Datos As New DPersona ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.ListarProveedores() ' Llamamos el metodo listar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function ListarClientes() As DataTable
        Try
            Dim Datos As New DPersona ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.ListarClientes() ' Llamamos el metodo listar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try

    End Function

    Public Function Buscar(Valor As String) As DataTable
        Try
            Dim Datos As New DPersona ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.Buscar(Valor) ' Llamamos el metodo Buscar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function BuscarProveedores(Valor As String) As DataTable
        Try
            Dim Datos As New DPersona ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.BuscarProveedores(Valor) ' Llamamos el metodo Buscar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function BuscarClientes(Valor As String) As DataTable
        Try
            Dim Datos As New DPersona ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.BuscarClientes(Valor) ' Llamamos el metodo Buscar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function


    Public Function Insertar(Obj As Persona) As Boolean ' Si logramos insertar devolvemos un True o sino un False 
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DPersona
            Datos.Insertar(Obj) ' Llamamos el metodo Insertar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function Actualizar(Obj As Persona) As Boolean ' Si logramos actualizar devolvemos un True o sino un False 
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DPersona
            Datos.Actualizar(Obj) ' Llamamos el metodo Actualizar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function Eliminar(Id As Integer) As Boolean
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DPersona
            Datos.Eliminar(Id) ' Llamamos el metodo Eliminar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
End Class
