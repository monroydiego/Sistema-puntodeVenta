Imports Sistema.Datos
Imports Sistema.Entidades

Public Class NIngreso
    Public Function Listar() As DataTable
        Try
            Dim Datos As New DIngreso ' Creamos una instancia de la clase DCategoria 
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
            Dim Datos As New DIngreso ' Creamos una instancia de la clase DCategoria 
            Dim Tabla As New DataTable ' Creamos una tabla en memoria para almacenar los datos de la consulta
            Tabla = Datos.Buscar(Valor) ' Llamamos el metodo Buscar de la clase de la capa datos
            Return Tabla
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function ConsultarDetallado(FechaInicio As Date, FechaFin As Date) As DataTable
        Try
            If FechaInicio > FechaFin Then
                Throw New Exception("La fecha de inicio no puede ser mayor a la fecha fin.")
            End If
            Dim Datos As New DIngreso
            Return Datos.ConsultarDetallado(FechaInicio, FechaFin)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function Insertar(Obj As Ingreso, Det As DataTable) As Boolean ' Si logramos insertar devolvemos un True o sino un False
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DIngreso
            Datos.Insertar(Obj, Det) ' Llamamos el metodo Insertar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function Anular(Id As Integer) As Boolean
        Try ' Si logramos Insertar devolvemos un True 
            Dim Datos As New DIngreso
            Datos.Anular(Id) ' Llamamos el metodo Desativar de la clase de la capa datos
            Return True
        Catch ex As Exception ' Si no logramos insertar Devolvemos un False
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

End Class
