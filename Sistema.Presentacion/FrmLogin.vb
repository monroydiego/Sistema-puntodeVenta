Imports Sistema.Entidades

Public Class FrmLogin
    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        End ' le indicamos que se finalice toda la aplicacion
    End Sub

    Private Sub BtnAcceder_Click(sender As Object, e As EventArgs) Handles BtnAcceder.Click
        Try
            Dim Email, Clave As String
            Dim Obj As New Entidades.Usuario
            Dim Neg As New Negocio.NUsuario

            Email = TxtEmail.Text.Trim() ' se elimnan los espacios en blanco
            Clave = TxtClave.Text.Trim()

            Obj = Neg.Login(Email, Clave) ' le enviamos los parametros al metodo login

            If (Obj Is Nothing) Then
                MsgBox("No existe un usuario con ese correo y/o contraseña", vbOKOnly + vbCritical, "Datos incorrectos") ' mensaje de error

            Else
                If (Obj.Estado = False) Then
                    MsgBox("El usuario no se encuentra activo", vbOKOnly + vbCritical, "Usuario inactivo")
                Else
                    Me.Hide() ' Ocultamos el formulario de login
                    MDIParent1.IdUsuario = Obj.IdUsuario ' Obtenemos los datos del usuario logueado
                    MDIParent1.IdRol = Obj.IdRol
                    MDIParent1.Rol = Obj.Rol
                    MDIParent1.Nombre = Obj.Nombre
                    MDIParent1.Show() ' mostramos el formulario principal
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class