<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmArticulo
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ErrorIcono = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.BtnActualizar = New System.Windows.Forms.Button()
        Me.TxtId = New System.Windows.Forms.TextBox()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.BtnInsertar = New System.Windows.Forms.Button()
        Me.TxtDescripcion = New System.Windows.Forms.TextBox()
        Me.TxtNombre = New System.Windows.Forms.TextBox()
        Me.LblDescripcion = New System.Windows.Forms.Label()
        Me.LblNombre = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TxtPrecioVenta = New System.Windows.Forms.TextBox()
        Me.TxtStock = New System.Windows.Forms.TextBox()
        Me.TxtCodigo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PicImagen = New System.Windows.Forms.PictureBox()
        Me.BtnCargarImagen = New System.Windows.Forms.Button()
        Me.txtImagen = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CboCategoria = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Seleccionar = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.BtnDesactivar = New System.Windows.Forms.Button()
        Me.BtnActivar = New System.Windows.Forms.Button()
        Me.BtnEliminar = New System.Windows.Forms.Button()
        Me.ChkSeleccionar = New System.Windows.Forms.CheckBox()
        Me.BtnBuscar = New System.Windows.Forms.Button()
        Me.TxtValor = New System.Windows.Forms.TextBox()
        Me.LblTotal = New System.Windows.Forms.Label()
        Me.DgvListado = New System.Windows.Forms.DataGridView()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.TabGeneral = New System.Windows.Forms.TabControl()
        CType(Me.ErrorIcono, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.PicImagen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvListado, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        Me.TabGeneral.SuspendLayout()
        Me.SuspendLayout()
        '
        'ErrorIcono
        '
        Me.ErrorIcono.ContainerControl = Me
        '
        'BtnActualizar
        '
        Me.BtnActualizar.Location = New System.Drawing.Point(122, 394)
        Me.BtnActualizar.Name = "BtnActualizar"
        Me.BtnActualizar.Size = New System.Drawing.Size(129, 23)
        Me.BtnActualizar.TabIndex = 7
        Me.BtnActualizar.Text = "Actualizar"
        Me.BtnActualizar.UseVisualStyleBackColor = True
        '
        'TxtId
        '
        Me.TxtId.Location = New System.Drawing.Point(303, 19)
        Me.TxtId.Name = "TxtId"
        Me.TxtId.Size = New System.Drawing.Size(100, 20)
        Me.TxtId.TabIndex = 6
        Me.TxtId.Visible = False
        '
        'BtnCancelar
        '
        Me.BtnCancelar.Location = New System.Drawing.Point(274, 393)
        Me.BtnCancelar.Name = "BtnCancelar"
        Me.BtnCancelar.Size = New System.Drawing.Size(129, 23)
        Me.BtnCancelar.TabIndex = 5
        Me.BtnCancelar.Text = "Cancelar"
        Me.BtnCancelar.UseVisualStyleBackColor = True
        '
        'BtnInsertar
        '
        Me.BtnInsertar.Location = New System.Drawing.Point(122, 393)
        Me.BtnInsertar.Name = "BtnInsertar"
        Me.BtnInsertar.Size = New System.Drawing.Size(129, 23)
        Me.BtnInsertar.TabIndex = 4
        Me.BtnInsertar.Text = "Insertar"
        Me.BtnInsertar.UseVisualStyleBackColor = True
        '
        'TxtDescripcion
        '
        Me.TxtDescripcion.Location = New System.Drawing.Point(122, 304)
        Me.TxtDescripcion.Multiline = True
        Me.TxtDescripcion.Name = "TxtDescripcion"
        Me.TxtDescripcion.Size = New System.Drawing.Size(281, 60)
        Me.TxtDescripcion.TabIndex = 3
        '
        'TxtNombre
        '
        Me.TxtNombre.Location = New System.Drawing.Point(122, 167)
        Me.TxtNombre.Name = "TxtNombre"
        Me.TxtNombre.Size = New System.Drawing.Size(281, 20)
        Me.TxtNombre.TabIndex = 2
        '
        'LblDescripcion
        '
        Me.LblDescripcion.AutoSize = True
        Me.LblDescripcion.Location = New System.Drawing.Point(35, 304)
        Me.LblDescripcion.Name = "LblDescripcion"
        Me.LblDescripcion.Size = New System.Drawing.Size(63, 13)
        Me.LblDescripcion.TabIndex = 1
        Me.LblDescripcion.Text = "Descripción"
        '
        'LblNombre
        '
        Me.LblNombre.AutoSize = True
        Me.LblNombre.Location = New System.Drawing.Point(35, 170)
        Me.LblNombre.Name = "LblNombre"
        Me.LblNombre.Size = New System.Drawing.Size(57, 13)
        Me.LblNombre.TabIndex = 0
        Me.LblNombre.Text = "Nombre (*)"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.TxtPrecioVenta)
        Me.TabPage2.Controls.Add(Me.TxtStock)
        Me.TabPage2.Controls.Add(Me.TxtCodigo)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Controls.Add(Me.PicImagen)
        Me.TabPage2.Controls.Add(Me.BtnCargarImagen)
        Me.TabPage2.Controls.Add(Me.txtImagen)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.CboCategoria)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.BtnActualizar)
        Me.TabPage2.Controls.Add(Me.TxtId)
        Me.TabPage2.Controls.Add(Me.BtnCancelar)
        Me.TabPage2.Controls.Add(Me.BtnInsertar)
        Me.TabPage2.Controls.Add(Me.TxtDescripcion)
        Me.TabPage2.Controls.Add(Me.TxtNombre)
        Me.TabPage2.Controls.Add(Me.LblDescripcion)
        Me.TabPage2.Controls.Add(Me.LblNombre)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1135, 422)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Mantenimiento"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(35, 255)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Precio venta (*)"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(35, 212)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 13)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Stock (*)"
        '
        'TxtPrecioVenta
        '
        Me.TxtPrecioVenta.Location = New System.Drawing.Point(122, 252)
        Me.TxtPrecioVenta.Name = "TxtPrecioVenta"
        Me.TxtPrecioVenta.Size = New System.Drawing.Size(144, 20)
        Me.TxtPrecioVenta.TabIndex = 17
        '
        'TxtStock
        '
        Me.TxtStock.Location = New System.Drawing.Point(122, 205)
        Me.TxtStock.Name = "TxtStock"
        Me.TxtStock.Size = New System.Drawing.Size(144, 20)
        Me.TxtStock.TabIndex = 16
        '
        'TxtCodigo
        '
        Me.TxtCodigo.Location = New System.Drawing.Point(122, 124)
        Me.TxtCodigo.Name = "TxtCodigo"
        Me.TxtCodigo.Size = New System.Drawing.Size(280, 20)
        Me.TxtCodigo.TabIndex = 15
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(35, 124)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Código"
        '
        'PicImagen
        '
        Me.PicImagen.Location = New System.Drawing.Point(626, 124)
        Me.PicImagen.Name = "PicImagen"
        Me.PicImagen.Size = New System.Drawing.Size(365, 240)
        Me.PicImagen.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicImagen.TabIndex = 13
        Me.PicImagen.TabStop = False
        '
        'BtnCargarImagen
        '
        Me.BtnCargarImagen.Location = New System.Drawing.Point(916, 71)
        Me.BtnCargarImagen.Name = "BtnCargarImagen"
        Me.BtnCargarImagen.Size = New System.Drawing.Size(75, 23)
        Me.BtnCargarImagen.TabIndex = 12
        Me.BtnCargarImagen.Text = "..."
        Me.BtnCargarImagen.UseVisualStyleBackColor = True
        '
        'txtImagen
        '
        Me.txtImagen.Enabled = False
        Me.txtImagen.Location = New System.Drawing.Point(626, 73)
        Me.txtImagen.Multiline = True
        Me.txtImagen.Name = "txtImagen"
        Me.txtImagen.Size = New System.Drawing.Size(274, 20)
        Me.txtImagen.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(548, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Imagen"
        '
        'CboCategoria
        '
        Me.CboCategoria.FormattingEnabled = True
        Me.CboCategoria.Location = New System.Drawing.Point(122, 76)
        Me.CboCategoria.Name = "CboCategoria"
        Me.CboCategoria.Size = New System.Drawing.Size(280, 21)
        Me.CboCategoria.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(35, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Categoria(*)"
        '
        'Seleccionar
        '
        Me.Seleccionar.HeaderText = "Seleccionar"
        Me.Seleccionar.Name = "Seleccionar"
        Me.Seleccionar.ReadOnly = True
        '
        'BtnDesactivar
        '
        Me.BtnDesactivar.Location = New System.Drawing.Point(375, 395)
        Me.BtnDesactivar.Name = "BtnDesactivar"
        Me.BtnDesactivar.Size = New System.Drawing.Size(120, 23)
        Me.BtnDesactivar.TabIndex = 8
        Me.BtnDesactivar.Text = "Desactivar"
        Me.BtnDesactivar.UseVisualStyleBackColor = True
        '
        'BtnActivar
        '
        Me.BtnActivar.Location = New System.Drawing.Point(249, 395)
        Me.BtnActivar.Name = "BtnActivar"
        Me.BtnActivar.Size = New System.Drawing.Size(120, 23)
        Me.BtnActivar.TabIndex = 7
        Me.BtnActivar.Text = "Activar"
        Me.BtnActivar.UseVisualStyleBackColor = True
        '
        'BtnEliminar
        '
        Me.BtnEliminar.Location = New System.Drawing.Point(114, 395)
        Me.BtnEliminar.Name = "BtnEliminar"
        Me.BtnEliminar.Size = New System.Drawing.Size(120, 23)
        Me.BtnEliminar.TabIndex = 6
        Me.BtnEliminar.Text = "Eliminar"
        Me.BtnEliminar.UseVisualStyleBackColor = True
        '
        'ChkSeleccionar
        '
        Me.ChkSeleccionar.AutoSize = True
        Me.ChkSeleccionar.Location = New System.Drawing.Point(9, 399)
        Me.ChkSeleccionar.Name = "ChkSeleccionar"
        Me.ChkSeleccionar.Size = New System.Drawing.Size(82, 17)
        Me.ChkSeleccionar.TabIndex = 5
        Me.ChkSeleccionar.Text = "Seleccionar"
        Me.ChkSeleccionar.UseVisualStyleBackColor = True
        '
        'BtnBuscar
        '
        Me.BtnBuscar.Location = New System.Drawing.Point(581, 5)
        Me.BtnBuscar.Name = "BtnBuscar"
        Me.BtnBuscar.Size = New System.Drawing.Size(204, 23)
        Me.BtnBuscar.TabIndex = 4
        Me.BtnBuscar.Text = "Buscar"
        Me.BtnBuscar.UseVisualStyleBackColor = True
        '
        'TxtValor
        '
        Me.TxtValor.Location = New System.Drawing.Point(6, 7)
        Me.TxtValor.Name = "TxtValor"
        Me.TxtValor.Size = New System.Drawing.Size(552, 20)
        Me.TxtValor.TabIndex = 3
        '
        'LblTotal
        '
        Me.LblTotal.AutoSize = True
        Me.LblTotal.Location = New System.Drawing.Point(1042, 400)
        Me.LblTotal.Name = "LblTotal"
        Me.LblTotal.Size = New System.Drawing.Size(31, 13)
        Me.LblTotal.TabIndex = 2
        Me.LblTotal.Text = "Total"
        '
        'DgvListado
        '
        Me.DgvListado.AllowUserToAddRows = False
        Me.DgvListado.AllowUserToDeleteRows = False
        Me.DgvListado.AllowUserToOrderColumns = True
        Me.DgvListado.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvListado.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Seleccionar})
        Me.DgvListado.Location = New System.Drawing.Point(6, 37)
        Me.DgvListado.Name = "DgvListado"
        Me.DgvListado.ReadOnly = True
        Me.DgvListado.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvListado.Size = New System.Drawing.Size(1123, 355)
        Me.DgvListado.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.BtnDesactivar)
        Me.TabPage1.Controls.Add(Me.BtnActivar)
        Me.TabPage1.Controls.Add(Me.BtnEliminar)
        Me.TabPage1.Controls.Add(Me.ChkSeleccionar)
        Me.TabPage1.Controls.Add(Me.BtnBuscar)
        Me.TabPage1.Controls.Add(Me.TxtValor)
        Me.TabPage1.Controls.Add(Me.LblTotal)
        Me.TabPage1.Controls.Add(Me.DgvListado)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1135, 422)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Listado"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'TabGeneral
        '
        Me.TabGeneral.Controls.Add(Me.TabPage1)
        Me.TabGeneral.Controls.Add(Me.TabPage2)
        Me.TabGeneral.Location = New System.Drawing.Point(3, 12)
        Me.TabGeneral.Name = "TabGeneral"
        Me.TabGeneral.SelectedIndex = 0
        Me.TabGeneral.Size = New System.Drawing.Size(1143, 448)
        Me.TabGeneral.TabIndex = 1
        '
        'FrmArticulo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1149, 463)
        Me.Controls.Add(Me.TabGeneral)
        Me.Name = "FrmArticulo"
        Me.Text = "Articulo"
        CType(Me.ErrorIcono, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.PicImagen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvListado, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabGeneral.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ErrorIcono As ErrorProvider
    Friend WithEvents TabGeneral As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents BtnDesactivar As Button
    Friend WithEvents BtnActivar As Button
    Friend WithEvents BtnEliminar As Button
    Friend WithEvents ChkSeleccionar As CheckBox
    Friend WithEvents BtnBuscar As Button
    Friend WithEvents TxtValor As TextBox
    Friend WithEvents LblTotal As Label
    Friend WithEvents DgvListado As DataGridView
    Friend WithEvents Seleccionar As DataGridViewCheckBoxColumn
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents BtnActualizar As Button
    Friend WithEvents TxtId As TextBox
    Friend WithEvents BtnCancelar As Button
    Friend WithEvents BtnInsertar As Button
    Friend WithEvents TxtDescripcion As TextBox
    Friend WithEvents TxtNombre As TextBox
    Friend WithEvents LblDescripcion As Label
    Friend WithEvents LblNombre As Label
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents CboCategoria As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents BtnCargarImagen As Button
    Friend WithEvents txtImagen As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents PicImagen As PictureBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TxtPrecioVenta As TextBox
    Friend WithEvents TxtStock As TextBox
    Friend WithEvents TxtCodigo As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label5 As Label
End Class
