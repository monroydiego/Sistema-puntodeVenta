<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmIngreso
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
        Me.TabGeneral = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.BtnAnular = New System.Windows.Forms.Button()
        Me.ChkSeleccionar = New System.Windows.Forms.CheckBox()
        Me.BtnBuscar = New System.Windows.Forms.Button()
        Me.TxtValor = New System.Windows.Forms.TextBox()
        Me.LblTotal = New System.Windows.Forms.Label()
        Me.DgvListado = New System.Windows.Forms.DataGridView()
        Me.Seleccionar = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.BtnBuscarArticulos = New System.Windows.Forms.Button()
        Me.TxtCodigo = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TxtTotal = New System.Windows.Forms.TextBox()
        Me.TxtTotalImpuesto = New System.Windows.Forms.TextBox()
        Me.TxtSubTotal = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DgvDetalle = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtImpuesto = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtNumComprobante = New System.Windows.Forms.TextBox()
        Me.TxtSerieComprobante = New System.Windows.Forms.TextBox()
        Me.CboTipoComprobante = New System.Windows.Forms.ComboBox()
        Me.BtnBuscarProveedor = New System.Windows.Forms.Button()
        Me.TxtNombreProveedor = New System.Windows.Forms.TextBox()
        Me.TxtIdProveedor = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtId = New System.Windows.Forms.TextBox()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.BtnInsertar = New System.Windows.Forms.Button()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ErrorIcono = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.PanelArticulos = New System.Windows.Forms.Panel()
        Me.TxtBuscarArticulos = New System.Windows.Forms.TextBox()
        Me.BtnBuscarArticulosDetalle = New System.Windows.Forms.Button()
        Me.DgvArticulos = New System.Windows.Forms.DataGridView()
        Me.LblTotalArticulos = New System.Windows.Forms.Label()
        Me.BtnCerrar = New System.Windows.Forms.Button()
        Me.TabGeneral.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvListado, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ErrorIcono, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelArticulos.SuspendLayout()
        CType(Me.DgvArticulos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabGeneral
        '
        Me.TabGeneral.Controls.Add(Me.TabPage1)
        Me.TabGeneral.Controls.Add(Me.TabPage2)
        Me.TabGeneral.Location = New System.Drawing.Point(12, 12)
        Me.TabGeneral.Name = "TabGeneral"
        Me.TabGeneral.SelectedIndex = 0
        Me.TabGeneral.Size = New System.Drawing.Size(1166, 617)
        Me.TabGeneral.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.BtnAnular)
        Me.TabPage1.Controls.Add(Me.ChkSeleccionar)
        Me.TabPage1.Controls.Add(Me.BtnBuscar)
        Me.TabPage1.Controls.Add(Me.TxtValor)
        Me.TabPage1.Controls.Add(Me.LblTotal)
        Me.TabPage1.Controls.Add(Me.DgvListado)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1158, 591)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Listado"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'BtnAnular
        '
        Me.BtnAnular.Location = New System.Drawing.Point(124, 395)
        Me.BtnAnular.Name = "BtnAnular"
        Me.BtnAnular.Size = New System.Drawing.Size(120, 23)
        Me.BtnAnular.TabIndex = 8
        Me.BtnAnular.Text = "Anular"
        Me.BtnAnular.UseVisualStyleBackColor = True
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
        'Seleccionar
        '
        Me.Seleccionar.HeaderText = "Seleccionar"
        Me.Seleccionar.Name = "Seleccionar"
        Me.Seleccionar.ReadOnly = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Controls.Add(Me.TxtId)
        Me.TabPage2.Controls.Add(Me.BtnCancelar)
        Me.TabPage2.Controls.Add(Me.BtnInsertar)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1158, 591)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Mantenimiento"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.PanelArticulos)
        Me.GroupBox2.Controls.Add(Me.BtnBuscarArticulos)
        Me.GroupBox2.Controls.Add(Me.TxtCodigo)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.TxtTotal)
        Me.GroupBox2.Controls.Add(Me.TxtTotalImpuesto)
        Me.GroupBox2.Controls.Add(Me.TxtSubTotal)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.DgvDetalle)
        Me.GroupBox2.Location = New System.Drawing.Point(26, 228)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1085, 328)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Detalles"
        '
        'BtnBuscarArticulos
        '
        Me.BtnBuscarArticulos.Location = New System.Drawing.Point(664, 28)
        Me.BtnBuscarArticulos.Name = "BtnBuscarArticulos"
        Me.BtnBuscarArticulos.Size = New System.Drawing.Size(75, 23)
        Me.BtnBuscarArticulos.TabIndex = 9
        Me.BtnBuscarArticulos.Text = "Buscar"
        Me.BtnBuscarArticulos.UseVisualStyleBackColor = True
        '
        'TxtCodigo
        '
        Me.TxtCodigo.Location = New System.Drawing.Point(108, 28)
        Me.TxtCodigo.Name = "TxtCodigo"
        Me.TxtCodigo.Size = New System.Drawing.Size(528, 20)
        Me.TxtCodigo.TabIndex = 8
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(21, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(77, 13)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Código articulo"
        '
        'TxtTotal
        '
        Me.TxtTotal.Location = New System.Drawing.Point(968, 299)
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(100, 20)
        Me.TxtTotal.TabIndex = 6
        '
        'TxtTotalImpuesto
        '
        Me.TxtTotalImpuesto.Location = New System.Drawing.Point(968, 268)
        Me.TxtTotalImpuesto.Name = "TxtTotalImpuesto"
        Me.TxtTotalImpuesto.ReadOnly = True
        Me.TxtTotalImpuesto.Size = New System.Drawing.Size(100, 20)
        Me.TxtTotalImpuesto.TabIndex = 5
        '
        'TxtSubTotal
        '
        Me.TxtSubTotal.Location = New System.Drawing.Point(968, 239)
        Me.TxtSubTotal.Name = "TxtSubTotal"
        Me.TxtSubTotal.ReadOnly = True
        Me.TxtSubTotal.Size = New System.Drawing.Size(100, 20)
        Me.TxtSubTotal.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(888, 299)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(31, 13)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "Total"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(888, 268)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Total Impuesto"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(888, 239)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "SubTotal"
        '
        'DgvDetalle
        '
        Me.DgvDetalle.AllowUserToAddRows = False
        Me.DgvDetalle.AllowUserToOrderColumns = True
        Me.DgvDetalle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvDetalle.Location = New System.Drawing.Point(13, 67)
        Me.DgvDetalle.Name = "DgvDetalle"
        Me.DgvDetalle.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvDetalle.Size = New System.Drawing.Size(1055, 166)
        Me.DgvDetalle.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtImpuesto)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.TxtNumComprobante)
        Me.GroupBox1.Controls.Add(Me.TxtSerieComprobante)
        Me.GroupBox1.Controls.Add(Me.CboTipoComprobante)
        Me.GroupBox1.Controls.Add(Me.BtnBuscarProveedor)
        Me.GroupBox1.Controls.Add(Me.TxtNombreProveedor)
        Me.GroupBox1.Controls.Add(Me.TxtIdProveedor)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(26, 45)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1085, 142)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Cabecera"
        '
        'TxtImpuesto
        '
        Me.TxtImpuesto.Location = New System.Drawing.Point(945, 65)
        Me.TxtImpuesto.Name = "TxtImpuesto"
        Me.TxtImpuesto.ReadOnly = True
        Me.TxtImpuesto.Size = New System.Drawing.Size(100, 20)
        Me.TxtImpuesto.TabIndex = 23
        Me.TxtImpuesto.Text = "0.16"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(868, 65)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Impuesto"
        '
        'TxtNumComprobante
        '
        Me.TxtNumComprobante.Location = New System.Drawing.Point(467, 84)
        Me.TxtNumComprobante.Name = "TxtNumComprobante"
        Me.TxtNumComprobante.Size = New System.Drawing.Size(136, 20)
        Me.TxtNumComprobante.TabIndex = 21
        '
        'TxtSerieComprobante
        '
        Me.TxtSerieComprobante.Location = New System.Drawing.Point(300, 84)
        Me.TxtSerieComprobante.Name = "TxtSerieComprobante"
        Me.TxtSerieComprobante.Size = New System.Drawing.Size(136, 20)
        Me.TxtSerieComprobante.TabIndex = 20
        '
        'CboTipoComprobante
        '
        Me.CboTipoComprobante.FormattingEnabled = True
        Me.CboTipoComprobante.Items.AddRange(New Object() {"Factura", "Boleta", "Ticket"})
        Me.CboTipoComprobante.Location = New System.Drawing.Point(108, 84)
        Me.CboTipoComprobante.Name = "CboTipoComprobante"
        Me.CboTipoComprobante.Size = New System.Drawing.Size(161, 21)
        Me.CboTipoComprobante.TabIndex = 19
        '
        'BtnBuscarProveedor
        '
        Me.BtnBuscarProveedor.Location = New System.Drawing.Point(620, 39)
        Me.BtnBuscarProveedor.Name = "BtnBuscarProveedor"
        Me.BtnBuscarProveedor.Size = New System.Drawing.Size(75, 23)
        Me.BtnBuscarProveedor.TabIndex = 18
        Me.BtnBuscarProveedor.Text = "..."
        Me.BtnBuscarProveedor.UseVisualStyleBackColor = True
        '
        'TxtNombreProveedor
        '
        Me.TxtNombreProveedor.Location = New System.Drawing.Point(214, 39)
        Me.TxtNombreProveedor.Name = "TxtNombreProveedor"
        Me.TxtNombreProveedor.ReadOnly = True
        Me.TxtNombreProveedor.Size = New System.Drawing.Size(389, 20)
        Me.TxtNombreProveedor.TabIndex = 17
        '
        'TxtIdProveedor
        '
        Me.TxtIdProveedor.Location = New System.Drawing.Point(108, 39)
        Me.TxtIdProveedor.Name = "TxtIdProveedor"
        Me.TxtIdProveedor.ReadOnly = True
        Me.TxtIdProveedor.Size = New System.Drawing.Size(100, 20)
        Me.TxtIdProveedor.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Proveedor(*)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(21, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Comprobante"
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
        Me.BtnCancelar.Location = New System.Drawing.Point(197, 562)
        Me.BtnCancelar.Name = "BtnCancelar"
        Me.BtnCancelar.Size = New System.Drawing.Size(129, 23)
        Me.BtnCancelar.TabIndex = 5
        Me.BtnCancelar.Text = "Cancelar"
        Me.BtnCancelar.UseVisualStyleBackColor = True
        '
        'BtnInsertar
        '
        Me.BtnInsertar.Location = New System.Drawing.Point(39, 562)
        Me.BtnInsertar.Name = "BtnInsertar"
        Me.BtnInsertar.Size = New System.Drawing.Size(129, 23)
        Me.BtnInsertar.TabIndex = 4
        Me.BtnInsertar.Text = "Insertar"
        Me.BtnInsertar.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'ErrorIcono
        '
        Me.ErrorIcono.ContainerControl = Me
        '
        'PanelArticulos
        '
        Me.PanelArticulos.BackColor = System.Drawing.SystemColors.Info
        Me.PanelArticulos.Controls.Add(Me.BtnCerrar)
        Me.PanelArticulos.Controls.Add(Me.LblTotalArticulos)
        Me.PanelArticulos.Controls.Add(Me.DgvArticulos)
        Me.PanelArticulos.Controls.Add(Me.BtnBuscarArticulosDetalle)
        Me.PanelArticulos.Controls.Add(Me.TxtBuscarArticulos)
        Me.PanelArticulos.Location = New System.Drawing.Point(108, 67)
        Me.PanelArticulos.Name = "PanelArticulos"
        Me.PanelArticulos.Size = New System.Drawing.Size(977, 261)
        Me.PanelArticulos.TabIndex = 10
        Me.PanelArticulos.Visible = False
        '
        'TxtBuscarArticulos
        '
        Me.TxtBuscarArticulos.Location = New System.Drawing.Point(17, 16)
        Me.TxtBuscarArticulos.Name = "TxtBuscarArticulos"
        Me.TxtBuscarArticulos.Size = New System.Drawing.Size(691, 20)
        Me.TxtBuscarArticulos.TabIndex = 0
        '
        'BtnBuscarArticulosDetalle
        '
        Me.BtnBuscarArticulosDetalle.Location = New System.Drawing.Point(714, 14)
        Me.BtnBuscarArticulosDetalle.Name = "BtnBuscarArticulosDetalle"
        Me.BtnBuscarArticulosDetalle.Size = New System.Drawing.Size(143, 23)
        Me.BtnBuscarArticulosDetalle.TabIndex = 1
        Me.BtnBuscarArticulosDetalle.Text = "Buscar"
        Me.BtnBuscarArticulosDetalle.UseVisualStyleBackColor = True
        '
        'DgvArticulos
        '
        Me.DgvArticulos.AllowUserToAddRows = False
        Me.DgvArticulos.AllowUserToDeleteRows = False
        Me.DgvArticulos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvArticulos.Location = New System.Drawing.Point(17, 42)
        Me.DgvArticulos.Name = "DgvArticulos"
        Me.DgvArticulos.ReadOnly = True
        Me.DgvArticulos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvArticulos.Size = New System.Drawing.Size(943, 187)
        Me.DgvArticulos.TabIndex = 2
        '
        'LblTotalArticulos
        '
        Me.LblTotalArticulos.AutoSize = True
        Me.LblTotalArticulos.Location = New System.Drawing.Point(760, 239)
        Me.LblTotalArticulos.Name = "LblTotalArticulos"
        Me.LblTotalArticulos.Size = New System.Drawing.Size(31, 13)
        Me.LblTotalArticulos.TabIndex = 3
        Me.LblTotalArticulos.Text = "Total"
        '
        'BtnCerrar
        '
        Me.BtnCerrar.Location = New System.Drawing.Point(875, 14)
        Me.BtnCerrar.Name = "BtnCerrar"
        Me.BtnCerrar.Size = New System.Drawing.Size(75, 23)
        Me.BtnCerrar.TabIndex = 4
        Me.BtnCerrar.Text = "Cerrar"
        Me.BtnCerrar.UseVisualStyleBackColor = True
        '
        'FrmIngreso
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1190, 641)
        Me.Controls.Add(Me.TabGeneral)
        Me.Name = "FrmIngreso"
        Me.Text = "Ingresos"
        Me.TabGeneral.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.DgvListado, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DgvDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.ErrorIcono, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelArticulos.ResumeLayout(False)
        Me.PanelArticulos.PerformLayout()
        CType(Me.DgvArticulos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabGeneral As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents BtnAnular As Button
    Friend WithEvents ChkSeleccionar As CheckBox
    Friend WithEvents BtnBuscar As Button
    Friend WithEvents TxtValor As TextBox
    Friend WithEvents LblTotal As Label
    Friend WithEvents DgvListado As DataGridView
    Friend WithEvents Seleccionar As DataGridViewCheckBoxColumn
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Label3 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents TxtId As TextBox
    Friend WithEvents BtnCancelar As Button
    Friend WithEvents BtnInsertar As Button
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ErrorIcono As ErrorProvider
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents BtnBuscarProveedor As Button
    Friend WithEvents TxtNombreProveedor As TextBox
    Friend WithEvents TxtIdProveedor As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TxtNumComprobante As TextBox
    Friend WithEvents TxtSerieComprobante As TextBox
    Friend WithEvents CboTipoComprobante As ComboBox
    Friend WithEvents TxtImpuesto As TextBox
    Friend WithEvents TxtTotal As TextBox
    Friend WithEvents TxtTotalImpuesto As TextBox
    Friend WithEvents TxtSubTotal As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents DgvDetalle As DataGridView
    Friend WithEvents BtnBuscarArticulos As Button
    Friend WithEvents TxtCodigo As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents PanelArticulos As Panel
    Friend WithEvents LblTotalArticulos As Label
    Friend WithEvents DgvArticulos As DataGridView
    Friend WithEvents BtnBuscarArticulosDetalle As Button
    Friend WithEvents TxtBuscarArticulos As TextBox
    Friend WithEvents BtnCerrar As Button
End Class
