<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmConsultaCompras
    Inherits System.Windows.Forms.Form

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

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBoxFiltro = New System.Windows.Forms.GroupBox()
        Me.LblDesde = New System.Windows.Forms.Label()
        Me.DtpFechaInicio = New System.Windows.Forms.DateTimePicker()
        Me.LblHasta = New System.Windows.Forms.Label()
        Me.DtpFechaFin = New System.Windows.Forms.DateTimePicker()
        Me.BtnBuscar = New System.Windows.Forms.Button()
        Me.BtnLimpiar = New System.Windows.Forms.Button()
        Me.DgvResultado = New System.Windows.Forms.DataGridView()
        Me.LblTotal = New System.Windows.Forms.Label()
        Me.LblSumaTotal = New System.Windows.Forms.Label()
        Me.GroupBoxFiltro.SuspendLayout()
        CType(Me.DgvResultado, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBoxFiltro
        '
        Me.GroupBoxFiltro.Controls.Add(Me.LblDesde)
        Me.GroupBoxFiltro.Controls.Add(Me.DtpFechaInicio)
        Me.GroupBoxFiltro.Controls.Add(Me.LblHasta)
        Me.GroupBoxFiltro.Controls.Add(Me.DtpFechaFin)
        Me.GroupBoxFiltro.Controls.Add(Me.BtnBuscar)
        Me.GroupBoxFiltro.Controls.Add(Me.BtnLimpiar)
        Me.GroupBoxFiltro.Location = New System.Drawing.Point(12, 12)
        Me.GroupBoxFiltro.Name = "GroupBoxFiltro"
        Me.GroupBoxFiltro.Size = New System.Drawing.Size(1150, 60)
        Me.GroupBoxFiltro.TabIndex = 0
        Me.GroupBoxFiltro.TabStop = False
        Me.GroupBoxFiltro.Text = "Filtro por Período"
        '
        'LblDesde
        '
        Me.LblDesde.AutoSize = True
        Me.LblDesde.Location = New System.Drawing.Point(10, 28)
        Me.LblDesde.Name = "LblDesde"
        Me.LblDesde.Size = New System.Drawing.Size(40, 13)
        Me.LblDesde.TabIndex = 0
        Me.LblDesde.Text = "Desde:"
        '
        'DtpFechaInicio
        '
        Me.DtpFechaInicio.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DtpFechaInicio.Location = New System.Drawing.Point(56, 24)
        Me.DtpFechaInicio.Name = "DtpFechaInicio"
        Me.DtpFechaInicio.Size = New System.Drawing.Size(120, 20)
        Me.DtpFechaInicio.TabIndex = 1
        '
        'LblHasta
        '
        Me.LblHasta.AutoSize = True
        Me.LblHasta.Location = New System.Drawing.Point(190, 28)
        Me.LblHasta.Name = "LblHasta"
        Me.LblHasta.Size = New System.Drawing.Size(38, 13)
        Me.LblHasta.TabIndex = 2
        Me.LblHasta.Text = "Hasta:"
        '
        'DtpFechaFin
        '
        Me.DtpFechaFin.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DtpFechaFin.Location = New System.Drawing.Point(234, 24)
        Me.DtpFechaFin.Name = "DtpFechaFin"
        Me.DtpFechaFin.Size = New System.Drawing.Size(120, 20)
        Me.DtpFechaFin.TabIndex = 3
        '
        'BtnBuscar
        '
        Me.BtnBuscar.Location = New System.Drawing.Point(370, 22)
        Me.BtnBuscar.Name = "BtnBuscar"
        Me.BtnBuscar.Size = New System.Drawing.Size(130, 28)
        Me.BtnBuscar.TabIndex = 4
        Me.BtnBuscar.Text = "Buscar"
        Me.BtnBuscar.UseVisualStyleBackColor = True
        '
        'BtnLimpiar
        '
        Me.BtnLimpiar.Location = New System.Drawing.Point(510, 22)
        Me.BtnLimpiar.Name = "BtnLimpiar"
        Me.BtnLimpiar.Size = New System.Drawing.Size(130, 28)
        Me.BtnLimpiar.TabIndex = 5
        Me.BtnLimpiar.Text = "Limpiar"
        Me.BtnLimpiar.UseVisualStyleBackColor = True
        '
        'DgvResultado
        '
        Me.DgvResultado.AllowUserToAddRows = False
        Me.DgvResultado.AllowUserToDeleteRows = False
        Me.DgvResultado.AllowUserToOrderColumns = True
        Me.DgvResultado.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvResultado.Location = New System.Drawing.Point(12, 85)
        Me.DgvResultado.Name = "DgvResultado"
        Me.DgvResultado.ReadOnly = True
        Me.DgvResultado.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvResultado.Size = New System.Drawing.Size(1150, 510)
        Me.DgvResultado.TabIndex = 1
        '
        'LblTotal
        '
        Me.LblTotal.AutoSize = True
        Me.LblTotal.Location = New System.Drawing.Point(12, 605)
        Me.LblTotal.Name = "LblTotal"
        Me.LblTotal.Size = New System.Drawing.Size(100, 13)
        Me.LblTotal.TabIndex = 2
        Me.LblTotal.Text = "Total Registros: 0"
        '
        'LblSumaTotal
        '
        Me.LblSumaTotal.AutoSize = True
        Me.LblSumaTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.LblSumaTotal.Location = New System.Drawing.Point(700, 605)
        Me.LblSumaTotal.Name = "LblSumaTotal"
        Me.LblSumaTotal.Size = New System.Drawing.Size(220, 15)
        Me.LblSumaTotal.TabIndex = 3
        Me.LblSumaTotal.Text = "Suma importes del período: S/. 0.00"
        '
        'FrmConsultaCompras
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1186, 630)
        Me.Controls.Add(Me.GroupBoxFiltro)
        Me.Controls.Add(Me.DgvResultado)
        Me.Controls.Add(Me.LblTotal)
        Me.Controls.Add(Me.LblSumaTotal)
        Me.Name = "FrmConsultaCompras"
        Me.Text = "Consulta de Compras por Período"
        Me.GroupBoxFiltro.ResumeLayout(False)
        Me.GroupBoxFiltro.PerformLayout()
        CType(Me.DgvResultado, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents GroupBoxFiltro As GroupBox
    Friend WithEvents LblDesde As Label
    Friend WithEvents DtpFechaInicio As DateTimePicker
    Friend WithEvents LblHasta As Label
    Friend WithEvents DtpFechaFin As DateTimePicker
    Friend WithEvents BtnBuscar As Button
    Friend WithEvents BtnLimpiar As Button
    Friend WithEvents DgvResultado As DataGridView
    Friend WithEvents LblTotal As Label
    Friend WithEvents LblSumaTotal As Label
End Class
