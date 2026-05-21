<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmReporteVentasPeriodo
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
    'Do not modify it by hand.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LblFechaInicio = New System.Windows.Forms.Label()
        Me.LblFechaFin = New System.Windows.Forms.Label()
        Me.DtpFechaInicio = New System.Windows.Forms.DateTimePicker()
        Me.DtpFechaFin = New System.Windows.Forms.DateTimePicker()
        Me.BtnGenerar = New System.Windows.Forms.Button()
        Me.DgvDetalle = New System.Windows.Forms.DataGridView()
        Me.DgvResumen = New System.Windows.Forms.DataGridView()
        Me.LblTotalVentas = New System.Windows.Forms.Label()
        Me.LblTotalMontoAlta = New System.Windows.Forms.Label()
        Me.LblTotalMontoMedia = New System.Windows.Forms.Label()
        Me.LblTotalMontoSBaja = New System.Windows.Forms.Label()
        CType(Me.DgvDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvResumen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblFechaInicio
        '
        Me.LblFechaInicio.AutoSize = True
        Me.LblFechaInicio.Location = New System.Drawing.Point(12, 15)
        Me.LblFechaInicio.Name = "LblFechaInicio"
        Me.LblFechaInicio.Size = New System.Drawing.Size(109, 13)
        Me.LblFechaInicio.TabIndex = 0
        Me.LblFechaInicio.Text = "Fecha Inicio (YYYY-MM-DD):"
        '
        'LblFechaFin
        '
        Me.LblFechaFin.AutoSize = True
        Me.LblFechaFin.Location = New System.Drawing.Point(350, 15)
        Me.LblFechaFin.Name = "LblFechaFin"
        Me.LblFechaFin.Size = New System.Drawing.Size(107, 13)
        Me.LblFechaFin.TabIndex = 1
        Me.LblFechaFin.Text = "Fecha Fin (YYYY-MM-DD):"
        '
        'DtpFechaInicio
        '
        Me.DtpFechaInicio.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DtpFechaInicio.Location = New System.Drawing.Point(127, 12)
        Me.DtpFechaInicio.Name = "DtpFechaInicio"
        Me.DtpFechaInicio.Size = New System.Drawing.Size(200, 20)
        Me.DtpFechaInicio.TabIndex = 2
        '
        'DtpFechaFin
        '
        Me.DtpFechaFin.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DtpFechaFin.Location = New System.Drawing.Point(457, 12)
        Me.DtpFechaFin.Name = "DtpFechaFin"
        Me.DtpFechaFin.Size = New System.Drawing.Size(200, 20)
        Me.DtpFechaFin.TabIndex = 3
        '
        'BtnGenerar
        '
        Me.BtnGenerar.Location = New System.Drawing.Point(670, 10)
        Me.BtnGenerar.Name = "BtnGenerar"
        Me.BtnGenerar.Size = New System.Drawing.Size(100, 25)
        Me.BtnGenerar.TabIndex = 4
        Me.BtnGenerar.Text = "Generar Reporte"
        Me.BtnGenerar.UseVisualStyleBackColor = True
        '
        'DgvDetalle
        '
        Me.DgvDetalle.AllowUserToAddRows = False
        Me.DgvDetalle.AllowUserToDeleteRows = False
        Me.DgvDetalle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvDetalle.Location = New System.Drawing.Point(12, 40)
        Me.DgvDetalle.Name = "DgvDetalle"
        Me.DgvDetalle.ReadOnly = True
        Me.DgvDetalle.Size = New System.Drawing.Size(758, 300)
        Me.DgvDetalle.TabIndex = 5
        '
        'DgvResumen
        '
        Me.DgvResumen.AllowUserToAddRows = False
        Me.DgvResumen.AllowUserToDeleteRows = False
        Me.DgvResumen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvResumen.Location = New System.Drawing.Point(12, 350)
        Me.DgvResumen.Name = "DgvResumen"
        Me.DgvResumen.ReadOnly = True
        Me.DgvResumen.Size = New System.Drawing.Size(758, 150)
        Me.DgvResumen.TabIndex = 6
        '
        'LblTotalVentas
        '
        Me.LblTotalVentas.AutoSize = True
        Me.LblTotalVentas.Location = New System.Drawing.Point(12, 515)
        Me.LblTotalVentas.Name = "LblTotalVentas"
        Me.LblTotalVentas.Size = New System.Drawing.Size(100, 13)
        Me.LblTotalVentas.TabIndex = 7
        Me.LblTotalVentas.Text = "Total Ventas: 0"
        '
        'LblTotalMontoAlta
        '
        Me.LblTotalMontoAlta.AutoSize = True
        Me.LblTotalMontoAlta.Location = New System.Drawing.Point(200, 515)
        Me.LblTotalMontoAlta.Name = "LblTotalMontoAlta"
        Me.LblTotalMontoAlta.Size = New System.Drawing.Size(100, 13)
        Me.LblTotalMontoAlta.TabIndex = 8
        Me.LblTotalMontoAlta.Text = "Monto Alta: S/. 0.00"
        '
        'LblTotalMontoMedia
        '
        Me.LblTotalMontoMedia.AutoSize = True
        Me.LblTotalMontoMedia.Location = New System.Drawing.Point(400, 515)
        Me.LblTotalMontoMedia.Name = "LblTotalMontoMedia"
        Me.LblTotalMontoMedia.Size = New System.Drawing.Size(110, 13)
        Me.LblTotalMontoMedia.TabIndex = 9
        Me.LblTotalMontoMedia.Text = "Monto Media: S/. 0.00"
        '
        'LblTotalMontoSBaja
        '
        Me.LblTotalMontoSBaja.AutoSize = True
        Me.LblTotalMontoSBaja.Location = New System.Drawing.Point(600, 515)
        Me.LblTotalMontoSBaja.Name = "LblTotalMontoSBaja"
        Me.LblTotalMontoSBaja.Size = New System.Drawing.Size(105, 13)
        Me.LblTotalMontoSBaja.TabIndex = 10
        Me.LblTotalMontoSBaja.Text = "Monto Baja: S/. 0.00"
        '
        'FrmReporteVentasPeriodo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(782, 545)
        Me.Controls.Add(Me.LblTotalMontoSBaja)
        Me.Controls.Add(Me.LblTotalMontoMedia)
        Me.Controls.Add(Me.LblTotalMontoAlta)
        Me.Controls.Add(Me.LblTotalVentas)
        Me.Controls.Add(Me.DgvResumen)
        Me.Controls.Add(Me.DgvDetalle)
        Me.Controls.Add(Me.BtnGenerar)
        Me.Controls.Add(Me.DtpFechaFin)
        Me.Controls.Add(Me.DtpFechaInicio)
        Me.Controls.Add(Me.LblFechaFin)
        Me.Controls.Add(Me.LblFechaInicio)
        Me.Name = "FrmReporteVentasPeriodo"
        Me.Text = "Reporte de Ventas por Período"
        CType(Me.DgvDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvResumen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblFechaInicio As Label
    Friend WithEvents LblFechaFin As Label
    Friend WithEvents DtpFechaInicio As DateTimePicker
    Friend WithEvents DtpFechaFin As DateTimePicker
    Friend WithEvents BtnGenerar As Button
    Friend WithEvents DgvDetalle As DataGridView
    Friend WithEvents DgvResumen As DataGridView
    Friend WithEvents LblTotalVentas As Label
    Friend WithEvents LblTotalMontoAlta As Label
    Friend WithEvents LblTotalMontoMedia As Label
    Friend WithEvents LblTotalMontoSBaja As Label
End Class
