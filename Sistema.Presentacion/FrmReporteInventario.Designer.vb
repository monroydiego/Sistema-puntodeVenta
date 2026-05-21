<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmReporteInventario
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
        Me.LblStockMinimo = New System.Windows.Forms.Label()
        Me.NumStockMinimo = New System.Windows.Forms.NumericUpDown()
        Me.BtnGenerar = New System.Windows.Forms.Button()
        Me.DgvInventario = New System.Windows.Forms.DataGridView()
        Me.LblTotalArticulos = New System.Windows.Forms.Label()
        Me.LblValorInventario = New System.Windows.Forms.Label()
        Me.LblArticulosCriticos = New System.Windows.Forms.Label()
        CType(Me.NumStockMinimo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvInventario, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblStockMinimo
        '
        Me.LblStockMinimo.AutoSize = True
        Me.LblStockMinimo.Location = New System.Drawing.Point(12, 15)
        Me.LblStockMinimo.Name = "LblStockMinimo"
        Me.LblStockMinimo.Size = New System.Drawing.Size(130, 13)
        Me.LblStockMinimo.TabIndex = 0
        Me.LblStockMinimo.Text = "Stock Mínimo (unidades):"
        '
        'NumStockMinimo
        '
        Me.NumStockMinimo.Location = New System.Drawing.Point(148, 12)
        Me.NumStockMinimo.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.NumStockMinimo.Name = "NumStockMinimo"
        Me.NumStockMinimo.Size = New System.Drawing.Size(100, 20)
        Me.NumStockMinimo.TabIndex = 1
        Me.NumStockMinimo.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'BtnGenerar
        '
        Me.BtnGenerar.Location = New System.Drawing.Point(260, 10)
        Me.BtnGenerar.Name = "BtnGenerar"
        Me.BtnGenerar.Size = New System.Drawing.Size(100, 25)
        Me.BtnGenerar.TabIndex = 2
        Me.BtnGenerar.Text = "Generar Reporte"
        Me.BtnGenerar.UseVisualStyleBackColor = True
        '
        'DgvInventario
        '
        Me.DgvInventario.AllowUserToAddRows = False
        Me.DgvInventario.AllowUserToDeleteRows = False
        Me.DgvInventario.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvInventario.Location = New System.Drawing.Point(12, 40)
        Me.DgvInventario.Name = "DgvInventario"
        Me.DgvInventario.ReadOnly = True
        Me.DgvInventario.Size = New System.Drawing.Size(760, 400)
        Me.DgvInventario.TabIndex = 3
        '
        'LblTotalArticulos
        '
        Me.LblTotalArticulos.AutoSize = True
        Me.LblTotalArticulos.Location = New System.Drawing.Point(12, 450)
        Me.LblTotalArticulos.Name = "LblTotalArticulos"
        Me.LblTotalArticulos.Size = New System.Drawing.Size(100, 13)
        Me.LblTotalArticulos.TabIndex = 4
        Me.LblTotalArticulos.Text = "Total Artículos: 0"
        '
        'LblValorInventario
        '
        Me.LblValorInventario.AutoSize = True
        Me.LblValorInventario.Location = New System.Drawing.Point(250, 450)
        Me.LblValorInventario.Name = "LblValorInventario"
        Me.LblValorInventario.Size = New System.Drawing.Size(150, 13)
        Me.LblValorInventario.TabIndex = 5
        Me.LblValorInventario.Text = "Valor Total Inventario: S/. 0.00"
        '
        'LblArticulosCriticos
        '
        Me.LblArticulosCriticos.AutoSize = True
        Me.LblArticulosCriticos.ForeColor = System.Drawing.Color.Red
        Me.LblArticulosCriticos.Location = New System.Drawing.Point(520, 450)
        Me.LblArticulosCriticos.Name = "LblArticulosCriticos"
        Me.LblArticulosCriticos.Size = New System.Drawing.Size(140, 13)
        Me.LblArticulosCriticos.TabIndex = 6
        Me.LblArticulosCriticos.Text = "Artículos Críticos: 0"
        '
        'FrmReporteInventario
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 475)
        Me.Controls.Add(Me.LblArticulosCriticos)
        Me.Controls.Add(Me.LblValorInventario)
        Me.Controls.Add(Me.LblTotalArticulos)
        Me.Controls.Add(Me.DgvInventario)
        Me.Controls.Add(Me.BtnGenerar)
        Me.Controls.Add(Me.NumStockMinimo)
        Me.Controls.Add(Me.LblStockMinimo)
        Me.Name = "FrmReporteInventario"
        Me.Text = "Reporte de Inventario Valorizado"
        CType(Me.NumStockMinimo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvInventario, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblStockMinimo As Label
    Friend WithEvents NumStockMinimo As NumericUpDown
    Friend WithEvents BtnGenerar As Button
    Friend WithEvents DgvInventario As DataGridView
    Friend WithEvents LblTotalArticulos As Label
    Friend WithEvents LblValorInventario As Label
    Friend WithEvents LblArticulosCriticos As Label
End Class
