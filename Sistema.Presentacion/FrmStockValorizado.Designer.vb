<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmStockValorizado
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
        Me.GroupBoxAcciones = New System.Windows.Forms.GroupBox()
        Me.BtnActualizar = New System.Windows.Forms.Button()
        Me.DgvResultado = New System.Windows.Forms.DataGridView()
        Me.LblTotal = New System.Windows.Forms.Label()
        Me.LblValorTotal = New System.Windows.Forms.Label()
        Me.GroupBoxAcciones.SuspendLayout()
        CType(Me.DgvResultado, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBoxAcciones
        '
        Me.GroupBoxAcciones.Controls.Add(Me.BtnActualizar)
        Me.GroupBoxAcciones.Location = New System.Drawing.Point(12, 12)
        Me.GroupBoxAcciones.Name = "GroupBoxAcciones"
        Me.GroupBoxAcciones.Size = New System.Drawing.Size(1150, 60)
        Me.GroupBoxAcciones.TabIndex = 0
        Me.GroupBoxAcciones.TabStop = False
        Me.GroupBoxAcciones.Text = "Inventario Valorizado"
        '
        'BtnActualizar
        '
        Me.BtnActualizar.Location = New System.Drawing.Point(10, 22)
        Me.BtnActualizar.Name = "BtnActualizar"
        Me.BtnActualizar.Size = New System.Drawing.Size(150, 28)
        Me.BtnActualizar.TabIndex = 0
        Me.BtnActualizar.Text = "Actualizar"
        Me.BtnActualizar.UseVisualStyleBackColor = True
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
        Me.LblTotal.Size = New System.Drawing.Size(120, 13)
        Me.LblTotal.TabIndex = 2
        Me.LblTotal.Text = "Total Artículos: 0"
        '
        'LblValorTotal
        '
        Me.LblValorTotal.AutoSize = True
        Me.LblValorTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.LblValorTotal.Location = New System.Drawing.Point(650, 605)
        Me.LblValorTotal.Name = "LblValorTotal"
        Me.LblValorTotal.Size = New System.Drawing.Size(260, 15)
        Me.LblValorTotal.TabIndex = 3
        Me.LblValorTotal.Text = "Valor total inventario: S/. 0.00"
        '
        'FrmStockValorizado
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1186, 630)
        Me.Controls.Add(Me.GroupBoxAcciones)
        Me.Controls.Add(Me.DgvResultado)
        Me.Controls.Add(Me.LblTotal)
        Me.Controls.Add(Me.LblValorTotal)
        Me.Name = "FrmStockValorizado"
        Me.Text = "Stock Valorizado"
        Me.GroupBoxAcciones.ResumeLayout(False)
        CType(Me.DgvResultado, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents GroupBoxAcciones As GroupBox
    Friend WithEvents BtnActualizar As Button
    Friend WithEvents DgvResultado As DataGridView
    Friend WithEvents LblTotal As Label
    Friend WithEvents LblValorTotal As Label
End Class
