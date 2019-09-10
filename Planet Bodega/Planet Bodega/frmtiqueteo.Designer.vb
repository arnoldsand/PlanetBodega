<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmtiqueteo
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.txtorden = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.btnAsignar = New System.Windows.Forms.Button()
        Me.btnFinalizar = New System.Windows.Forms.Button()
        Me.btnPreTransferencia = New System.Windows.Forms.Button()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.grdOrden = New System.Windows.Forms.DataGridView()
        Me.grdasignados = New System.Windows.Forms.DataGridView()
        CType(Me.grdOrden, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdasignados, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtorden
        '
        Me.txtorden.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtorden.Location = New System.Drawing.Point(151, 52)
        Me.txtorden.Name = "txtorden"
        Me.txtorden.Size = New System.Drawing.Size(97, 26)
        Me.txtorden.TabIndex = 49
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label2.ForeColor = System.Drawing.Color.Navy
        Me.label2.Location = New System.Drawing.Point(12, 56)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(142, 18)
        Me.label2.TabIndex = 48
        Me.label2.Text = "Orden de Producción:"
        '
        'btnAsignar
        '
        Me.btnAsignar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAsignar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAsignar.ForeColor = System.Drawing.Color.Navy
        Me.btnAsignar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAsignar.Location = New System.Drawing.Point(15, 12)
        Me.btnAsignar.Name = "btnAsignar"
        Me.btnAsignar.Size = New System.Drawing.Size(97, 33)
        Me.btnAsignar.TabIndex = 86
        Me.btnAsignar.Text = "Asignar"
        Me.btnAsignar.UseVisualStyleBackColor = True
        '
        'btnFinalizar
        '
        Me.btnFinalizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFinalizar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFinalizar.ForeColor = System.Drawing.Color.Navy
        Me.btnFinalizar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnFinalizar.Location = New System.Drawing.Point(114, 12)
        Me.btnFinalizar.Name = "btnFinalizar"
        Me.btnFinalizar.Size = New System.Drawing.Size(97, 33)
        Me.btnFinalizar.TabIndex = 87
        Me.btnFinalizar.Text = "Finalizar"
        Me.btnFinalizar.UseVisualStyleBackColor = True
        '
        'btnPreTransferencia
        '
        Me.btnPreTransferencia.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPreTransferencia.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPreTransferencia.ForeColor = System.Drawing.Color.Navy
        Me.btnPreTransferencia.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPreTransferencia.Location = New System.Drawing.Point(212, 12)
        Me.btnPreTransferencia.Name = "btnPreTransferencia"
        Me.btnPreTransferencia.Size = New System.Drawing.Size(97, 33)
        Me.btnPreTransferencia.TabIndex = 88
        Me.btnPreTransferencia.Text = "Pre-Transf"
        Me.btnPreTransferencia.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCerrar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCerrar.ForeColor = System.Drawing.Color.Navy
        Me.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCerrar.Location = New System.Drawing.Point(747, 12)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(97, 33)
        Me.btnCerrar.TabIndex = 89
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'grdOrden
        '
        Me.grdOrden.AllowUserToAddRows = False
        Me.grdOrden.AllowUserToDeleteRows = False
        Me.grdOrden.AllowUserToResizeColumns = False
        Me.grdOrden.AllowUserToResizeRows = False
        Me.grdOrden.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdOrden.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdOrden.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdOrden.DefaultCellStyle = DataGridViewCellStyle2
        Me.grdOrden.Location = New System.Drawing.Point(15, 84)
        Me.grdOrden.MultiSelect = False
        Me.grdOrden.Name = "grdOrden"
        Me.grdOrden.RowHeadersWidth = 4
        Me.grdOrden.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.grdOrden.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdOrden.Size = New System.Drawing.Size(829, 314)
        Me.grdOrden.TabIndex = 90
        '
        'grdasignados
        '
        Me.grdasignados.AllowUserToAddRows = False
        Me.grdasignados.AllowUserToDeleteRows = False
        Me.grdasignados.AllowUserToResizeColumns = False
        Me.grdasignados.AllowUserToResizeRows = False
        Me.grdasignados.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdasignados.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.grdasignados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdasignados.DefaultCellStyle = DataGridViewCellStyle4
        Me.grdasignados.Location = New System.Drawing.Point(15, 404)
        Me.grdasignados.MultiSelect = False
        Me.grdasignados.Name = "grdasignados"
        Me.grdasignados.RowHeadersWidth = 4
        Me.grdasignados.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.grdasignados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdasignados.Size = New System.Drawing.Size(829, 314)
        Me.grdasignados.TabIndex = 91
        '
        'frmtiqueteo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1026, 723)
        Me.Controls.Add(Me.grdasignados)
        Me.Controls.Add(Me.grdOrden)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.btnPreTransferencia)
        Me.Controls.Add(Me.btnFinalizar)
        Me.Controls.Add(Me.btnAsignar)
        Me.Controls.Add(Me.txtorden)
        Me.Controls.Add(Me.label2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmtiqueteo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Modulo de Asignación de Orden de Producción"
        CType(Me.grdOrden, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdasignados, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents txtorden As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Friend WithEvents btnAsignar As System.Windows.Forms.Button
    Friend WithEvents btnFinalizar As System.Windows.Forms.Button
    Friend WithEvents btnPreTransferencia As System.Windows.Forms.Button
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents grdOrden As System.Windows.Forms.DataGridView
    Friend WithEvents grdasignados As System.Windows.Forms.DataGridView
End Class
