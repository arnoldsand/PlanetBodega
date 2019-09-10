<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPreTransferencia
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
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.grdTiendas = New System.Windows.Forms.DataGridView()
        Me.grdPreTransfer = New System.Windows.Forms.DataGridView()
        CType(Me.grdTiendas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdPreTransfer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCerrar
        '
        Me.btnCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCerrar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCerrar.ForeColor = System.Drawing.Color.Navy
        Me.btnCerrar.Image = Global.WindowsApplication1.My.Resources.Resources.remove1
        Me.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCerrar.Location = New System.Drawing.Point(584, 12)
        Me.btnCerrar.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(85, 27)
        Me.btnCerrar.TabIndex = 7
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'grdTiendas
        '
        Me.grdTiendas.AllowUserToAddRows = False
        Me.grdTiendas.AllowUserToDeleteRows = False
        Me.grdTiendas.AllowUserToResizeColumns = False
        Me.grdTiendas.AllowUserToResizeRows = False
        Me.grdTiendas.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.grdTiendas.BackgroundColor = System.Drawing.Color.White
        Me.grdTiendas.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdTiendas.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdTiendas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdTiendas.Location = New System.Drawing.Point(6, 44)
        Me.grdTiendas.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grdTiendas.Name = "grdTiendas"
        Me.grdTiendas.ReadOnly = True
        Me.grdTiendas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdTiendas.Size = New System.Drawing.Size(196, 575)
        Me.grdTiendas.TabIndex = 9
        '
        'grdPreTransfer
        '
        Me.grdPreTransfer.AllowUserToAddRows = False
        Me.grdPreTransfer.AllowUserToDeleteRows = False
        Me.grdPreTransfer.AllowUserToResizeColumns = False
        Me.grdPreTransfer.AllowUserToResizeRows = False
        Me.grdPreTransfer.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.grdPreTransfer.BackgroundColor = System.Drawing.Color.White
        Me.grdPreTransfer.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdPreTransfer.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grdPreTransfer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdPreTransfer.Location = New System.Drawing.Point(208, 44)
        Me.grdPreTransfer.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grdPreTransfer.Name = "grdPreTransfer"
        Me.grdPreTransfer.ReadOnly = True
        Me.grdPreTransfer.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdPreTransfer.Size = New System.Drawing.Size(461, 575)
        Me.grdPreTransfer.TabIndex = 10
        '
        'frmPreTransferencia
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnCerrar
        Me.ClientSize = New System.Drawing.Size(682, 630)
        Me.Controls.Add(Me.grdPreTransfer)
        Me.Controls.Add(Me.grdTiendas)
        Me.Controls.Add(Me.btnCerrar)
        Me.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Navy
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.Name = "frmPreTransferencia"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PRE TRANSFERENCIA"
        CType(Me.grdTiendas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdPreTransfer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents grdTiendas As System.Windows.Forms.DataGridView
    Friend WithEvents grdPreTransfer As System.Windows.Forms.DataGridView
End Class
