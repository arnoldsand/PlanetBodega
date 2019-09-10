<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInformePatron
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
        Me.grdTiendas = New System.Windows.Forms.DataGridView()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.DocToPrint = New System.Drawing.Printing.PrintDocument()
        Me.SelectPrint = New System.Windows.Forms.PrintDialog()
        Me.grdPatron = New System.Windows.Forms.DataGridView()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.txtcriterio = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.grdTiendas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdPatron, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdTiendas
        '
        Me.grdTiendas.AllowUserToAddRows = False
        Me.grdTiendas.AllowUserToDeleteRows = False
        Me.grdTiendas.AllowUserToResizeColumns = False
        Me.grdTiendas.AllowUserToResizeRows = False
        Me.grdTiendas.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.grdTiendas.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.grdTiendas.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdTiendas.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdTiendas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdTiendas.Location = New System.Drawing.Point(302, 49)
        Me.grdTiendas.Margin = New System.Windows.Forms.Padding(2)
        Me.grdTiendas.MultiSelect = False
        Me.grdTiendas.Name = "grdTiendas"
        Me.grdTiendas.ReadOnly = True
        Me.grdTiendas.RowHeadersWidth = 4
        Me.grdTiendas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdTiendas.Size = New System.Drawing.Size(316, 491)
        Me.grdTiendas.TabIndex = 26
        '
        'btnImprimir
        '
        Me.btnImprimir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnImprimir.ForeColor = System.Drawing.Color.Navy
        Me.btnImprimir.Image = Global.WindowsApplication1.My.Resources.Resources.printer
        Me.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnImprimir.Location = New System.Drawing.Point(13, 6)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(97, 33)
        Me.btnImprimir.TabIndex = 27
        Me.btnImprimir.Text = "Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = True
        '
        'DocToPrint
        '
        '
        'SelectPrint
        '
        Me.SelectPrint.UseEXDialog = True
        '
        'grdPatron
        '
        Me.grdPatron.AllowUserToAddRows = False
        Me.grdPatron.AllowUserToDeleteRows = False
        Me.grdPatron.AllowUserToResizeColumns = False
        Me.grdPatron.AllowUserToResizeRows = False
        Me.grdPatron.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.grdPatron.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.grdPatron.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdPatron.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.grdPatron.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdPatron.Location = New System.Drawing.Point(11, 77)
        Me.grdPatron.Margin = New System.Windows.Forms.Padding(2)
        Me.grdPatron.MultiSelect = False
        Me.grdPatron.Name = "grdPatron"
        Me.grdPatron.ReadOnly = True
        Me.grdPatron.RowHeadersWidth = 4
        Me.grdPatron.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdPatron.Size = New System.Drawing.Size(287, 463)
        Me.grdPatron.TabIndex = 28
        '
        'btnCerrar
        '
        Me.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCerrar.ForeColor = System.Drawing.Color.Navy
        Me.btnCerrar.Image = Global.WindowsApplication1.My.Resources.Resources.remove
        Me.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCerrar.Location = New System.Drawing.Point(520, 6)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(97, 33)
        Me.btnCerrar.TabIndex = 29
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'txtcriterio
        '
        Me.txtcriterio.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcriterio.Location = New System.Drawing.Point(90, 49)
        Me.txtcriterio.Name = "txtcriterio"
        Me.txtcriterio.Size = New System.Drawing.Size(208, 23)
        Me.txtcriterio.TabIndex = 31
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 53)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 14)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "Nombre Patron:"
        '
        'frmInformePatron
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(632, 561)
        Me.Controls.Add(Me.txtcriterio)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.grdPatron)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.grdTiendas)
        Me.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Navy
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmInformePatron"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "INFORME PATRON DE DISTRIBUCION"
        CType(Me.grdTiendas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdPatron, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grdTiendas As System.Windows.Forms.DataGridView
    Friend WithEvents btnImprimir As System.Windows.Forms.Button
    Friend WithEvents DocToPrint As System.Drawing.Printing.PrintDocument
    Friend WithEvents SelectPrint As System.Windows.Forms.PrintDialog
    Friend WithEvents grdPatron As System.Windows.Forms.DataGridView
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents txtcriterio As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
