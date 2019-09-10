<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIngresoNoImportados
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
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.btnDeshacer = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtArticulo = New System.Windows.Forms.TextBox()
        Me.grdArticulo = New System.Windows.Forms.DataGridView()
        Me.btnConsulta = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnProveedor = New System.Windows.Forms.Button()
        Me.btnImportacion = New System.Windows.Forms.Button()
        Me.txtProveedor = New System.Windows.Forms.TextBox()
        Me.txtConsMov = New System.Windows.Forms.TextBox()
        Me.txtConsOrden = New System.Windows.Forms.TextBox()
        Me.txtOperario = New System.Windows.Forms.TextBox()
        Me.txtImportacion = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.grdArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCerrar
        '
        Me.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCerrar.ForeColor = System.Drawing.Color.Navy
        Me.btnCerrar.Image = Global.WindowsApplication1.My.Resources.Resources.remove
        Me.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCerrar.Location = New System.Drawing.Point(522, 10)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(97, 33)
        Me.btnCerrar.TabIndex = 22
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnDeshacer
        '
        Me.btnDeshacer.Enabled = False
        Me.btnDeshacer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDeshacer.ForeColor = System.Drawing.Color.Navy
        Me.btnDeshacer.Image = Global.WindowsApplication1.My.Resources.Resources.repeat
        Me.btnDeshacer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDeshacer.Location = New System.Drawing.Point(107, 12)
        Me.btnDeshacer.Name = "btnDeshacer"
        Me.btnDeshacer.Size = New System.Drawing.Size(97, 33)
        Me.btnDeshacer.TabIndex = 2
        Me.btnDeshacer.Text = "Deshacer"
        Me.btnDeshacer.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Enabled = False
        Me.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGuardar.ForeColor = System.Drawing.Color.Navy
        Me.btnGuardar.Image = Global.WindowsApplication1.My.Resources.Resources.save
        Me.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGuardar.Location = New System.Drawing.Point(205, 12)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(97, 33)
        Me.btnGuardar.TabIndex = 3
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(17, 150)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 14)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Articulo:"
        '
        'txtArticulo
        '
        Me.txtArticulo.Enabled = False
        Me.txtArticulo.Location = New System.Drawing.Point(100, 145)
        Me.txtArticulo.Name = "txtArticulo"
        Me.txtArticulo.Size = New System.Drawing.Size(131, 22)
        Me.txtArticulo.TabIndex = 19
        '
        'grdArticulo
        '
        Me.grdArticulo.AllowUserToAddRows = False
        Me.grdArticulo.AllowUserToDeleteRows = False
        Me.grdArticulo.AllowUserToResizeColumns = False
        Me.grdArticulo.AllowUserToResizeRows = False
        Me.grdArticulo.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.grdArticulo.BackgroundColor = System.Drawing.Color.White
        Me.grdArticulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdArticulo.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdArticulo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdArticulo.Location = New System.Drawing.Point(11, 171)
        Me.grdArticulo.Name = "grdArticulo"
        Me.grdArticulo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdArticulo.Size = New System.Drawing.Size(611, 442)
        Me.grdArticulo.TabIndex = 20
        '
        'btnConsulta
        '
        Me.btnConsulta.Enabled = False
        Me.btnConsulta.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConsulta.Location = New System.Drawing.Point(235, 145)
        Me.btnConsulta.Name = "btnConsulta"
        Me.btnConsulta.Size = New System.Drawing.Size(29, 22)
        Me.btnConsulta.TabIndex = 21
        Me.btnConsulta.Text = "..."
        Me.btnConsulta.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.Navy
        Me.Button1.Image = Global.WindowsApplication1.My.Resources.Resources.accept
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(9, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 33)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Aceptar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnProveedor)
        Me.GroupBox1.Controls.Add(Me.btnImportacion)
        Me.GroupBox1.Controls.Add(Me.txtProveedor)
        Me.GroupBox1.Controls.Add(Me.txtConsMov)
        Me.GroupBox1.Controls.Add(Me.txtConsOrden)
        Me.GroupBox1.Controls.Add(Me.txtOperario)
        Me.GroupBox1.Controls.Add(Me.txtImportacion)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(11, 51)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(608, 88)
        Me.GroupBox1.TabIndex = 23
        Me.GroupBox1.TabStop = False
        '
        'btnProveedor
        '
        Me.btnProveedor.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnProveedor.Location = New System.Drawing.Point(165, 36)
        Me.btnProveedor.Name = "btnProveedor"
        Me.btnProveedor.Size = New System.Drawing.Size(29, 22)
        Me.btnProveedor.TabIndex = 17
        Me.btnProveedor.Text = "..."
        Me.btnProveedor.UseVisualStyleBackColor = True
        '
        'btnImportacion
        '
        Me.btnImportacion.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnImportacion.Location = New System.Drawing.Point(257, 12)
        Me.btnImportacion.Name = "btnImportacion"
        Me.btnImportacion.Size = New System.Drawing.Size(29, 22)
        Me.btnImportacion.TabIndex = 12
        Me.btnImportacion.Text = "..."
        Me.btnImportacion.UseVisualStyleBackColor = True
        '
        'txtProveedor
        '
        Me.txtProveedor.BackColor = System.Drawing.Color.White
        Me.txtProveedor.Location = New System.Drawing.Point(89, 36)
        Me.txtProveedor.Name = "txtProveedor"
        Me.txtProveedor.ReadOnly = True
        Me.txtProveedor.Size = New System.Drawing.Size(73, 22)
        Me.txtProveedor.TabIndex = 16
        '
        'txtConsMov
        '
        Me.txtConsMov.BackColor = System.Drawing.Color.White
        Me.txtConsMov.Location = New System.Drawing.Point(544, 36)
        Me.txtConsMov.Name = "txtConsMov"
        Me.txtConsMov.ReadOnly = True
        Me.txtConsMov.Size = New System.Drawing.Size(58, 22)
        Me.txtConsMov.TabIndex = 19
        Me.txtConsMov.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtConsOrden
        '
        Me.txtConsOrden.BackColor = System.Drawing.Color.White
        Me.txtConsOrden.Location = New System.Drawing.Point(544, 60)
        Me.txtConsOrden.Name = "txtConsOrden"
        Me.txtConsOrden.ReadOnly = True
        Me.txtConsOrden.Size = New System.Drawing.Size(58, 22)
        Me.txtConsOrden.TabIndex = 21
        Me.txtConsOrden.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtOperario
        '
        Me.txtOperario.Location = New System.Drawing.Point(89, 60)
        Me.txtOperario.Name = "txtOperario"
        Me.txtOperario.Size = New System.Drawing.Size(166, 22)
        Me.txtOperario.TabIndex = 18
        '
        'txtImportacion
        '
        Me.txtImportacion.BackColor = System.Drawing.Color.White
        Me.txtImportacion.Location = New System.Drawing.Point(89, 12)
        Me.txtImportacion.Name = "txtImportacion"
        Me.txtImportacion.ReadOnly = True
        Me.txtImportacion.Size = New System.Drawing.Size(166, 22)
        Me.txtImportacion.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(431, 65)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(111, 14)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Orden de Producción:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(475, 39)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(70, 14)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Movimiento:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(5, 39)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 14)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Proveedor:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(4, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 14)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Operario:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(5, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 14)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Importación:"
        '
        'frmIngresoNoImportados
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCerrar
        Me.ClientSize = New System.Drawing.Size(627, 625)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnConsulta)
        Me.Controls.Add(Me.grdArticulo)
        Me.Controls.Add(Me.txtArticulo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnDeshacer)
        Me.Controls.Add(Me.btnCerrar)
        Me.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Navy
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmIngresoNoImportados"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "INGRESO NO IMPORTADOS"
        CType(Me.grdArticulo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents btnDeshacer As System.Windows.Forms.Button
    Friend WithEvents btnGuardar As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtArticulo As System.Windows.Forms.TextBox
    Friend WithEvents grdArticulo As System.Windows.Forms.DataGridView
    Friend WithEvents btnConsulta As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnProveedor As System.Windows.Forms.Button
    Friend WithEvents btnImportacion As System.Windows.Forms.Button
    Friend WithEvents txtProveedor As System.Windows.Forms.TextBox
    Friend WithEvents txtConsMov As System.Windows.Forms.TextBox
    Friend WithEvents txtConsOrden As System.Windows.Forms.TextBox
    Friend WithEvents txtOperario As System.Windows.Forms.TextBox
    Friend WithEvents txtImportacion As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
