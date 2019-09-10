<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBusqueda
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.grdConsulta = New System.Windows.Forms.DataGridView()
        Me.txtConsulta = New System.Windows.Forms.TextBox()
        Me.chkReferencia = New System.Windows.Forms.CheckBox()
        Me.chkDescripcion = New System.Windows.Forms.CheckBox()
        Me.FOTO = New System.Windows.Forms.PictureBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.grdConsulta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FOTO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnAceptar
        '
        Me.btnAceptar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnAceptar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAceptar.ForeColor = System.Drawing.Color.Navy
        Me.btnAceptar.Image = Global.WindowsApplication1.My.Resources.Resources.accept
        Me.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAceptar.Location = New System.Drawing.Point(12, 9)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(97, 33)
        Me.btnAceptar.TabIndex = 21
        Me.btnAceptar.Text = "Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'grdConsulta
        '
        Me.grdConsulta.AllowUserToAddRows = False
        Me.grdConsulta.AllowUserToDeleteRows = False
        Me.grdConsulta.AllowUserToResizeColumns = False
        Me.grdConsulta.AllowUserToResizeRows = False
        Me.grdConsulta.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdConsulta.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.grdConsulta.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.Color.Navy
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdConsulta.DefaultCellStyle = DataGridViewCellStyle6
        Me.grdConsulta.Location = New System.Drawing.Point(12, 72)
        Me.grdConsulta.MultiSelect = False
        Me.grdConsulta.Name = "grdConsulta"
        Me.grdConsulta.ReadOnly = True
        Me.grdConsulta.RowHeadersWidth = 4
        Me.grdConsulta.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.grdConsulta.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdConsulta.Size = New System.Drawing.Size(509, 443)
        Me.grdConsulta.TabIndex = 25
        '
        'txtConsulta
        '
        Me.txtConsulta.Location = New System.Drawing.Point(12, 47)
        Me.txtConsulta.Name = "txtConsulta"
        Me.txtConsulta.Size = New System.Drawing.Size(378, 22)
        Me.txtConsulta.TabIndex = 26
        '
        'chkReferencia
        '
        Me.chkReferencia.AutoSize = True
        Me.chkReferencia.ForeColor = System.Drawing.Color.Navy
        Me.chkReferencia.Location = New System.Drawing.Point(6, 16)
        Me.chkReferencia.Name = "chkReferencia"
        Me.chkReferencia.Size = New System.Drawing.Size(79, 18)
        Me.chkReferencia.TabIndex = 27
        Me.chkReferencia.Text = "Referencia"
        Me.chkReferencia.UseVisualStyleBackColor = True
        '
        'chkDescripcion
        '
        Me.chkDescripcion.AutoSize = True
        Me.chkDescripcion.ForeColor = System.Drawing.Color.Navy
        Me.chkDescripcion.Location = New System.Drawing.Point(6, 41)
        Me.chkDescripcion.Name = "chkDescripcion"
        Me.chkDescripcion.Size = New System.Drawing.Size(83, 18)
        Me.chkDescripcion.TabIndex = 28
        Me.chkDescripcion.Text = "Descripción"
        Me.chkDescripcion.UseVisualStyleBackColor = True
        '
        'FOTO
        '
        Me.FOTO.Location = New System.Drawing.Point(527, 74)
        Me.FOTO.Name = "FOTO"
        Me.FOTO.Size = New System.Drawing.Size(278, 228)
        Me.FOTO.TabIndex = 29
        Me.FOTO.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkDescripcion)
        Me.GroupBox1.Controls.Add(Me.chkReferencia)
        Me.GroupBox1.Location = New System.Drawing.Point(396, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(125, 66)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.Navy
        Me.Button1.Image = Global.WindowsApplication1.My.Resources.Resources.remove
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(709, 9)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 33)
        Me.Button1.TabIndex = 31
        Me.Button1.Text = "Cerrar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmBusqueda
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(810, 534)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.FOTO)
        Me.Controls.Add(Me.txtConsulta)
        Me.Controls.Add(Me.grdConsulta)
        Me.Controls.Add(Me.btnAceptar)
        Me.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBusqueda"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Busqueda"
        CType(Me.grdConsulta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FOTO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnAceptar As System.Windows.Forms.Button
    Friend WithEvents grdConsulta As System.Windows.Forms.DataGridView
    Friend WithEvents txtConsulta As System.Windows.Forms.TextBox
    Friend WithEvents chkReferencia As System.Windows.Forms.CheckBox
    Friend WithEvents chkDescripcion As System.Windows.Forms.CheckBox
    Friend WithEvents FOTO As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
