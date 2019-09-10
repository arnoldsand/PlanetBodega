<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_busqueda_articulos
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
        Me.chkDescripcion = New System.Windows.Forms.CheckBox()
        Me.chkReferencia = New System.Windows.Forms.CheckBox()
        Me.txtConsulta = New System.Windows.Forms.TextBox()
        Me.grdConsulta = New System.Windows.Forms.DataGridView()
        Me.Precios = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.FOTO = New System.Windows.Forms.PictureBox()
        Me.btnAceptar = New System.Windows.Forms.Button()
        CType(Me.grdConsulta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FOTO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkDescripcion
        '
        Me.chkDescripcion.AutoSize = True
        Me.chkDescripcion.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDescripcion.ForeColor = System.Drawing.Color.Navy
        Me.chkDescripcion.Location = New System.Drawing.Point(725, 46)
        Me.chkDescripcion.Name = "chkDescripcion"
        Me.chkDescripcion.Size = New System.Drawing.Size(88, 19)
        Me.chkDescripcion.TabIndex = 34
        Me.chkDescripcion.Text = "Descripcion"
        Me.chkDescripcion.UseVisualStyleBackColor = True
        '
        'chkReferencia
        '
        Me.chkReferencia.AutoSize = True
        Me.chkReferencia.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkReferencia.ForeColor = System.Drawing.Color.Navy
        Me.chkReferencia.Location = New System.Drawing.Point(725, 21)
        Me.chkReferencia.Name = "chkReferencia"
        Me.chkReferencia.Size = New System.Drawing.Size(84, 19)
        Me.chkReferencia.TabIndex = 33
        Me.chkReferencia.Text = "Referencia"
        Me.chkReferencia.UseVisualStyleBackColor = True
        '
        'txtConsulta
        '
        Me.txtConsulta.Location = New System.Drawing.Point(12, 45)
        Me.txtConsulta.Name = "txtConsulta"
        Me.txtConsulta.Size = New System.Drawing.Size(707, 20)
        Me.txtConsulta.TabIndex = 32
        '
        'grdConsulta
        '
        Me.grdConsulta.AllowUserToAddRows = False
        Me.grdConsulta.AllowUserToDeleteRows = False
        Me.grdConsulta.AllowUserToResizeColumns = False
        Me.grdConsulta.AllowUserToResizeRows = False
        Me.grdConsulta.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdConsulta.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdConsulta.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdConsulta.Location = New System.Drawing.Point(12, 68)
        Me.grdConsulta.MultiSelect = False
        Me.grdConsulta.Name = "grdConsulta"
        Me.grdConsulta.ReadOnly = True
        Me.grdConsulta.RowHeadersWidth = 4
        Me.grdConsulta.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.grdConsulta.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdConsulta.Size = New System.Drawing.Size(801, 443)
        Me.grdConsulta.TabIndex = 31
        '
        'Precios
        '
        Me.Precios.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Precios.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Precios.ForeColor = System.Drawing.Color.Navy
        Me.Precios.Image = Global.WindowsApplication1.My.Resources.Resources.application_form_edit
        Me.Precios.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Precios.Location = New System.Drawing.Point(207, 3)
        Me.Precios.Name = "Precios"
        Me.Precios.Size = New System.Drawing.Size(97, 33)
        Me.Precios.TabIndex = 38
        Me.Precios.Text = "Precios"
        Me.Precios.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Navy
        Me.Button2.Image = Global.WindowsApplication1.My.Resources.Resources.add
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(109, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(97, 33)
        Me.Button2.TabIndex = 37
        Me.Button2.Text = "Nuevo"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Navy
        Me.Button1.Image = Global.WindowsApplication1.My.Resources.Resources.remove
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(1000, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 33)
        Me.Button1.TabIndex = 36
        Me.Button1.Text = "Cerrar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FOTO
        '
        Me.FOTO.Location = New System.Drawing.Point(819, 68)
        Me.FOTO.Name = "FOTO"
        Me.FOTO.Size = New System.Drawing.Size(278, 228)
        Me.FOTO.TabIndex = 35
        Me.FOTO.TabStop = False
        '
        'btnAceptar
        '
        Me.btnAceptar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAceptar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAceptar.ForeColor = System.Drawing.Color.Navy
        Me.btnAceptar.Image = Global.WindowsApplication1.My.Resources.Resources.accept
        Me.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAceptar.Location = New System.Drawing.Point(11, 3)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(97, 33)
        Me.btnAceptar.TabIndex = 30
        Me.btnAceptar.Text = "Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'frm_busqueda_articulos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1104, 514)
        Me.Controls.Add(Me.Precios)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.FOTO)
        Me.Controls.Add(Me.chkDescripcion)
        Me.Controls.Add(Me.chkReferencia)
        Me.Controls.Add(Me.txtConsulta)
        Me.Controls.Add(Me.grdConsulta)
        Me.Controls.Add(Me.btnAceptar)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_busqueda_articulos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Busqueda de Articulos"
        CType(Me.grdConsulta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FOTO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FOTO As System.Windows.Forms.PictureBox
    Friend WithEvents chkDescripcion As System.Windows.Forms.CheckBox
    Friend WithEvents chkReferencia As System.Windows.Forms.CheckBox
    Friend WithEvents txtConsulta As System.Windows.Forms.TextBox
    Friend WithEvents grdConsulta As System.Windows.Forms.DataGridView
    Friend WithEvents btnAceptar As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Precios As System.Windows.Forms.Button
End Class
