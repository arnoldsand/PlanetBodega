<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFotosImportacion
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFotosImportacion))
        Me.txtNombreImportacion = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.btnImportacion = New System.Windows.Forms.Button()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.btnFactura = New System.Windows.Forms.Button()
        Me.txtProveedor = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtgGrillaImportacion = New System.Windows.Forms.DataGridView()
        Me.VisorPDF = New AxAcroPDFLib.AxAcroPDF()
        CType(Me.dtgGrillaImportacion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VisorPDF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtNombreImportacion
        '
        Me.txtNombreImportacion.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNombreImportacion.Location = New System.Drawing.Point(127, 50)
        Me.txtNombreImportacion.Name = "txtNombreImportacion"
        Me.txtNombreImportacion.Size = New System.Drawing.Size(541, 26)
        Me.txtNombreImportacion.TabIndex = 49
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label2.ForeColor = System.Drawing.Color.Navy
        Me.label2.Location = New System.Drawing.Point(12, 53)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(87, 18)
        Me.label2.TabIndex = 48
        Me.label2.Text = "Importacion:"
        '
        'btnImportacion
        '
        Me.btnImportacion.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnImportacion.Location = New System.Drawing.Point(674, 50)
        Me.btnImportacion.Name = "btnImportacion"
        Me.btnImportacion.Size = New System.Drawing.Size(29, 26)
        Me.btnImportacion.TabIndex = 50
        Me.btnImportacion.Text = "..."
        Me.btnImportacion.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCerrar.Image = Global.WindowsApplication1.My.Resources.Resources.remove
        Me.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCerrar.Location = New System.Drawing.Point(606, 11)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(97, 33)
        Me.btnCerrar.TabIndex = 51
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnFactura
        '
        Me.btnFactura.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFactura.Image = Global.WindowsApplication1.My.Resources.Resources.database_up
        Me.btnFactura.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnFactura.Location = New System.Drawing.Point(12, 11)
        Me.btnFactura.Name = "btnFactura"
        Me.btnFactura.Size = New System.Drawing.Size(97, 33)
        Me.btnFactura.TabIndex = 52
        Me.btnFactura.Text = "Factura"
        Me.btnFactura.UseVisualStyleBackColor = True
        '
        'txtProveedor
        '
        Me.txtProveedor.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProveedor.Location = New System.Drawing.Point(127, 82)
        Me.txtProveedor.Name = "txtProveedor"
        Me.txtProveedor.Size = New System.Drawing.Size(541, 26)
        Me.txtProveedor.TabIndex = 55
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 18)
        Me.Label1.TabIndex = 54
        Me.Label1.Text = "Filtrar Proveedor:"
        '
        'dtgGrillaImportacion
        '
        Me.dtgGrillaImportacion.AllowUserToAddRows = False
        Me.dtgGrillaImportacion.AllowUserToDeleteRows = False
        Me.dtgGrillaImportacion.AllowUserToOrderColumns = True
        Me.dtgGrillaImportacion.AllowUserToResizeColumns = False
        Me.dtgGrillaImportacion.AllowUserToResizeRows = False
        Me.dtgGrillaImportacion.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dtgGrillaImportacion.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dtgGrillaImportacion.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtgGrillaImportacion.Location = New System.Drawing.Point(12, 114)
        Me.dtgGrillaImportacion.MultiSelect = False
        Me.dtgGrillaImportacion.Name = "dtgGrillaImportacion"
        Me.dtgGrillaImportacion.RowHeadersWidth = 4
        Me.dtgGrillaImportacion.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dtgGrillaImportacion.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dtgGrillaImportacion.Size = New System.Drawing.Size(691, 554)
        Me.dtgGrillaImportacion.TabIndex = 157
        '
        'VisorPDF
        '
        Me.VisorPDF.Enabled = True
        Me.VisorPDF.Location = New System.Drawing.Point(709, 53)
        Me.VisorPDF.Name = "VisorPDF"
        Me.VisorPDF.OcxState = CType(resources.GetObject("VisorPDF.OcxState"), System.Windows.Forms.AxHost.State)
        Me.VisorPDF.Size = New System.Drawing.Size(539, 615)
        Me.VisorPDF.TabIndex = 158
        '
        'frmFotosImportacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1260, 676)
        Me.Controls.Add(Me.VisorPDF)
        Me.Controls.Add(Me.dtgGrillaImportacion)
        Me.Controls.Add(Me.txtProveedor)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnFactura)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.btnImportacion)
        Me.Controls.Add(Me.txtNombreImportacion)
        Me.Controls.Add(Me.label2)
        Me.Name = "frmFotosImportacion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fotos Importación"
        CType(Me.dtgGrillaImportacion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VisorPDF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtNombreImportacion As TextBox
    Private WithEvents label2 As Label
    Friend WithEvents btnImportacion As Button
    Friend WithEvents btnCerrar As Button
    Friend WithEvents btnFactura As Button
    Friend WithEvents txtProveedor As TextBox
    Private WithEvents Label1 As Label
    Friend WithEvents dtgGrillaImportacion As DataGridView
    Friend WithEvents VisorPDF As AxAcroPDFLib.AxAcroPDF
End Class
