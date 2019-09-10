<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_vista_codigos_de_barra
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
        Me.components = New System.ComponentModel.Container()
        Dim ReportDataSource2 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.etiquetasBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtplanetbodega = New WindowsApplication1.dtplanetbodega()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.codigodebarrasBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.codigodebarrasTableAdapter = New WindowsApplication1.dtplanetbodegaTableAdapters.codigodebarrasTableAdapter()
        Me.etiquetasTableAdapter = New WindowsApplication1.dtplanetbodegaTableAdapters.etiquetasTableAdapter()
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.etiquetasBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtplanetbodega, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.codigodebarrasBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'etiquetasBindingSource
        '
        Me.etiquetasBindingSource.DataMember = "etiquetas"
        Me.etiquetasBindingSource.DataSource = Me.dtplanetbodega
        '
        'dtplanetbodega
        '
        Me.dtplanetbodega.DataSetName = "dtplanetbodega"
        Me.dtplanetbodega.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ReportViewer1
        '
        ReportDataSource2.Name = "DataSet1"
        ReportDataSource2.Value = Me.etiquetasBindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource2)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "WindowsApplication1.Report2.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(12, 44)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(766, 478)
        Me.ReportViewer1.TabIndex = 0
        '
        'codigodebarrasBindingSource
        '
        Me.codigodebarrasBindingSource.DataMember = "codigodebarras"
        Me.codigodebarrasBindingSource.DataSource = Me.dtplanetbodega
        '
        'codigodebarrasTableAdapter
        '
        Me.codigodebarrasTableAdapter.ClearBeforeFill = True
        '
        'etiquetasTableAdapter
        '
        Me.etiquetasTableAdapter.ClearBeforeFill = True
        '
        'Button3
        '
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Navy
        Me.Button3.Image = Global.WindowsApplication1.My.Resources.Resources.remove
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(681, 7)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(97, 33)
        Me.Button3.TabIndex = 78
        Me.Button3.Text = "Cerrar"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'frm_vista_codigos_de_barra
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 530)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.ReportViewer1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_vista_codigos_de_barra"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Vista Preliminar de Codigos de Barra"
        CType(Me.etiquetasBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtplanetbodega, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.codigodebarrasBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents codigodebarrasBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dtplanetbodega As WindowsApplication1.dtplanetbodega
    Friend WithEvents codigodebarrasTableAdapter As WindowsApplication1.dtplanetbodegaTableAdapters.codigodebarrasTableAdapter
    Friend WithEvents etiquetasBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents etiquetasTableAdapter As WindowsApplication1.dtplanetbodegaTableAdapters.etiquetasTableAdapter
    Friend WithEvents Button3 As System.Windows.Forms.Button
End Class
