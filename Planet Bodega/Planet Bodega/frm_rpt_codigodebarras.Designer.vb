<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_rpt_codigodebarras
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
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.codigodebarrasBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtplanetbodega = New WindowsApplication1.dtplanetbodega()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.codigodebarrasTableAdapter = New WindowsApplication1.dtplanetbodegaTableAdapters.codigodebarrasTableAdapter()
        Me.button5 = New System.Windows.Forms.Button()
        CType(Me.codigodebarrasBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtplanetbodega, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'codigodebarrasBindingSource
        '
        Me.codigodebarrasBindingSource.DataMember = "codigodebarras"
        Me.codigodebarrasBindingSource.DataSource = Me.dtplanetbodega
        '
        'dtplanetbodega
        '
        Me.dtplanetbodega.DataSetName = "dtplanetbodega"
        Me.dtplanetbodega.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ReportViewer1
        '
        ReportDataSource1.Name = "DataSet1"
        ReportDataSource1.Value = Me.codigodebarrasBindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource1)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "WindowsApplication1.Report1.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(0, 46)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(1072, 529)
        Me.ReportViewer1.TabIndex = 0
        '
        'codigodebarrasTableAdapter
        '
        Me.codigodebarrasTableAdapter.ClearBeforeFill = True
        '
        'button5
        '
        Me.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.button5.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button5.ForeColor = System.Drawing.Color.Navy
        Me.button5.Image = Global.WindowsApplication1.My.Resources.Resources.remove1
        Me.button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.button5.Location = New System.Drawing.Point(975, 8)
        Me.button5.Name = "button5"
        Me.button5.Size = New System.Drawing.Size(97, 33)
        Me.button5.TabIndex = 82
        Me.button5.Text = "Cerrar"
        Me.button5.UseVisualStyleBackColor = True
        '
        'frm_rpt_codigodebarras
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1074, 576)
        Me.Controls.Add(Me.button5)
        Me.Controls.Add(Me.ReportViewer1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_rpt_codigodebarras"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Codigos de Barra para Preinspección"
        CType(Me.codigodebarrasBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtplanetbodega, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents codigodebarrasBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dtplanetbodega As WindowsApplication1.dtplanetbodega
    Friend WithEvents codigodebarrasTableAdapter As WindowsApplication1.dtplanetbodegaTableAdapters.codigodebarrasTableAdapter
    Private WithEvents button5 As System.Windows.Forms.Button
End Class
