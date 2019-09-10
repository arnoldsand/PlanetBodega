Public Class frm_reporte_codigos_barra1

    Private Sub frm_reporte_codigos_barra1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'DataSet1.DataTable' Puede moverla o quitarla según sea necesario.
        Me.DataTableTableAdapter.Fill(Me.DataSet1.DataTable)

        Me.ReportViewer1.RefreshReport()
    End Sub
End Class