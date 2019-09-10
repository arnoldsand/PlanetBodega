Public Class frm_reporte_envio

    Private Sub frm_reporte_envio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'pl_cajasDataSet.DataTable1' Puede moverla o quitarla según sea necesario.
        Me.DataTable1TableAdapter.Fill(Me.pl_cajasDataSet.DataTable1)

        Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub
End Class