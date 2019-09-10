Public Class frm_rpt_codigodebarras

    Private Sub frm_rpt_codigodebarras_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'dtplanetbodega.codigodebarras' Puede moverla o quitarla según sea necesario.
        Me.codigodebarrasTableAdapter.Fill(Me.dtplanetbodega.codigodebarras)

        Me.ReportViewer1.RefreshReport()

    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub
End Class