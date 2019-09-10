Public Class frm_vista_codigos_de_barra

    Private Sub frm_vista_codigos_de_barra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'dtplanetbodega.etiquetas' Puede moverla o quitarla según sea necesario.
        Me.etiquetasTableAdapter.Fill(Me.dtplanetbodega.etiquetas)
        'TODO: esta línea de código carga datos en la tabla 'dtplanetbodega.codigodebarras' Puede moverla o quitarla según sea necesario.
        Me.codigodebarrasTableAdapter.Fill(Me.dtplanetbodega.codigodebarras)

        Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class