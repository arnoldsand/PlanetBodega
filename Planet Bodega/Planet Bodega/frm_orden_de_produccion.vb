Public Class frm_orden_de_produccion
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_orden_de_produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT cabpatrondist.pat_codigo, cabpatrondist.pat_nombre, cabpatrondist.pat_fecha, SUM(detpatrondist.dp_cantidad) FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo) GROUP BY cabpatrondist.pat_codigo ORDER BY cabpatrondist.pat_nombre ASC, cabpatrondist.pat_fecha DESC", "patron")
        If clase.dt.Tables("patron").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("patron")
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Enabled = False
        TextBox1.Focus()
        Button2.Enabled = False
        clase.consultar("SELECT tienda FROM tiendas WHERE (estado =TRUE) ORDER BY tienda ASC", "tiendas")
        '  If clase.dt.Tables("tiendas").Rows.Count Then
    End Sub
End Class