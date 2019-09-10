Public Class frm_patron_distribicion
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
        Button3.Enabled = True
        clase.consultar("SELECT* FROM tiendas WHERE (estado =TRUE) ORDER BY tienda ASC", "tiendas")
        Dim a As Integer
        If clase.dt.Tables("tiendas").Rows.Count Then
            For a = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
                With DataGridView2
                    .RowCount = .RowCount + 1
                    .Item(0, a).Value = clase.dt.Tables("tiendas").Rows(a)("Id")
                    .Item(1, a).Value = clase.dt.Tables("tiendas").Rows(a)("tienda")

                End With
            Next
        Else

        End If
    End Sub
End Class