Public Class frm_listado_proveedores6
    Dim clase As class_library = New class_library
    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub



    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.GotFocus
        textBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub


    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.LostFocus
        textBox2.BackColor = Color.White
    End Sub

    Private Sub frm_proveedores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With dataGridView1
            .RowCount = 0
            .RowHeadersWidth = 4
        End With

        clase.consultar("select prv_codigoasignado, prv_codigo from proveedores", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("tbl")
        End If
        dataGridView1.Columns(0).Width = 300
    End Sub

    Private Sub textBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.TextChanged
        Dim criterio As String
        criterio = textBox2.Text
        clase.consultar("select prv_codigoasignado, prv_codigo from proveedores where prv_codigoasignado like '" & criterio & "%' order by prv_codigoasignado asc", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("tbl")
            dataGridView1.Columns(0).Width = 300
        Else
            dataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsNumeric(dataGridView1.CurrentCell.ColumnIndex) And IsNumeric(dataGridView1.CurrentCell.RowIndex) Then
            frm_ver_detalle_factura.TextBox1.Text = dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value
            Me.Close()
        End If
    End Sub
End Class