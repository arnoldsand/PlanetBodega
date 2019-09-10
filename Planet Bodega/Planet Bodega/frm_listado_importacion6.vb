Public Class frm_listado_importacion6
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_listado_importacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_combo()
        With DataGridView1
            .Columns(0).Width = 350

        End With
    End Sub

    Sub llenar_combo()
        Dim c1 As New MySqlDataAdapter("select imp_nombrefecha as NOMBRE_IMPORTACION, imp_codigo from cabimportacion where imp_estado = TRUE order by imp_fecha desc", clase.conex)
        Dim dt As New DataTable
        clase.Conectar()
        c1.SelectCommand.Connection = clase.conex
        c1.Fill(dt)
        clase.desconectar()
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            'cod_importacion = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value  
            'si se necesita otra variable con el codigo de la importacion no usar "cod_importacion"
            'tabulacion_de_facturas.textBox2.Text = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
            frm_facturas_tabuladas.textBox2.Text = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
            Me.Close()
        End If
    End Sub
End Class