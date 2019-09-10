Public Class frm_listado_importacion
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_listado_importacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_combo()

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
            DataGridView1.Columns(0).Width = 350
        Else
            DataGridView1.DataSource = Nothing

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        frm_crear_importacion.ShowDialog()
        frm_crear_importacion.Dispose()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            cod_importacion = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
            frm_tabulacion_importacion.textBox2.Text = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value
            frm_tabulacion_importacion.Button3.Enabled = True
            frm_tabulacion_importacion.TextBox1.Enabled = True
            frm_tabulacion_importacion.ComboBox1.Enabled = True
            frm_tabulacion_importacion.procedimiento_para_llenar_grilla_cajas()
            frm_tabulacion_importacion.llenar_listado_referencias("", "", "", "", "")
            frm_tabulacion_importacion.ComboBox1.Text = "Referencia"
            frmPreInspeccion.txtImportacion.Text = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value
            Me.Close()
        End If
    End Sub
End Class