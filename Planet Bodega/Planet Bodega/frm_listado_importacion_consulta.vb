Public Class frm_listado_importacion_consulta
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_listado_importacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_combo("")
    End Sub

    Sub llenar_combo(parametro As String)
        Dim c1 As New MySqlDataAdapter("select imp_nombrefecha as NOMBRE_IMPORTACION, imp_codigo from cabimportacion where imp_estado = TRUE AND imp_nombrefecha LIKE '%" & parametro & "%' order by imp_fecha desc", clase.conex)
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

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        frm_crear_importacion.ShowDialog()
        frm_crear_importacion.Dispose()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        frmInforme.ConsImport = DataGridView1(1, DataGridView1.CurrentCell.RowIndex).Value.ToString
        frmInforme.txtImportacion.Text = DataGridView1(0, DataGridView1.CurrentCell.RowIndex).Value.ToString
        Me.Close()
    End Sub

    Private Sub txtfiltradoimportacion_TextChanged(sender As Object, e As EventArgs) Handles txtfiltradoimportacion.TextChanged
        llenar_combo(txtfiltradoimportacion.Text)
    End Sub
End Class