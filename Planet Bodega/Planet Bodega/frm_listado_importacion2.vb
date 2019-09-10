Public Class frm_listado_importacion2
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_listado_importacion2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_combo()
        With DataGridView1
            .Columns(0).Width = 350

        End With
    End Sub

    Sub llenar_combo()
        Dim c1 As New MySqlDataAdapter("select imp_nombrefecha as NOMBRE_IMPORTACIÓN, imp_codigo from cabimportacion order by imp_fecha desc LIMIT 0, 10", clase.conex)
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

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        frm_crear_importacion.ShowDialog()
        frm_crear_importacion.Dispose()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            importacion = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
            frm_recibo_cajas.TextBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value
            frm_recibo_cajas.llenar_listado_cajas(importacion, frm_recibo_cajas.ComboBox1.Text)
            Me.Close()
        End If
    End Sub
End Class