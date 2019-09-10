Public Class frm_seleccionarimportacion
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_seleccionarimportacion_Load(sender As Object, e As EventArgs) Handles Me.Load
        llenar_combo()

    End Sub

    Sub llenar_combo()
        Dim c1 As New MySqlDataAdapter("select imp_nombrefecha as NOMBRE_IMPORTACION, imp_codigo from cabimportacion order by imp_fecha desc LIMIT 0, 10", clase.conex)
        Dim dt As New DataTable
        clase.Conectar()
        c1.SelectCommand.Connection = clase.conex
        c1.Fill(dt)
        clase.desconectar()
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
            With DataGridView1
                .Columns(0).Width = 350

            End With
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            Me.Close()
        End If
    End Sub
End Class