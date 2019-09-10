Public Class frm_tallas_colores
    Dim clase As New class_library
    Private Sub frm_tallas_colores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_colores()
        llenar_tallas()
    End Sub

    Sub llenar_colores()
        clase.consultar("SELECT colornombre AS Color FROM colores ORDER BY colornombre ASC", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("tabla")
            dataGridView1.Columns(0).Width = 300
        Else
            dataGridView1.DataSource = Nothing
        End If
    End Sub

    Sub llenar_tallas()        '
        clase.consultar("SELECT nombretalla AS Talla FROM tallas ORDER BY nombretalla ASC", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            DataGridView2.DataSource = clase.dt.Tables("tabla")
            DataGridView2.Columns(0).Width = 300
        Else
            DataGridView2.DataSource = Nothing
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        crear_colores.ShowDialog()
        crear_colores.Dispose()
        llenar_colores()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        crear_tallas.ShowDialog()
        crear_tallas.Dispose()
        llenar_tallas()
    End Sub
End Class