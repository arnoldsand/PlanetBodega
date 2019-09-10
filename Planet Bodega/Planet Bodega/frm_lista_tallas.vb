Public Class frm_lista_tallas
    Dim clase As New class_library
    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_lista_colores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        llenar_datagrid()
    End Sub

    Sub llenar_datagrid()
        clase.consultar("select nombretalla, codigo_talla from tallas order by nombretalla asc", "tbl")
        If clase.dt.Tables(0).Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables(0)
        End If
        DataGridView1.Columns(0).Width = 300
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (IsDBNull(DataGridView1.CurrentCell.RowIndex)) And (IsDBNull(DataGridView1.CurrentCell.ColumnIndex)) Then
            Exit Sub
        End If
        '  frm_ver_articulo.TextBox13.Text = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
        Me.Close()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        crear_tallas.ShowDialog()
        crear_tallas.Dispose()
        llenar_datagrid()
    End Sub
End Class