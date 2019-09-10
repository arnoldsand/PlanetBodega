Public Class frm_listado_de_registros_importaciones
    Dim clase As New class_library

    Private Sub frm_listado_de_registros_importaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT registro, fecha FROM registros_de_importaciones WHERE (articulo =" & frm_articulos.dataGridView1.Item(0, frm_ver_articulo.registro_actual).Value & ") ORDER BY codigo DESC", "list")
        If clase.dt.Tables("list").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("list")
            preparar_colummnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 2
            preparar_colummnas()
        End If
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub preparar_colummnas()
        With DataGridView1
            .Columns(0).HeaderText = "Registro de Importación"
            .Columns(1).HeaderText = "Fecha"
            .Columns(0).Width = 200
            .Columns(1).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub
End Class