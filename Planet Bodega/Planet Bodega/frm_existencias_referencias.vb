Public Class frm_existencias_referencias
    Dim clase As New class_library
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub frm_existencias_referencias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ref As String = frm_stock_bodega_muerta.datagridreferencia.Item(0, frm_stock_bodega_muerta.datagridreferencia.CurrentRow.Index).Value
        clase.consultar("select detcab_referencia as Referencia, detcab_cantidad AS Cant, detcab_codigocaja AS Caja, detcab_unimedida AS Medida from detalle_importacion_detcajas where detcab_referencia = '" & ref & "' and detcab_cantidad > 0", "rest")
        If clase.dt.Tables("rest").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("rest")
        Else
            dataGridView1.DataSource = Nothing
        End If
    End Sub


End Class