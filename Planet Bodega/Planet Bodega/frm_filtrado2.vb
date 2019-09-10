Public Class frm_filtrado2
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

    Private Sub frm_filtrado2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With dataGridView1
            .RowCount = 0
            .RowHeadersWidth = 4
        End With
        Dim consulta As String = ""
        Select Case frm_stock_bodega_muerta.ComboBox1.Text
            Case "Caja"
                consulta = "SELECT DISTINCT detalle_importacion_cabcajas.det_caja AS CAJA, detalle_importacion_cabcajas.det_caja FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_detcajas.detcab_cantidad <>0  ) ORDER BY det_caja ASC"
            Case "Compañia"
                consulta = "SELECT DISTINCT proveedores.prv_codigoasignado AS COMPAÑIA, proveedores.prv_codigo FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores  ON (detalle_importacion_cabcajas.det_codigoproveedor = proveedores.prv_codigo) WHERE (detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_detcajas.detcab_cantidad <>0 ) ORDER BY proveedores.prv_codigoasignado ASC"
            Case "Referencia"
                consulta = "SELECT  detalle_importacion_detcajas.detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_descripcion FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_detcajas.detcab_cantidad <>0 ) order by detalle_importacion_detcajas.detcab_referencia ASC"
        End Select
        clase.consultar(consulta, "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("tbl")
        End If
        dataGridView1.Columns(0).Width = 300
    End Sub

    Private Sub textBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.TextChanged
        Dim criterio As String
        criterio = textBox2.Text
        Dim consulta As String = ""
        Select Case frm_stock_bodega_muerta.ComboBox1.Text
            Case "Caja"
                consulta = "SELECT DISTINCT detalle_importacion_cabcajas.det_caja AS CAJA, detalle_importacion_cabcajas.det_caja FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_detcajas.detcab_cantidad <>0 AND detalle_importacion_detcajas.det_caja like '" & criterio & "%' ) ORDER BY det_caja ASC"
            Case "Compañia"
                consulta = "SELECT DISTINCT proveedores.prv_codigoasignado AS COMPAÑIA, proveedores.prv_codigo FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores  ON (detalle_importacion_cabcajas.det_codigoproveedor = proveedores.prv_codigo) WHERE (detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_detcajas.detcab_cantidad <>0 AND proveedores.prv_codigoasignado like '" & criterio & "%') ORDER BY proveedores.prv_codigoasignado ASC"
            Case "Referencia"
                consulta = "SELECT  detalle_importacion_detcajas.detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_descripcion FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_detcajas.detcab_cantidad <>0 AND detalle_importacion_detcajas.detcab_referencia LIKE '" & criterio & "%' ) order by detalle_importacion_detcajas.detcab_referencia ASC"
        End Select
        clase.consultar(consulta, "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("tbl")
            dataGridView1.Columns(0).Width = 300
        Else
            dataGridView1.DataSource = Nothing
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsNumeric(dataGridView1.CurrentCell.ColumnIndex) And IsNumeric(dataGridView1.CurrentCell.RowIndex) Then
            frm_stock_bodega_muerta.TextBox1.Text = dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value
           
            Me.Close()
        End If
    End Sub
End Class