Public Class frm_filtrado
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

    Private Sub frm_filtrado_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With dataGridView1
            .RowCount = 0
            .RowHeadersWidth = 4
        End With
        Dim consulta As String = ""
        Select Case frm_tabulacion_importacion.ComboBox1.Text
            Case "Caja"
                consulta = "select det_caja as CAJA, det_caja from detalle_importacion_cabcajas where det_codigoimportacion = " & cod_importacion & " ORDER BY det_caja ASC"
            Case "Compañia"
                consulta = "SELECT DISTINCT proveedores.prv_codigoasignado AS COMPAÑIA, proveedores.prv_codigo FROM detalle_importacion_cabcajas INNER JOIN proveedores ON (detalle_importacion_cabcajas.det_codigoproveedor = proveedores.prv_codigo) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & ") ORDER BY proveedores.prv_nombre ASC"
            Case "Hoja"
                consulta = "SELECT  DISTINCT det_hoja AS HOJA, det_hoja FROM detalle_importacion_cabcajas WHERE (det_codigoimportacion =" & cod_importacion & ") ORDER BY det_hoja ASC"
            Case "Referencia"
                consulta = "SELECT DISTINCT detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_referencia FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & ") ORDER BY detcab_referencia ASC"
            Case "Subpartida"
                consulta = "SELECT DISTINCT detalle_importacion_detcajas.detcab_subpartida AS SUBPARTIDA, detalle_importacion_detcajas.detcab_subpartida FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & " AND detalle_importacion_detcajas.detcab_subpartida IS NOT NULL) ORDER BY detcab_subpartida ASC"
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
        Select Case frm_tabulacion_importacion.ComboBox1.Text
            Case "Caja"
                consulta = "select det_caja as CAJA, det_caja from detalle_importacion_cabcajas where det_codigoimportacion = " & cod_importacion & " and det_caja like '" & criterio & "%' ORDER BY det_caja ASC"
            Case "Compañia"
                consulta = "SELECT DISTINCT proveedores.prv_codigoasignado AS COMPAÑIA, proveedores.prv_codigo FROM detalle_importacion_cabcajas INNER JOIN proveedores ON (detalle_importacion_cabcajas.det_codigoproveedor = proveedores.prv_codigo) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & " AND proveedores.prv_codigoasignado like '" & criterio & "%') ORDER BY proveedores.prv_nombre ASC"
            Case "Hoja"
                consulta = "SELECT  DISTINCT det_hoja AS HOJA, det_hoja FROM detalle_importacion_cabcajas WHERE (det_codigoimportacion =" & cod_importacion & " AND det_hoja like '" & criterio & "%') ORDER BY det_hoja ASC"
            Case "Referencia"
                consulta = "SELECT DISTINCT detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_referencia FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & " and detalle_importacion_detcajas.detcab_referencia like '" & criterio & "%') ORDER BY detcab_referencia ASC"
            Case "Subpartida"
                consulta = "SELECT DISTINCT detalle_importacion_detcajas.detcab_subpartida AS SUBPARTIDA, detalle_importacion_detcajas.detcab_subpartida FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & " and detalle_importacion_detcajas.detcab_subpartida like '" & criterio & "%') ORDER BY detcab_referencia ASC"
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
            frm_tabulacion_importacion.TextBox1.Text = dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value
            Select Case frm_tabulacion_importacion.ComboBox1.Text
                Case "Caja"
                    frm_tabulacion_importacion.llenar_listado_referencias("", "", dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value, "", "")
                Case "Hoja"
                    frm_tabulacion_importacion.llenar_listado_referencias("", dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value, "", "", "")
                Case "Compañia"
                    frm_tabulacion_importacion.llenar_listado_referencias("", "", "", dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value, "")
                Case "Referencia"
                    frm_tabulacion_importacion.llenar_listado_referencias(dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value, "", "", "", "")
                Case "Subpartida"
                    frm_tabulacion_importacion.llenar_listado_referencias("", "", "", "", dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value)
            End Select
            Me.Close()
        End If
    End Sub
End Class