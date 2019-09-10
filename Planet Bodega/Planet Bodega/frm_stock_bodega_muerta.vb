Public Class frm_stock_bodega_muerta
    Dim clase As New class_library
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frm_filtrado2.ShowDialog()
        frm_filtrado2.Dispose()
    End Sub

    Private Sub frm_stock_bodega_muerta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Referencia"
        llenar_listado_referencias2("", "", "", "")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        If datagridreferencia.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(datagridreferencia.CurrentRow.Index) Then
            frm_existencias_referencias.ShowDialog()
            frm_existencias_referencias.Dispose()
        End If
    End Sub

    Private Sub llenar_listado_referencias2(ref As String, marca As String, caja As String, proveedor As String)
        clase.consultar("SELECT detcab_referencia AS Referencia, detcab_descripcion AS Descripcion, detcab_marca AS Marca, detcab_composicion AS Composicion, SUM(detcab_cantidad) AS Cant, detcab_unimedida AS Medida FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detcab_cantidad <> 0 AND detcab_referencia LIKE '" & ref & "%' AND detcab_marca  LIKE '" & marca & "%' AND detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_cabcajas.det_caja LIKE '" & caja & "%' AND proveedores.prv_codigoasignado LIKE '" & proveedor & "%') GROUP BY detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, detalle_importacion_detcajas.detcab_marca, detalle_importacion_detcajas.detcab_unimedida ORDER BY detcab_referencia ASC", "resultados_tab")
        If clase.dt.Tables("resultados_tab").Rows.Count > 0 Then
            Application.DoEvents()
            datagridreferencia.Columns.Clear()
            datagridreferencia.DataSource = clase.dt.Tables("resultados_tab")
            preparar_lista_grilla()
        Else
            datagridreferencia.DataSource = Nothing
            datagridreferencia.ColumnCount = 6
            preparar_lista_grilla()
        End If
    End Sub

    Private Sub preparar_lista_grilla()
        With datagridreferencia
            .Columns(0).Width = 100
            .Columns(1).Width = 150
            .Columns(2).Width = 100
            .Columns(3).Width = 250
            .Columns(4).Width = 50
            .Columns(5).Width = 80
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripcion"
            .Columns(2).HeaderText = "Marca"
            .Columns(3).HeaderText = "Composicion"
            .Columns(4).HeaderText = "Cant"
            .Columns(5).HeaderText = "Medida"
        End With
    End Sub

   

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Select Case ComboBox1.Text
            Case "Caja"
                llenar_listado_referencias2("", "", TextBox1.Text, "")
            Case "Referencia"
                llenar_listado_referencias2(TextBox1.Text, "", "", "")
            Case "Compañia"
                llenar_listado_referencias2("", "", "", TextBox1.Text)
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        frm_entrada_mercancia.ShowDialog()
        frm_entrada_mercancia.Dispose()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm_salidas_mercancia.ShowDialog()
        frm_salidas_mercancia.Dispose()
    End Sub
End Class