Public Class frm_agregar_item_salida
    Dim clase As New class_library
    Dim columnacheck As New System.Windows.Forms.DataGridViewCheckBoxColumn()

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frm_listado_proveedores7.ShowDialog()
        frm_listado_proveedores7.Dispose()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Sub llenar_grilla()
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, detalle_importacion_detcajas.detcab_marca, proveedores.prv_codigoasignado, SUM(detalle_importacion_detcajas.detcab_cantidad) AS cant, detalle_importacion_detcajas.detcab_unimedida, detalle_importacion_detcajas.detcab_coditem FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores ON (detalle_importacion_cabcajas.det_codigoproveedor = proveedores.prv_codigo) WHERE (detalle_importacion_detcajas.detcab_referencia LIKE '" & UCase(TextBox2.Text) & "%' AND detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_cabcajas.det_codigoproveedor =" & TextBox1.Text & ") GROUP BY detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, detalle_importacion_detcajas.detcab_marca, detalle_importacion_detcajas.detcab_unimedida HAVING (cant <>0)", "referencias")
        '  clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, detalle_importacion_detcajas.detcab_marca, proveedores.prv_codigoasignado, detalle_importacion_detcajas.detcab_cantidad, detalle_importacion_detcajas.detcab_unimedida, detalle_importacion_detcajas.detcab_coditem FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores ON (detalle_importacion_cabcajas.det_codigoproveedor = proveedores.prv_codigo) WHERE (detalle_importacion_cabcajas.det_codigoproveedor = " & TextBox1.Text & " AND detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL AND detalle_importacion_detcajas.detcab_referencia like '" & UCase(TextBox2.Text) & "%' AND detalle_importacion_detcajas.detcab_cantidad <> 0)", "referencias")
        If clase.dt.Tables("referencias").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("referencias")
            DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewCheckBoxColumn() {Me.columnacheck})
            Dim a As Short
            For a = 0 To 6
                DataGridView1.Columns(a).ReadOnly = True
            Next
            DataGridView1.Columns(7).ReadOnly = False
            CheckBox1.Visible = True
            preparar_columnas()
            TextBox2.Enabled = True
            TextBox2.Focus()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 6
            ' preparar_columnas()
            CheckBox1.Visible = False
        End If
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Marca"
            .Columns(3).HeaderText = "Proveedor"
            .Columns(4).HeaderText = "Cant"
            .Columns(5).HeaderText = "Medida"
            .Columns(6).Visible = False
            .Columns(0).Width = 120
            .Columns(1).Width = 200
            .Columns(2).Width = 100
            .Columns(3).Width = 200
            .Columns(4).Width = 50
            .Columns(5).Width = 100
            .Columns(7).Width = 20
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub



    Private Sub frm_agregar_item_salida_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        llenar_grilla()
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            Dim a, z As Short
            Dim cont As Short = 0
            For a = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Item(7, a).Value = True Then
                    cont += 1
                End If
            Next
            If cont = 0 Then
                MessageBox.Show("Debe seleccionar por lo menos un item para agregar al movimiento.", "SELECCIONAR ITEM", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Dim ind As Boolean = False
            For z = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Item(7, z).Value = True Then
                    ind = False
                    For a = 0 To frm_salidas_mercancia.DataGridView1.RowCount - 1
                        If frm_salidas_mercancia.DataGridView1.Item(0, a).Value = Me.DataGridView1.Item(6, z).Value Then
                            MessageBox.Show("La refererencia: " & DataGridView1.Item(0, z).Value & " - " & DataGridView1.Item(1, z).Value & " ya se encuentra en el listado, no se agregará al listado.", "REFERENCIA YA AGREGADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ind = True
                            Exit For
                        End If
                    Next
                    If ind = False Then
                        With frm_salidas_mercancia.DataGridView1
                            .RowCount = .RowCount + 1
                            .Item(0, .RowCount - 1).Value = Me.DataGridView1.Item(6, z).Value
                            .Item(1, .RowCount - 1).Value = Me.DataGridView1.Item(0, z).Value
                            .Item(2, .RowCount - 1).Value = Me.DataGridView1.Item(1, z).Value
                            .Item(3, .RowCount - 1).Value = Me.DataGridView1.Item(3, z).Value
                            .Item(4, .RowCount - 1).Value = Me.DataGridView1.Item(4, z).Value
                            .Item(5, .RowCount - 1).Value = Me.DataGridView1.Item(5, z).Value
                            .Item(6, .RowCount - 1).Value = Me.TextBox1.Text
                        End With
                    End If
                End If
            Next
            deseleccionar()
        End If
    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        Dim a As Short
        For a = 0 To DataGridView1.RowCount - 1
            If CheckBox1.Checked = True Then
                DataGridView1.Item(7, a).Value = True
            Else
                DataGridView1.Item(7, a).Value = False
            End If
        Next
    End Sub

    Private Sub deseleccionar()
        Dim a As Short
        For a = 0 To DataGridView1.RowCount - 1
            DataGridView1.Item(7, a).Value = False
        Next
        CheckBox1.Checked = False
    End Sub
End Class