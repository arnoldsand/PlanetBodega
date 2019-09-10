Public Class frm_proveedores_mercancia_no_importada
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

    Private Sub frm_proveedores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With dataGridView1
            .RowCount = 0
            .RowHeadersWidth = 4
        End With

        clase.consultar("select prv_codigoasignado, prv_nombre, prv_codigo from proveedores", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("tbl")
        End If
        preparar_columnas()
    End Sub

    Private Sub preparar_columnas()
        dataGridView1.Columns(0).Width = 80
        dataGridView1.Columns(0).HeaderText = "Codigo"
        dataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dataGridView1.Columns(1).Width = 250
        dataGridView1.Columns(1).HeaderText = "Nombre"
        dataGridView1.Columns(2).Visible = False
    End Sub

    Private Sub textBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.TextChanged
        Dim criterio As String
        criterio = textBox2.Text
        clase.consultar("select prv_codigoasignado, prv_nombre, prv_codigo from proveedores where prv_nombre like '" & criterio & "%' order by prv_nombre asc", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt.Tables("tbl")
            preparar_columnas()
        Else
            dataGridView1.DataSource = Nothing
        End If

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsNumeric(dataGridView1.CurrentCell.ColumnIndex) And IsNumeric(dataGridView1.CurrentCell.RowIndex) Then
            frm_ingreso_no_importado.IdProveedor = dataGridView1.Item(2, dataGridView1.CurrentCell.RowIndex).Value
            frm_ingreso_no_importado.NombreProveedor = dataGridView1.Item(1, dataGridView1.CurrentCell.RowIndex).Value
            Me.Close()
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' frm_crear_proveedores.Showdialog()
        ' frm_crear_proveedores.Dispose()
        'actualizar_lista_de proveedoresfrm_crear_proveedores
    End Sub
End Class