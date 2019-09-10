Public Class frm_consolidado_codigo
    Dim clase As New class_library
    Private Sub frm_consolidado_codigo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = frm_ajustes.TextBox1.Text
        llenar_listado(TextBox1.Text)

        TextBox6.Text = FormatCurrency(frm_ajustes.precio_costo(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value), 0)
        TextBox7.Text = FormatCurrency(frm_ajustes.precio_venta1(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value), 0)
        TextBox8.Text = FormatCurrency(frm_ajustes.precio_venta2(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value), 0)
        TextBox2.Text = frm_ajustes.total_items_ajustados(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value)
        TextBox1.Focus()
    End Sub

    Sub llenar_listado(criterio As String) '0 referencia
        If Len(criterio) = 13 Then
            criterio = generar_codigo_ean13(criterio)
        End If
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, SUM(detaj_cantidad) AS cantidad, FORMAT(SUM(detaj_precio_costo * detaj_cantidad), 'Currency') AS costo, FORMAT(SUM(detaj_precio_venta1 * detaj_cantidad), 'Currency') AS venta1, FORMAT(SUM(detaj_precio_venta2 * detaj_cantidad), 'Currency') AS venta2 FROM detajuste INNER JOIN articulos ON (detajuste.detaj_articulo = articulos.ar_codigo) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (detajuste.detaj_codigo_ajuste =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND articulos.ar_codigo like '" & criterio & "%') GROUP BY articulos.ar_codigo ORDER BY detajuste.detaj_codigo_ajuste ASC", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("tabla")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 11
            preparar_columnas()
        End If
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "Precio Costo"
            .Columns(7).HeaderText = "Precio Venta1"
            .Columns(8).HeaderText = "Precio Venta2"
            .Columns(0).Width = 80
            .Columns(1).Width = 100
            .Columns(2).Width = 200
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 80
            .Columns(6).Width = 120
            .Columns(7).Width = 120
            .Columns(8).Width = 120
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ' .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As EventArgs) Handles TextBox7.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As EventArgs) Handles TextBox8.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' If clase.validar_cajas_text(TextBox1, "Filtro") = False Then Exit Sub
        llenar_listado(TextBox1.Text)
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.validar_numeros(e)
    End Sub
End Class