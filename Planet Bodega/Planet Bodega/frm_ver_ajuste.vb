Public Class frm_ver_ajuste
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ver_ajuste_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT bodegas.bod_nombre FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) INNER JOIN bodegas ON (cabajuste.cabaj_bodega = bodegas.bod_codigo) WHERE (detajuste.detaj_codigo_ajuste =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detajuste.detaj_gondola ='" & frm_detalle_ajuste.dataGridView1.Item(0, frm_detalle_ajuste.dataGridView1.CurrentCell.RowIndex).Value & "')", "tablita")
        TextBox5.Text = clase.dt.Tables("tablita").Rows(0)("bod_nombre")
        textBox2.Text = frm_detalle_ajuste.dataGridView1.Item(0, frm_detalle_ajuste.dataGridView1.CurrentCell.RowIndex).Value
        unidades_items()
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, detajuste.detaj_cantidad_anterior, detajuste.detaj_cantidad FROM detajuste INNER JOIN articulos ON (detajuste.detaj_articulo = articulos.ar_codigo) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (detajuste.detaj_codigo_ajuste =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detajuste.detaj_gondola ='" & frm_detalle_ajuste.dataGridView1.Item(0, frm_detalle_ajuste.dataGridView1.CurrentCell.RowIndex).Value & "')", "listaajustes")
        If clase.dt.Tables("listaajustes").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("listaajustes")
        Else
            DataGridView1.DataSource = Nothing
        End If
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "Ajuste"
            .Columns(0).Width = 60
            .Columns(1).Width = 120
            .Columns(2).Width = 180
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 80
            .Columns(6).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub unidades_items()
        clase.consultar1("SELECT COUNT(detaj_codigo) AS items, SUM(detaj_cantidad) As cantidades FROM detajuste WHERE (detaj_codigo_ajuste =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detaj_gondola ='" & UCase(textBox2.Text) & "') GROUP BY detaj_codigo_ajuste", "tbl")
        If clase.dt1.Tables("tbl").Rows.Count > 0 Then
            TextBox1.Text = clase.dt1.Tables("tbl").Rows(0)("items")
            TextBox3.Text = clase.dt1.Tables("tbl").Rows(0)("cantidades")
        End If
    End Sub
End Class