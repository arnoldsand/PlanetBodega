Public Class frm_ver_gondola_almacenar
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"
            .Columns(5).HeaderText = "Cant"
            .Columns(0).Width = 80
            .Columns(1).Width = 120
            .Columns(2).Width = 180
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub frm_ver_gondola_almacenar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_grilla(frm_almacenaje_articulos.textBox2.Text, frm_almacenaje_articulos.DataGridView1.Item(0, frm_almacenaje_articulos.DataGridView1.CurrentRow.Index).Value)
        preparar_columnas()
        textBox2.Text = UCase(frm_almacenaje_articulos.DataGridView1.Item(0, frm_almacenaje_articulos.DataGridView1.CurrentRow.Index).Value)
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Sub llenar_grilla(entrada As Integer, gondola As String)
        clase.consultar("SELECT detalle_entrada.detent_articulo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, detalle_entrada.detent_cantidad FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (detalle_entrada.detent_cod_entrada =" & entrada & " AND detalle_entrada.detent_gondola ='" & gondola & "')", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("tbl")
            preparar_columnas()
            TextBox3.Text = clase.dt.Tables("tbl").Rows.Count
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 6
            preparar_columnas()
            TextBox3.Text = 0
        End If
        clase.consultar("SELECT SUM(detent_cantidad) AS totalitems FROM detalle_entrada WHERE (detent_cod_entrada =" & entrada & " AND detent_gondola ='" & gondola & "')", "tbl1")
        If clase.dt.Tables("tbl1").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tbl1").Rows(0)("totalitems")) Then
                TextBox1.Text = 0
                Exit Sub
            End If
            TextBox1.Text = clase.dt.Tables("tbl1").Rows(0)("totalitems")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm_eliminar2.ShowDialog()
        frm_eliminar2.Dispose()
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub
End Class