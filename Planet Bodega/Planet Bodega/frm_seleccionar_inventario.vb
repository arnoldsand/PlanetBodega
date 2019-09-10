Public Class frm_seleccionar_inventario
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_seleccionar_inventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar1("SELECT  cabinventario.cabinv_codigo, cabinventario.cabinv_fecha, cabinventario.cabinv_hora, cabinventario.cabinv_tipo_inventario, cabinventario.cabinv_operario, cabinventario.cabinv_procesado FROM  cabinventario ORDER BY cabinventario.cabinv_codigo DESC", "lis-inv")
        If clase.dt1.Tables("lis-inv").Rows.Count > 0 Then
            dataGridView1.DataSource = clase.dt1.Tables("lis-inv")
            preparar_columna()
        Else
            dataGridView1.DataSource = Nothing
        End If
    End Sub


    Private Sub preparar_columna()
        With dataGridView1
            .Columns(0).Width = 50
            .Columns(1).Width = 80
            .Columns(2).Width = 50
            .Columns(3).Width = 150
            .Columns(4).Width = 150
            .Columns(5).Width = 80
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Fecha"
            .Columns(2).HeaderText = "Hora"
            .Columns(3).HeaderText = "Tipo de Inventario"
            .Columns(4).HeaderText = "Operario"
            .Columns(5).HeaderText = "Procesado"
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        clase.consultar1("SELECT cabinventario.cabinv_tipo_inventario, cabinventario.cabinv_operario, SUM(detinv_cant_definitiva * ar_precio1) as precio1, SUM(detinv_cant_definitiva * ar_precio2) as precio2, SUM(detinv_cant_definitiva * ar_costo) as costo, SUM(detinventario.detinv_cant_definitiva) as unidades FROM detinventario INNER JOIN articulos ON (detinventario.detinv_articulo = articulos.ar_codigo) INNER JOIN cabinventario ON (detinventario.detinv_cod_inv = cabinventario.cabinv_codigo) WHERE (cabinventario.cabinv_codigo =" & dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value & " AND detinventario.detinv_procesado =TRUE) GROUP BY detinventario.detinv_cod_inv", "invent")
        If clase.dt1.Tables("invent").Rows.Count > 0 Then

            frm_modulo_inventario.TextBox1.Text = dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value
            frm_modulo_inventario.TextBox4.Text = clase.dt1.Tables("invent").Rows(0)("cabinv_tipo_inventario")
            frm_modulo_inventario.TextBox18.Text = clase.dt1.Tables("invent").Rows(0)("cabinv_operario")
            frm_modulo_inventario.TextBox5.Text = FormatCurrency(clase.dt1.Tables("invent").Rows(0)("precio1"))
            frm_modulo_inventario.TextBox2.Text = FormatCurrency(clase.dt1.Tables("invent").Rows(0)("precio2"))
            frm_modulo_inventario.TextBox3.Text = FormatCurrency(clase.dt1.Tables("invent").Rows(0)("costo"))
            frm_modulo_inventario.TextBox7.Text = clase.dt1.Tables("invent").Rows(0)("unidades")
            clase.consultar2("SELECT COUNT(DISTINCT detinv_articulo) AS referencias FROM detinventario WHERE (detinv_cod_inv =" & dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value & ")", "inven")
            frm_modulo_inventario.TextBox8.Text = clase.dt2.Tables("inven").Rows(0)("referencias")

        Else
            frm_modulo_inventario.TextBox1.Text = dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value
            frm_modulo_inventario.TextBox4.Text = dataGridView1.Item(3, dataGridView1.CurrentCell.RowIndex).Value
            frm_modulo_inventario.TextBox18.Text = dataGridView1.Item(4, dataGridView1.CurrentCell.RowIndex).Value
            frm_modulo_inventario.TextBox5.Text = FormatCurrency(0)
            frm_modulo_inventario.TextBox2.Text = FormatCurrency(0)
            frm_modulo_inventario.TextBox3.Text = FormatCurrency(0)
            frm_modulo_inventario.TextBox7.Text = 0
            frm_modulo_inventario.TextBox8.Text = 0

        End If
        frm_modulo_inventario.llenar_inventario_gondola(dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value)
        Me.Close()
    End Sub
End Class