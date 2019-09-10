Public Class frm_modulo_inventario
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frm_crear_nuevo_inventario.ShowDialog()
        frm_crear_nuevo_inventario.Dispose()
    End Sub

    Private Sub frm_modulo_inventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox1, "select* from bodegas order by bod_codigo asc", "bod_nombre", "bod_codigo")
        dataGridView1.ColumnCount = 5
        llenar_inventario_gondola(17)

    End Sub

    Sub llenar_inventario_gondola(cod As Short)
        clase.consultar1("SELECT detinventario.detinv_bodega, bodegas.bod_nombre, detinventario.detinv_gondola, SUM(detinventario.detinv_cant_cont1), SUM(detinventario.detinv_cant_cont2), SUM(detinventario.detinv_cant_cont3) FROM detinventario INNER JOIN bodegas ON (detinventario.detinv_bodega = bodegas.bod_codigo) INNER JOIN cabinventario ON (detinventario.detinv_cod_inv = cabinventario.cabinv_codigo) WHERE (cabinventario.cabinv_codigo =" & cod & ") GROUP BY detinventario.detinv_bodega, bodegas.bod_nombre, detinventario.detinv_gondola ORDER BY detinventario.detinv_bodega ASC, detinventario.detinv_gondola ASC", "inventario")
        If clase.dt1.Tables("inventario").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt1.Tables("inventario")
            Select Case ComboBox2.Text
                Case "Todos"
                Case "Procesado"
                Case "No Procesado"
            End Select
            Dim columnacheck As New DataGridViewCheckBoxColumn
            dataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {columnacheck})
            preparar_columnas()
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 7
            preparar_columnas()
        End If
    End Sub

    Private Sub TextBox18_GotFocus(sender As Object, e As EventArgs) Handles TextBox18.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        frm_seleccionar_inventario.ShowDialog()
        frm_seleccionar_inventario.Dispose()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frm_generar_secciones.ShowDialog()
        frm_generar_secciones.Dispose()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clase.consultar("select* from cabinventario where cabinv_codigo = " & TextBox1.Text & "", "inv")
        If clase.dt.Tables("inv").Rows(0)("cabinv_procesado") = True Then
            MessageBox.Show("El inventario ya fue procesado no se puede añadir nuevas secciones.", "AÑADIR SECCIONES", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        frm_recibir_datos_inventario.ShowDialog()
        frm_recibir_datos_inventario.Dispose()
    End Sub

    Private Sub preparar_columnas()
        With dataGridView1
            .Columns(0).Visible = False
            .Columns(1).HeaderText = "Bodega"
            .Columns(2).HeaderText = "Góndola"
            .Columns(3).HeaderText = "Cant Conteo 1"
            .Columns(4).HeaderText = "Cant Conteo 2"
            .Columns(5).HeaderText = "Cant Conteo 3"
            .Columns(6).HeaderText = "Procesado"
            .Columns(1).Width = 150
            .Columns(2).Width = 80
            .Columns(6).Width = 80
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = False
        End With
        ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As EventArgs) Handles TextBox8.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As EventArgs) Handles TextBox7.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub
End Class