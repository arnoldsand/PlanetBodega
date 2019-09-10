Public Class frm_entradas_mercancias_suspendidas
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ajustes_importacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bodegaalmacenaje = vbEmpty
        clase.consultar("SELECT cabentrada_bodega.cabent_codigo, bodegas.bod_nombre, cabentrada_bodega.cabent_fecha, cabentrada_bodega.cabent_hora, cabentrada_bodega.cabent_operario FROM cabentrada_bodega INNER JOIN bodegas ON (cabentrada_bodega.cabent_bodega = bodegas.bod_codigo) WHERE (cabentrada_bodega.cabent_estado =FALSE AND cabentrada_bodega.cabent_operario IS NOT NULL) ORDER BY cabentrada_bodega.cabent_codigo DESC", "rest")
        If clase.dt.Tables("rest").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("rest")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 5
            preparar_columnas()
        End If
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 50
            .Columns(1).Width = 120
            .Columns(2).Width = 80
            .Columns(3).Width = 50
            .Columns(4).Width = 200
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Bodega"
            .Columns(2).HeaderText = "Fecha"
            .Columns(3).HeaderText = "Hora"
            .Columns(4).HeaderText = "Operario"
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            frm_almacenaje_x_hand_held.TextBox6.Text = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value
            clase.consultar1("select cabent_bodega from cabentrada_bodega where cabent_codigo = " & DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value & "", "consulta")
            frm_almacenaje_x_hand_held.ComboBox1.SelectedValue = clase.dt1.Tables("consulta").Rows(0)("cabent_bodega")
            Me.Close()
        End If
    End Sub

    
End Class