Public Class frm_suspedidas_almacenaje
    Dim clase As New class_library

    Private Sub frm_suspedidas_almacenaje_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT cabent_codigo AS NumeroEntrada FROM cabentrada_bodega WHERE (cabent_estado = FALSE AND cabent_operario IS NOT NULL) ORDER BY NumeroEntrada DESC", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("tbl")
            DataGridView1.Columns(0).Width = 300
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.Columns(0).Width = 300
            DataGridView1.Columns(0).HeaderText = "NumeroEntrada"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            frm_almacenaje_articulos.textBox2.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            frm_almacenaje_articulos.Button2.Enabled = False
            frm_almacenaje_articulos.Button1.Enabled = True
            frm_almacenaje_articulos.Button3.Enabled = True
            frm_almacenaje_articulos.Button6.Enabled = True
            frm_almacenaje_articulos.Button4.Enabled = True
            clase.consultar("select* from cabentrada_bodega where cabent_codigo = " & frm_almacenaje_articulos.textBox2.Text, "almacenaje")
            If clase.dt.Tables("almacenaje").Rows.Count > 0 Then
                frm_almacenaje_articulos.ComboBox1.SelectedValue = clase.dt.Tables("almacenaje").Rows(0)("cabent_bodega")
                frm_almacenaje_articulos.ComboBox1.Enabled = False
                frm_almacenaje_articulos.TextBox3.Text = clase.dt.Tables("almacenaje").Rows(0)("cabent_operario")
                frm_almacenaje_articulos.TextBox3.Enabled = False
                frm_almacenaje_articulos.llenar_lista_gondolas(DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value)
            End If

          
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class