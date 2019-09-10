Public Class frm_transferir_desde_gondola
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(textBox2, "Góndola") = False Then Exit Sub
        If frm_generar_ajuste_gondola.verificar_existencia_gondola(UCase(textBox2.Text), frm_transferencias_desde_bodega.ComboBox1.SelectedValue) = False Then
            MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
            Exit Sub
        End If
        Button3.Enabled = True
        Button4.Enabled = True
        Button2.Enabled = True
        Button1.Enabled = False
        TextBox4.Enabled = True
        Me.AcceptButton = Button3
        TextBox4.Focus()
        textBox2.Enabled = False
        llenar_grilla(frm_transferencias_desde_bodega.TextBox1.Text, UCase(textBox2.Text))
    End Sub

    Sub llenar_grilla(transferencia As Long, gondola As String)
        clase.consultar("SELECT dettransferencia.dt_codarticulo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, dettransferencia.dt_cantidad FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (dettransferencia.dt_trnumero =" & transferencia & " AND dettransferencia.dt_gondola ='" & UCase(gondola) & "')", "tbl")
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
        clase.consultar("SELECT SUM(dt_cantidad) AS totalitems FROM dettransferencia WHERE (dt_trnumero =" & transferencia & " AND dt_gondola ='" & UCase(gondola) & "')", "tbl1")
        If clase.dt.Tables("tbl1").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tbl1").Rows(0)("totalitems")) Then
                TextBox1.Text = 0
                Exit Sub
            End If
            TextBox1.Text = clase.dt.Tables("tbl1").Rows(0)("totalitems")
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox4.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox4.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox4.Text)
        End If
        If Len(TextBox4.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox4.Text & ")", "tabla11")
            codigo_articulo = TextBox4.Text
        End If
        If Len(TextBox4.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            clase.consultar1("SELECT dt_cantidad FROM dettransferencia WHERE (dt_trnumero =" & frm_transferencias_desde_bodega.TextBox1.Text & " AND dt_gondola ='" & UCase(textBox2.Text) & "' AND dt_codarticulo =" & codigo_articulo & ")", "erp")
            If clase.dt1.Tables("erp").Rows.Count > 0 Then
                Dim canti As Integer = clase.dt1.Tables("erp").Rows(0)("dt_cantidad")
                clase.actualizar("UPDATE dettransferencia SET dt_cantidad = " & canti + 1 & " WHERE (dt_trnumero =" & frm_transferencias_desde_bodega.TextBox1.Text & " AND dt_gondola ='" & UCase(textBox2.Text) & "' AND dt_codarticulo =" & codigo_articulo & ")")
            Else
                clase.agregar_registro("INSERT INTO `dettransferencia`(`dt_trnumero`,`dt_gondola`,`dt_codarticulo`,`dt_cantidad`,`dt_costo`,`dt_venta1`,`dt_venta2`) VALUES ( '" & frm_transferencias_desde_bodega.TextBox1.Text & "','" & UCase(textBox2.Text) & "','" & codigo_articulo & "','1','" & Str(precio_costo(codigo_articulo)) & "','" & Str(precio_venta1(codigo_articulo)) & "','" & Str(precio_venta2(codigo_articulo)) & "')")
            End If
            llenar_grilla(frm_transferencias_desde_bodega.TextBox1.Text, UCase(textBox2.Text))
            establecer_fila_actual(codigo_articulo)
            TextBox4.Text = ""
            TextBox4.Focus()
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If
    End Sub

    Function precio_costo(articulo As Double) As Double
        clase.consultar("select ar_costo from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("ar_costo")
    End Function

    Function precio_venta1(articulo As Double) As Double
        clase.consultar("select ar_precio1 from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("ar_precio1")
    End Function

    Function precio_venta2(articulo As Double) As Double
        clase.consultar("select ar_precio2 from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("ar_precio2")
    End Function

    Private Sub establecer_fila_actual(codigo As Long)
        Dim x As Integer
        For x = 0 To DataGridView1.RowCount - 1
            If codigo = DataGridView1.Item(0, x).Value Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, x)
                Exit Sub
            End If
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        clase.borradoautomatico("delete from dettransferencia WHERE (dt_trnumero =" & frm_transferencias_desde_bodega.TextBox1.Text & " AND dt_gondola ='" & UCase(textBox2.Text) & "')")
        llenar_grilla(frm_transferencias_desde_bodega.TextBox1.Text, UCase(textBox2.Text))
        TextBox4.Text = ""
        TextBox4.Enabled = False
        textBox2.Text = ""
        textBox2.Enabled = True
        textBox2.Focus()
        Button3.Enabled = False
        Button4.Enabled = False
        Button2.Enabled = False
        Button1.Enabled = True
        Me.AcceptButton = Button1
    End Sub

    Private Sub frm_transferir_desde_gondola_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        textBox2.Focus()
        DataGridView1.ColumnCount = 6
        preparar_columnas()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm_eliminar_item2.ShowDialog()
        frm_eliminar_item2.Dispose()
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub
End Class