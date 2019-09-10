Public Class frm_almacenar_gondola
    Dim clase As New class_library
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_almacenar_gondola_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        textBox2.Focus()
        DataGridView1.ColumnCount = 6
        preparar_columnas()

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


    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(textBox2, "Góndola") = False Then Exit Sub
        If verificar_existencia_gondola(UCase(textBox2.Text), frm_almacenaje_articulos.ComboBox1.SelectedValue) = False Then
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
        llenar_grilla(frm_almacenaje_articulos.textBox2.Text, UCase(textBox2.Text))
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
            clase.consultar1("SELECT detent_cantidad FROM detalle_entrada WHERE (detent_cod_entrada =" & frm_almacenaje_articulos.textBox2.Text & " AND detent_gondola ='" & UCase(textBox2.Text) & "' AND detent_articulo =" & codigo_articulo & ")", "erp")
            If clase.dt1.Tables("erp").Rows.Count > 0 Then
                Dim canti As Integer = clase.dt1.Tables("erp").Rows(0)("detent_cantidad")
                clase.actualizar("UPDATE detalle_entrada SET detent_cantidad = " & canti + 1 & " WHERE (detent_cod_entrada =" & frm_almacenaje_articulos.textBox2.Text & " AND detent_gondola ='" & UCase(textBox2.Text) & "' AND detent_articulo =" & codigo_articulo & ")")
            Else
                clase.agregar_registro("INSERT INTO `detalle_entrada`(`detent_cod_entrada`,`detent_gondola`,`detent_articulo`,`detent_cantidad`) VALUES ( '" & frm_almacenaje_articulos.textBox2.Text & "','" & UCase(textBox2.Text) & "','" & codigo_articulo & "','1')")
            End If
            llenar_grilla(frm_almacenaje_articulos.textBox2.Text, UCase(textBox2.Text))
            establecer_fila_actual(codigo_articulo)
            TextBox4.Text = ""
            TextBox4.Focus()
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If
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
        clase.borradoautomatico("delete from detalle_entrada WHERE (detent_cod_entrada =" & frm_almacenaje_articulos.textBox2.Text & " AND detent_gondola ='" & UCase(textBox2.Text) & "')")
        llenar_grilla(frm_almacenaje_articulos.textBox2.Text, UCase(textBox2.Text))
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

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm_eliminar_item1.ShowDialog()
        frm_eliminar_item1.Dispose()
    End Sub


    Function verificar_existencia_gondola(gondo As String, bodega As Short) As Boolean
        Dim gondola(cant_gondola_bodega(bodega)) As String
        clase.consultar("SELECT detbodega.dtbod_bloque, detbodega.dtbod_cant_gondola FROM detbodega INNER JOIN bodegas ON (detbodega.dtbod_bodega = bodegas.bod_codigo) WHERE (bodegas.bod_codigo =" & bodega & ")", "tabla1")
        If clase.dt.Tables("tabla1").Rows.Count > 0 Then
            Dim a, b As Short
            Dim cont As Integer = -1
            For a = 0 To clase.dt.Tables("tabla1").Rows.Count - 1
                For b = 1 To clase.dt.Tables("tabla1").Rows(a)("dtbod_cant_gondola")
                    cont += 1
                    gondola(cont) = bloque(clase.dt.Tables("tabla1").Rows(a)("dtbod_bloque") - 1) & b
                Next
            Next
            Dim i As Integer
            For i = 0 To cont
                If gondo = gondola(i) Then
                    Return True
                    Exit Function
                End If
            Next
            Return False
        End If
    End Function

    Function cant_gondola_bodega(bodega As Short) As Integer
        clase.consultar("SELECT SUM(detbodega.dtbod_cant_gondola) AS totalgondola FROM detbodega INNER JOIN bodegas ON (detbodega.dtbod_bodega = bodegas.bod_codigo) WHERE (bodegas.bod_codigo =" & bodega & ")", "tablatotal")
        If clase.dt.Tables("tablatotal").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tablatotal").Rows(0)("totalgondola")) Then
                Return 0
                Exit Function
            End If
            Return clase.dt.Tables("tablatotal").Rows(0)("totalgondola")
        End If
    End Function
End Class