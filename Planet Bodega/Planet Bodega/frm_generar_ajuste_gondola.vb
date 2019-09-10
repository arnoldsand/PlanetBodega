Public Class frm_generar_ajuste_gondola
    Dim clase As New class_library
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Dim codigoajuste As Short
    Dim clase_ajuste() As Short = {1, -1}
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_combobox(ComboBox6, "Bodega Almacenaje") = False Then Exit Sub
        If clase.validar_cajas_text(textBox2, "Góndola") = False Then Exit Sub
        If verificar_existencia_gondola(UCase(textBox2.Text.Trim), ComboBox6.SelectedValue) = False Then
            MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
            Exit Sub
        End If
        clase.consultar("SELECT* FROM detajuste WHERE (detaj_gondola ='" & UCase(textBox2.Text) & "' AND detaj_codigo_ajuste =" & codigoajuste & ")", "tbl123")
        If clase.dt.Tables("tbl123").Rows.Count > 0 Then
            Dim v As String = MessageBox.Show("Ya existe un ajuste hecho para esta góndola, se borrará el ajuste hecho anteriormente cuando guarde el ajuste a realizar ¿Desea Continuar?", "AJUSTE YA REALIZADO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                cuadrar_formulario()
            End If
            If v = 7 Then
                textBox2.Text = ""
                textBox2.Focus()
                Exit Sub
            End If
        Else
            cuadrar_formulario()
        End If
    End Sub

    Private Sub cuadrar_formulario()
        TextBox4.Enabled = True
        TextBox4.Focus()
        TextBox1.Text = 0
        TextBox3.Text = 0
        ComboBox6.Enabled = False
        textBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button6.Enabled = True
        Me.AcceptButton = Button3
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

    Private Sub frm_generar_ajuste_gondola_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox6, "SELECT bod_codigo, bod_nombre FROM bodegas ORDER BY bod_codigo ASC", "bod_nombre", "bod_codigo")
        codigoajuste = frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value
        ComboBox6.SelectedValue = codigo_bodega(codigoajuste)
        ComboBox6.Enabled = False
        Select Case tipo_ajuste(codigoajuste)
            Case "A"
                ComboBox1.SelectedIndex = 0
                ComboBox1.Enabled = True
            Case "S"
                ComboBox1.SelectedIndex = 0
                ComboBox1.Enabled = False
            Case "R"
                ComboBox1.SelectedIndex = 1
                ComboBox1.Enabled = False
        End Select
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
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
            Dim z As Short
            For z = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Item(0, z).Value = codigo_articulo Then
                    Dim cantid As Integer = DataGridView1.Item(6, z).Value
                    DataGridView1.Item(6, z).Value = cantid + (1 * DataGridView1.Item(7, z).Value)
                    TextBox4.Text = ""
                    TextBox4.Focus()
                    cantidad_unidades()
                    Exit Sub
                End If
            Next
            DataGridView1.RowCount = DataGridView1.RowCount + 1
            With DataGridView1
                .Item(0, .RowCount - 1).Value = codigo_articulo
                .Item(1, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_referencia")
                .Item(2, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_descripcion")
                .Item(3, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("colornombre")
                .Item(4, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("nombretalla")
                .Item(5, .RowCount - 1).Value = cant_articulos_x_gondola(UCase(textBox2.Text), ComboBox6.SelectedValue, codigo_articulo)
                .Item(6, .RowCount - 1).Value = 1 * clase_ajuste(ComboBox1.SelectedIndex)
                .Item(7, .RowCount - 1).Value = clase_ajuste(ComboBox1.SelectedIndex)
                .CurrentCell = .Item(0, .RowCount - 1)
            End With

            cantidad_unidades()
            TextBox4.Text = ""
            TextBox4.Focus()
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If

    End Sub

    Sub cantidad_unidades()
        Dim i As Integer
        Dim cant_filas As Integer = 0
        For i = 0 To DataGridView1.RowCount - 1
            cant_filas = cant_filas + DataGridView1.Item(6, i).Value
        Next
        TextBox1.Text = cant_filas
        TextBox3.Text = DataGridView1.RowCount
    End Sub

    Function cant_articulos_x_gondola(gondola As String, bodega As Short, articulo As Integer) As Integer
        clase.consultar("SELECT inv_cantidad FROM inventario_bodega WHERE (inv_bodega =" & bodega & " AND inv_gondola ='" & gondola & "' AND inv_codigoart =" & articulo & ")", "rest")
        If clase.dt.Tables("rest").Rows.Count > 0 Then
            Return clase.dt.Tables("rest").Rows(0)("inv_cantidad")
        Else
            Return 0
        End If
    End Function

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        restablecer()
    End Sub

    Private Sub restablecer()
        '  ComboBox6.SelectedIndex = -1
        '   ComboBox6.Enabled = True
        textBox2.Text = ""
        textBox2.Enabled = True
        textBox2.Focus()
        TextBox1.Text = 0
        TextBox3.Text = 0
        TextBox4.Enabled = False
        TextBox4.Text = ""
        DataGridView1.RowCount = 0
        Button1.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Me.AcceptButton = Button1
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DataGridView1.RowCount = 0 Then
            MessageBox.Show("Debe agregar por los menos un item para guardar el ajuste", "AGREGAR ITEMS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim v As String = MessageBox.Show("¿Desea guardar el ajuste hecho en este momento?", "GUARDAR AJUSTE", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            clase.borradoautomatico("delete from detajuste where detaj_codigo_ajuste = " & codigoajuste & " AND detaj_gondola = '" & UCase(textBox2.Text) & "'")
            Dim i As Short
            For i = 0 To DataGridView1.RowCount - 1
                With DataGridView1
                    clase.agregar_registro("INSERT INTO `detajuste`(`detaj_codigo_ajuste`,`detaj_gondola`,`detaj_articulo`,`detaj_cantidad`,`detaj_cantidad_anterior`,`detaj_precio_costo`,`detaj_precio_venta1`,`detaj_precio_venta2`) VALUES ('" & codigoajuste & "','" & UCase(textBox2.Text) & "','" & .Item(0, i).Value & "','" & .Item(6, i).Value & "','" & .Item(5, i).Value & "','" & Str(precio_costo(.Item(0, i).Value)) & "','" & Str(precio_venta1(.Item(0, i).Value)) & "','" & Str(precio_venta2(.Item(0, i).Value)) & "')")
                End With
            Next
            restablecer()
            frm_detalle_ajuste.llenar_grid(codigoajuste, frm_detalle_ajuste.ComboBox6.Text)
            frm_detalle_ajuste.cargar_datos()
        End If
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Function precio_costo(articulo As Integer) As Double
        clase.consultar("select ar_costo from articulos where ar_codigo = " & articulo & "", "costo")
        If clase.dt.Tables("costo").Rows.Count > 0 Then
            Return clase.dt.Tables("costo").Rows(0)("ar_costo")
        Else
            Return 0
        End If
    End Function

    Function precio_venta1(articulo As Integer) As Double
        clase.consultar("select ar_precio1 from articulos where ar_codigo = " & articulo & "", "venta")
        If clase.dt.Tables("venta").Rows.Count > 0 Then
            Return clase.dt.Tables("venta").Rows(0)("ar_precio1")
        Else
            Return 0
        End If
    End Function

    Function precio_venta2(articulo As Integer) As Double
        clase.consultar("select ar_precio2 from articulos where ar_codigo = " & articulo & "", "venta")
        If clase.dt.Tables("venta").Rows.Count > 0 Then
            Return clase.dt.Tables("venta").Rows(0)("ar_precio2")
        Else
            Return 0
        End If
    End Function

    Function tipo_ajuste(ajuste As Short) As String
        clase.consultar("SELECT tipos_ajustes.tip_tipo FROM cabajuste INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE (cabajuste.cabaj_codigo =" & ajuste & ")", "tipo")
        If clase.dt.Tables("tipo").Rows.Count > 0 Then
            Return clase.dt.Tables("tipo").Rows(0)("tip_tipo")
        Else
            Return ""
        End If
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        frm_eliminar_item.ShowDialog()
        frm_eliminar_item.Dispose()
    End Sub

    Function codigo_bodega(ajuste As Short) As Short
        clase.consultar("select cabaj_bodega from cabajuste where cabaj_codigo = " & ajuste & "", "tabla3")
        Return clase.dt.Tables("tabla3").Rows(0)("cabaj_bodega")
    End Function
End Class