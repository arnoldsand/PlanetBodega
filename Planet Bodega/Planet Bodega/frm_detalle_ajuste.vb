Public Class frm_detalle_ajuste
    Dim clase As New class_library
    Dim codigo_articulo As Long
    'Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    'Dim arreglobloque(26, 2) As String

    Private Sub frm_detalle_ajuste_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frm_ajustes.llenar_ajustes(frm_ajustes.ComboBox6.SelectedIndex, frm_ajustes.DateTimePicker1.Value.ToString("yyyy-MM-dd"), frm_ajustes.DateTimePicker2.Value.ToString("yyyy-MM-dd"))
    End Sub

    Private Sub frm_detalle_ajuste_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_listado(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value)
        'dataGridView1.ColumnCount = 9
        'preparar_columnas()

        cargar_datos()
        TextBox4.Focus()
        'Dim a As Short
        'For a = 0 To 25
        '    arreglobloque(a, 0) = bloque(a)
        '    arreglobloque(a, 1) = a + 1
        'Next
    End Sub

    'Function convertir_letra_en_numero(letra As String) As Short
    '    Dim b As Short
    '    For b = 0 To 25
    '        If arreglobloque(b, 0) = letra Then
    '            Return arreglobloque(b, 1)
    '        End If
    '    Next
    'End Function

    Sub cargar_datos()
        clase.consultar("SELECT cabajuste.cabaj_codigo, tipos_ajustes.tip_nombre, cabajuste.cabaj_fecha, cabajuste.cabaj_operario, cabajuste.cabaj_observaciones, cabajuste.cabaj_procesado FROM cabajuste INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE (cabajuste.cabaj_codigo =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & ")", "tabla1")
        If clase.dt.Tables("tabla1").Rows.Count > 0 Then
            With clase.dt.Tables("tabla1")
                TextBox3.Text = .Rows(0)("tip_nombre")
                TextBox1.Text = .Rows(0)("cabaj_operario")
                TextBox5.Text = .Rows(0)("cabaj_fecha")
                TextBox2.Text = .Rows(0)("cabaj_observaciones") & ""
                TextBox6.Text = FormatCurrency(frm_ajustes.precio_costo(.Rows(0)("cabaj_codigo")), 2)
                TextBox7.Text = FormatCurrency(frm_ajustes.precio_venta1(.Rows(0)("cabaj_codigo")), 2)
                TextBox8.Text = FormatCurrency(frm_ajustes.precio_venta2(.Rows(0)("cabaj_codigo")), 2)
                'dataGridView1.ColumnCount = 6             

            End With
        End If
    End Sub

    'Private Sub preparar_columnas()
    '    With dataGridView1
    '        .Columns(0).HeaderText = "Góndola"
    '        .Columns(1).HeaderText = "Codigos"
    '        .Columns(2).HeaderText = "Cant Ajustar"
    '        .Columns(3).HeaderText = "Saldo Anterior"
    '        .Columns(4).HeaderText = "Total Costo"
    '        .Columns(5).HeaderText = "Total Venta1"
    '        .Columns(6).HeaderText = "Total Venta2"
    '    End With
    'End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_eliminar_desde_ajuste.ShowDialog()
        frm_eliminar_desde_ajuste.Dispose()
        Me.AcceptButton = Button3
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        If dataGridView1.Rows.Count = 0 Then
            Exit Sub
        End If
        If IsNumeric(dataGridView1.CurrentCell.RowIndex) Then

        End If
    End Sub

    Sub llenar_listado(codajuste As Integer)
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, detajuste.detaj_cantidad, FORMAT(ar_costo * detaj_cantidad, 'Currency'), FORMAT(ar_precio1 * detaj_cantidad, 'Currency'), FORMAT(ar_precio2 * detaj_cantidad, 'Currency') FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) INNER JOIN articulos  ON (detajuste.detaj_articulo = articulos.ar_codigo) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (cabajuste.cabaj_codigo =" & codajuste & ")", "grilla")
        If clase.dt.Tables("grilla").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("grilla")
            preparar_columnas()
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 9
            preparar_columnas()
        End If
        Dim a As Short
        For a = 0 To dataGridView1.RowCount - 1
            If codigo_articulo = dataGridView1.Item(0, a).Value Then
                dataGridView1.CurrentCell = dataGridView1.Item(0, a)
                Exit For
            End If
        Next
    End Sub

    Private Sub preparar_columnas()
        With dataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "PrecioCosto"
            .Columns(7).HeaderText = "PrecioVenta 1"
            .Columns(8).HeaderText = "PrecioVenta 2"
            .Columns(0).Width = 80
            .Columns(1).Width = 150
            .Columns(2).Width = 180
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 50
            .Columns(6).Width = 120
            .Columns(7).Width = 120
            .Columns(8).Width = 120
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Codigo Articulo") = False Then Exit Sub
        If Len(TextBox4.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox4.Text & "')", "tabla11")
            '  codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox5.Text)            esta linea seguramente debe ser quitada
        End If
        If Len(TextBox4.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox4.Text & ")", "tabla11")
            ' codigo_articulo = TextBox5.Text     esta linea seguramente debe ser quitada
        End If
        If Len(TextBox4.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            Dim tipoajuste As String = frm_ajustes.naturaleza_del_ajuste(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value)
            codigo_articulo = clase.dt.Tables("tabla11").Rows(0)("ar_codigo")
            clase.consultar("select* from detajuste where detaj_codigo_ajuste = " & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detaj_articulo = " & codigo_articulo & "", "ajuste")
            If clase.dt.Tables("ajuste").Rows.Count > 0 Then
                Dim cant As Short = clase.dt.Tables("ajuste").Rows(0)("detaj_cantidad")
                Select Case tipoajuste
                    Case "R"
                        clase.actualizar("UPDATE detajuste set detaj_cantidad = " & cant - 1 & " where detaj_codigo_ajuste = " & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detaj_articulo = " & codigo_articulo & " ")
                    Case "S"
                        clase.actualizar("UPDATE detajuste set detaj_cantidad = " & cant + 1 & " where detaj_codigo_ajuste = " & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detaj_articulo = " & codigo_articulo & " ")
                End Select
            Else  ' evaluar si se debe colocar cantidad anterior en el detalle del ajuste, terminar la linea de abajo
                Select Case tipoajuste
                    Case "R"
                        clase.agregar_registro("insert into detajuste (detaj_codigo_ajuste, detaj_articulo, detaj_cantidad, detaj_precio_costo, detaj_precio_venta1, detaj_precio_venta2) VALUES ('" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & "', '" & codigo_articulo & "', '-1','" & Str(precio_costo(codigo_articulo)) & "', '" & Str(precio_venta1(codigo_articulo)) & "', '" & Str(precio_venta2(codigo_articulo)) & "')")
                    Case "S"
                        clase.agregar_registro("insert into detajuste (detaj_codigo_ajuste, detaj_articulo, detaj_cantidad, detaj_precio_costo, detaj_precio_venta1, detaj_precio_venta2) VALUES ('" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & "', '" & codigo_articulo & "', '1','" & Str(precio_costo(codigo_articulo)) & "', '" & Str(precio_venta1(codigo_articulo)) & "', '" & Str(precio_venta2(codigo_articulo)) & "')")
                End Select
            End If
            llenar_listado(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value)
            TextBox4.Text = ""
            TextBox4.Focus()
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If
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

    Private Sub TextBox4_GotFocus1(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        Me.AcceptButton = Button3
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub
End Class