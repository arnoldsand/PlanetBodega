Public Class frm_almacenaje_articulos
    Dim clase As New class_library
    Dim consecutivo As Integer
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_almacenaje_articulos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.borradoautomatico("delete from cabentrada_bodega where cabent_bodega IS NULL AND cabent_operario IS NULL AND cabent_estado = FALSE")
        clase.llenar_combo(ComboBox1, "select* from bodegas order by bod_codigo asc", "bod_nombre", "bod_codigo")
        codiimportation = vbEmpty
        DataGridView1.ColumnCount = 6
        preparar_columnas()
        encabezado_entrada()
        TextBox4.Text = FormatCurrency(0, 0)
        TextBox5.Text = FormatCurrency(0, 0)
        TextBox6.Text = FormatCurrency(0, 0)
    End Sub

    Private Sub encabezado_entrada()
        consecutivo = consecutivo_entrada()
        clase.agregar_registro("INSERT INTO cabentrada_bodega (cabent_codigo, cabent_fecha, cabent_hora, cabent_estado) VALUES ('" & consecutivo & "', '" & Now.ToString("yyyy-MM-dd") & "', '" & Now.ToString("HH:mm:ss") & "', FALSE)")
        textBox2.Text = consecutivo
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Góndola"
            .Columns(1).HeaderText = "Referencias"
            .Columns(2).HeaderText = "Cant"
            .Columns(3).HeaderText = "Precio Costo"
            .Columns(4).HeaderText = "Precio Venta 1"
            .Columns(5).HeaderText = "Precio Venta 2"
            .Columns(0).Width = 100
            .Columns(1).Width = 100
            .Columns(2).Width = 100
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 100
        End With
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Function consecutivo_entrada() As Integer
        clase.consultar("SELECT MAX(cabent_codigo) AS maximo FROM cabentrada_bodega", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("maximo")) Then
                Return 1
            End If
            Return clase.dt.Tables("tbl").Rows(0)("maximo") + 1
        End If
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        frm_seleccionar_importacion1.ShowDialog()
        frm_seleccionar_importacion1.Dispose()
    End Sub

    Function precion_costo(entrada As Short) As Double
        clase.consultar("SELECT SUM(articulos.ar_costo) AS costo FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) WHERE (detalle_entrada.detent_cod_entrada =" & entrada & ")", "table")
        If IsDBNull(clase.dt.Tables("table").Rows(0)("costo")) Then
            Return 0
        End If
        Return clase.dt.Tables("table").Rows(0)("costo")
    End Function

    Function precion_venta1(entrada As Short) As Double
        clase.consultar("SELECT SUM(articulos.ar_precio1) AS venta1 FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) WHERE (detalle_entrada.detent_cod_entrada =" & entrada & ")", "table")
        If IsDBNull(clase.dt.Tables("table").Rows(0)("venta1")) Then
            Return 0
        End If
        Return clase.dt.Tables("table").Rows(0)("venta1")
    End Function

    Function precion_venta2(entrada As Short) As Double
        clase.consultar("SELECT SUM(articulos.ar_precio2) AS venta2 FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) WHERE (detalle_entrada.detent_cod_entrada =" & entrada & ")", "table")
        If IsDBNull(clase.dt.Tables("table").Rows(0)("venta2")) Then
            Return 0
        End If
        Return clase.dt.Tables("table").Rows(0)("venta2")
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_combobox(ComboBox1, "Bodega") = False Then Exit Sub

        If clase.validar_cajas_text(TextBox3, "Realizado Por") = False Then Exit Sub
        Button1.Enabled = True
        Button3.Enabled = True
        Button6.Enabled = True
        Button4.Enabled = True

        Button2.Enabled = False
        ComboBox1.Enabled = False
        TextBox3.Enabled = False
        Dim fecha As Date = Now
        clase.actualizar("UPDATE cabentrada_bodega set cabent_operario = '" & UCase(TextBox3.Text) & "', cabent_bodega = " & ComboBox1.SelectedValue & " where cabent_codigo = " & textBox2.Text)
    End Sub

    Sub llenar_lista_gondolas(codalmacenaje As Integer)
        clase.consultar("SELECT detalle_entrada.detent_gondola, COUNT(detalle_entrada.detent_articulo), SUM(detalle_entrada.detent_cantidad), FORMAT(SUM(ar_costo * detent_cantidad), 'currency') AS costo, FORMAT(SUM(ar_precio1 * detent_cantidad), 'currency') AS venta1, FORMAT(SUM(ar_precio2 * detent_cantidad), 'currency') AS venta2 FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) WHERE (detalle_entrada.detent_cod_entrada =" & codalmacenaje & ") GROUP BY detalle_entrada.detent_gondola ORDER BY detalle_entrada.detent_gondola ASC", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("tabla")
            preparar_columnas()
            TextBox4.Text = FormatCurrency(valor_costo(textBox2.Text), 0)
            TextBox5.Text = FormatCurrency(valor_precio1(textBox2.Text), 0)
            TextBox6.Text = FormatCurrency(valor_precio2(textBox2.Text), 0)
            DataGridView1.CurrentCell = DataGridView1.Item(0, DataGridView1.RowCount - 1)
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 6
            preparar_columnas()
            TextBox4.Text = FormatCurrency(0, 0)
            TextBox5.Text = FormatCurrency(0, 0)
            TextBox6.Text = FormatCurrency(0, 0)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_almacenar_gondola.ShowDialog()
        frm_almacenar_gondola.Dispose()
        llenar_lista_gondolas(textBox2.Text)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            Dim v As String = MessageBox.Show("¿Desea Borrar los articulos capturados en la góndola especificada?", "ELIMINAR GÓNDOLA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                clase.borradoautomatico("delete from detalle_entrada where detent_cod_entrada = " & textBox2.Text & " AND detent_gondola = '" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "'")
                llenar_lista_gondolas(textBox2.Text)
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            frm_ver_gondola_almacenar.ShowDialog()
            frm_ver_gondola_almacenar.Dispose()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            Dim v As String = MessageBox.Show("Se cargará al inventario los articulos capturados en este momento ¿Desea continuar?", "GUARDAR ENTRADA DE ARTICULOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                Dim cant As Integer = 0
                'Dim cantingresada As Integer = 0
                clase.consultar("select detent_articulo, SUM(detent_cantidad) as cantidad from detalle_entrada where detent_cod_entrada = " & textBox2.Text & " group by detent_articulo ", "tbl1")
                If clase.dt.Tables("tbl1").Rows.Count > 0 Then
                    Dim x As Integer
                    Dim cantingresada As Integer = 0
                    For x = 0 To clase.dt.Tables("tbl1").Rows.Count - 1
                        With clase.dt.Tables("tbl1")
                            clase.consultar1("SELECT inv_cantidad, inv_ingresado FROM inventario_bodega WHERE (inv_codigoart =" & .Rows(x)("detent_articulo") & ")", "rest")
                            If clase.dt1.Tables("rest").Rows.Count > 0 Then
                                cant = .Rows(x)("cantidad")
                                cantingresada = comprobar_nulidad_de_integer(clase.dt1.Tables("rest").Rows(0)("inv_ingresado"))
                                clase.actualizar("UPDATE inventario_bodega SET inv_cantidad = " & cant + clase.dt1.Tables("rest").Rows(0)("inv_cantidad") & ", inv_ingresado = " & cant + cantingresada & " WHERE (inv_codigoart =" & .Rows(x)("detent_articulo") & ")")
                            Else
                                clase.agregar_registro("INSERT INTO `inventario_bodega`(`inv_codigoart`,`inv_bodega`,`inv_gondola`,`inv_ingresado`,`inv_cantidad`) VALUES ('" & .Rows(x)("detent_articulo") & "','" & ComboBox1.SelectedValue & "',NULL,'" & .Rows(x)("cantidad") & "','" & .Rows(x)("cantidad") & "')")
                            End If
                        End With
                    Next
                    'guardado de ubicaciones
                    clase.consultar("select* from detalle_entrada where detent_cod_entrada = " & textBox2.Text & "", "ubicaciones")
                    Dim z As Short
                    For z = 0 To clase.dt.Tables("ubicaciones").Rows.Count - 1
                        guardar_ubicacion(clase.dt.Tables("ubicaciones").Rows(z)("detent_articulo"), clase.dt.Tables("ubicaciones").Rows(z)("detent_cod_entrada"))
                    Next
                    clase.actualizar("UPDATE cabentrada_bodega SET cabent_estado = TRUE, cabent_hora = '" & Now.ToString("HH:mm:ss") & "' WHERE cabent_codigo = " & textBox2.Text & "")
                    MessageBox.Show("La entrada de articulos se cargó satisfactoriamente al inventario.", "ARTICULOS CARGADOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DataGridView1.DataSource = Nothing
                    DataGridView1.ColumnCount = 6
                    preparar_columnas()
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    ComboBox1.SelectedIndex = -1
                    ComboBox1.Enabled = True
                    TextBox3.Text = ""
                    TextBox3.Enabled = True
                    Button1.Enabled = False
                    Button3.Enabled = False
                    Button6.Enabled = False
                    Button4.Enabled = False
                    Button2.Enabled = True
                    TextBox4.Text = FormatCurrency(0, 0)
                    TextBox5.Text = FormatCurrency(0, 0)
                    TextBox6.Text = FormatCurrency(0, 0)
                    encabezado_entrada()
                End If
            End If
        End If
    End Sub

    Private Sub guardar_ubicacion(articulo As Integer, cabentrada As Integer)
        clase.consultar1("SELECT cabentrada_bodega.cabent_bodega, detalle_entrada.detent_gondola FROM detalle_entrada INNER JOIN cabentrada_bodega ON (detalle_entrada.detent_cod_entrada = cabentrada_bodega.cabent_codigo) WHERE (detalle_entrada.detent_articulo =" & articulo & " AND cabentrada_bodega.cabent_codigo =" & cabentrada & ")", "ubicaciones")
        If clase.dt1.Tables("ubicaciones").Rows.Count > 0 Then
            Dim z As Short
            For z = 0 To clase.dt1.Tables("ubicaciones").Rows.Count - 1
                clase.consultar2("SELECT articulo, bodega, gondola FROM articulos_gondolas WHERE (articulo =" & articulo & " AND bodega =" & clase.dt1.Tables("ubicaciones").Rows(z)("cabent_bodega") & " AND gondola ='" & clase.dt1.Tables("ubicaciones").Rows(z)("detent_gondola") & "')", "gondolas-articulos")
                If clase.dt2.Tables("gondolas-articulos").Rows.Count = 0 Then
                    clase.agregar_registro("insert into articulos_gondolas (articulo, bodega, gondola) values ('" & articulo & "', '" & clase.dt1.Tables("ubicaciones").Rows(z)("cabent_bodega") & "', '" & clase.dt1.Tables("ubicaciones").Rows(z)("detent_gondola") & "')")
                End If
            Next
        End If
    End Sub


    Function valor_costo(entrada As Integer) As Double
        clase.consultar("SELECT SUM(ar_costo * detent_cantidad) AS costo FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) WHERE (detalle_entrada.detent_cod_entrada =" & entrada & ")", "123")
        If IsDBNull(clase.dt.Tables("123").Rows(0)("costo")) Then
            Return 0
            Exit Function
        End If
        Return clase.dt.Tables("123").Rows(0)("costo")
    End Function

    Function valor_precio1(entrada As Integer) As Double
        clase.consultar("SELECT SUM(ar_precio1 * detent_cantidad) AS precio1 FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) WHERE (detalle_entrada.detent_cod_entrada =" & entrada & ")", "123")
        If IsDBNull(clase.dt.Tables("123").Rows(0)("precio1")) Then
            Return 0
            Exit Function
        End If
        Return clase.dt.Tables("123").Rows(0)("precio1")
    End Function

    Function valor_precio2(entrada As Integer) As Double
        clase.consultar("SELECT SUM(ar_precio2 * detent_cantidad) AS precio2 FROM detalle_entrada INNER JOIN articulos ON (detalle_entrada.detent_articulo = articulos.ar_codigo) WHERE (detalle_entrada.detent_cod_entrada =" & entrada & ")", "123")
        If IsDBNull(clase.dt.Tables("123").Rows(0)("precio2")) Then
            Return 0
            Exit Function
        End If
        Return clase.dt.Tables("123").Rows(0)("precio2")
    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frm_suspedidas_almacenaje.ShowDialog()
        frm_suspedidas_almacenaje.Dispose()
    End Sub
End Class
