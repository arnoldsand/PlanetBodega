Public Class frm_creacion_articulos_lotes
    Dim clase As New class_library
    Dim ind_carga As Boolean
    Dim encabezados(30) As String
    Dim encabezados_en_letras(30) As String
    Dim listado(30) As String
    Dim registro_actual As Integer
    Dim cantidad_registros As Integer
    Dim codigo_registro As Integer
    Dim unidades As Integer
    Dim parameter As Integer 'referencia
    Dim parameter_descripcion As Integer
    Dim parameter_unidades As Integer

    Dim ind_operacion As String
    Dim precio1 As Double
    Dim precio2 As Double
    Dim precioant1 As String
    Dim precioant2 As String
    Dim precioanttixy1 As String
    Dim precioanttixy2 As String
    Dim campos As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_creacion_articulos_lotes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        clase.consultar_global("select* from imp_parametros where codigo_importacion = " & cod_importacion & "", "table")
        If clase.dt_global.Tables("table").Rows.Count > 0 Then
            Dim a As Integer
            For a = 0 To frm_crear_articulos_lotes.dataGridView1.Columns.Count - 1
                encabezados(a) = frm_crear_articulos_lotes.dataGridView1.Columns(a).HeaderText
                encabezados_en_letras(a) = convertir_encabezado_comumna_enletra(encabezados(a))
            Next
            cantidad_registros = frm_crear_articulos_lotes.dataGridView1.RowCount
            registro_actual = frm_crear_articulos_lotes.dataGridView1.CurrentCell.RowIndex
            ind_carga = False
            clase.llenar_combo(ComboBox1, "select* from linea1 order by ln1_nombre asc", "ln1_nombre", "ln1_codigo")
            ind_carga = True
            parameter = determinar_columna_referencia(clase.dt_global.Tables("table").Rows(0)("referencia"))
            '______ este procedimiento es solo para mirar cuantos articulos estan asociados a la misma referencia

              'antes de realizar la busqueda completa debo verificar que la referencia escrita
            procesar_referencia(registro_actual)
            '_____

        Else
            MessageBox.Show("No se han creados parametros para generar la ficha de articulos. Por favor creelos y vuelva a intentarlo.", "CREAR PARAMETROS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            GroupBox1.Enabled = False
        End If
        ind_creacion_colores = False
    End Sub

    Sub mostrar_datos(ByVal fila As Integer)
        Dim i As Integer
        i = 0
        ind_operacion = "crear"
        textBox2.Text = frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value & ""
        TextBox1.Text = frm_crear_articulos_lotes.dataGridView1.Item(parameter_descripcion, fila).Value & ""
        If IsDBNull(frm_crear_articulos_lotes.dataGridView1.Item(parameter_unidades, fila).Value) Then
            unidades = vbEmpty
        Else
            unidades = frm_crear_articulos_lotes.dataGridView1.Item(parameter_unidades, fila).Value
        End If
        GroupBox1.Enabled = True
        TextBox7.Text = cod_importacion
        Label17.Visible = False
        ComboBox6.Visible = False
        ' codigo_registro = clase.dt_global.Tables("table")(0)("codigo_detalle")
        If fila + 1 = cantidad_registros Then Button4.Enabled = False
        If fila = 0 Then Button3.Enabled = False
    End Sub

    Function convertir_encabezado_comumna_enletra(ByVal txt As String) As String
        If ind_cargador_lote = True Then
            Dim a As Short
            convertir_encabezado_comumna_enletra = ""
            For a = 0 To frm_crear_articulos_lotes.dataGridView1.Columns.Count - 1
                listado(a) = frm_crear_articulos_lotes.dataGridView1.Columns(a).HeaderText
            Next
            a = 0
            Do While a <= frm_crear_articulos_lotes.dataGridView1.Columns.Count - 1
                If listado(a) = txt Then
                    convertir_encabezado_comumna_enletra = campos(a)
                    Exit Do
                End If
                a += 1
            Loop
        Else
            convertir_encabezado_comumna_enletra = txt
        End If
    End Function


    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        clase.enter(e)
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        clase.enter(e)
    End Sub

    Private Sub ComboBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox3.KeyPress
        clase.enter(e)
    End Sub

    Private Sub ComboBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox4.KeyPress
        clase.enter(e)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox2.Enabled = True
            clase.llenar_combo(ComboBox2, "SELECT* FROM sublinea1 WHERE sl1_ln1codigo = " & ComboBox1.SelectedValue & " ORDER BY sl1_nombre asc", "sl1_nombre", "sl1_codigo")
            ComboBox3.DataSource = Nothing
            ComboBox3.Enabled = False
            ComboBox4.DataSource = Nothing
            ComboBox4.Enabled = False
            ComboBox5.DataSource = Nothing
            ComboBox5.Enabled = False
            ind_carga = True
        End If
    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox3.Enabled = True
            clase.llenar_combo(ComboBox3, "SELECT* FROM sublinea2 WHERE sl2_sl1codigo = " & ComboBox2.SelectedValue & "  ORDER BY sl2_nombre asc", "sl2_nombre", "sl2_codigo")
            ComboBox4.DataSource = Nothing
            ComboBox4.Enabled = False
            ComboBox5.DataSource = Nothing
            ComboBox5.Enabled = False
            ind_carga = True
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox4.Enabled = True
            clase.llenar_combo(ComboBox4, "SELECT* FROM sublinea3 WHERE sl3_sl2codigo = " & ComboBox3.SelectedValue & "  ORDER BY sl3_nombre asc", "sl3_nombre", "sl3_codigo")
            ComboBox5.DataSource = Nothing
            ComboBox5.Enabled = False
            ind_carga = True
        End If
    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.GotFocus
        TextBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textBox2.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox8.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.LostFocus
        textBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub TextBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.GotFocus
        TextBox4.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.GotFocus
        TextBox5.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox7_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.BackColor = Color.White
    End Sub

    Private Sub TextBox8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox8_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub TextBox9_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.GotFocus
        TextBox9.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox9_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        limpiar()
        ind_creacion_colores = False
    End Sub

    Sub limpiar()
        ind_carga = False
        TextBox1.Text = ""
        textBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
        textBox2.Focus()
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ReDim colores(0)
        ReDim arraytallas(0)
        ReDim cantidades(0)
        ind_carga = True
    End Sub

    Private Sub dateTimePicker1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dateTimePicker1.KeyPress
        clase.enter(e)
    End Sub

   
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Referencia") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Descripción") = False Then Exit Sub
        If clase.validar_combobox(ComboBox1, "Linea") = False Then Exit Sub
        If clase.validar_combobox(ComboBox2, "Sub-Linea1") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox3, "Precio Costo") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox4, "Precio Planet") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox5, "Precio Tixy") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox6, "Creado Por") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox8, "Tasa Dcto") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox9, "Proveedor") = False Then Exit Sub
        If ind_creacion_colores = False Then
            MessageBox.Show("No se han establecido patrones de colores para esta referencia. Por favor establezcalos y vuelva a intentarlo.", "PATRONES DE COLORES", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If Val(TextBox8.Text) > 100 Then
            MessageBox.Show("La tasa de descuento no puedo ser superior al 100%.", "DESCUENTO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox8.Focus()
            Exit Sub
        End If
        Dim c As Short
        Dim sum As Integer = 0
        For c = 0 To cantidades.Length - 1
            sum = sum + cantidades(c)
        Next
        If sum = 0 Then
            MessageBox.Show("Debe especificar las cantidades para las tallas y/o colores que desea crear-editar.", "ESPECIFICAR CANTIDADES", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim consecutivo As Long
        Dim fecha As String = dateTimePicker1.Value.ToString("yyyy-MM-dd")
        Dim cmb3, cmb4, cmb5 As String
        If ComboBox3.Text = "" Then
            cmb3 = "NULL"
        Else
            cmb3 = ComboBox3.SelectedValue
        End If
        If ComboBox4.Text = "" Then
            cmb4 = "NULL"
        Else
            cmb4 = ComboBox4.SelectedValue
        End If
        If ComboBox5.Text = "" Then
            cmb5 = "NULL"
        Else
            cmb5 = ComboBox5.SelectedValue
        End If
        Select Case ind_operacion
            Case "crear"
                Dim a As Integer
                For a = 0 To colores.Length - 1
                    consecutivo = clase.generar_codigo_tabla_articulos
                    clase.agregar_registro("INSERT INTO `articulos`(`ar_codigo`,`ar_codigobarras`,`ar_referencia`,`ar_descripcion`,`ar_color`,`ar_talla`,`ar_linea`,`ar_sublinea1`,`ar_sublinea2`,`ar_sublinea3`,`ar_sublinea4`,`ar_costo`,`ar_precio1`,`ar_precio2`,`ar_precioanterior`,`ar_fechaingreso`,`ar_precioanterior2`,`ar_creadopor`,`ar_fecdescontinua`,`ar_activo`,`ar_tasadesc`) VALUES ('" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & UCase(textBox2.Text) & "','" & UCase(TextBox1.Text) & "','" & colores(a) & "','" & arraytallas(a) & "','" & ComboBox1.SelectedValue & "','" & ComboBox2.SelectedValue & "'," & cmb3 & "," & cmb4 & "," & cmb5 & ",'" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "',NULL,'" & fecha & "',NULL,'" & TextBox6.Text & "',NULL, True,'" & TextBox8.Text & "')")
                    clase.agregar_registro("INSERT INTO entradamercancia (com_codigoimp, com_codigoart, com_unidades) values (" & cod_importacion & ", " & consecutivo & ", " & cantidades(a) & ")")
                    clase.agregar_registro("INSERT INTO detalle_proveedores_articulos (codigo_articulo, codigo_proveedor, codigo_importacion) values ('" & consecutivo & "', '" & TextBox9.Text & "', '" & TextBox7.Text & "')")
                Next
                clase.actualizar("UPDATE `detalleimportacion` SET `procesado`='1' WHERE `item`=" & frm_crear_articulos_lotes.dataGridView1.Item(26, frm_crear_articulos_lotes.dataGridView1.CurrentCell.RowIndex).Value & "")
            Case "actualizar"
                Dim z As Short 'aqui voy crear codigos nuevos x si las cantidades y las tallas/colores son aumentadas al actualioar la referencia
                For z = 0 To colores.Length - 1
                    clase.consultar("select* from articulos where ar_color = " & colores(z) & " and ar_talla = " & arraytallas(z) & " and ar_linea = " & ComboBox6.Items(ComboBox6.SelectedIndex).linea & " and ar_sublinea1 = " & ComboBox6.Items(ComboBox6.SelectedIndex).sublinea1 & " and ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, registro_actual).Value & "'" & generar_cadena_sql_apartir_del_combobox_de_coincidencias(ComboBox6.Items(ComboBox6.SelectedIndex).sublinea2, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea3, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea4), "busqueda")
                    If clase.dt.Tables("busqueda").Rows.Count = 0 Then
                        consecutivo = clase.generar_codigo_tabla_articulos
                        clase.agregar_registro("insert into articulos (`ar_codigo`,`ar_codigobarras`,`ar_referencia`,`ar_descripcion`,`ar_color`,`ar_talla`,`ar_linea`,`ar_sublinea1`,`ar_sublinea2`,`ar_sublinea3`,`ar_sublinea4`,`ar_costo`,`ar_precio1`,`ar_precio2`,`ar_precioanterior`,`ar_fechaingreso`,`ar_precioanterior2`,`ar_creadopor`,`ar_fecdescontinua`,`ar_activo`,`ar_tasadesc`,`ar_precioanttixy1`,`ar_precioanttixy2`) values ('" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & UCase(textBox2.Text) & "','" & UCase(TextBox1.Text) & "','" & colores(z) & "','" & arraytallas(z) & "','" & ComboBox1.SelectedValue & "','" & ComboBox2.SelectedValue & "'," & cmb3 & "," & cmb4 & "," & cmb5 & ",'" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "'," & precioant1 & ",'" & fecha & "'," & precioant2 & ",'" & TextBox6.Text & "',NULL, True,'" & TextBox8.Text & "'," & precioanttixy1 & "," & precioanttixy2 & ")")
                    End If
                Next
                redimencionar_array_de_codigosdebarras()

                Dim preciosactuales() As Double = {clase.dt_global.Tables("resultados").Rows(0)("ar_precio1"), clase.dt_global.Tables("resultados").Rows(0)("ar_precio2")}
                Dim preciosafijar() As Double = {TextBox4.Text, TextBox5.Text}
                Dim f As Short
                For f = 0 To codigos_de_barra.Length - 1
                    clase.actualizar("update articulos set ar_tasadesc = " & TextBox8.Text & ", ar_costo = " & TextBox3.Text & ", ar_precio1 = " & TextBox4.Text & ", ar_precio2 = " & TextBox5.Text & precios_anteriores(preciosactuales, clase.dt_global.Tables("resultados"), preciosafijar) & " where ar_codigo = " & codigos_de_barra(f) & "")
                    clase.agregar_registro("INSERT INTO entradamercancia (com_codigoimp, com_codigoart, com_unidades) values (" & cod_importacion & ", " & codigos_de_barra(f) & ", " & cantidades(f) & ")")
                    clase.agregar_registro("INSERT INTO detalle_proveedores_articulos (codigo_articulo, codigo_proveedor, codigo_importacion) values ('" & codigos_de_barra(f) & "', '" & TextBox9.Text & "', '" & cod_importacion & "')")
                Next
                clase.actualizar("UPDATE `detalleimportacion` SET `procesado`='1' WHERE `item`=" & frm_crear_articulos_lotes.dataGridView1.Item(26, frm_crear_articulos_lotes.dataGridView1.CurrentCell.RowIndex).Value & "")
        End Select
        frm_crear_articulos_lotes.llenar_datagrid()
        cantidad_registros = frm_crear_articulos_lotes.dataGridView1.RowCount
        Button4_Click(Nothing, Nothing)
        ind_creacion_colores = False
        'guardar los proveedores
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox5.Enabled = True
            clase.llenar_combo(ComboBox5, "SELECT* FROM sublinea4 WHERE sl4_sl3codigo = " & ComboBox4.SelectedValue & "  ORDER BY sl4_nombre asc", "sl4_nombre", "sl4_codigo")
            ind_carga = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        registro_actual += 1
        Button3.Enabled = True
        If registro_actual > cantidad_registros - 1 Then
            registro_actual = cantidad_registros - 1
            Button4.Enabled = False
            limpiar()
            procesar_referencia(registro_actual)
            ind_creacion_colores = False
        Else
            limpiar()
            procesar_referencia(registro_actual)
            ind_creacion_colores = False
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        registro_actual -= 1
        Button4.Enabled = True
        If registro_actual < 0 Then
            registro_actual = 0
            Button3.Enabled = False
            limpiar()
            procesar_referencia(registro_actual)
            ind_creacion_colores = False
        Else
            limpiar()
            procesar_referencia(registro_actual)
            ind_creacion_colores = False
        End If
    End Sub

    Private Sub button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button7.Click
        frm_proveedores.ShowDialog()
        frm_proveedores.Dispose()
        Button2.Focus()
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        frm_colores.ShowDialog()
        frm_colores.Dispose()
    End Sub

    Class items
        Dim dt7 As DataRow
        Dim mostar As String


        Public Sub New(ByVal fila As System.Data.DataRow)
            dt7 = fila
        End Sub

        Public Function sublinea1() As String
            Return dt7("ar_sublinea1")
        End Function

        Public Function sublinea2() As String
            If IsDBNull(dt7("ar_sublinea2")) Then
                Return Nothing
            Else
                Return dt7("ar_sublinea2")
            End If
        End Function

        Public Function sublinea3() As String
            If IsDBNull(dt7("ar_sublinea3")) Then
                Return Nothing
            Else
                Return dt7("ar_sublinea3")
            End If
        End Function

        Public Function sublinea4() As String
            If IsDBNull(dt7("ar_sublinea4")) Then
                Return Nothing
            Else
                Return dt7("ar_sublinea4")
            End If
        End Function

        Public Function linea() As String
            Return dt7("ar_linea")
        End Function

        Public Overrides Function ToString() As String
            Return dt7("refart")
        End Function
    End Class

    Private Sub redimencionar_array_de_codigosdebarras()
        clase.consultar("select* from articulos where ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1(parameter, registro_actual).Value & "' and ar_linea = " & ComboBox6.Items(ComboBox6.SelectedIndex).linea & " and ar_sublinea1 = " & ComboBox6.Items(ComboBox6.SelectedIndex).sublinea1 & generar_cadena_sql_apartir_del_combobox_de_coincidencias(ComboBox6.Items(ComboBox6.SelectedIndex).sublinea2, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea3, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea4), "resultadosbusqueda")
        If clase.dt.Tables("resultadosbusqueda").Rows.Count > 0 Then
            Dim b As Short
            ReDim codigos_de_barra(clase.dt.Tables("resultadosbusqueda").Rows.Count - 1)
            For b = 0 To clase.dt.Tables("resultadosbusqueda").Rows.Count - 1
                codigos_de_barra(b) = clase.dt.Tables("resultadosbusqueda").Rows(0)("ar_codigo")
            Next
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        If ComboBox6.Text = "" Then
            Exit Sub
        End If
        Dim con_sql As String = "select* from articulos where ar_linea = " & ComboBox6.Items(ComboBox6.SelectedIndex).linea & " and ar_sublinea1 = " & ComboBox6.Items(ComboBox6.SelectedIndex).sublinea1 & generar_cadena_sql_apartir_del_combobox_de_coincidencias(ComboBox6.Items(ComboBox6.SelectedIndex).sublinea2, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea3, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea4) & " and ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, registro_actual).Value & "'"
        mostrar_referencia_ya_existente(con_sql)
    End Sub


    Private Sub mostrar_referencia_ya_existente(ByVal consulta As String)
        ind_operacion = "actualizar"
        limpiar()
        GroupBox1.Enabled = True
        clase.dt_global.Tables.Clear()
        clase.consultar_global(consulta, "resultados")
        If clase.dt_global.Tables("resultados").Rows.Count > 0 Then
            ind_carga = False
            textBox2.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_referencia")
            TextBox1.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_descripcion")
            ComboBox1.SelectedValue = clase.dt_global.Tables("resultados").Rows(0)("ar_linea")
            ComboBox2.Enabled = True
            clase.llenar_combo(ComboBox2, "SELECT* FROM sublinea1 WHERE sl1_ln1codigo = " & ComboBox1.SelectedValue & " ORDER BY sl1_nombre asc", "sl1_nombre", "sl1_codigo")
            ComboBox2.SelectedValue = clase.dt_global.Tables("resultados").Rows(0)("ar_sublinea1")
            If IsDBNull(clase.dt_global.Tables("resultados").Rows(0)("ar_sublinea2")) Then
                ComboBox3.SelectedIndex = -1
            Else
                ComboBox3.Enabled = True
                clase.llenar_combo(ComboBox3, "SELECT* FROM sublinea2 WHERE sl2_sl1codigo = " & ComboBox2.SelectedValue & "  ORDER BY sl2_nombre asc", "sl2_nombre", "sl2_codigo")
                ComboBox3.SelectedValue = clase.dt_global.Tables("resultados").Rows(0)("ar_sublinea2")
            End If
            If IsDBNull(clase.dt_global.Tables("resultados").Rows(0)("ar_sublinea3")) Then
                ComboBox4.SelectedIndex = -1
            Else
                ComboBox4.Enabled = True
                clase.llenar_combo(ComboBox4, "SELECT* FROM sublinea3 WHERE sl3_sl2codigo = " & ComboBox3.SelectedValue & "  ORDER BY sl3_nombre asc", "sl3_nombre", "sl3_codigo")
                ComboBox4.SelectedValue = clase.dt_global.Tables("resultados").Rows(0)("ar_sublinea3")
            End If
            If IsDBNull(clase.dt_global.Tables("resultados").Rows(0)("ar_sublinea4")) Then
                ComboBox5.SelectedIndex = -1
            Else
                ComboBox5.Enabled = True
                clase.llenar_combo(ComboBox5, "SELECT* FROM sublinea4 WHERE sl4_sl3codigo = " & ComboBox4.SelectedValue & "  ORDER BY sl4_nombre asc", "sl4_nombre", "sl4_codigo")
                ComboBox5.SelectedValue = clase.dt_global.Tables("resultados").Rows(0)("ar_sublinea4")
            End If
            ind_carga = True
            TextBox3.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_costo")
            TextBox4.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_precio1")
            TextBox5.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_precio2")
            dateTimePicker1.Value = clase.dt_global.Tables("resultados").Rows(0)("ar_fechaingreso")
            TextBox6.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_creadopor")
            TextBox8.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_tasadesc")
            TextBox7.Text = cod_importacion
            precio1 = clase.dt_global.Tables("resultados").Rows(0)("ar_precio1")
            precio2 = clase.dt_global.Tables("resultados").Rows(0)("ar_precio2")
            precioant1 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanterior"))
            precioant2 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanterior2"))
            precioanttixy1 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy1"))
            precioanttixy2 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy2"))
            '   precioanttixy1 = clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy1")
            '  precioanttixy2 = clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy2")
            construir_matrices_colores_tallas_codigosdebarra(clase.dt_global.Tables("resultados"))
        End If
    End Sub

   

    Private Function determinar_columna_referencia(ByVal referencia As String) As String
        Dim a As Integer
        Dim columna As Integer
        For a = 0 To 30
            If encabezados_en_letras(a) = referencia Then
                columna = a
            End If
        Next
        Return columna
    End Function

    Private Function determinar_columna_descripcion(ByVal descripcion As String) As String
        Dim a As Integer
        Dim columna As Integer
        For a = 0 To 30
            If encabezados_en_letras(a) = descripcion Then
                columna = a
            End If
        Next
        Return columna
    End Function

    Private Function determinar_columna_unidades(ByVal unidades As String) As String
        Dim a As Integer
        Dim columna As Integer
        For a = 0 To 30
            If encabezados_en_letras(a) = unidades Then
                columna = a
            End If
        Next
        Return columna
    End Function

    Private Sub procesar_referencia(ByVal fila As Integer)
        clase.consultar("select* from articulos where ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value & "'", "result")
        If clase.dt.Tables("result").Rows.Count > 0 Then
            GroupBox1.Enabled = False

            'aqui realizo la busqueda completa
            clase.consultar("SELECT DISTINCT articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4, CONCAT(linea1.ln1_nombre, ', ', sublinea1.sl1_nombre) AS refart FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (sublinea2.sl2_sl1codigo = sublinea1.sl1_codigo) AND (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) AND (sublinea1.sl1_ln1codigo = linea1.ln1_codigo) LEFT JOIN sublinea3 ON (sublinea3.sl3_sl2codigo = sublinea2.sl2_codigo) AND (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (sublinea4.sl4_sl3codigo = sublinea3.sl3_codigo) AND (articulos.ar_sublinea4 = sublinea4.sl4_codigo) WHERE (articulos.ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value & "')", "tableresult11")
            If clase.dt.Tables("tableresult11").Rows.Count > 0 Then
                Dim cantidad_registros As Integer = clase.dt.Tables("tableresult11").Rows.Count
                Select Case cantidad_registros
                    Case Is > 1
                        Label17.Visible = True
                        ComboBox6.Visible = True
                        Dim f As Short
                        ComboBox6.Items.Clear()
                        For f = 0 To clase.dt.Tables("tableresult11").Rows.Count - 1
                            ComboBox6.Items.Add(New items(clase.dt.Tables("tableresult11").Rows(f)))
                        Next
                        MessageBox.Show("Se encontraron varios articulos diferentes con la misma referencia, por favor especifique a cual hace referencia antes de continuar.", "REFERENCIA YA EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Case Is = 1
                        Label17.Visible = False
                        ComboBox6.Visible = False
                        Dim complementosql2, complementosql3, complementosql4 As String
                        If IsDBNull(clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea2")) Then
                            complementosql2 = " AND ar_sublinea2 IS NULL"
                        Else
                            complementosql2 = " AND ar_sublinea2 = " & clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea2")
                        End If
                        If IsDBNull(clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea3")) Then
                            complementosql3 = " AND ar_sublinea3 IS NULL"
                        Else
                            complementosql3 = " AND ar_sublinea3 = " & clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea3")
                        End If
                        If IsDBNull(clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea4")) Then
                            complementosql4 = " AND ar_sublinea4 IS NULL"
                        Else
                            complementosql4 = " AND ar_sublinea4 = " & clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea4")
                        End If
                        Dim con_sql As String = "select* from articulos where ar_linea = " & clase.dt.Tables("tableresult11").Rows(0)("ar_linea") & " and ar_sublinea1 = " & clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea1") & complementosql2 & complementosql3 & complementosql4 & " and ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value & "'"
                        mostrar_referencia_ya_existente(con_sql)
                End Select
                ind_edicion_colores = False
            End If
        Else
            ind_edicion_colores = True
            mostrar_datos(registro_actual)
        End If

    End Sub
End Class