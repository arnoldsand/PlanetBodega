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
    Dim articulo As New tipoarticulo
    Dim campos As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_creacion_articulos_lotes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ind_creacion_colores = False
        cod_importacion = hallar_codigo_importacion(frm_crear_articulos_lotes.textBox2.Text)
        cantidad_registros = frm_crear_articulos_lotes.dataGridView1.RowCount
            registro_actual = frm_crear_articulos_lotes.dataGridView1.CurrentCell.RowIndex
            ind_carga = False
        clase.llenar_combo(ComboBox1, "select* from linea1 order by ln1_nombre asc", "ln1_nombre", "ln1_codigo")
        ind_carga = True
        '______ este procedimiento es solo para mirar cuantos articulos estan asociados a la misma referencia
              'antes de realizar la busqueda completa debo verificar que la referencia escrita
        procesar_referencia(registro_actual)
        '_____     
    End Sub

    Private Function codigo_proveedor_apartir_del_mov_salida(movimiento As Integer) As Short 'esta funcion obtiene el codigo del proveedor apartir del codigo de la salida
        clase.consultar("SELECT detalle_importacion_cabcajas.det_codigoproveedor FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN detsalidas_mercancia ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN cabsalidas_mercancia ON (detsalidas_mercancia.det_salidacodigo = cabsalidas_mercancia.cabsal_cod) WHERE (detsalidas_mercancia.det_codigo =" & movimiento & ")", "codproveedor")
        Return clase.dt.Tables("codproveedor").Rows(0)("det_codigoproveedor")
    End Function

    Sub mostrar_datos(ByVal fila As Integer)
        Dim i As Integer
        i = 0
        ind_operacion = "crear"
        Button7.Enabled = False
        textBox2.Text = frm_crear_articulos_lotes.dataGridView1.Item(0, fila).Value & ""
        TextBox1.Text = frm_crear_articulos_lotes.dataGridView1.Item(1, fila).Value & ""
        If IsDBNull(frm_crear_articulos_lotes.dataGridView1.Item(3, fila).Value) Then
            unidades = vbEmpty
        Else
            unidades = frm_crear_articulos_lotes.dataGridView1.Item(3, fila).Value
        End If
        GroupBox1.Enabled = True
        TextBox7.Text = cod_importacion
        TextBox8.Text = "0"
        TextBox9.Text = codigo_proveedor_apartir_del_mov_salida(frm_crear_articulos_lotes.dataGridView1.Item(5, frm_crear_articulos_lotes.dataGridView1.CurrentRow.Index).Value)
        TextBox11.Text = "0"
        Label17.Visible = False
        ComboBox6.Visible = False
        TextBox12.Text = extraer_costo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)
        establecer_estado_para_controles_para_impedir_edicion(True)
        ' codigo_registro = clase.dt_global.Tables("table")(0)("codigo_detalle")
        If fila + 1 = cantidad_registros Then Button4.Enabled = False
        If fila = 0 Then Button3.Enabled = False
        PictureBox1.Image = Nothing

        establecer_colores_tallas(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)     ''creador automatico de tallas y colores
        If frm_crear_articulos_lotes.dataGridView1.Item(4, registro_actual).Value = False Then
            Button2.Enabled = True

            Button6.Enabled = True
            Button9.Enabled = True
            ' button7.Enabled = True
        End If
        'impuestos
        ReDim impuestos(0)
        impuestos(0) = hallar_impuesto_predeterminado()
    End Sub

    Function extraer_costo(precodigo As Short) As Double
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_costopesos_x_pieza FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detsalidas_mercancia.det_codigo =" & precodigo & ")", "costo2")
        Return clase.dt.Tables("costo2").Rows(0)("detcab_costopesos_x_pieza")
    End Function

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
        textBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
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

    Private Sub TextBox12_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox12.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox12_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox12.LostFocus
        TextBox12.BackColor = Color.White
    End Sub

    Private Sub TextBox12_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox12.GotFocus
        TextBox12.BackColor = System.Drawing.SystemColors.GradientActiveCaption
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
        PictureBox1.Image = Nothing
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
        If clase.validar_cajas_text(TextBox3, "Precio Costo 1") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox12, "Precio Costo 2") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox4, "Precio Planet") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox5, "Precio Tixy") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox6, "Creado Por") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox8, "Tasa Dcto") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox9, "Proveedor") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox11, "Dañado") = False Then Exit Sub
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
        Dim fechaultimamodificacion As String = Now.ToString("yyyy-MM-dd")
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
        Dim extension As String = System.IO.Path.GetExtension(TextBox10.Text)
        Select Case ind_operacion
            Case "crear"
                Dim a As Integer
                For a = 0 To colores.Length - 1
                    consecutivo = clase.generar_codigo_tabla_articulos
                    If TextBox10.Text = ruta_foto() & "\" & consecutivo & extension Then
                        MessageBox.Show("La ruta de la foto de destino no puede ser igual a la de la foto de origen. Especifique otra y vuelva a intentarlo.", "CREACIÓN DE FOTOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If
                    ' creo la foto si el usuario especifíca alguna
                    If TextBox10.Text <> "" Then
                        crear_foto_grande(TextBox10.Text, ruta_foto() & "\" & consecutivo & extension)     'linea de desbloqueo
                        crear_foto_pequeño(ruta_foto() & "\" & consecutivo & extension, ruta_foto() & "\" & consecutivo & "mini.jpg")      'linea de desbloqueo
                        ' crear_foto_grande(TextBox10.Text, ruta_google_drive() & "\" & consecutivo & extension)
                    End If
                    clase.agregar_registro("INSERT INTO `articulos`(`ar_codigo`,`ar_codigobarras`,`ar_codigo2`,`ar_codigobarras2`,`ar_referencia`,`ar_descripcion`,`ar_color`,`ar_talla`,`ar_linea`,`ar_sublinea1`,`ar_sublinea2`,`ar_sublinea3`,`ar_sublinea4`,`ar_costo`,`ar_costo2`,`ar_precio1`,`ar_precio2`,`ar_precioanterior`,`ar_fechaingreso`,`ar_precioanterior2`,`ar_creadopor`,`ar_fecdescontinua`,`ar_activo`,`ar_tasadesc`,`ar_ultimamodificacion`,`ar_foto`) VALUES ('" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & UCase(textBox2.Text) & "','" & UCase(TextBox1.Text) & "','" & colores(a) & "','" & arraytallas(a) & "','" & ComboBox1.SelectedValue & "','" & ComboBox2.SelectedValue & "'," & cmb3 & "," & cmb4 & "," & cmb5 & ",'" & Str(TextBox12.Text) & "','" & Str(TextBox3.Text) & "', '" & Str(TextBox4.Text) & "','" & Str(TextBox5.Text) & "',NULL,'" & fecha & "',NULL,'" & TextBox6.Text & "',NULL, True,'" & TextBox8.Text & "','" & fechaultimamodificacion & "','" & Replace(ruta_foto() & "\" & consecutivo & extension, "\", "\\") & "')")
                    If cantidades(a) > 0 Then

                        clase.agregar_registro("INSERT INTO entradamercancia (com_codigoimp, com_codigoart, com_unidades, com_danado, com_operario) values (" & cod_importacion & ", " & consecutivo & ", " & cantidades(a) & ", " & TextBox11.Text & ", '" & UCase(TextBox6.Text) & "')")
                        clase.agregar_registro("INSERT INTO detalle_proveedores_articulos (codigo_articulo, codigo_proveedor) values ('" & consecutivo & "', '" & TextBox9.Text & "')")
                        clase.agregar_registro("INSERT INTO `asociaciones_codigos`(`asc_precodref`,`asc_postcodart`) VALUES ( '" & obtener_codigo_apartir_de_consecutivo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value) & "','" & consecutivo & "')")
                        ' MsgBox(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)
                        clase.agregar_registro("INSERT INTO registros_de_importaciones (articulo, registro, fecha) VALUES ('" & consecutivo & "', " & obtener_registro_importacion_apartir_de_consecutivo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value) & ", " & obtener_fecha_importacion_apartir_de_consecutivo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value) & ")")
                        'creo la foto solo si el usuario ha espcificado alguna
                        If impuestos.Length > 0 Then
                            Dim x As Short
                            For x = 0 To impuestos.Length - 1
                                clase.consultar("select* from impuestos_articulos where impart_articulo = " & consecutivo & " and impart_impuestos = " & impuestos(x) & "", "buscarimp")
                                If clase.dt.Tables("buscarimp").Rows.Count = 0 Then
                                    clase.agregar_registro("INSERT INTO `impuestos_articulos`(`impart_codigo`,`impart_articulo`,`impart_impuestos`) VALUES ( NULL,'" & consecutivo & "','" & impuestos(x) & "')")
                                End If
                            Next
                        Else
                            clase.borradoautomatico("delete from impuestos_articulos where impart_articulo = " & consecutivo & "")
                        End If
                        CrearMovimientoCompraLovePOSImportacion(cod_importacion, TextBox9.Text, consecutivo, cantidades(a), UCase(TextBox6.Text))
                    End If
                Next
                marcar_item_procesado(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)
            Case "actualizar"
                Dim z As Short 'aqui voy crear codigos nuevos x si las cantidades y las tallas/colores son aumentadas al actualizar la referencia
                For z = 0 To colores.Length - 1
                    If ComboBox6.Visible = True Then ' aquí cuando hay varios tipos de articulos con la misma referencia (osea el combobox esta visible)
                        clase.consultar("select* from articulos where ar_color = " & colores(z) & " and ar_talla = " & arraytallas(z) & " and ar_linea = " & ComboBox6.Items(ComboBox6.SelectedIndex).linea & " and ar_sublinea1 = " & ComboBox6.Items(ComboBox6.SelectedIndex).sublinea1 & " and ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, registro_actual).Value & "'" & generar_cadena_sql_apartir_del_combobox_de_coincidencias(ComboBox6.Items(ComboBox6.SelectedIndex).sublinea2, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea3, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea4), "busqueda")
                    Else                             ' aqui cuando hay un (1) solo tipo de articulos con la misma referencia (osea cuando el combobox no esta visible)
                        clase.consultar("select* from articulos where ar_color = " & colores(z) & " and ar_talla = " & arraytallas(z) & " and ar_linea = " & articulo.linea1 & " and ar_sublinea1 = " & articulo.sublinea1 & " and ar_referencia = '" & articulo.referencia & "'" & generar_cadena_sql_apartir_del_combobox_de_coincidencias(verificar_nulidad_vacio(articulo.sublinea2), verificar_nulidad_vacio(articulo.sublinea3), verificar_nulidad_vacio(articulo.sublinea4)), "busqueda")
                    End If
                    If clase.dt.Tables("busqueda").Rows.Count = 0 Then
                        consecutivo = clase.generar_codigo_tabla_articulos
                        clase.agregar_registro("insert into articulos (`ar_codigo`,`ar_codigobarras`,`ar_codigo2`,`ar_codigobarras2`,`ar_referencia`,`ar_descripcion`,`ar_color`,`ar_talla`,`ar_linea`,`ar_sublinea1`,`ar_sublinea2`,`ar_sublinea3`,`ar_sublinea4`,`ar_costo`,`ar_costo2`,`ar_precio1`,`ar_precio2`,`ar_precioanterior`,`ar_fechaingreso`,`ar_precioanterior2`,`ar_creadopor`,`ar_fecdescontinua`,`ar_activo`,`ar_tasadesc`,`ar_precioanttixy1`,`ar_precioanttixy2`,`ar_ultimamodificacion`,`ar_foto`) values ('" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & UCase(textBox2.Text) & "','" & UCase(TextBox1.Text) & "','" & colores(z) & "','" & arraytallas(z) & "','" & ComboBox1.SelectedValue & "','" & ComboBox2.SelectedValue & "'," & cmb3 & "," & cmb4 & "," & cmb5 & ",'" & Str(TextBox12.Text) & "','" & Str(TextBox3.Text) & "','" & TextBox4.Text & "','" & TextBox5.Text & "'," & precioant1 & ",'" & fecha & "'," & precioant2 & ",'" & TextBox6.Text & "',NULL, True,'" & TextBox8.Text & "'," & precioanttixy1 & "," & precioanttixy2 & ",'" & fechaultimamodificacion & "','" & Replace(ruta_foto() & "\" & consecutivo & extension, "\", "\\") & "')")
                        If impuestos.Length > 0 Then
                            Dim x As Short
                            For x = 0 To impuestos.Length - 1
                                clase.consultar("select* from impuestos_articulos where impart_articulo = " & consecutivo & " and impart_impuestos = " & impuestos(x) & "", "buscarimp")
                                If clase.dt.Tables("buscarimp").Rows.Count = 0 Then
                                    clase.agregar_registro("INSERT INTO `impuestos_articulos`(`impart_codigo`,`impart_articulo`,`impart_impuestos`) VALUES ( NULL,'" & consecutivo & "','" & impuestos(x) & "')")
                                End If
                            Next
                        Else
                            clase.borradoautomatico("delete from impuestos_articulos where impart_articulo = " & consecutivo & "")
                        End If
                    End If
                Next
                redimencionar_array_de_codigosdebarras() ' aqui rescato los codigos de barra que pude haber creado nuevos (x si los hubiese hecho)(osea incluyo los codigos de barras nuevos que creé en el insert articulos anterior)
                Dim preciosactuales() As Double = {clase.dt_global.Tables("resultados").Rows(0)("ar_precio1"), clase.dt_global.Tables("resultados").Rows(0)("ar_precio2")}
                Dim preciosafijar() As Double = {TextBox4.Text, TextBox5.Text}
                Dim f As Short
                For f = 0 To codigos_de_barra.Length - 1
                    'creo la foto solo si el usuario ha espcificado alguna
                    If TextBox10.Text <> "" Then
                        If TextBox10.Text <> ruta_foto() & "\" & codigos_de_barra(f) & extension Then
                            crear_foto_grande(TextBox10.Text, ruta_foto() & "\" & codigos_de_barra(f) & extension)
                            crear_foto_pequeño(ruta_foto() & "\" & codigos_de_barra(f) & extension, ruta_foto() & "\" & codigos_de_barra(f) & "mini.jpg")
                        End If
                    End If
                    clase.actualizar("update articulos set ar_tasadesc = " & TextBox8.Text & ", ar_costo = " & Str(TextBox12.Text) & ", ar_costo2 = " & Str(TextBox3.Text) & ", ar_ultimamodificacion = '" & fechaultimamodificacion & "', ar_precio1 = " & TextBox4.Text & ", ar_precio2 = " & TextBox5.Text & precios_anteriores(preciosactuales, clase.dt_global.Tables("resultados"), preciosafijar) & " where ar_codigo = " & codigos_de_barra(f) & "")
                    'entrada de mercancia
                    If cantidades(f) > 0 Then
                        'verifico si existe en la tabla entrada de mercancia, si existe actualizo cantidad y sino lo agrego
                        clase.consultar2("select* FROM entradamercancia WHERE com_codigoart = " & codigos_de_barra(f) & " AND com_codigoimp = " & cod_importacion & "", "entrada")
                        If clase.dt2.Tables("entrada").Rows.Count > 0 Then
                            Dim cantidad As Short = clase.dt2.Tables("entrada").Rows(0)("com_unidades")
                            clase.actualizar("UPDATE entradamercancia SET com_unidades = " & cantidad + cantidades(f) & " WHERE com_codigoart = " & codigos_de_barra(f) & " AND com_codigoimp = " & cod_importacion & "")
                        Else
                            clase.agregar_registro("INSERT INTO entradamercancia (com_codigoimp, com_codigoart, com_unidades, com_danado, com_operario) values (" & cod_importacion & ", " & codigos_de_barra(f) & ", " & cantidades(f) & ", " & TextBox11.Text & ", '" & UCase(TextBox6.Text) & "')")
                        End If
                    End If
                    'detalle_proveedores_articulos
                    If cantidades(f) > 0 Then
                        clase.consultar2("select* FROM detalle_proveedores_articulos WHERE codigo_articulo = " & codigos_de_barra(f) & "", "detalle_proveedores")
                        If clase.dt2.Tables("detalle_proveedores").Rows.Count = 0 Then
                            clase.agregar_registro("INSERT INTO detalle_proveedores_articulos (codigo_articulo, codigo_proveedor) values ('" & codigos_de_barra(f) & "', '" & TextBox9.Text & "')")
                        End If
                        CrearMovimientoCompraLovePOSImportacion(cod_importacion, TextBox9.Text, codigos_de_barra(f), cantidades(f), UCase(TextBox6.Text))
                    End If
                    'asociaciones_codigos   !!!!!!!!  (MIRAR si hay que actualizar el precodigo para que un postcoidgo no este amarrado a un precodigo viejo (esto se hace x lo de los registros de importacion))
                    If cantidades(f) > 0 Then             'CREO QUE YA ESTA HECHO
                        clase.consultar2("select* from asociaciones_codigos where asc_postcodart = " & codigos_de_barra(f) & "", "asociaciones")
                        If clase.dt2.Tables("asociaciones").Rows.Count > 0 Then
                            clase.actualizar("UPDATE asociaciones_codigos SET asc_precodref = " & obtener_codigo_apartir_de_consecutivo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value) & " WHERE asc_postcodart = " & codigos_de_barra(f) & "")
                        Else
                            clase.agregar_registro("INSERT INTO `asociaciones_codigos`(`asc_precodref`,`asc_postcodart`) VALUES ( '" & obtener_codigo_apartir_de_consecutivo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value) & "','" & codigos_de_barra(f) & "')")
                        End If
                    End If
                    If cantidades(f) > 0 Then
                        clase.agregar_registro("INSERT INTO registros_de_importaciones (articulo, registro, fecha) VALUES ('" & codigos_de_barra(f) & "', " & obtener_registro_importacion_apartir_de_consecutivo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value) & ", " & obtener_fecha_importacion_apartir_de_consecutivo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value) & ")")
                    End If
                Next
                marcar_item_procesado(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)
        End Select
        frm_crear_articulos_lotes.llenar_grilla(frm_crear_articulos_lotes.TextBox4.Text)
        cantidad_registros = frm_crear_articulos_lotes.dataGridView1.RowCount
        Button4_Click(Nothing, Nothing)
        '  ind_creacion_colores = False
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
        Button7.Enabled = False
        If registro_actual > cantidad_registros - 1 Then
            registro_actual = cantidad_registros - 1
            Button4.Enabled = False
            limpiar()
            ind_creacion_colores = False
            procesar_referencia(registro_actual)
        Else
            limpiar()
            ind_creacion_colores = False
            procesar_referencia(registro_actual)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        registro_actual -= 1
        Button4.Enabled = True
        Button7.Enabled = False
        If registro_actual < 0 Then
            registro_actual = 0
            Button3.Enabled = False
            limpiar()
            ind_creacion_colores = False
            procesar_referencia(registro_actual)
        Else
            limpiar()
            ind_creacion_colores = False
            procesar_referencia(registro_actual)
        End If
    End Sub

    Private Sub button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Structure tipoarticulo
        Dim referencia As String
        Dim linea1 As Short
        Dim sublinea1 As Short
        Dim sublinea2 As Object
        Dim sublinea3 As Object
        Dim sublinea4 As Object
    End Structure

    Private Sub redimencionar_array_de_codigosdebarras()
        If ComboBox6.Visible = True Then
            clase.consultar("select* from articulos where ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1(parameter, registro_actual).Value & "' and ar_linea = " & ComboBox6.Items(ComboBox6.SelectedIndex).linea & " and ar_sublinea1 = " & ComboBox6.Items(ComboBox6.SelectedIndex).sublinea1 & generar_cadena_sql_apartir_del_combobox_de_coincidencias(ComboBox6.Items(ComboBox6.SelectedIndex).sublinea2, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea3, ComboBox6.Items(ComboBox6.SelectedIndex).sublinea4), "resultadosbusqueda")
        Else
            clase.consultar("select* from articulos where ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1(parameter, registro_actual).Value & "' and ar_linea = " & articulo.linea1 & " and ar_sublinea1 = " & articulo.sublinea1 & generar_cadena_sql_apartir_del_combobox_de_coincidencias(verificar_nulidad_vacio(articulo.sublinea2), verificar_nulidad_vacio(articulo.sublinea3), verificar_nulidad_vacio(articulo.sublinea4)), "resultadosbusqueda")
        End If
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
            TextBox12.Text = extraer_costo(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)
            TextBox3.Text = ""
            TextBox4.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_precio1")
            TextBox5.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_precio2")
            dateTimePicker1.Value = clase.dt_global.Tables("resultados").Rows(0)("ar_fechaingreso")
            TextBox6.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_creadopor")
            TextBox8.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_tasadesc")
            TextBox7.Text = cod_importacion
            TextBox9.Text = codigo_proveedor_apartir_del_mov_salida(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)
            precio1 = clase.dt_global.Tables("resultados").Rows(0)("ar_precio1")
            precio2 = clase.dt_global.Tables("resultados").Rows(0)("ar_precio2")
            If System.IO.File.Exists(clase.dt_global.Tables("resultados").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt_global.Tables("resultados").Rows(0)("ar_foto"))
                TextBox10.Text = clase.dt_global.Tables("resultados").Rows(0)("ar_foto")
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
                TextBox10.Text = ""
            End If
            TextBox11.Text = "0"
            SetImage(PictureBox1)
            precioant1 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanterior"))
            precioant2 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanterior2"))
            precioanttixy1 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy1"))
            precioanttixy2 = verificar_nulidad_nulo(clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy2"))
            '  precioanttixy1 = clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy1")
            '  precioanttixy2 = clase.dt_global.Tables("resultados").Rows(0)("ar_precioanttixy2")
            construir_matrices_colores_tallas_codigosdebarra(clase.dt_global.Tables("resultados"))
            ReDim impuestos(0)
            impuestos(0) = hallar_impuesto_predeterminado()
            If clase.dt_global.Tables("resultados").Rows.Count = 1 Then
                ' MsgBox(obtener_cantidad(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value))
                cantidades(0) = obtener_cantidad(frm_crear_articulos_lotes.dataGridView1.Item(5, registro_actual).Value)
            End If
            establecer_estado_para_controles_para_impedir_edicion(False)
            If frm_crear_articulos_lotes.dataGridView1.Item(4, registro_actual).Value = True Then
                Button2.Enabled = False
                Button6.Enabled = False
                Button9.Enabled = False
            Else
                Button2.Enabled = True
                Button6.Enabled = True
                Button9.Enabled = True
            End If
        End If

    End Sub

    Function hallar_impuesto_predeterminado() As Short
        clase.consultar2("select impuesto_predeterminado from informacion", "impuesto")
        If IsDBNull(clase.dt2.Tables("impuesto").Rows(0)("impuesto_predeterminado")) Then
            Return 0
        Else
            Return clase.dt2.Tables("impuesto").Rows(0)("impuesto_predeterminado")
        End If
    End Function

    Private Sub procesar_referencia(ByVal fila As Integer)                                                                                          ' and ar_descripcion = '" & frm_crear_articulos_lotes.dataGridView1.Item(1, fila).Value & "'
        clase.consultar("select* from articulos where ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value & "'", "result")
        If clase.dt.Tables("result").Rows.Count > 0 Then
            GroupBox1.Enabled = False
            'aqui realizo la busqueda completa                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            '  and ar_descripcion = '" & frm_crear_articulos_lotes.dataGridView1.Item(1, fila).Value & "'
            clase.consultar("SELECT DISTINCT articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4, CONCAT(linea1.ln1_nombre, ', ', sublinea1.sl1_nombre) AS refart FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (sublinea2.sl2_sl1codigo = sublinea1.sl1_codigo) AND (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) AND (sublinea1.sl1_ln1codigo = linea1.ln1_codigo) LEFT JOIN sublinea3 ON (sublinea3.sl3_sl2codigo = sublinea2.sl2_codigo) AND (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (sublinea4.sl4_sl3codigo = sublinea3.sl3_codigo) AND (articulos.ar_sublinea4 = sublinea4.sl4_codigo) WHERE (articulos.ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value & "')", "tableresult11")
            If clase.dt.Tables("tableresult11").Rows.Count > 0 Then
                Dim cantidad_registros As Integer = clase.dt.Tables("tableresult11").Rows.Count
                Select Case cantidad_registros
                    Case Is > 1
                        Label17.Visible = True
                        Dim f As Short
                        '--------------------------------aqui el combobox esta visible cuando hay varios tipos de articulos con la misma referencia
                        ComboBox6.Visible = True
                        ComboBox6.Items.Clear()
                        For f = 0 To clase.dt.Tables("tableresult11").Rows.Count - 1
                            ComboBox6.Items.Add(New items(clase.dt.Tables("tableresult11").Rows(f)))
                        Next
                        '__________
                        MessageBox.Show("Se encontraron varios articulos diferentes con la misma referencia, por favor especifique a cual hace referencia antes de continuar.", "REFERENCIA YA EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If frm_crear_articulos_lotes.dataGridView1.Item(4, registro_actual).Value = True Then
                            Button7.Enabled = False
                        Else
                            Button7.Enabled = True
                        End If
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
                        articulo.referencia = frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value
                        articulo.linea1 = clase.dt.Tables("tableresult11").Rows(0)("ar_linea")
                        articulo.sublinea1 = clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea1")
                        articulo.sublinea2 = clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea2")
                        articulo.sublinea3 = clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea3")
                        articulo.sublinea4 = clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea4")
                        Dim con_sql As String = "select* from articulos where ar_linea = " & clase.dt.Tables("tableresult11").Rows(0)("ar_linea") & " and ar_sublinea1 = " & clase.dt.Tables("tableresult11").Rows(0)("ar_sublinea1") & complementosql2 & complementosql3 & complementosql4 & " and ar_referencia = '" & frm_crear_articulos_lotes.dataGridView1.Item(parameter, fila).Value & "'"
                        If frm_crear_articulos_lotes.dataGridView1.Item(4, registro_actual).Value = True Then
                            Button7.Enabled = False
                        Else
                            Button7.Enabled = True
                        End If
                        mostrar_referencia_ya_existente(con_sql)
                End Select
                ind_edicion_colores = False
            End If
        Else
            ind_edicion_colores = True
            mostrar_datos(registro_actual)
        End If
    End Sub

    Function hallar_codigo_importacion(codmov As Integer) As Integer
        clase.consultar("SELECT detalle_importacion_cabcajas.det_codigoimportacion FROM detsalidas_mercancia INNER JOIN cabsalidas_mercancia ON (detsalidas_mercancia.det_salidacodigo = cabsalidas_mercancia.cabsal_cod) INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (cabsalidas_mercancia.cabsal_cod =" & codmov & ")", "tablita")
        If clase.dt.Tables("tablita").Rows.Count > 0 Then
            Return clase.dt.Tables("tablita").Rows(0)("det_codigoimportacion")
        End If
    End Function

    Private Sub marcar_item_procesado(item As Integer)
        clase.actualizar("UPDATE `detsalidas_mercancia` SET `det_procesado`=True WHERE `det_codigo`=" & item & "")
    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As EventArgs) Handles TextBox10.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        ' System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim openFileDialog1 As System.Windows.Forms.OpenFileDialog
        openFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Dim excelPathName As String
        With openFileDialog1
            .Title = "Cargar Foto"
            .FileName = ""
            .DefaultExt = ".jpg"
            .AddExtension = True
            .Filter = "Formato de Intercambio de Archivos JPEG|*.jpg; *.jpeg;*.jpe;*.jfif"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                excelPathName = (CType(.FileName, String))
                If (excelPathName.Length) <> 0 Then
                    TextBox10.Text = excelPathName
                    PictureBox1.Image = Image.FromFile(excelPathName)
                    SetImage(PictureBox1)
                Else
                    MessageBox.Show("No se ha seleccionado ningun archivo compatible. Pulse aceptar para volverlo a intentar.", "SELECCIONAR ARCHIVO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End With
    End Sub

    Private Sub establecer_estado_para_controles_para_impedir_edicion(estado As Boolean)
        textBox2.Enabled = estado
        TextBox1.Enabled = estado
        ComboBox1.Enabled = estado
        ComboBox2.Enabled = estado
        ComboBox3.Enabled = estado
        ComboBox4.Enabled = estado
        ComboBox5.Enabled = estado
        'Button8.Enabled = estado
    End Sub


    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        frm_cruze_facturas.ShowDialog()
        frm_cruze_facturas.Dispose()
    End Sub

    Private Sub TextBox11_GotFocus(sender As Object, e As EventArgs) Handles TextBox11.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox11_LostFocus(sender As Object, e As EventArgs) Handles TextBox11.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub asociar_codigos(precodigo As Short, postcodigo As Short, registrodian As String, fecharegistro As Date) ' esta funcion asocia el codigo que tenia el articulo antes de liquidarse con el tiene despues de ser liquidado
        clase.agregar_registro("INSERT INTO `asociaciones_codigos`(`asc_precodref`,`asc_postcodart`,`asc_registrodian`,`asc_fecharegistro`) VALUES ( '" & precodigo & "','" & postcodigo & "','" & registrodian & "','" & fecharegistro & "')")
    End Sub

    Private Function obtener_codigo_apartir_de_consecutivo(consecutivo As Short) As Short ' esta funcion me devuelve el codigo de la referencia (el precodigo antes de ser liquidado) apartir del consecutivo de la tabla de detalle de salida de mercancia (detsalida_mercancia)
        clase.consultar("SELECT det_codref FROM detsalidas_mercancia WHERE (det_codigo =" & consecutivo & ")", "codref")
        Return clase.dt.Tables("codref").Rows(0)("det_codref")
    End Function

    Private Function obtener_registro_importacion_apartir_de_consecutivo(consecutivo As Short) As String
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_registrodian FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detsalidas_mercancia.det_codigo =" & consecutivo & ")", "regimp")
        If IsDBNull(clase.dt.Tables("regimp").Rows(0)("detcab_registrodian")) Then
            Return "NULL"
        Else
            Return "'" & clase.dt.Tables("regimp").Rows(0)("detcab_registrodian") & "'"
        End If
    End Function

    Private Function obtener_fecha_importacion_apartir_de_consecutivo(consecutivo As Short) As String
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_fecharegistrodian FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detsalidas_mercancia.det_codigo =" & consecutivo & ")", "fechimp")
        If IsDBNull(clase.dt.Tables("fechimp").Rows(0)("detcab_fecharegistrodian")) Then
            Return "NULL"
        Else
            Dim fecha As Date = clase.dt.Tables("fechimp").Rows(0)("detcab_fecharegistrodian")
            Return "'" & fecha.ToString("yyyy-MM-dd") & "'"
        End If
    End Function

    Private Function equivalencia_medida_en_unidades(medida As String) As Short
        clase.consultar("select unidades from unidadesmedida where nombremedida = '" & medida & "'", "medida")
        Return clase.dt.Tables("medida").Rows(0)("unidades")
    End Function

    Function obtener_cantidad(precodigo As Short) As Short
        clase.consultar("SELECT detsalidas_mercancia.det_cant, detalle_importacion_detcajas.detcab_unimedida FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detsalidas_mercancia.det_codigo =" & precodigo & ")", "cant")
        Return clase.dt.Tables("cant").Rows(0)("det_cant") * equivalencia_medida_en_unidades(clase.dt.Tables("cant").Rows(0)("detcab_unimedida"))
    End Function

    Private Sub establecer_colores_tallas(pcod As Short)
        ReDim colores(0)
        ReDim cantidades(0)
        ReDim arraytallas(0)
        colores(0) = 18
        arraytallas(0) = 6
        cantidades(0) = obtener_cantidad(pcod)
        ind_creacion_colores = True
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        limpiar()
        ind_edicion_colores = True
        mostrar_datos(registro_actual)
        Button7.Enabled = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        frm_impuestos.ShowDialog()
        frm_impuestos.Dispose()
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs)
       
    End Sub
End Class