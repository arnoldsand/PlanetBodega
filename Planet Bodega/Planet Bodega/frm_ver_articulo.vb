Public Class frm_ver_articulo
    Dim clase As New class_library
    Dim ind_carga As Boolean
    Dim encabezados(30) As String
    Dim listado(30) As String
    Public registro_actual As Integer
    Dim cantidad_registros As Integer
    Dim codigo_registro As Integer
    Dim unidades As Integer
    Dim precioplanet As Double
    'Dim precioantplanet2 As Double
    Dim preciotixy As Double
    'Dim precioanttixy2 As Double
    Dim rutaantigua As String
    Dim ind_actualizacion As Boolean 'esta variable me indica caundo se descargue el formulario si va a llenar otra vez el listado solo si hubo actualizacion
    Dim campos As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ver_articulo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If ind_actualizacion = True Then
            frm_articulos.llenar_datagrid(frm_articulos.textBox2.Text, frm_articulos.ComboBox1.Text)
        End If
    End Sub

    Private Sub frm_creacion_articulos_lotes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cantidad_registros = frm_articulos.dataGridView1.RowCount()
        registro_actual = frm_articulos.dataGridView1.CurrentCell.RowIndex
        ind_carga = False
        clase.llenar_combo(ComboBox1, "select* from linea1 order by ln1_nombre asc", "ln1_nombre", "ln1_codigo")
        clase.llenar_combo(cmbClasificacion, "select* from CabClasificacion order by IdClasificacion ASC", "Clasificacion", "IdClasificacion")
        mostrar_datos(registro_actual)
        ind_carga = True
        SetImage(PictureBox1)
        ind_actualizacion = False
    End Sub

    Private Sub mostrar_datos(ByVal fila As Integer)
        clase.dt_global.Tables.Clear()
        'Dim parametro As Long = frm_articulos.dataGridView1.Item(0, fila).Value
        '"select* from articulos where ar_referencia = '" & frm_articulos.dataGridView1.Item(0, fila).Value & "' and ar_linea = " & frm_articulos.dataGridView1.Item(18, fila).Value & " and ar_sublinea1 = " & frm_articulos.dataGridView1.Item(19, fila).Value & generar_cadena_sql_apartir_del_combobox_de_coincidencias(verificar_nulidad_vacio(frm_articulos.dataGridView1.Item(20, fila).Value), verificar_nulidad_vacio(frm_articulos.dataGridView1.Item(21, fila).Value), verificar_nulidad_vacio(frm_articulos.dataGridView1.Item(22, fila).Value))
        clase.consultar_global("select* from articulos where ar_codigo = " & frm_articulos.dataGridView1.Item(0, fila).Value & "", "tblresult")
        If clase.dt_global.Tables("tblresult").Rows.Count > 0 Then
            textBox2.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_referencia")
            TextBox1.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_descripcion")
            ComboBox1.SelectedValue = clase.dt_global.Tables("tblresult").Rows(0)("ar_linea")
            ComboBox2.Enabled = True
            clase.llenar_combo(ComboBox2, "SELECT* FROM sublinea1 WHERE sl1_ln1codigo = " & ComboBox1.SelectedValue & " ORDER BY sl1_nombre asc", "sl1_nombre", "sl1_codigo")
            ComboBox2.SelectedValue = clase.dt_global.Tables("tblresult").Rows(0)("ar_sublinea1")
            If IsDBNull(clase.dt_global.Tables("tblresult").Rows(0)("ar_sublinea2")) Then
                ComboBox3.SelectedIndex = -1
            Else
                ComboBox3.Enabled = True
                clase.llenar_combo(ComboBox3, "SELECT* FROM sublinea2 WHERE sl2_sl1codigo = " & ComboBox2.SelectedValue & "  ORDER BY sl2_nombre asc", "sl2_nombre", "sl2_codigo")
                ComboBox3.SelectedValue = clase.dt_global.Tables("tblresult").Rows(0)("ar_sublinea2")
            End If
            If IsDBNull(clase.dt_global.Tables("tblresult").Rows(0)("ar_sublinea3")) Then
                ComboBox4.SelectedIndex = -1
            Else
                ComboBox4.Enabled = True
                clase.llenar_combo(ComboBox4, "SELECT* FROM sublinea3 WHERE sl3_sl2codigo = " & ComboBox3.SelectedValue & "  ORDER BY sl3_nombre asc", "sl3_nombre", "sl3_codigo")
                ComboBox4.SelectedValue = clase.dt_global.Tables("tblresult").Rows(0)("ar_sublinea3")
            End If
            If IsDBNull(clase.dt_global.Tables("tblresult").Rows(0)("ar_sublinea4")) Then
                ComboBox5.SelectedIndex = -1
            Else
                ComboBox5.Enabled = True
                clase.llenar_combo(ComboBox5, "SELECT* FROM sublinea4 WHERE sl4_sl3codigo = " & ComboBox4.SelectedValue & "  ORDER BY sl4_nombre asc", "sl4_nombre", "sl4_codigo")
                ComboBox5.SelectedValue = clase.dt_global.Tables("tblresult").Rows(0)("ar_sublinea4")
            End If
            If IsDBNull(clase.dt_global.Tables("tblresult").Rows(0)("ar_IdClasificacion")) Then
                cmbClasificacion.SelectedIndex = -1
            Else
                cmbClasificacion.Enabled = True
                clase.llenar_combo(cmbClasificacion, "select* from CabClasificacion order by IdClasificacion ASC", "Clasificacion", "IdClasificacion")
                cmbClasificacion.SelectedValue = clase.dt_global.Tables("tblresult").Rows(0)("ar_IdClasificacion")
            End If
            TextBox3.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_costo2")
            TextBox7.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_costo")
            TextBox4.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_precio1")
            precioplanet = clase.dt_global.Tables("tblresult").Rows(0)("ar_precio1")
            TextBox5.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_precio2")
            preciotixy = clase.dt_global.Tables("tblresult").Rows(0)("ar_precio2")
            TextBox9.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_fechaingreso")
            TextBox6.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_creadopor")
            TextBox8.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_tasadesc")
            TextBox11.Text = verificar_nulidad_vacio(clase.dt_global.Tables("tblresult").Rows(0)("ar_ultimamodificacion"))
            TextBox12.Text = comprobar_nulidad_de_integer(clase.dt_global.Tables("tblresult").Rows(0)("ar_precioanterior"))
            TextBox13.Text = comprobar_nulidad_de_integer(clase.dt_global.Tables("tblresult").Rows(0)("ar_precioanterior2"))
            TextBox15.Text = comprobar_nulidad_de_integer(clase.dt_global.Tables("tblresult").Rows(0)("ar_precioanttixy1"))
            TextBox14.Text = comprobar_nulidad_de_integer(clase.dt_global.Tables("tblresult").Rows(0)("ar_precioanttixy2"))
            ind_edicion_colores = False
            If System.IO.File.Exists(clase.dt_global.Tables("tblresult").Rows(0)("ar_foto")) Then
                Try
                    Dim myimage As Image = Image.FromFile(clase.dt_global.Tables("tblresult").Rows(0)("ar_foto"))
                    Dim mybitmap As Bitmap = New Bitmap(myimage)
                    myimage.Dispose()
                    myimage = Nothing
                    PictureBox1.Image = mybitmap
                Catch
                    PictureBox1.Image = Global.WindowsApplication1.My.Resources.sinfoto  'Image.FromFile("\\servidor\Fotos Articulos\sinfoto.jpg")
                    rutaantigua = ""
                    TextBox10.Text = ""
                End Try
                'PictureBox1.Image = Image.FromFile(clase.dt_global.Tables("tblresult").Rows(0)("ar_foto"))
                ' MsgBox(System.IO.File.Exists(clase.dt_global.Tables("tblresult").Rows(0)("ar_foto")))
                TextBox10.Text = clase.dt_global.Tables("tblresult").Rows(0)("ar_foto")
                rutaantigua = clase.dt_global.Tables("tblresult").Rows(0)("ar_foto")
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.sinfoto  'Image.FromFile("\\servidor\Fotos Articulos\sinfoto.jpg")
                rutaantigua = ""
                TextBox10.Text = ""
            End If
            SetImage(PictureBox1)
            clase.consultar1("select* from impuestos_articulos where impart_articulo = " & frm_articulos.dataGridView1.Item(0, registro_actual).Value & "", "imp")
            If clase.dt1.Tables("imp").Rows.Count > 0 Then
                Dim z As Short
                ReDim impuestos2(clase.dt1.Tables("imp").Rows.Count - 1)
                For z = 0 To clase.dt1.Tables("imp").Rows.Count - 1
                    impuestos2(z) = clase.dt1.Tables("imp").Rows(z)("impart_impuestos")
                Next
            End If
        End If
        ' codigo_registro = clase.dt_global.Tables("table")(0)("codigo_detalle")
        If fila + 1 = cantidad_registros Then Button4.Enabled = False
        If fila = 0 Then Button3.Enabled = False
        construir_matrices_colores_tallas_codigosdebarra(clase.dt_global.Tables("tblresult"))
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


    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.BackColor = Color.White
    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub


    Private Sub TextBox7_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox8_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        limpiar()
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
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        textBox2.Focus()
    End Sub

    Sub validar_solo_numeros()

    End Sub

    Private Sub dateTimePicker1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        clase.enter(e)
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Referencia") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Descripción") = False Then Exit Sub
        If clase.validar_combobox(ComboBox1, "Linea") = False Then Exit Sub
        If clase.validar_combobox(ComboBox2, "Sub-Linea1") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox3, "Precio Costo 1") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox7, "Precio Costo 2") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox4, "Precio Planet") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox5, "Precio Tixy") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox6, "Creado Por") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox8, "Tasa Dcto") = False Then Exit Sub
        'If clase.validar_cajas_text(TextBox12, "Color") = False Then Exit Sub
        '  If clase.validar_cajas_text(TextBox13, "Talla") = False Then Exit Sub
        If Val(TextBox8.Text) > 100 Then
            MessageBox.Show("La tasa de descuento no puedo ser superior al 100%.", "DESCUENTO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox8.Focus()
            Exit Sub
        End If
        Dim consecutivo As Long = clase.generar_codigo_tabla_articulos
        Dim fecha As Date = TextBox9.Text
        Dim fechaultimamodificacion As String = Now.ToString("yyyy-MM-dd")
        Dim cmb3, cmb4, cmb5, cmb6 As String
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
        If cmbClasificacion.Text = "" Then
            cmb6 = "NULL"
        Else
            cmb6 = cmbClasificacion.SelectedValue
        End If
        'voy por aqui
        Dim extension As String = System.IO.Path.GetExtension(TextBox10.Text)
        Dim preciosactuales() As Double = {clase.dt_global.Tables("tblresult").Rows(0)("ar_precio1"), clase.dt_global.Tables("tblresult").Rows(0)("ar_precio2")}
        Dim preciosafijar() As Double = {TextBox4.Text, TextBox5.Text}
        Dim p As Short
        Dim nombrearchivo As String = ""
        Dim fotoActualizada As Boolean
        For p = 0 To codigos_de_barra.Length - 1
            fotoActualizada = False
            If TextBox10.Text <> "" Then
                If rutaantigua <> TextBox10.Text Then
                    fotoActualizada = True
                    nombrearchivo = extraer_nombre_archivo(clase.dt_global.Tables("tblresult").Rows(0)("ar_foto"))
                    If nombrearchivo = "" Then
                        nombrearchivo = codigos_de_barra(p)
                    End If
                    PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
                    nombrearchivo = nombrearchivo & "-" & "1"
                    crear_foto_grande(TextBox10.Text, ruta_foto() & "\" & nombrearchivo & extension)
                    crear_foto_pequeño(ruta_foto() & "\" & nombrearchivo & extension, ruta_foto() & "\" & nombrearchivo & "mini.jpg")
                    '    crear_foto_grande(TextBox10.Text, ruta_google_drive() & "\" & nombrearchivo & extension)
                End If
            End If
            If fotoActualizada = True Then
                clase.actualizar("UPDATE `articulos` SET `ar_referencia`='" & UCase(textBox2.Text) & "',`ar_descripcion`='" & UCase(TextBox1.Text) & "',`ar_linea`='" & ComboBox1.SelectedValue & "',`ar_sublinea1`='" & ComboBox2.SelectedValue & "',`ar_sublinea2`=" & cmb3 & ",`ar_sublinea3`=" & cmb4 & ",`ar_sublinea4`=" & cmb5 & ",`ar_costo`='" & Str(TextBox7.Text) & "', ar_costo2 = " & Str(TextBox3.Text) & ", `ar_precio1`='" & TextBox4.Text & "',`ar_precio2`='" & TextBox5.Text & "',`ar_creadopor`='" & TextBox6.Text & "',`ar_tasadesc`='" & TextBox8.Text & "',`ar_color`='" & colores(p) & "',`ar_talla`='" & arraytallas(p) & "'" & precios_anteriores(preciosactuales, clase.dt_global.Tables("tblresult"), preciosafijar) & ", ar_ultimamodificacion = '" & fechaultimamodificacion & "', ar_foto = '" & Replace(ruta_foto() & "\" & nombrearchivo & extension, "\", "\\") & "', `ar_migrado`=FALSE, ar_IdClasificacion = " & cmb6 & "  WHERE `ar_codigo`='" & codigos_de_barra(p) & "'")
            Else
                clase.actualizar("UPDATE `articulos` SET `ar_referencia`='" & UCase(textBox2.Text) & "',`ar_descripcion`='" & UCase(TextBox1.Text) & "',`ar_linea`='" & ComboBox1.SelectedValue & "',`ar_sublinea1`='" & ComboBox2.SelectedValue & "',`ar_sublinea2`=" & cmb3 & ",`ar_sublinea3`=" & cmb4 & ",`ar_sublinea4`=" & cmb5 & ",`ar_costo`='" & Str(TextBox7.Text) & "', ar_costo2 = " & Str(TextBox3.Text) & ", `ar_precio1`='" & TextBox4.Text & "',`ar_precio2`='" & TextBox5.Text & "',`ar_creadopor`='" & TextBox6.Text & "',`ar_tasadesc`='" & TextBox8.Text & "',`ar_color`='" & colores(p) & "',`ar_talla`='" & arraytallas(p) & "'" & precios_anteriores(preciosactuales, clase.dt_global.Tables("tblresult"), preciosafijar) & ", ar_ultimamodificacion = '" & fechaultimamodificacion & "', `ar_migrado`=FALSE, ar_IdClasificacion = " & cmb6 & "  WHERE `ar_codigo`='" & codigos_de_barra(p) & "'")
            End If
            If impuestos2.Length > 0 Then
                clase.borradoautomatico("delete from impuestos_articulos where impart_articulo = " & codigos_de_barra(p) & "")
                Dim x As Short
                For x = 0 To impuestos2.Length - 1
                    clase.consultar("select* from impuestos_articulos where impart_articulo = " & codigos_de_barra(p) & " and impart_impuestos = " & impuestos2(x) & "", "buscarimp")
                    If clase.dt.Tables("buscarimp").Rows.Count = 0 Then
                        clase.agregar_registro("INSERT INTO `impuestos_articulos`(`impart_codigo`,`impart_articulo`,`impart_impuestos`) VALUES ( NULL,'" & codigos_de_barra(p) & "','" & impuestos2(x) & "')")
                    End If
                Next
            Else
                clase.borradoautomatico("delete from impuestos_articulos where impart_articulo = " & codigos_de_barra(p) & "")
            End If
        Next
        '  mostrar_datos(registro_actual)
        MessageBox.Show("La actualización se realizó satisfactoriamente.", "OPERACIÓN EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ' frm_articulos.llenar_datagrid(frm_articulos.textBox2.Text, frm_articulos.ComboBox1.Text)
        cantidad_registros = frm_articulos.dataGridView1.RowCount
        Button4_Click(Nothing, Nothing)
        ind_actualizacion = True
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

    Function extraer_nombre_archivo(cadena As String) As String

        Dim strArray() As String
        strArray = Split(cadena, "\")
        Dim strArray1() As String
        strArray1 = Split(strArray(4), ".")
        Return strArray1(0)
    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        registro_actual += 1
        Button3.Enabled = True
        If registro_actual > cantidad_registros - 1 Then
            registro_actual = cantidad_registros - 1
            Button4.Enabled = False
            limpiar()
            mostrar_datos(registro_actual)
            '  ind_creacion_colores = False
        Else
            limpiar()
            mostrar_datos(registro_actual)
            ' ind_creacion_colores = False
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        registro_actual -= 1
        Button4.Enabled = True
        If registro_actual < 0 Then
            registro_actual = 0
            Button3.Enabled = False
            limpiar()
            mostrar_datos(registro_actual)
            ' ind_creacion_colores = False
        Else
            limpiar()
            mostrar_datos(registro_actual)
            'ind_creacion_colores = False
        End If
    End Sub

   

   
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frm_lista_colores.ShowDialog()
        frm_lista_colores.Dispose()
    End Sub

    


    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frm_lista_tallas.ShowDialog()
        frm_lista_tallas.Dispose()
    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        frm_lista_colores.ShowDialog()
        frm_lista_colores.Dispose()
    End Sub


    Private Sub Button7_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        frm_codigos_asociados.ShowDialog()
        frm_codigos_asociados.Dispose()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        frm_codigosdebarraasociados.ShowDialog()
        frm_codigosdebarraasociados.Dispose()
    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As EventArgs) Handles TextBox10.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
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

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        frm_impuestos2.ShowDialog()
        frm_impuestos2.Dispose()
    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox11_GotFocus(sender As Object, e As EventArgs) Handles TextBox11.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As EventArgs) Handles TextBox12.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As EventArgs) Handles TextBox13.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox14_GotFocus(sender As Object, e As EventArgs) Handles TextBox14.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As EventArgs) Handles TextBox15.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        frm_listado_de_registros_importaciones.ShowDialog()
        frm_listado_de_registros_importaciones.Dispose()
    End Sub
End Class