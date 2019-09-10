Public Class frm_crear_articulo
    Dim clase As New class_library
    Dim ind_carga As Boolean
    Dim encabezados(30) As String
    Dim listado(30) As String
    Dim registro_actual As Integer
    Dim cantidad_registros As Integer
    Dim codigo_registro As Integer
    Dim unidades As Integer
    Dim campos As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_crear_articulos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox9.Text = Now.ToString("yyyy-MM-dd")
        ind_carga = False
        clase.llenar_combo(ComboBox1, "select* from linea1 order by ln1_nombre asc", "ln1_nombre", "ln1_codigo")
        clase.llenar_combo(cmbCabClasificacion, "select* from CabClasificacion order by IdClasificacion ASC", "Clasificacion", "IdClasificacion")
        ind_carga = True
        ind_creacion_colores = False
        TextBox8.Text = "0"
        If indicador_de_formulario = True Then
            TextBox6.Text = frm_ingreso_no_importado.txtOperario.Text
            If frm_busqueda_articulos.chkReferencia.Checked = True Then
                textBox2.Text = UCase(frm_busqueda_articulos.txtConsulta.Text)
            End If
            If frm_busqueda_articulos.chkDescripcion.Checked = True Then
                TextBox1.Text = UCase(frm_busqueda_articulos.txtConsulta.Text)
            End If
        End If
        If indicador_de_formulario2 = True Then
            TextBox6.Text = frm_ingreso_puerta_puerta_puerta_puerta.txtOperario.Text
            If frm_busqueda_articulos_puerta_puerta.chkReferencia.Checked = True Then
                textBox2.Text = UCase(frm_busqueda_articulos_puerta_puerta.txtConsulta.Text)
            End If
            If frm_busqueda_articulos_puerta_puerta.chkDescripcion.Checked = True Then
                TextBox1.Text = UCase(frm_busqueda_articulos_puerta_puerta.txtConsulta.Text)
            End If
        End If
        ReDim impuestos1(0)
        impuestos1(0) = frm_creacion_articulos_lotes.hallar_impuesto_predeterminado()
    End Sub

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

    Private Sub TextBox8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox8_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox7_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
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
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
        cmbCabClasificacion.SelectedIndex = -1
        textBox2.Focus()
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        PictureBox1.Image = Nothing
        TextBox10.Text = ""
        ReDim colores(Nothing)
        ReDim arraytallas(Nothing)
        CheckBox1.Checked = False
        ind_carga = True
        ind_creacion_colores = False
    End Sub

    Private Sub dateTimePicker1_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
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
        If ind_creacion_colores = False Then
            MessageBox.Show("No se han establecido patrones de colores para esta referencia. Por favor establezcalos y vuelva a intentarlo.", "PATRONES DE COLORES", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If Val(TextBox8.Text) > 100 Then
            MessageBox.Show("La tasa de descuento no puedo ser superior al 100%.", "DESCUENTO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox8.Focus()
            Exit Sub
        End If
        Dim consecutivo As Long
        Dim fecha As String = Now.ToString("yyyy-MM-dd")
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
        If cmbCabClasificacion.Text = "" Then
            cmb6 = "NULL"
        Else
            cmb6 = cmbCabClasificacion.SelectedValue
        End If
        Dim extension As String = System.IO.Path.GetExtension(TextBox10.Text)
        Dim a As Integer
        For a = 0 To colores.Length - 1
            consecutivo = clase.generar_codigo_tabla_articulos
            'creo la foto solo si el usuario ha espcificado alguna
            If TextBox10.Text <> "" Then
                crear_foto_grande(TextBox10.Text, ruta_foto() & "\" & consecutivo & extension)
                crear_foto_pequeño(ruta_foto() & "\" & consecutivo & extension, ruta_foto() & "\" & consecutivo & "mini.jpg")
                '     crear_foto_grande(TextBox10.Text, ruta_google_drive() & "\" & consecutivo & extension)
            End If
            clase.agregar_registro("INSERT INTO `articulos`(`ar_codigo`,`ar_codigobarras`,`ar_codigo2`,`ar_codigobarras2`,`ar_referencia`,`ar_descripcion`,`ar_color`,`ar_talla`,`ar_linea`,`ar_sublinea1`,`ar_sublinea2`,`ar_sublinea3`,`ar_sublinea4`,`ar_costo`,`ar_costo2`,`ar_precio1`,`ar_precio2`,`ar_precioanterior`,`ar_fechaingreso`,`ar_precioanterior2`,`ar_creadopor`,`ar_fecdescontinua`,`ar_activo`,`ar_tasadesc`,`ar_ultimamodificacion`,`ar_foto`,`ar_IdClasificacion`) VALUES ('" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & consecutivo & "','" & generar_codigo_ean13(consecutivo) & "','" & UCase(textBox2.Text) & "','" & UCase(TextBox1.Text) & "','" & colores(a) & "','" & arraytallas(a) & "','" & ComboBox1.SelectedValue & "','" & ComboBox2.SelectedValue & "'," & cmb3 & "," & cmb4 & "," & cmb5 & ",'" & Str(TextBox7.Text) & "','" & Str(TextBox3.Text) & "','" & TextBox4.Text & "','" & TextBox5.Text & "',NULL,'" & fecha & "',NULL,'" & TextBox6.Text & "',NULL, True,'" & TextBox8.Text & "','" & fechaultimamodificacion & "','" & Replace(ruta_foto() & "\" & consecutivo & extension, "\", "\\") & "'," & cmb6 & ")")
            If impuestos1.Length > 0 Then
                Dim x As Short
                For x = 0 To impuestos1.Length - 1
                    clase.consultar("select* from impuestos_articulos where impart_articulo = " & consecutivo & " and impart_impuestos = " & impuestos1(x) & "", "buscarimp")
                    If clase.dt.Tables("buscarimp").Rows.Count = 0 Then
                        clase.agregar_registro("INSERT INTO `impuestos_articulos`(`impart_codigo`,`impart_articulo`,`impart_impuestos`) VALUES ( NULL,'" & consecutivo & "','" & impuestos1(x) & "')")
                    End If
                Next
            Else
                clase.borradoautomatico("delete from impuestos_articulos where impart_articulo = " & consecutivo & "")
            End If
        Next
        ReDim impuestos1(0)
        impuestos1(0) = frm_creacion_articulos_lotes.hallar_impuesto_predeterminado()
        If indicador_de_formulario = False Then
            frm_articulos.llenar_datagrid(frm_articulos.textBox2.Text, "")
        End If
        limpiar()
        ind_creacion_colores = False
        If indicador_de_formulario = True Then
            If colores.Length = 1 Then
                frm_ingreso_no_importado.txtArticulo.Text = consecutivo
                frm_busqueda_articulos.Close()
                Me.Close()
            End If
        End If
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        frm_colores_tallas.ShowDialog()
        frm_colores_tallas.Dispose()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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

    Private Sub TextBox10_GotFocus(sender As Object, e As EventArgs) Handles TextBox10.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Select Case CheckBox1.Checked
            Case False
                ReDim colores(Nothing)
                ReDim arraytallas(Nothing)
                
                ind_creacion_colores = False
            Case True
                ReDim colores(0)
                ReDim arraytallas(0)
                colores(0) = 18
                arraytallas(0) = 6
                ind_creacion_colores = True
        End Select
    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        frm_impuestos1.ShowDialog()
        frm_impuestos1.Dispose()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub TextBox9_GotFocus1(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub cmbCabClasificacion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cmbCabClasificacion.KeyPress
        clase.enter(e)
    End Sub
End Class