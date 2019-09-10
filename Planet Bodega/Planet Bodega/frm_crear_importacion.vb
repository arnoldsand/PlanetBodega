Public Class frm_crear_importacion
    Dim clase As class_library = New class_library
    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Lugar") = False Then Exit Sub
        If clase.validar_cajas_text(textBox3, "Descripción") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Comentarios") = False Then Exit Sub
        Dim fecha As Date = dateTimePicker1.Value
        Dim fech As String = fecha.ToString("yyyy-MM-dd")
        Dim nombrefecha As String = fecha.ToString("MMMM-yyyy") & "(" & textBox2.Text & ")"
        clase.agregar_registro("INSERT INTO `cabimportacion`(`imp_fecha`,`imp_descripcion`,`imp_lugar`,`imp_nombrefecha`,`imp_estado`) VALUES ( '" & fech & "','" & UCase(TextBox1.Text) & "','" & UCase(textBox2.Text) & "', '" & UCase(textBox3.Text) & "', TRUE)")
        ''buscar codigo del registro que creé para compartirlo en una varibale publica
        'Dim c1 As New MySqlDataAdapter("Select* from cabimportacion where imp_fecha = '" & fech & "' and imp_descripcion = '" & UCase(TextBox1.Text) & "' and imp_lugar = '" & UCase(textBox2.Text) & "' and imp_nombrefecha like '" & UCase(textBox3.Text) & "%'", clase.conex)
        'Dim dt As New DataTable
        'clase.Conectar()
        'c1.Fill(dt)
        'clase.desconectar()
        '   frm_crear_articulos_lotes.Button9.Enabled = True
        '  frm_crear_articulos_lotes.Button10.Enabled = True
        ' frm_crear_articulos_lotes.Button12.Enabled = True
        frm_listado_importacion.llenar_combo()
        ' frm_crear_articulos_lotes.ComboBox1.Text = nombrefecha
        Me.Close()
    End Sub
End Class