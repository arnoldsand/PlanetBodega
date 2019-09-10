Public Class frm_crear_importacion
    Dim clase As class_library = New class_library
    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        Me.Close()
    End Sub

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        If clase.validar_cajas_text(textBox2, "Lugar") = False Then Exit Sub
        Dim fecha As Date = dateTimePicker1.Value
        Dim fech As String = fecha.ToString("yyyy-MM-dd")
        Dim nombrefecha As String = fecha.ToString("MMMM-yyyy") & "(" & textBox2.Text & ")"
        clase.agregar_registro("INSERT INTO `cabimportacion`(`imp_fecha`,`imp_descripcion`,`imp_lugar`,`imp_nombrefecha`) VALUES ( '" & fech & "','" & textBox3.Text & "','" & textBox2.Text & "', '" & nombrefecha & "')")
        'buscar codigo del registro que creé para compartirlo en una varibale publica
        Dim c1 As New MySqlDataAdapter("Select* from cabimportacion where imp_fecha = '" & fech & "' and imp_descripcion = '" & textBox3.Text & "' and imp_lugar = '" & textBox2.Text & "'", clase.conex)
        Dim dt As New DataTable
        clase.Conectar()
        c1.Fill(dt)
        clase.desconectar()
        cod_importacion = dt.Rows(0)("imp_codigo")
        frm_crear_articulos_lotes.Button9.Enabled = True
        frm_crear_articulos_lotes.Button10.Enabled = True
        frm_crear_articulos_lotes.Button12.Enabled = True
        frm_crear_articulos_lotes.llenar_combo()
        frm_crear_articulos_lotes.ComboBox1.Text = nombrefecha
        Me.Close()
    End Sub

   
End Class