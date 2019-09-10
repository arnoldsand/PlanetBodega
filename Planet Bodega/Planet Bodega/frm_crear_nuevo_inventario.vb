Public Class frm_crear_nuevo_inventario
    Dim clase As New class_library
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_combobox(ComboBox9, "Tipo Inventario") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox18, "Operario") = False Then Exit Sub
        Dim fecha As Date = Now
        clase.agregar_registro("INSERT INTO `cabinventario`(`cabinv_codigo`,`cabinv_fecha`,`cabinv_hora`,`cabinv_operario`,`cabinv_observaciones`,`cabinv_tipo_inventario`,`cabinv_procesado`) VALUES ( NULL,'" & fecha.ToString("yyyy-MM-dd") & "','" & fecha.ToString("HH:mm:ss") & "','" & UCase(TextBox18.Text) & "','" & UCase(TextBox1.Text) & "','" & ComboBox9.Text & "',FALSE)")
        clase.consultar1("select max(cabinv_codigo) as maximo from cabinventario", "consulta")
        frm_modulo_inventario.TextBox4.Text = ComboBox9.Text
        frm_modulo_inventario.TextBox1.Text = clase.dt1.Tables("consulta").Rows(0)("maximo")
        frm_modulo_inventario.TextBox18.Text = TextBox18.Text
        frm_modulo_inventario.TextBox5.Text = FormatCurrency(0)
        frm_modulo_inventario.TextBox2.Text = FormatCurrency(0)
        frm_modulo_inventario.TextBox3.Text = FormatCurrency(0)
        clase.consultar("select* from secciones_inventario where secc_inventario = " & frm_modulo_inventario.TextBox1.Text & "", "bus")
        If clase.dt.Tables("bus").Rows.Count = 0 Then
            clase.consultar1("SELECT bodegas.bod_codigo, bodegas.bod_abreviatura, bodegas.bod_nombre, COUNT(detbodega.dtbod_bodega) AS cantidad FROM detbodega INNER JOIN bodegas ON (detbodega.dtbod_bodega = bodegas.bod_codigo) GROUP BY detbodega.dtbod_bodega", "detbod")
            If clase.dt1.Tables("detbod").Rows.Count > 0 Then
                Dim a, b, c As Short

                Dim consulta As String = "INSERT INTO secciones_inventario (secc_codigo, secc_inventario, secc_bodega, secc_conteo, secc_gondola, secc_operario, secc_procesado) VALUES "
                Dim ind1 As Boolean = False
                For a = 0 To clase.dt1.Tables("detbod").Rows.Count - 1
                    For b = 1 To clase.dt1.Tables("detbod").Rows(a)("cantidad")
                        clase.consultar2("SELECT dtbod_cant_gondola FROM detbodega WHERE (dtbod_bodega =" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & " AND dtbod_bloque =" & b & ")", "cant-gond")
                        For c = 1 To clase.dt2.Tables("cant-gond").Rows(0)("dtbod_cant_gondola")
                            Application.DoEvents()
                            If ind1 = False Then     'falta agregar el campo conteo en las dos consultas de abajo
                                consulta = consulta & "(NULL, '" & frm_modulo_inventario.TextBox1.Text & "', '" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & "', '1', '" & bloque(b - 1) & c & "', '', FALSE)"
                                consulta = consulta & ", (NULL, '" & frm_modulo_inventario.TextBox1.Text & "', '" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & "', '2', '" & bloque(b - 1) & c & "', '', FALSE)"
                                consulta = consulta & ", (NULL, '" & frm_modulo_inventario.TextBox1.Text & "', '" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & "', '3', '" & bloque(b - 1) & c & "', '', FALSE)"
                                ind1 = True
                            Else
                                consulta = consulta & ", " & "(NULL, '" & frm_modulo_inventario.TextBox1.Text & "', '" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & "', '1','" & bloque(b - 1) & c & "', '', FALSE)"
                                consulta = consulta & ", " & "(NULL, '" & frm_modulo_inventario.TextBox1.Text & "', '" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & "', '2','" & bloque(b - 1) & c & "', '', FALSE)"
                                consulta = consulta & ", " & "(NULL, '" & frm_modulo_inventario.TextBox1.Text & "', '" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & "', '3','" & bloque(b - 1) & c & "', '', FALSE)"
                            End If

                            'clase.agregar_registro("INSERT INTO secciones_inventario (secc_codigo, secc_inventario, secc_bodega, secc_gondola, secc_operario, secc_procesado) VALUES (NULL, '" & frm_modulo_inventario.TextBox1.Text & "', '" & clase.dt1.Tables("detbod").Rows(a)("bod_codigo") & "', '" & bloque(b - 1) & c & "', '', FALSE)")
                        Next
                    Next
                Next
                clase.agregar_registro(consulta)
            End If
        End If
        MessageBox.Show("Registro creado exitosamente.", "OPERACIÓN SATISFACTORIA", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ComboBox9.SelectedIndex = -1
        TextBox18.Text = ""
        TextBox1.Text = ""
        ComboBox9.Focus()
    End Sub
End Class