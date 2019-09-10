Public Class frm_ver_tienda
    Dim clase As New class_library
    Dim registro_actutal As Integer
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ver_tienda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox1, "select* from empresas order by nombre_empresa asc", "nombre_empresa", "cod_empresa")
        registro_actutal = frm_tienda.datagridreferencia.Item(9, frm_tienda.datagridreferencia.CurrentRow.Index).Value
        clase.consultar("select* from tiendas where id = " & registro_actutal & "", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            textBox2.Text = clase.dt.Tables("tbl").Rows(0)("tienda")
            TextBox1.Text = clase.dt.Tables("tbl").Rows(0)("ciudad")
            ComboBox1.SelectedValue = clase.dt.Tables("tbl").Rows(0)("empresa")
            TextBox3.Text = clase.dt.Tables("tbl").Rows(0)("email")
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("telefono1")) Then
                TextBox4.Text = ""
            Else
                TextBox4.Text = clase.dt.Tables("tbl").Rows(0)("telefono1")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("telefono2")) Then
                TextBox5.Text = ""
            Else
                TextBox5.Text = clase.dt.Tables("tbl").Rows(0)("telefono2")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("direccion")) Then
                TextBox6.Text = ""
            Else
                TextBox6.Text = clase.dt.Tables("tbl").Rows(0)("direccion")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("administrador")) Then
                TextBox7.Text = ""
            Else
                TextBox7.Text = clase.dt.Tables("tbl").Rows(0)("administrador")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("local")) Then
                TextBox8.Text = ""
            Else
                TextBox8.Text = clase.dt.Tables("tbl").Rows(0)("local")
            End If
            If clase.dt.Tables("tbl").Rows(0)("estado") = True Then
                ComboBox2.SelectedIndex = 0
            End If
            If clase.dt.Tables("tbl").Rows(0)("estado") = False Then
                ComboBox2.SelectedIndex = 1
            End If
            clase.consultar1("select* from usuarios_tienda where tienda = " & registro_actutal & "", "usuario-catalogo")
            TextBox10.Text = clase.dt1.Tables("usuario-catalogo").Rows(0)("usuario")

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Nombre Tienda") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Ciudad") = False Then Exit Sub
        If clase.validar_combobox(ComboBox1, "Empresa") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox3, "E-mail") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox10, "Usuario Catálogo") = False Then Exit Sub
        Dim telefono1, telefono2, direccion, administrador, local As String
        Dim estado As Boolean
        If TextBox4.Text.Trim = "" Then
            telefono1 = "NULL"
        Else
            telefono1 = "'" & UCase(TextBox4.Text) & "'"
        End If
        If TextBox5.Text.Trim = "" Then
            telefono2 = "NULL"
        Else
            telefono2 = "'" & UCase(TextBox5.Text) & "'"
        End If
        If TextBox6.Text.Trim = "" Then
            direccion = "NULL"
        Else
            direccion = "'" & UCase(TextBox6.Text) & "'"
        End If
        If TextBox7.Text.Trim = "" Then
            administrador = "NULL"
        Else
            administrador = "'" & UCase(TextBox7.Text) & "'"
        End If
        If TextBox8.Text.Trim = "" Then
            local = "NULL"
        Else
            local = "'" & UCase(TextBox8.Text) & "'"
        End If
        If ComboBox2.Text = "Activa" Then
            estado = True
        End If
        If ComboBox2.Text = "Inactiva" Then
            estado = False
        End If
        clase.actualizar("UPDATE `tiendas` SET `tienda`='" & textBox2.Text & "',`ciudad`='" & TextBox1.Text & "',`email`='" & TextBox3.Text & "',`direccion`=" & direccion & ",`local`=" & local & ",`empresa`='" & ComboBox1.SelectedValue & "',`telefono1`=" & telefono1 & ",`telefono2`=" & telefono2 & ",`administrador`=" & administrador & ",`estado`=" & estado & " WHERE `id`=" & registro_actutal & "")
        If TextBox11.Text <> "" Then
            clase.actualizar("UPDATE usuarios_tienda SET password = '" & frm_nueva_tienda.MD5EncryptPass(TextBox11.Text) & "' WHERE tienda = " & registro_actutal & "")
        End If
        '  restablecer()
        frm_tienda.llenar_tienda(frm_tienda.devolver_valor(frm_tienda.ComboBox2.Text))
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        textBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        ComboBox1.SelectedIndex = -1
        textBox2.Focus()
        ComboBox2.SelectedIndex = -1
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        textBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox2_LostFocus(sender As Object, e As EventArgs) Handles textBox2.LostFocus
        textBox2.BackColor = Color.White
    End Sub

    Private Sub textBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs)
        clase.enter(e)
    End Sub

    Private Sub textBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub textBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        TextBox3.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub textBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        TextBox4.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub textBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        TextBox5.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub textBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        TextBox6.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox6_LostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        TextBox6.BackColor = Color.White
    End Sub

    Private Sub textBox7_GotFocus(sender As Object, e As EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox7_LostFocus(sender As Object, e As EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub textBox8_GotFocus(sender As Object, e As EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox8_LostFocus(sender As Object, e As EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub
End Class