Public Class frm_ver_proveedor
    Dim clase As New class_library
    Dim registro_actual As Integer
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Nombre Proveedor") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox6, "Codigo Proveedor") = False Then Exit Sub
        clase.consultar("SELECT prv_codigoasignado FROM proveedores WHERE (prv_codigoasignado =" & TextBox6.Text.Trim & "  AND prv_codigo <> " & registro_actual & ")", "prov")
        If clase.dt.Tables("prov").Rows.Count > 0 Then
            MessageBox.Show("El codigo especificado ya fue asignado a otra compañia.", "CODIGO YA EN USO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim ciudad, direccion, contacto, tel1, tel2, email1, email2, web As String
        ciudad = comprobar_nulidad(TextBox1.Text)
        direccion = comprobar_nulidad(TextBox3.Text)
        contacto = comprobar_nulidad(TextBox4.Text)
        tel1 = comprobar_nulidad(TextBox5.Text)
        tel2 = comprobar_nulidad(TextBox10.Text)
        email1 = comprobar_nulidad(TextBox9.Text)
        email2 = comprobar_nulidad(TextBox8.Text)
        web = comprobar_nulidad(TextBox7.Text)
        clase.actualizar("UPDATE `proveedores` SET `prv_codigoasignado`=" & TextBox6.Text & ", `prv_nombre`='" & UCase(textBox2.Text) & "',`prv_ciudad`=" & UCase(ciudad) & ",`prv_direccion`=" & UCase(direccion) & ",`prv_contacto`=" & UCase(contacto) & ",`prv_telefono1`=" & UCase(tel1) & ",`prv_telefono2`=" & UCase(tel2) & ",`prv_email1`=" & LCase(email1) & ",`prv_email2`=" & LCase(email2) & ",`prv_web`=" & LCase(web) & ", `prv_migrado`=FALSE WHERE `prv_codigo`=" & registro_actual & "")
        restablecer()
        frm_maestro_proveedores.llenar_proveedores(frm_maestro_proveedores.textBox2.Text)
        Me.Close()
    End Sub

    Private Sub restablecer()
        textBox2.Text = ""
        textBox2.Focus()
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox10.Text = ""
        TextBox9.Text = ""
        TextBox8.Text = ""
        TextBox7.Text = ""
        TextBox6.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        restablecer()
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        textBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles textBox2.LostFocus
        textBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        TextBox3.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        TextBox4.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        TextBox5.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As EventArgs) Handles TextBox10.GotFocus
        TextBox10.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox10_LostFocus(sender As Object, e As EventArgs) Handles TextBox10.LostFocus
        TextBox10.BackColor = Color.White
    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        TextBox9.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub


    Private Sub TextBox9_LostFocus(sender As Object, e As EventArgs) Handles TextBox9.LostFocus
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox8_LostFocus(sender As Object, e As EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.enter(e)
    End Sub

    
   

    Private Sub textBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        clase.enter(e)
    End Sub

    Private Sub frm_ver_proveedor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        registro_actual = frm_maestro_proveedores.datagridreferencia.Item(10, frm_maestro_proveedores.datagridreferencia.CurrentRow.Index).Value
        clase.consultar("select* from proveedores where prv_codigo = " & registro_actual & "", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            textBox2.Text = clase.dt.Tables("tbl").Rows(0)("prv_nombre")
            TextBox6.Text = clase.dt.Tables("tbl").Rows(0)("prv_codigoasignado")
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_ciudad")) Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = clase.dt.Tables("tbl").Rows(0)("prv_ciudad")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_direccion")) Then
                TextBox3.Text = ""
            Else
                TextBox3.Text = clase.dt.Tables("tbl").Rows(0)("prv_direccion")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_contacto")) Then
                TextBox4.Text = ""
            Else
                TextBox4.Text = clase.dt.Tables("tbl").Rows(0)("prv_contacto")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_telefono1")) Then
                TextBox5.Text = ""
            Else
                TextBox5.Text = clase.dt.Tables("tbl").Rows(0)("prv_telefono1")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_telefono2")) Then
                TextBox10.Text = ""
            Else
                TextBox10.Text = clase.dt.Tables("tbl").Rows(0)("prv_telefono2")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_email1")) Then
                TextBox9.Text = ""
            Else
                TextBox9.Text = clase.dt.Tables("tbl").Rows(0)("prv_email1")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_email2")) Then
                TextBox8.Text = ""
            Else
                TextBox8.Text = clase.dt.Tables("tbl").Rows(0)("prv_email2")
            End If
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("prv_web")) Then
                TextBox7.Text = ""
            Else
                TextBox7.Text = clase.dt.Tables("tbl").Rows(0)("prv_web")
            End If
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        TextBox6.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        TextBox6.BackColor = Color.White
    End Sub
End Class