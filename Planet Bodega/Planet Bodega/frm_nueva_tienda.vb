Imports System.Security.Cryptography
Public Class frm_nueva_tienda
    Dim clase As New class_library


    Function MD5EncryptPass(ByVal StrPass As String) As String
        Dim PasConMd5 As String
        PasConMd5 = ""
        Dim md5 As New MD5CryptoServiceProvider
        Dim bytValue() As Byte
        Dim bytHash() As Byte
        Dim i As Integer
        bytValue = System.Text.Encoding.UTF8.GetBytes(StrPass)
        bytHash = md5.ComputeHash(bytValue)
        md5.Clear()
        For i = 0 To bytHash.Length - 1
            PasConMd5 &= bytHash(i).ToString("x").PadLeft(2, "0")
        Next
        Return PasConMd5
    End Function

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_nueva_tienda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox1, "select* from empresas order by nombre_empresa asc", "nombre_empresa", "cod_empresa")
        clase.consultar1("select* from informacion", "inf")
        TextBox11.Text = clase.dt1.Tables("inf").Rows(0)("password_catalogo_tiendas")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Nombre Tienda") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Ciudad") = False Then Exit Sub
        If clase.validar_combobox(ComboBox1, "Empresa") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox3, "E-mail") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox10, "Usuario") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox11, "Contraseña") = False Then Exit Sub
        Dim telefono1, telefono2, direccion, administrador, local As String
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
        clase.agregar_registro("INSERT INTO `tiendas`(`tienda`,`ciudad`,`email`,`direccion`,`local`,`empresa`,`telefono1`,`telefono2`,`administrador`,`estado`) VALUES ('" & UCase(textBox2.Text) & "','" & UCase(TextBox1.Text) & "','" & UCase(TextBox3.Text) & "'," & direccion & "," & local & ",'" & ComboBox1.SelectedValue & "'," & telefono1 & "," & telefono2 & "," & administrador & ",True)")
        clase.agregar_registro("INSERT INTO usuarios_tienda (tienda, usuario, password) VALUES ('" & ultima_tienda() & "', '" & LCase(TextBox3.Text) & "', '" & MD5EncryptPass(TextBox11.Text) & "') ")
        restablecer()
        frm_tienda.llenar_tienda(frm_tienda.devolver_valor(frm_tienda.ComboBox2.Text))
    End Sub

    Function ultima_tienda() As Short
        clase.consultar1("select * from tiendas", "ultimo")
        Return clase.dt1.Tables("ultimo").Rows(clase.dt1.Tables("ultimo").Rows.Count - 1)("id")
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        restablecer()
    End Sub

    Private Sub restablecer()
        TextBox1.Text = ""
        textBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        ComboBox1.SelectedIndex = -1
        textBox2.Focus()
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
        TextBox10.Text = TextBox3.Text
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

    Private Sub TextBox9_KeyPress1(sender As Object, e As KeyPressEventArgs)
        clase.validar_numeros(e)
    End Sub
End Class