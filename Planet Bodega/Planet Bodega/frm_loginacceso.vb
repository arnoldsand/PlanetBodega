Public Class frm_loginacceso
    Dim clase As class_library = New class_library
    Dim ind_cierre As Boolean = False
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox1, "Usuario") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox2, "Contraseña") = False Then Exit Sub
        clase.consultar("select* from usuarios where usuario = '" & TextBox1.Text & "'", "resultados")
        If clase.dt.Tables("resultados").Rows.Count > 0 Then
            If TextBox2.Text = clase.dt.Tables("resultados").Rows(0)("pwd") Then
                ind_cierre = True
                Me.Close()
            Else
                MessageBox.Show("El Usuario o la Contraseña especificados son incorrectos. Pulse aceptar para voverlo a intentar.", "USUARIO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.Focus()
            End If
        Else
            MessageBox.Show("El Usuario o la Contraseña especificados son incorrectos. Pulse aceptar para voverlo a intentar.", "USUARIO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If
    End Sub


    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub


    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub


    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub frm_login_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If ind_cierre = False Then
            End
        End If
    End Sub

   

  
End Class