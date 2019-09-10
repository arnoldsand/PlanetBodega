Public Class frm_password_ajustes
    Dim clase As New class_library
    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtPassword, "Contraseña Ajustes") = False Then Exit Sub
        clase.consultar1("select password_ajustes from informacion", "pw")
        If clase.dt1.Tables("pw").Rows.Count > 0 Then
            If clase.dt1.Tables("pw").Rows(0)("password_ajustes") = txtPassword.Text Then
                passtrue = True
                Me.Close()
            Else
                MessageBox.Show("La contraseña digitada es incorrecta.", "CONTRASEÑA INCORRECTA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPassword.Focus()
                txtPassword.Text = ""
            End If

        End If
    End Sub
End Class