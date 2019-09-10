Public Class frm_contrasena_revision
    Dim clase As New class_library
    Public validado As Boolean

    Private Sub frm_contrasena_revision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        validado = False
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtcontrasena, "Contraseña") = False Then Exit Sub
        clase.consultar2("select password_revision from informacion", "contrasena")
        If clase.dt2.Tables("contrasena").Rows.Count > 0 Then
            If IsDBNull(clase.dt2.Tables("contrasena").Rows(0)("password_revision")) Then
                MessageBox.Show("No hay ninguna información establecida.", "CONTRASEÑA NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If txtcontrasena.Text = clase.dt2.Tables("contrasena").Rows(0)("password_revision") Then
                    validado = True
                    Me.Close()
                Else
                    MessageBox.Show("La contraseña digitada es incorrecta. Vuelva a intentarlo.", "CONTRASEÑA INCORRECTA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtcontrasena.Text = ""
                    txtcontrasena.Focus()
                End If
            End If
        Else
            MessageBox.Show("No hay ninguna información establecida.", "CONTRASEÑA NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub
End Class