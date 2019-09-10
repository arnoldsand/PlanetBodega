Public Class frm_password_liquidacion
    Dim clase As New class_library
    Public ValidaPassword1 As String

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtPassword, "Password") = False Then Exit Sub
        ValidaPassword1 = txtPassword.Text
        'frmRevision.ValidaPassword = txtPassword.Text
        'frm_diferencias_transferencia.ValidaPassword = txtPassword.Text
        Me.Close()
    End Sub

    Private Sub frm_password_revision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtPassword.Focus()
    End Sub

    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        clase.enter(e)
    End Sub
End Class