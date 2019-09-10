Public Class frmCantidad
    Dim clase As New class_library
    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If txtCantidad.Text = "" Then
            MessageBox.Show("DEBES COLOCAR LA CANTIDAD REVISADA PARA CONTINUAR", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Len(txtCantidad.Text) > 3 Then
            MessageBox.Show("Revise la cantidad", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCantidad.Text = ""
            txtCantidad.Focus()
            Exit Sub
        End If
        frmRevision.CantidadRev = txtCantidad.Text
        frmRevision.txtRevisadas.Text = Val(frmRevision.txtRevisadas.Text) + Val(txtCantidad.Text)

        Me.Dispose()
    End Sub

    Private Sub txtCantidad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCantidad.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub
End Class