Public Class frmDanado
    Dim clase As New class_library

    Public Danado As String
    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If txtDanado.Text = "" Then
            MessageBox.Show("DEBE COLOCAR UNA CANTIDAD, PUEDE SER 0", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Danado = txtDanado.Text
        Me.Close()
    End Sub

    Private Sub txtDanado_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDanado.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub
End Class