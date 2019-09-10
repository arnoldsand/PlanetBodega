Public Class asignar_cant_patron_dist
    Dim clase As New class_library
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Cant") = False Then Exit Sub
        Dim b As Short
        Dim ind As Boolean = False
        For b = 0 To frm_patron_mercancia_no_importada.DataGridView2.RowCount - 1
            If frm_patron_mercancia_no_importada.DataGridView2.Item(3, b).Value = True Then
                ind = True
                Exit For
            End If
        Next
        If ind = True Then

            For b = 0 To frm_patron_mercancia_no_importada.DataGridView2.RowCount - 1
                If frm_patron_mercancia_no_importada.DataGridView2.Item(3, b).Value = True Then
                    frm_patron_mercancia_no_importada.DataGridView2.Item(2, b).Value = textBox2.Text
                End If
            Next
            frm_patron_mercancia_no_importada.calcular_total()
            Me.Close()
        Else
            MessageBox.Show("Debe seleccionar por lo menos un item para agregar cantidades multiples.", "CANTIDADES MULTIPLES", MessageBoxButtons.OK, MessageBoxIcon.Information)
            textBox2.Focus()
        End If
    End Sub
End Class