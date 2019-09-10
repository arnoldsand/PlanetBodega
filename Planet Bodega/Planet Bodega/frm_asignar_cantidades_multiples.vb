Public Class frm_asignar_cantidades_multiples
    Dim clase As New class_library
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Cantidad") = False Then Exit Sub
        Dim b As Short
        Dim ind As Boolean = False
        For b = 0 To frm_impresion_codigos_de_barra.dataGridView1.RowCount - 1
            If frm_impresion_codigos_de_barra.dataGridView1.Item(0, b).Value = True Then
                ind = True
                Exit For
            End If
        Next
        If ind = True Then
            Select Case RadioButton1.Checked
                Case False
                    For b = 0 To frm_impresion_codigos_de_barra.dataGridView1.RowCount - 1
                        If frm_impresion_codigos_de_barra.dataGridView1.Item(0, b).Value = True Then
                            With frm_impresion_codigos_de_barra.dataGridView1
                                .Item(8, b).Value = textBox2.Text
                                .Item(9, b).Value = .Item(7, b).Value - textBox2.Text
                            End With
                        End If
                    Next
                Case True
                    For b = 0 To frm_impresion_codigos_de_barra.dataGridView1.RowCount - 1
                        If frm_impresion_codigos_de_barra.dataGridView1.Item(0, b).Value = True Then
                            With frm_impresion_codigos_de_barra.dataGridView1
                                .Item(8, b).Value = .Item(7, b).Value - textBox2.Text
                                .Item(9, b).Value = textBox2.Text
                            End With
                        End If
                    Next
            End Select
        Else
            MessageBox.Show("Debe seleccionar por lo menos una referencia para asignaciones multiples.", "ASIGNACIONES MULTIPLES", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If
        Me.Close()
    End Sub
End Class