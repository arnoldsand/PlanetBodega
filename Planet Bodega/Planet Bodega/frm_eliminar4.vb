Public Class frm_eliminar4
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_eliminar_item_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox4.Focus()
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox4.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox4.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox4.Text)
        End If
        If Len(TextBox4.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox4.Text & ")", "tabla11")
            codigo_articulo = TextBox4.Text
        End If
        If Len(TextBox4.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            Dim v As String = MessageBox.Show("¿Desea borrar el codigo digitado?", "BORRAR CODIGO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                With frm_generar_ajuste_para_gondola2.DataGridView1
                    Dim z As Short
                    For z = 0 To .RowCount - 1
                        If codigo_articulo = .Item(0, z).Value Then
                            Dim cant As Integer = .Item(6, z).Value
                            If cant = 0 Then
                                MessageBox.Show("No se puede eliminar el articulo digitado porque no ha sido agregado antes a la lista de ajustes.", "CODIGO NO AGREGADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                TextBox4.Text = ""
                                TextBox4.Focus()
                                Exit Sub
                            End If
                            Select Case .Item(7, z).Value
                                Case 1
                                    .Item(6, z).Value = cant - 1
                                Case -1
                                    .Item(6, z).Value = cant + 1
                            End Select
                            quitar_articulos_con_cantidades_cero()
                            TextBox4.Text = ""
                            TextBox4.Focus()
                            frm_generar_ajuste_para_gondola2.cantidad_unidades()
                            Exit Sub
                        End If
                    Next
                    MessageBox.Show("No se puede eliminar el articulo digitado porque no ha sido agregado antes a la lista de ajustes.", "CODIGO NO AGREGADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox4.Text = ""
                    TextBox4.Focus()
                    Exit Sub
                End With
            End If
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If
    End Sub

    Private Sub quitar_articulos_con_cantidades_cero()
        With frm_generar_ajuste_para_gondola2.DataGridView1
            For Each p As DataGridViewRow In .Rows
                If p.Cells(6).Value = 0 Then
                    .Rows.Remove(p)
                End If
            Next
        End With
    End Sub
End Class