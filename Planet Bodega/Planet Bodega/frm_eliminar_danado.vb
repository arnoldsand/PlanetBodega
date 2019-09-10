Public Class frm_eliminar_danado
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox4.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox4.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox4.Text)
        End If
        If Len(TextBox4.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox4.Text & ")", "tabla11")
            codigo_articulo = TextBox4.Text
        End If
        If Len(TextBox4.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            Dim v As String = MessageBox.Show("¿Desea ")
            Dim a As Short
            Dim indfind As Boolean = False 'esta variable me indica si el articulo en mencion fue encontrado o no en la lista de captura
            For a = 0 To frm_reporte_de_danado.DataGridView1.RowCount - 1
                If codigo_articulo = frm_reporte_de_danado.DataGridView1.Item(0, a).Value Then
                    indfind = True
                    Dim cant As Short = frm_reporte_de_danado.DataGridView1.Item(6, a).Value
                    If cant > 1 Then
                        frm_reporte_de_danado.DataGridView1.Item(6, a).Value = cant - 1
                    Else
                        frm_reporte_de_danado.DataGridView1.Rows.Remove(frm_reporte_de_danado.DataGridView1.Rows(a))
                    End If
                    frm_reporte_de_danado.DataGridView1.CurrentCell = frm_reporte_de_danado.DataGridView1.Item(0, a - 1)
                    Exit For
                End If
            Next
            If indfind = False Then
                MessageBox.Show("El articulos especificado no fue agregado antes a la lista de captura.", "ARTICULO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            TextBox4.Text = ""
            TextBox4.Focus()

        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If
    End Sub
End Class