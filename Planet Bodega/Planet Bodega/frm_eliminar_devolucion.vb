Public Class frm_eliminar_devolucion
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub




    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox4.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos." & campo_codigobarra_largo(frmTransferenciaDevoluciones.ComboBox9.SelectedValue) & " ='" & TextBox4.Text & "')", "tabla11")
            ' codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox4.Text)
        End If
        If Len(TextBox4.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos." & campo_codigobarra_corto(frmTransferenciaDevoluciones.ComboBox9.SelectedValue) & " =" & TextBox4.Text & ")", "tabla11")
            'codigo_articulo = TextBox4.Text
        End If
        If Len(TextBox4.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            codigo_articulo = clase.dt.Tables("tabla11").Rows(0)("ar_codigo")
            clase.consultar("SELECT dt_cantidad FROM dettransferencia WHERE (dt_trnumero = " & frmTransferenciaDevoluciones.TextBox1.Text & " AND dt_codarticulo =" & codigo_articulo & ")", "tbl")
            If clase.dt.Tables("tbl").Rows.Count > 0 Then
                Dim v As String = MessageBox.Show("¿Desea borrar el codigo digitado?", "BORRAR CODIGO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If v = 6 Then
                    Dim totcant As Integer = clase.dt.Tables("tbl").Rows(0)("dt_cantidad")
                    If totcant > 1 Then
                        clase.actualizar("UPDATE dettransferencia SET dt_cantidad = " & totcant - 1 & " WHERE (dt_trnumero = " & frmTransferenciaDevoluciones.TextBox1.Text & " AND dt_codarticulo =" & codigo_articulo & ")")

                    Else
                        clase.borradoautomatico("Delete from dettransferencia WHERE (dt_trnumero = " & frmTransferenciaDevoluciones.TextBox1.Text & " AND dt_codarticulo =" & codigo_articulo & ")")
                    End If
                    frmTransferenciaDevoluciones.llenar_grilla(frmTransferenciaDevoluciones.TextBox1.Text)
                    Dim x As Short
                    For x = 0 To frmTransferenciaDevoluciones.DataGridView1.RowCount - 1
                        If codigo_articulo = frmTransferenciaDevoluciones.DataGridView1.Item(0, x).Value Then
                            frmTransferenciaDevoluciones.DataGridView1.CurrentCell = frmTransferenciaDevoluciones.DataGridView1.Item(0, x)
                            Exit For
                        End If
                    Next
                    TextBox4.Text = ""
                    TextBox4.Focus()
                End If
            Else
                MessageBox.Show("El codigo digitado no se puede eliminar porque no ha sido agregado antes a la lista de captura.", "CODIGO NO AGREGADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox4.Text = ""
                TextBox4.Focus()
            End If
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub
End Class