Public Class frm_eliminar_desde_ajuste
    Dim clase As New class_library
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
            clase.consultar1("select detaj_cantidad from detajuste WHERE (detaj_codigo_ajuste =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detaj_articulo = " & codigo_articulo & ")", "123")
            If clase.dt1.Tables("123").Rows.Count > 0 Then
                Dim v As String = MessageBox.Show("¿Desea borrar el codigo digitado?", "BORRAR CODIGO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If v = 6 Then
                    Dim totcant As Integer = clase.dt1.Tables("123").Rows(0)("detaj_cantidad")
                    If totcant > 1 Then
                        clase.actualizar("UPDATE detajuste SET detaj_cantidad = " & totcant - 1 & " WHERE (detaj_codigo_ajuste =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detaj_articulo = " & codigo_articulo & ")")
                    Else
                        clase.borradoautomatico("Delete from detajuste WHERE (detaj_codigo_ajuste =" & frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value & " AND detaj_articulo = " & codigo_articulo & ")")
                    End If
                    frm_detalle_ajuste.llenar_listado(frm_ajustes.dataGridView1.Item(0, frm_ajustes.dataGridView1.CurrentRow.Index).Value)
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

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_eliminar_desde_ajuste_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub
End Class