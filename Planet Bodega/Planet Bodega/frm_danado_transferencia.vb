Public Class frm_danado_transferencia
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox4.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos." & campo_codigobarra_largo(frmTransferencia.ComboBox9.SelectedValue) & " ='" & TextBox4.Text & "')", "tabla11")
            ' codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox4.Text)
        End If
        If Len(TextBox4.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos." & campo_codigobarra_corto(frmTransferencia.ComboBox9.SelectedValue) & " =" & TextBox4.Text & ")", "tabla11")
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
            clase.consultar1("SELECT dt_danado FROM dettransferencia WHERE (dt_trnumero =" & frmTransferencia.TextBox1.Text & " AND dt_codarticulo =" & codigo_articulo & ")", "trans")
            If clase.dt1.Tables("trans").Rows.Count > 0 Then
                Dim cant As Short = clase.dt1.Tables("trans").Rows(0)("dt_danado")
                clase.actualizar("UPDATE dettransferencia SET dt_danado = " & cant + 1 & " WHERE (dt_trnumero =" & frmTransferencia.TextBox1.Text & " AND dt_codarticulo =" & codigo_articulo & ")")
            Else
                clase.agregar_registro("INSERT INTO `dettransferencia`(`dt_numero`,`dt_trnumero`,`dt_gondola`,`dt_codarticulo`,`dt_cantidad`,`dt_faltante`,`dt_danado`,`dt_costo`,`dt_venta1`,`dt_venta2`) VALUES ( NULL,'" & frmTransferencia.TextBox1.Text & "',NULL,'" & codigo_articulo & "','0','0','1','" & Str(clase.dt.Tables("tabla11").Rows(0)("ar_costo")) & "','" & Str(clase.dt.Tables("tabla11").Rows(0)("ar_precio1")) & "','" & Str(clase.dt.Tables("tabla11").Rows(0)("ar_precio2")) & "')")
            End If
            frmTransferencia.llenar_grilla(frmTransferencia.TextBox1.Text)
            TextBox4.Text = ""
            TextBox4.Focus()
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