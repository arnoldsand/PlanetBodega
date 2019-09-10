Public Class crear_tallas
    Dim clase As class_library = New class_library
    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(textBox2, "Nombre Tallas") = False Then Exit Sub
        clase.consultar("select* from tallas where nombretalla = '" & textBox2.Text & "'", "tabla2")
        If clase.dt.Tables("tabla2").Rows.Count > 0 Then
            MessageBox.Show("La talla que intenta crear ya existe y no se puede volver a insertar. Pulse aceptar para volverlo a intentar.")
            Exit Sub
        Else
            clase.agregar_registro("insert into tallas (nombretalla) values ('" & UCase(textBox2.Text) & "')")
            clase.llenar_combo_datagrid(frm_colores.tallas, "select* from tallas order by nombretalla asc", "nombretalla", "codigo_talla")
            Me.Close()
        End If
    End Sub
End Class