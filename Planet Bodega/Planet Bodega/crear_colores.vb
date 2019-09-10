Public Class crear_colores
    Dim clase As class_library = New class_library
    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(textBox2, "Nombre Color") = False Then Exit Sub
        clase.consultar("select* from colores where colornombre = '" & textBox2.Text & "'", "tabla2")
        If clase.dt.Tables("tabla2").Rows.Count > 0 Then
            MessageBox.Show("El color que intenta crear ya existe y no se puede volver a insertar. Pulse aceptar para volverlo a intentar.")
            Exit Sub
        Else
            clase.agregar_registro("insert into colores (colornombre) values ('" & UCase(textBox2.Text) & "')")
            clase.llenar_combo_datagrid(frm_colores.cmbcolores, "select* from colores order by colornombre asc", "colornombre", "cod_color")
            Me.Close()
        End If
    End Sub

    Private Sub crear_colores_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class