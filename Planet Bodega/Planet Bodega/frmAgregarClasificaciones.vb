Public Class frmAgregarClasificaciones
    Dim clase As New class_library

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(txtClasificaciones, "Clasificacion") = False Then Exit Sub
        clase.agregar_registro("insert into CabClasificacion (Clasificacion) values ('" & txtClasificaciones.Text.ToUpper().Trim() & "')")
        txtClasificaciones.Text = ""
        txtClasificaciones.Focus()
    End Sub
End Class