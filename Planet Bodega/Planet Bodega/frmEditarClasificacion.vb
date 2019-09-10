Public Class frmEditarClasificacion
    Dim clase As New class_library
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(txtClasificaciones, "Clasificacion") = False Then Exit Sub
        clase.actualizar("UPDATE CabClasificacion SET Clasificacion = '" & txtClasificaciones.Text.ToUpper().Trim() & "' WHERE IdClasificacion = " & frmClasificaciones.DataGridView1.Item(0, frmClasificaciones.DataGridView1.CurrentRow.Index).Value & "")
        Me.Close()
    End Sub

    Private Sub frmEditarClasificacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtClasificaciones.Text = frmClasificaciones.DataGridView1.Item(1, frmClasificaciones.DataGridView1.CurrentRow.Index).Value
    End Sub
End Class