Public Class frm_cambiar_numero_seccion
    Dim clase As New class_library
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frm_cambiar_numero_seccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox1, "select* from bodegas order by bod_codigo asc", "bod_nombre", "bod_codigo")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_combobox(ComboBox3, "Conteo") = False Then Exit Sub
        If clase.validar_combobox(ComboBox1, "Bodega") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox18, "Góndola") = False Then Exit Sub
        If frm_generar_ajuste_para_gondola2.verificar_existencia_gondola(TextBox18.Text, ComboBox1.SelectedValue) = False Then
            MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox18.Text = ""
            TextBox18.Focus()
            Exit Sub
        End If
        clase.consultar("SELECT secc_codigo FROM secciones_inventario WHERE (secc_bodega =" & ComboBox1.SelectedValue & "  AND secc_gondola ='" & UCase(TextBox18.Text) & "' AND secc_conteo =" & ComboBox3.SelectedIndex + 1 & " AND secc_inventario = " & frm_modulo_inventario.TextBox1.Text & ")", "busqueda")
        If clase.dt.Tables("busqueda").Rows.Count > 0 Then
            frm_recibir_datos_inventario.TextBox1.Text = clase.dt.Tables("busqueda").Rows(0)("secc_codigo")
            Me.Close()
        End If
    End Sub
End Class