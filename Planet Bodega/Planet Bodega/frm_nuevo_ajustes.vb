Public Class frm_nuevo_ajustes
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_nuevo_ajustes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox6, "select* from tipos_ajustes where tip_tipo IS NOT NULL order by tip_nombre asc", "tip_nombre", "tip_codigo")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ComboBox6.SelectedIndex = -1
        TextBox2.Text = ""
        TextBox1.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_combobox(ComboBox6, "Tipo Ajuste") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Operario") = False Then Exit Sub
        Dim fecha As Date = Now
        clase.agregar_registro("INSERT INTO `cabajuste`(`cabaj_fecha`,`cabaj_hora`,`cabaj_tipo_ajuste`,`cabaj_operario`,`cabaj_observaciones`,`cabaj_procesado`) VALUES ('" & fecha.ToString("yyyy-MM-dd") & "','" & Now.ToString("HH:mm:ss") & "','" & ComboBox6.SelectedValue & "','" & UCase(TextBox1.Text) & "','" & UCase(TextBox2.Text) & "" & "',False)")
        frm_ajustes.ComboBox6.SelectedIndex = 2
        frm_ajustes.llenar_ajustes(2, frm_ajustes.DateTimePicker1.Value.ToString("yyyy-MM-dd"), frm_ajustes.DateTimePicker2.Value.ToString("yyyy-MM-dd"))
        Me.Close()
    End Sub
End Class