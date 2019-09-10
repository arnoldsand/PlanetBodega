Public Class frm_dian_registro_dian
    Dim clase As New class_library
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox9, "Registro DIAN") = False Then Exit Sub
        clase.actualizar("UPDATE detalle_importacion_detcajas SET detcab_registrodian = '" & TextBox9.Text & "', detcab_fecharegistrodian = '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' WHERE detcab_coditem = " & frm_registro_dian_referencia.DataGridView1.Item(9, frm_registro_dian_referencia.DataGridView1.CurrentCell.RowIndex).Value & "")
        frm_registro_dian_referencia.llenar_grilla(frm_registro_dian_referencia.TextBox1.Text)
        Me.Close()
    End Sub


    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub frm_dian_registro_dian_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT detcab_fecharegistrodian, detcab_registrodian FROM detalle_importacion_detcajas WHERE (detcab_coditem =" & frm_registro_dian_referencia.DataGridView1.Item(9, frm_registro_dian_referencia.DataGridView1.CurrentCell.RowIndex).Value & ")", "registro")
        If clase.dt.Tables("registro").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("registro").Rows(0)("detcab_fecharegistrodian")) = False Then
                DateTimePicker1.Value = clase.dt.Tables("registro").Rows(0)("detcab_fecharegistrodian")
            End If
            If IsDBNull(clase.dt.Tables("registro").Rows(0)("detcab_registrodian")) = False Then
                TextBox9.Text = clase.dt.Tables("registro").Rows(0)("detcab_registrodian")
            End If
        End If
    End Sub
End Class