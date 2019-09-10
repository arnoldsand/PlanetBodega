Public Class frm_faltante_liquidacion
    Dim clase As New class_library
    Private Sub frm_faltante_liquidacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
            textBox2.Text = frm_cruze_facturas.dtgridcajas.Item(0, frm_cruze_facturas.dtgridcajas.CurrentRow.Index).Value
            TextBox1.Text = frm_cruze_facturas.dtgridcajas.Item(1, frm_cruze_facturas.dtgridcajas.CurrentRow.Index).Value
            TextBox3.Text = frm_cruze_facturas.dtgridcajas.Item(2, frm_cruze_facturas.dtgridcajas.CurrentRow.Index).Value
            TextBox5.Text = frm_cruze_facturas.dtgridcajas.Item(3, frm_cruze_facturas.dtgridcajas.CurrentRow.Index).Value
        
    End Sub


   

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub


    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox4, "Faltante") = False Then Exit Sub
        Dim v As String = MessageBox.Show("No podrá modificar posteriormente el dato de faltante. ¿Desea Continuar?", "REPORTAR FALTANTE", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            clase.actualizar("UPDATE detfacturas SET detfact_faltante = " & TextBox4.Text & " WHERE detfact_codigo = " & frm_cruze_facturas.dtgridcajas.Item(5, frm_cruze_facturas.dtgridcajas.CurrentCell.RowIndex).Value & "")
            frm_cruze_facturas.llenar_dtgrid_referencias(frm_cruze_facturas.TextBox1.Text, frm_cruze_facturas.DataGridView1.Item(0, frm_cruze_facturas.DataGridView1.CurrentCell.RowIndex).Value)
            Me.Close()
        End If
    End Sub
End Class