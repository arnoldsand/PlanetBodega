Public Class Codigos_de_Barra_para_Preinspección
    Dim clase As New class_library
    Private Sub Codigos_de_Barra_para_Preinspección_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Sub textBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textBox2.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(textBox2, "Cantidad de Codigos") = False Then Exit Sub

    End Sub
End Class