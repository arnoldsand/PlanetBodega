Public Class frm_captura_multiple
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Public Sub New(Cantidad As Integer)

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()
        TextBox4.Text = Cantidad
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub



    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Cantidad") = False Then Exit Sub
        If Val(TextBox4.Text) = 0 Then
            MessageBox.Show("Debe digitar una cantidad válida diferente de 0.")
            Exit Sub
        End If
        frmTransferencia.cant_captura = TextBox4.Text
        Me.Close()
    End Sub

    Private Sub frm_captura_multiple_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class