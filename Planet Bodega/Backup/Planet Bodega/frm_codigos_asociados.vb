Public Class frm_codigos_asociados

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_codigos_asociados_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.RowCount = codigos_de_barra.Length
        Dim c As Integer
        For c = 0 To codigos_de_barra.Length - 1
            DataGridView1.Item(0, c).Value = codigos_de_barra(c)
            DataGridView1.Item(1, c).Value = color_apartirdecodigo(colores(c)) & " - " & arraytallas(c)
        Next
    End Sub
End Class