Public Class frm_codigos_asociados
    Dim clase As New class_library
    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_codigos_asociados_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.RowCount = codigos_de_barra.Length
        Dim c As Integer
        For c = 0 To codigos_de_barra.Length - 1
            DataGridView1.Item(0, c).Value = codigos_de_barra(c)
            DataGridView1.Item(1, c).Value = color_apartirdecodigo(colores(c)) & " - " & nombre_talla(arraytallas(c))
        Next
    End Sub

    Function nombre_talla(codigotalla) As String
        clase.consultar("select* from tallas where codigo_talla = " & codigotalla & "", "tablita")
        Return clase.dt.Tables("tablita").Rows(0)("nombretalla")
    End Function
End Class