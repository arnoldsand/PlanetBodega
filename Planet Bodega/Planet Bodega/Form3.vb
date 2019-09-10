Public Class Form3
    Dim clase As New class_library

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        clase.consultar("select* from articulos where ar_codigo between 128343 and 128484", "consulta")
        If clase.dt.Tables("consulta").Rows.Count > 0 Then
            Dim a As Integer
            For a = 0 To clase.dt.Tables("consulta").Rows.Count - 1
                Application.DoEvents()
                clase.actualizar("update articulos set ar_codigo2 = " & clase.dt.Tables("consulta").Rows(a)("ar_codigo") & ", ar_codigobarras2 = '" & clase.dt.Tables("consulta").Rows(a)("ar_codigobarras") & "' where ar_codigo = " & clase.dt.Tables("consulta").Rows(a)("ar_codigo") & "")
                TextBox1.Text = a + 1
            Next
        End If
    End Sub
End Class