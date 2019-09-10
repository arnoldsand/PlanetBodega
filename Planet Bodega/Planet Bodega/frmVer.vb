Public Class frmVer
    Public Codigo As String
    Dim clase As New class_library

    Private Sub btnRegresar_Click(sender As Object, e As EventArgs) Handles btnRegresar.Click
        Me.Close()
        frmConsulta_x_foto.Show()
    End Sub

    Private Sub frmVer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = False
        frmConsulta_x_foto.Show()
    End Sub

    Private Sub frmVer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT ar_codigo,ar_referencia,ar_descripcion,ar_foto FROM articulos WHERE ar_codigo='" & Codigo & "'", "consulta")

        txtCodigo.Text = clase.dt.Tables("consulta").Rows(0)("ar_codigo")
        txtReferencia.Text = clase.dt.Tables("consulta").Rows(0)("ar_referencia")
        txtDescripcion.Text = clase.dt.Tables("consulta").Rows(0)("ar_descripcion")

        Try
            Foto.Image = Image.FromFile(clase.dt.Tables("consulta").Rows(0)("ar_foto"))
        Catch ex As Exception
            Foto.Image = Global.WindowsApplication1.My.Resources.sinfoto
        End Try
        Foto.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub
End Class