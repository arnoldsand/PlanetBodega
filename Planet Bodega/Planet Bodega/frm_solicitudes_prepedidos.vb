Public Class frm_solicitudes_prepedidos
    Dim clase As New class_library

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_solicitudes_prepedidos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT solicitudes_productos.solpro_fecha, solicitudes_productos.solpro_hora, tiendas.tienda, solicitudes_productos.solpro_pregunta1, solicitudes_productos.solpro_pregunta2, solicitudes_productos.solpro_pregunta3, solicitudes_productos.solpro_foto1, solicitudes_productos.solpro_foto2, solicitudes_productos.solpro_foto3, solicitudes_productos.solpro_operario1 FROM tiendas INNER JOIN solicitudes_productos ON (tiendas.id = solicitudes_productos.solpro_tienda) WHERE (solicitudes_productos.solpro_codigo =" & frm_mantenimiento_prepedidos.DataGridView1.Item(0, frm_mantenimiento_prepedidos.DataGridView1.CurrentRow.Index).Value & ")", "solcitud")
        If clase.dt.Tables("solcitud").Rows.Count > 0 Then
            textBox2.Text = clase.dt.Tables("solcitud").Rows(0)("solpro_pregunta1")
            TextBox1.Text = clase.dt.Tables("solcitud").Rows(0)("solpro_pregunta2")
            TextBox3.Text = clase.dt.Tables("solcitud").Rows(0)("solpro_pregunta3")
            Dim campofoto As String
            If clase.dt.Tables("solcitud").Rows(0)("solpro_foto1") <> "" Then

                campofoto = clase.dt.Tables("solcitud").Rows(0)("solpro_foto1")
                PictureBox1.Image = Image.FromFile("\\MAQUINAUBUNTU\Fotos\" & campofoto)
                SetImage(PictureBox1)
           
            End If
            If clase.dt.Tables("solcitud").Rows(0)("solpro_foto2") <> "" Then
                campofoto = clase.dt.Tables("solcitud").Rows(0)("solpro_foto2")
                PictureBox2.Image = Image.FromFile("\\MAQUINAUBUNTU\Fotos\" & campofoto)
                SetImage(PictureBox2)
            End If
            If clase.dt.Tables("solcitud").Rows(0)("solpro_foto3") <> "" Then
                campofoto = clase.dt.Tables("solcitud").Rows(0)("solpro_foto3")

                PictureBox3.Image = Image.FromFile("\\MAQUINAUBUNTU\Fotos\" & campofoto)
                SetImage(PictureBox3)
            End If
        End If
    End Sub

End Class