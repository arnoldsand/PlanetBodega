Public Class frm_detalles_pedidos
    Dim clase As New class_library
    Private Sub frm_detalles_pedidos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, detpedido.detped_cant, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM articulos INNER JOIN detpedido ON (articulos.ar_codigo = detpedido.detped_articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (detpedido.detped_codpedido =" & frm_mantenimiento_pedidos.DataGridView1.Item(0, frm_mantenimiento_pedidos.DataGridView1.CurrentRow.Index).Value & ")", "detpedido")
        If clase.dt.Tables("detpedido").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("detpedido")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 80
            .Columns(1).Width = 140
            .Columns(2).Width = 180
            .Columns(3).Width = 130
            .Columns(4).Width = 130
            .Columns(5).Width = 60
            .Columns(0).HeaderText = "Codigos"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Linea"
            .Columns(4).HeaderText = "Sublinea"
            .Columns(5).HeaderText = "Cant"
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            clase.consultar("select ar_foto from articulos where ar_codigo = " & .Item(0, .CurrentCell.RowIndex).Value & "", "tablita")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        End With
    End Sub
End Class