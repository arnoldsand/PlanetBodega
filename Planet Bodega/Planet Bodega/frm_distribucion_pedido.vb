Public Class frm_distribucion_pedido
    Dim codpedido As String
    Dim clase As New class_library
    Dim codtienda As Integer
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_distribucion_pedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        codpedido = frm_patron_despacho.TextBox18.Text
        clase.consultar("SELECT cabped_tienda FROM cabpedido WHERE (cabped_codigo =" & codpedido & ")", "codtienda")
        codtienda = clase.dt.Tables("codtienda").Rows(0)("cabped_tienda")
        TextBox18.Text = codpedido
        TextBox1.Text = nombretienda(codtienda)
        clase.consultar("SELECT CONCAT(bodegas.bod_abreviatura, ' - ',  patron_despacho.desp_gondola) AS bodegagondola, patron_despacho.desp_articulo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, patron_despacho.desp_cant, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM bodegas  INNER JOIN patron_despacho ON (bodegas.bod_codigo = patron_despacho.desp_bodega) INNER JOIN articulos ON (articulos.ar_codigo = patron_despacho.desp_articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (patron_despacho.desp_pedido = " & TextBox18.Text & ") ORDER BY patron_despacho.desp_bodega, patron_despacho.desp_gondola ASC", "listapatron")
        If clase.dt.Tables("listapatron").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("listapatron")
            prepara_columnas()
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Sub prepara_columnas()
        With DataGridView1
            .Columns(7).Visible = False
            .Columns(8).Visible = False
            .Columns(9).Visible = False
            .Columns(10).Visible = False
            .Columns(11).Visible = False
            .Columns(0).HeaderText = "Góndola"
            .Columns(1).HeaderText = "Articulo"
            .Columns(2).HeaderText = "Referencia"
            .Columns(3).HeaderText = "Descripción"
            .Columns(4).HeaderText = "Línea"
            .Columns(5).HeaderText = "Sublinea"
            .Columns(6).HeaderText = "Cant"
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim v As String
        v = MessageBox.Show("¿Desea imprimir el pedido actual?", "IMPRIMIR PEDIDO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            frm_patron_despacho.imprimir_patron_pedido()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        clase.consultar("select ar_foto from articulos where ar_codigo = " & DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value & "", "tablita")
        If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
            PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
        Else
            PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
        End If
        SetImage(PictureBox1)
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox18_GotFocus(sender As Object, e As EventArgs) Handles TextBox18.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub
End Class