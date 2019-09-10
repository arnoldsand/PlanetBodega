Public Class frm_busqueda_articulos
    Dim clase As New class_library
    Dim codigo As String
    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If codigo = "" Then
            Me.Close()
        Else
            frm_ingreso_no_importado.txtArticulo.Text = grdConsulta(0, grdConsulta.CurrentCell.RowIndex).Value.ToString
            Me.Close()
        End If

    End Sub

    Private Sub txtConsulta_TextChanged(sender As Object, e As EventArgs) Handles txtConsulta.TextChanged
        llenar_rejilla()
    End Sub


    Sub llenar_rejilla()
        If chkReferencia.Checked = False And chkDescripcion.Checked = False Then
            MessageBox.Show("DEBE SELECCIONAR UN CRITERIO DE BUSQUEDA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If chkReferencia.Checked = True Then
            clase.consultar("SELECT articulos.ar_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, linea1.ln1_nombre AS LINEA, ar_foto AS FOTO, articulos.ar_costo AS Costo, articulos.ar_precio1 As Precio1 , articulos.ar_precio2 AS Precio2 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) WHERE articulos.ar_referencia LIKE '%" & txtConsulta.Text & "%';", "consulta")
        End If
        If chkDescripcion.Checked = True Then
            clase.consultar("SELECT articulos.ar_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, linea1.ln1_nombre AS LINEA, ar_foto AS FOTO, articulos.ar_costo AS Costo, articulos.ar_precio1 As Precio1, articulos.ar_precio2 AS Precio2 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) WHERE articulos.ar_descripcion LIKE '%" & txtConsulta.Text & "%';", "consulta")
        End If
        With grdConsulta
            .DataSource = clase.dt.Tables("consulta")
            .RowHeadersWidth = 4
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(1).Width = 150
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(2).Width = 150
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).Width = 150
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).Visible = False
        End With
    End Sub

    Private Sub chkReferencia_CheckedChanged(sender As Object, e As EventArgs) Handles chkReferencia.CheckedChanged
        chkDescripcion.Checked = False
    End Sub

    Private Sub chkDescripcion_CheckedChanged(sender As Object, e As EventArgs) Handles chkDescripcion.CheckedChanged
        chkReferencia.Checked = False
    End Sub

    Private Sub grdConsulta_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdConsulta.CellClick
        Dim Ruta As String = grdConsulta(4, grdConsulta.CurrentCell.RowIndex).Value.ToString
        codigo = grdConsulta(0, grdConsulta.CurrentCell.RowIndex).Value.ToString
        If Ruta <> "" Then
            Try
                FOTO.Image = Image.FromFile(Ruta)
                FOTO.SizeMode = PictureBoxSizeMode.StretchImage
            Catch ex As Exception
                FOTO.Image = Nothing
            End Try
        End If
    End Sub

    Private Sub frmBusqueda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        chkReferencia.Checked = True
        llenar_rejilla()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        indicador_de_formulario = True
        frm_crear_articulo.ShowDialog()
        frm_crear_articulo.Dispose()
        indicador_de_formulario = False
    End Sub

    Private Sub Precios_Click(sender As Object, e As EventArgs) Handles Precios.Click
        If grdConsulta.RowCount = 0 Then
            Exit Sub
        End If
        frm_cambio_precios.txtCodigo.Text = grdConsulta(0, grdConsulta.CurrentCell.RowIndex).Value.ToString
        frm_cambio_precios.txtReferencia.Text = grdConsulta(1, grdConsulta.CurrentCell.RowIndex).Value.ToString
        frm_cambio_precios.txtDescripcion.Text = grdConsulta(2, grdConsulta.CurrentCell.RowIndex).Value.ToString
        frm_cambio_precios.txtCosto.Text = extraer_costo2(grdConsulta(0, grdConsulta.CurrentCell.RowIndex).Value)
        frm_cambio_precios.txtCosto2.Text = grdConsulta(5, grdConsulta.CurrentCell.RowIndex).Value.ToString
        frm_cambio_precios.txtPrecio1.Text = grdConsulta(6, grdConsulta.CurrentCell.RowIndex).Value.ToString
        frm_cambio_precios.txtPrecio2.Text = grdConsulta(7, grdConsulta.CurrentCell.RowIndex).Value.ToString
        frm_cambio_precios.ShowDialog()
        frm_cambio_precios.Dispose()
    End Sub

    Function extraer_costo2(articulo As Long) As Double
        clase.consultar("SELECT ar_costo2 FROM articulos WHERE (ar_codigo =" & articulo & ")", "costo2")
        Return clase.dt.Tables("costo2").Rows(0)("ar_costo2")
    End Function
End Class