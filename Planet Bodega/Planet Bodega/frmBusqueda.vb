Public Class frmBusqueda
    Dim clase As New class_library
    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        frmIngresoNoImportados.txtArticulo.Text = grdConsulta(0, grdConsulta.CurrentCell.RowIndex).Value.ToString
        Me.Close()
    End Sub

    Private Sub txtConsulta_TextChanged(sender As Object, e As EventArgs) Handles txtConsulta.TextChanged
        If chkReferencia.Checked = False And chkDescripcion.Checked = False Then
            MessageBox.Show("DEBE SELECCIONAR UN CRITERIO DE BUSQUEDA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If chkReferencia.Checked = True Then
            clase.consultar("SELECT articulos.ar_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, linea1.ln1_nombre AS LINEA, ar_foto AS FOTO FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) WHERE articulos.ar_referencia LIKE '%" & txtConsulta.Text & "%';", "consulta")
        End If
        If chkDescripcion.Checked = True Then
            clase.consultar("SELECT articulos.ar_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, linea1.ln1_nombre AS LINEA, ar_foto AS FOTO FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) WHERE articulos.ar_descripcion LIKE '%" & txtConsulta.Text & "%';", "consulta")
        End If
        With grdConsulta
            .DataSource = clase.dt.Tables("consulta")
            .RowHeadersWidth = 4
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Width = 50
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).Width = 150
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(2).Width = 150
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).Width = 150
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
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
        clase.consultar("SELECT articulos.ar_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, linea1.ln1_nombre AS LINEA, ar_foto AS FOTO FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) WHERE articulos.ar_descripcion LIKE '%" & txtConsulta.Text & "%';", "consulta")
        If clase.dt.Tables("consulta").Rows.Count > 0 Then
            With grdConsulta
                .DataSource = clase.dt.Tables("consulta")
                .RowHeadersWidth = 4
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(0).Width = 50
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(1).Width = 150
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(2).Width = 150
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(3).Width = 150
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(4).Visible = False
            End With
        Else
            With grdConsulta
                .DataSource = Nothing
                .RowHeadersWidth = 4
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(0).Width = 50
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(1).Width = 150
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(2).Width = 150
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(3).Width = 150
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(4).Visible = False
            End With
        End If
    End Sub
End Class