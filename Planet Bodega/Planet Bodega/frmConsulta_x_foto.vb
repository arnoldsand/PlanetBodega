Public Class frmConsulta_x_foto
    Dim clase As New class_library
    Dim Ini As Boolean = True
    Dim Sl1 As Boolean = True
    Dim Sl2 As Boolean = True
    Dim Sl3 As Boolean = True
    Dim Sl4 As Boolean = True
    Dim ConsImport, ConsL, ConsSub1, ConsSub2, ConsSub3, ConsSub4 As Boolean
    Dim VariasPaginas As Boolean = False
    Dim Page As Integer
    Dim Importacion, Linea, Sublinea1, Sublinea2, Sublinea3, Sublinea4 As String

    Dim Column0 As New System.Windows.Forms.DataGridViewImageColumn
    Dim Column1 As New System.Windows.Forms.DataGridViewTextBoxColumn

    Dim Column2 As New System.Windows.Forms.DataGridViewImageColumn
    Dim Column3 As New System.Windows.Forms.DataGridViewTextBoxColumn

    Dim Column4 As New System.Windows.Forms.DataGridViewImageColumn
    Dim Column5 As New System.Windows.Forms.DataGridViewTextBoxColumn

    Dim Column6 As New System.Windows.Forms.DataGridViewImageColumn
    Dim Column7 As New System.Windows.Forms.DataGridViewTextBoxColumn

    Dim Column8 As New System.Windows.Forms.DataGridViewImageColumn
    Dim Column9 As New System.Windows.Forms.DataGridViewTextBoxColumn

    Private Sub frmConsulta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(cbLinea, "SELECT * FROM linea1", "ln1_nombre", "ln1_codigo")
        clase.llenar_combo(cbImportacion, "SELECT * FROM cabimportacion", "imp_nombrefecha", "imp_codigo")
        PrepararColumnas()
        ConsImport = True
        ConsL = False
        ConsSub1 = False
        ConsSub2 = False
        ConsSub3 = False
        ConsSub4 = False
        Ini = False
    End Sub

    Public Sub LlenarGrid()
        grdConsulta.Rows.Clear()

        If cbImportacion.SelectedValue = Nothing Or cbImportacion.Text = "" Then
            LlenarGrid2()
            Exit Sub
        End If

        If ConsImport = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (cabimportacion.imp_codigo ='" & cbImportacion.SelectedValue.ToString & "')  ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsL = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (cabimportacion.imp_codigo ='" & cbImportacion.SelectedValue.ToString & "' AND articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' ) ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub1 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (cabimportacion.imp_codigo ='" & cbImportacion.SelectedValue.ToString & "' AND articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "') ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub2 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (cabimportacion.imp_codigo ='" & cbImportacion.SelectedValue.ToString & "' AND articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "' AND articulos.ar_sublinea2='" & cbSublinea2.SelectedValue.ToString & "' ) ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub3 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (cabimportacion.imp_codigo ='" & cbImportacion.SelectedValue.ToString & "' AND articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "' AND articulos.ar_sublinea2='" & cbSublinea2.SelectedValue.ToString & "' AND articulos.ar_sublinea3='" & cbSublinea3.SelectedValue.ToString & "' ) ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub4 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (cabimportacion.imp_codigo ='" & cbImportacion.SelectedValue.ToString & "' AND articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "' AND articulos.ar_sublinea2='" & cbSublinea2.SelectedValue.ToString & "' AND articulos.ar_sublinea3='" & cbSublinea3.SelectedValue.ToString & "' AND articulos.ar_sublinea4='" & cbSublinea4.SelectedValue.ToString & "' ) ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If clase.dt.Tables("consulta").Rows.Count = 0 Then
            MessageBox.Show("No se encontraron resultados para este criterio de busqueda. Gracias.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            grdConsulta.DataSource = Nothing
            Exit Sub
        End If

        With grdConsulta
            .RowHeadersVisible = False
            .ColumnHeadersVisible = False
            .Columns(1).Visible = False
            .Columns(3).Visible = False
            .Columns(5).Visible = False
            .Columns(7).Visible = False
            .Columns(9).Visible = False
        End With
        Dim j As Integer

        Dim Conteo As Integer = clase.dt.Tables("consulta").Rows.Count - 1
        If Conteo > 20 Then
            VariasPaginas = True
            btnSiguiente.Enabled = True
        Else
            btnSiguiente.Enabled = False
            ConsL = True
            Page = 0
        End If

        For i = 0 To clase.dt.Tables("consulta").Rows.Count - 1
            j = grdConsulta.Rows.Add
            Try
                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i)("ar_foto")) Then
                    grdConsulta(0, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i)("ar_foto").ToString)
                Else
                    grdConsulta.Item(0, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(1, j).Value = clase.dt.Tables("consulta").Rows(i)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 1)("ar_foto")) Then
                    grdConsulta(2, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 1)("ar_foto").ToString)
                Else
                    grdConsulta.Item(2, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(3, j).Value = clase.dt.Tables("consulta").Rows(i + 1)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 2)("ar_foto")) Then
                    grdConsulta(4, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 2)("ar_foto").ToString)
                Else
                    grdConsulta.Item(4, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(5, j).Value = clase.dt.Tables("consulta").Rows(i + 2)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 3)("ar_foto")) Then
                    grdConsulta(6, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 3)("ar_foto").ToString)
                Else
                    grdConsulta.Item(6, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(7, j).Value = clase.dt.Tables("consulta").Rows(i + 3)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 4)("ar_foto")) Then
                    grdConsulta(8, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 4)("ar_foto").ToString)
                Else
                    grdConsulta.Item(8, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(9, j).Value = clase.dt.Tables("consulta").Rows(i + 4)("ar_codigo").ToString

            Catch ex As Exception

            End Try
            i = i + 4
        Next
        For i = 0 To grdConsulta.RowCount - 1
            Try
                SetImage1(grdConsulta(0, i))
                SetImage1(grdConsulta(2, i))
                SetImage1(grdConsulta(4, i))
                SetImage1(grdConsulta(6, i))
                SetImage1(grdConsulta(8, i))
            Catch ex As Exception

            End Try
        Next
        'ConsL = True
        'ConsSub1 = False
        'ConsSub2 = False
        'ConsSub3 = False
        'ConsSub4 = False
    End Sub

    Public Sub LlenarGrid2()
        grdConsulta.Rows.Clear()

        If ConsImport = True Then
            ConsL = True
        End If

        If ConsL = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "')  ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub1 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "')  ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub2 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "' AND articulos.ar_sublinea2='" & cbSublinea2.SelectedValue.ToString & "') ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub3 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "' AND articulos.ar_sublinea2='" & cbSublinea2.SelectedValue.ToString & "' AND articulos.ar_sublinea3='" & cbSublinea3.SelectedValue.ToString & "')  ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        If ConsSub4 = True Then
            clase.consultar("SELECT articulos.ar_foto , articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE articulos.ar_linea='" & cbLinea.SelectedValue.ToString & "' AND articulos.ar_sublinea1='" & cbSublinea1.SelectedValue.ToString & "' AND articulos.ar_sublinea2='" & cbSublinea2.SelectedValue.ToString & "' AND articulos.ar_sublinea3='" & cbSublinea3.SelectedValue.ToString & "' AND articulos.ar_sublinea4='" & cbSublinea4.SelectedValue.ToString & "')   ORDER BY ar_codigo DESC LIMIT " & Page & ",25;", "consulta")
        End If

        Try
            If clase.dt.Tables("consulta").Rows.Count = 0 Then
                MessageBox.Show("No se encontraron resultados para este criterio de busqueda. Gracias.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                grdConsulta.DataSource = Nothing
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Debe elegir un criterio de busqueda.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        
        With grdConsulta
            .RowHeadersVisible = False
            .ColumnHeadersVisible = False
            .Columns(1).Visible = False
            .Columns(3).Visible = False
            .Columns(5).Visible = False
            .Columns(7).Visible = False
            .Columns(9).Visible = False
        End With
        Dim j As Integer

        Dim Conteo As Integer = clase.dt.Tables("consulta").Rows.Count - 1
        If Conteo > 20 Then
            VariasPaginas = True
            btnSiguiente.Enabled = True
        Else
            btnSiguiente.Enabled = False
            ConsL = True
            Page = 0
        End If

        For i = 0 To clase.dt.Tables("consulta").Rows.Count - 1
            j = grdConsulta.Rows.Add

            Try
                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i)("ar_foto")) Then
                    grdConsulta(0, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i)("ar_foto").ToString)
                Else
                    grdConsulta.Item(0, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(1, j).Value = clase.dt.Tables("consulta").Rows(i)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 1)("ar_foto")) Then
                    grdConsulta(2, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 1)("ar_foto").ToString)
                Else
                    grdConsulta.Item(2, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(3, j).Value = clase.dt.Tables("consulta").Rows(i + 1)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 2)("ar_foto")) Then
                    grdConsulta(4, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 2)("ar_foto").ToString)
                Else
                    grdConsulta.Item(4, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(5, j).Value = clase.dt.Tables("consulta").Rows(i + 2)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 3)("ar_foto")) Then
                    grdConsulta(6, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 3)("ar_foto").ToString)
                Else
                    grdConsulta.Item(6, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(7, j).Value = clase.dt.Tables("consulta").Rows(i + 3)("ar_codigo").ToString

                If System.IO.File.Exists(clase.dt.Tables("consulta").Rows(i + 4)("ar_foto")) Then
                    grdConsulta(8, j).Value = System.Drawing.Image.FromFile(clase.dt.Tables("consulta").Rows(i + 4)("ar_foto").ToString)
                Else
                    grdConsulta.Item(8, j).Value = Global.WindowsApplication1.My.Resources.sinfoto
                End If
                grdConsulta(9, j).Value = clase.dt.Tables("consulta").Rows(i + 4)("ar_codigo").ToString

            Catch ex As Exception

            End Try
            i = i + 4
        Next

        For i = 0 To grdConsulta.RowCount - 1
            Try
                SetImage1(grdConsulta(0, i))
                SetImage1(grdConsulta(2, i))
                SetImage1(grdConsulta(4, i))
                SetImage1(grdConsulta(6, i))
                SetImage1(grdConsulta(8, i))
            Catch ex As Exception

            End Try
        Next
    End Sub

    Public Sub PrepararColumnas()
        grdConsulta.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column0, Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column6, Me.Column7, Me.Column8, Me.Column9})
    End Sub

    Private Sub grdConsulta_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdConsulta.CellClick
        Dim x As Integer = grdConsulta.CurrentCell.RowIndex
        Dim y As Integer = grdConsulta.CurrentCell.ColumnIndex
        Try
            frmVer.Codigo = grdConsulta(y + 1, x).Value.ToString

            frmVer.ShowDialog()
            frmVer.Dispose()
        Catch ex As Exception
            MessageBox.Show("ESTAS CELDAS NO CORRESPONDEN A NINGUN PRODUCTO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

    End Sub

    Private Sub cbLinea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbLinea.SelectedIndexChanged
        If Ini = False Then
            clase.llenar_combo(cbSublinea1, "SELECT * FROM sublinea1 WHERE sl1_ln1codigo='" & cbLinea.SelectedValue.ToString & "'", "sl1_nombre", "sl1_codigo")
            Sl1 = False
            ConsImport = False
            ConsSub1 = False
            ConsSub2 = False
            ConsSub3 = False
            ConsSub4 = False
            ConsL = True
            Page = 0
        End If
        grdConsulta.Focus()
    End Sub

    Private Sub cbSublinea1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSublinea1.SelectedIndexChanged
        If Sl1 = False Then
            Try
                clase.llenar_combo(cbSublinea2, "SELECT * FROM sublinea2 WHERE sl2_sl1codigo='" & cbSublinea1.SelectedValue.ToString & "'", "sl2_nombre", "sl2_codigo")
                Sl2 = False
                ConsSub1 = True
                ConsSub2 = False
                ConsSub3 = False
                ConsSub4 = False
                ConsL = False
                Page = 0
            Catch ex As Exception

            End Try
        End If
        grdConsulta.Focus()
    End Sub

    Private Sub cbSublinea2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSublinea2.SelectedIndexChanged
        If Sl2 = False Then
            Try
                clase.llenar_combo(cbSublinea3, "SELECT * FROM sublinea3 WHERE sl3_sl2codigo='" & cbSublinea2.SelectedValue.ToString & "'", "sl3_nombre", "sl3_codigo")
                Sl3 = False
                ConsSub2 = True
                ConsSub1 = False
                ConsSub3 = False
                ConsSub4 = False
                ConsL = False
                Page = 0
            Catch ex As Exception

            End Try
        End If
        grdConsulta.Focus()
    End Sub

    Private Sub cbSublinea3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSublinea3.SelectedIndexChanged
        If Sl3 = False Then
            Try
                clase.llenar_combo(cbSublinea4, "SELECT * FROM sublinea4 WHERE sl4_sl3codigo='" & cbSublinea3.SelectedValue.ToString & "'", "sl4_nombre", "sl4_codigo")
                Sl4 = False

                ConsSub3 = True
                ConsSub1 = False
                ConsSub2 = False
                ConsSub4 = False
                ConsL = False
                Page = 0
            Catch ex As Exception

            End Try
        End If
        grdConsulta.Focus()
    End Sub

    Private Sub cbImportacion_MouseWheel(sender As Object, e As MouseEventArgs) Handles cbImportacion.MouseWheel
        Dim disable As HandledMouseEventArgs = e
        disable.Handled = True
    End Sub

    Private Sub cbImportacion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbImportacion.SelectedIndexChanged
        If Ini = False Then
            cbLinea.Enabled = True

            ConsSub3 = False
            ConsSub1 = False
            ConsSub2 = False
            ConsSub4 = False
            ConsL = False

            cbSublinea1.Enabled = True
            cbSublinea2.Enabled = True
            cbSublinea3.Enabled = True
            cbSublinea4.Enabled = True
            Page = 0
        End If
        grdConsulta.Focus()
    End Sub

    Private Sub btnDeshacer_Click(sender As Object, e As EventArgs) Handles btnDeshacer.Click
        grdConsulta.Rows.Clear()
        cbImportacion.Text = ""
        cbLinea.Text = ""
        cbSublinea1.Text = ""
        cbSublinea2.Text = ""
        cbSublinea3.Text = ""
        cbSublinea4.Text = ""

        cbLinea.Text = ""
        cbSublinea1.Text = ""
        cbSublinea2.Text = ""
        cbSublinea3.Text = ""
        cbSublinea4.Text = ""

        ConsImport = True
        ConsL = False
        ConsSub1 = False
        ConsSub2 = False
        ConsSub3 = False
        ConsSub4 = False

        Label7.Text = 1
    End Sub

    Private Sub cbSublinea4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSublinea4.SelectedIndexChanged
        ConsSub4 = True
        ConsSub1 = False
        ConsSub2 = False
        ConsSub3 = False
        ConsL = False
        Page = 0
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnSiguiente_Click(sender As Object, e As EventArgs) Handles btnSiguiente.Click
        Page = Page + 25
        If ConsImport = True Then
            ConsSub1 = False
            ConsSub2 = False
            ConsSub3 = False
            ConsSub4 = False
            ConsL = False
        End If
        If ConsL = True Then
            ConsImport = False
            ConsSub1 = False
            ConsSub2 = False
            ConsSub3 = False
            ConsSub4 = False
        End If
        If ConsSub1 = True Then
            ConsImport = False
            ConsSub2 = False
            ConsSub3 = False
            ConsSub4 = False
            ConsL = False
        End If
        If ConsSub2 = True Then
            ConsImport = False
            ConsSub1 = False
            ConsSub3 = False
            ConsSub4 = False
            ConsL = False
        End If
        If ConsSub3 = True Then
            ConsImport = False
            ConsSub1 = False
            ConsSub2 = False
            ConsSub4 = False
            ConsL = False
        End If
        LlenarGrid()
        btnAtras.Enabled = True
        Label7.Text = (Page + 25) / 25
    End Sub

    Private Sub btnConsulta_Click(sender As Object, e As EventArgs) Handles btnConsulta.Click
        'If clase.validar_combobox(cbImportacion, "Importación") = False Then Exit Sub
        LlenarGrid()
        Page = 0
        Label7.Text = 1
    End Sub

    Private Sub btnAtras_Click(sender As Object, e As EventArgs) Handles btnAtras.Click

        Page = Page - 25
        If ConsImport = True Then
            ConsSub1 = False
            ConsSub2 = False
            ConsSub3 = False
            ConsSub4 = False
            ConsL = False
        End If
        If ConsL = True Then
            ConsImport = False
            ConsSub1 = False
            ConsSub2 = False
            ConsSub3 = False
            ConsSub4 = False
        End If
        If ConsSub1 = True Then
            ConsImport = False
            ConsSub2 = False
            ConsSub3 = False
            ConsSub4 = False
            ConsL = False
        End If
        If ConsSub2 = True Then
            ConsImport = False
            ConsSub1 = False
            ConsSub3 = False
            ConsSub4 = False
            ConsL = False
        End If
        If ConsSub3 = True Then
            ConsImport = False
            ConsSub1 = False
            ConsSub2 = False
            ConsSub4 = False
            ConsL = False
        End If
        LlenarGrid()
        Label7.Text = (Page + 25) / 25

        If Page = 0 Then
            btnAtras.Enabled = False
            Label7.Text = (Page + 25) / 25
        End If
    End Sub
End Class

