Public Class frm_stock_articulos
    Dim clase As New class_library
    Dim ind_carga As Boolean
    Dim codigoarticulo As Long

    Private Sub frm_stock_articulos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox9.SelectedIndex = 0
        Label18.Text = "Codigo"
        ind_carga = False
        clase.llenar_combo(ComboBox1, "select* from linea1 order by ln1_nombre asc", "ln1_nombre", "ln1_codigo")
        clase.llenar_combo(ComboBox7, "select* from  colores order by colornombre asc", "colornombre", "cod_color")
        clase.llenar_combo(ComboBox6, "select* from tallas order by nombretalla asc", "nombretalla", "codigo_talla")
        ind_carga = True
        DataGridView2.ColumnCount = 13
        columnas_datagridview2()
        TextBox1.Text = 0
        TextBox2.Text = 0
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        '  DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Private Sub columnas_datagridview2()
        With DataGridView2
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Linea"
            .Columns(4).HeaderText = "Sublinea 1"
            .Columns(5).HeaderText = "Sublinea 2"
            .Columns(6).HeaderText = "Sublinea 3"
            .Columns(7).HeaderText = "Sublinea 4"
            .Columns(8).HeaderText = "Cant"
            .Columns(0).Width = 60

            .Columns(1).Width = 130
            .Columns(2).Width = 180
            .Columns(8).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub


    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        Select Case ComboBox9.SelectedIndex
            Case 0
                Label18.Text = "Codigo"
                TextBox18.Visible = True
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox3.Enabled = False
                ComboBox4.Enabled = False
                ComboBox5.Enabled = False
                ComboBox6.Enabled = False
                ComboBox7.Enabled = False
                ind_carga = False
                ComboBox1.SelectedIndex = -1
                ComboBox2.DataSource = Nothing
                ComboBox3.DataSource = Nothing
                ComboBox4.DataSource = Nothing
                ComboBox5.DataSource = Nothing
                ind_carga = True
                TextBox18.Text = ""
                ComboBox6.SelectedIndex = -1
                ComboBox7.SelectedIndex = -1
                ComboBox8.Enabled = False
                ComboBox8.SelectedIndex = -1
            Case 1
                Label18.Text = "Referencia"
                TextBox18.Visible = True
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox3.Enabled = False
                ComboBox4.Enabled = False
                ComboBox5.Enabled = False
                ComboBox6.Enabled = False
                ComboBox7.Enabled = False
                ind_carga = False
                ComboBox1.SelectedIndex = -1
                ComboBox2.DataSource = Nothing
                ComboBox3.DataSource = Nothing
                ComboBox4.DataSource = Nothing
                ComboBox5.DataSource = Nothing
                ind_carga = True
                TextBox18.Text = ""
                ComboBox6.SelectedIndex = -1
                ComboBox7.SelectedIndex = -1
                ComboBox8.Enabled = False
                ComboBox8.SelectedIndex = -1
            Case 2
                Label18.Text = ""
                TextBox18.Visible = False
                ComboBox1.Enabled = True
                ComboBox2.Enabled = False
                ComboBox3.Enabled = False
                ComboBox4.Enabled = False
                ComboBox5.Enabled = False
                ComboBox6.Enabled = True
                ComboBox7.Enabled = True
                ind_carga = False
                ComboBox1.SelectedIndex = -1
                ComboBox2.DataSource = Nothing
                ComboBox3.DataSource = Nothing
                ComboBox4.DataSource = Nothing
                ComboBox5.DataSource = Nothing
                ind_carga = True
                TextBox18.Text = ""
                ComboBox6.SelectedIndex = -1
                ComboBox7.SelectedIndex = -1
                ComboBox8.Enabled = True
                ComboBox8.SelectedIndex = 0
        End Select
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

   

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox2.Enabled = True
            clase.llenar_combo(ComboBox2, "SELECT* FROM sublinea1 WHERE sl1_ln1codigo = " & ComboBox1.SelectedValue & " ORDER BY sl1_nombre asc", "sl1_nombre", "sl1_codigo")
            ComboBox3.DataSource = Nothing
            ComboBox3.Enabled = False
            ComboBox4.DataSource = Nothing
            ComboBox4.Enabled = False
            ComboBox5.DataSource = Nothing
            ComboBox5.Enabled = False
            ind_carga = True
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox3.Enabled = True
            clase.llenar_combo(ComboBox3, "SELECT* FROM sublinea2 WHERE sl2_sl1codigo = " & ComboBox2.SelectedValue & "  ORDER BY sl2_nombre asc", "sl2_nombre", "sl2_codigo")
            ComboBox4.DataSource = Nothing
            ComboBox4.Enabled = False
            ComboBox5.DataSource = Nothing
            ComboBox5.Enabled = False
            ind_carga = True
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox4.Enabled = True
            clase.llenar_combo(ComboBox4, "SELECT* FROM sublinea3 WHERE sl3_sl2codigo = " & ComboBox3.SelectedValue & "  ORDER BY sl3_nombre asc", "sl3_nombre", "sl3_codigo")
            ComboBox5.DataSource = Nothing
            ComboBox5.Enabled = False
            ind_carga = True
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ind_carga = True Then
            ind_carga = False
            ComboBox5.Enabled = True
            clase.llenar_combo(ComboBox5, "SELECT* FROM sublinea4 WHERE sl4_sl3codigo = " & ComboBox4.SelectedValue & "  ORDER BY sl4_nombre asc", "sl4_nombre", "sl4_codigo")
            ind_carga = True
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        codigoarticulo = vbEmpty
        Dim sql1 As String = ""
        '   encabezado_movimientos()
        Select Case ComboBox9.SelectedIndex
            Case 0
                If clase.validar_cajas_text(TextBox18, "Codigo de Articulo") = False Then Exit Sub
                If Len(TextBox18.Text) = 13 Then
                    sql1 = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, SUM(inventario_bodega.inv_cantidad), articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN inventario_bodega ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) WHERE (articulos.ar_codigobarras ='" & TextBox18.Text & "') GROUP BY inventario_bodega.inv_codigoart"
                    codigoarticulo = convertir_codigobarra_a_codigo_normal(TextBox18.Text)
                End If
                If Len(TextBox18.Text) < 13 Then
                    sql1 = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, SUM(inventario_bodega.inv_cantidad), articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN inventario_bodega ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) WHERE (articulos.ar_codigo =" & TextBox18.Text & ") GROUP BY inventario_bodega.inv_codigoart"
                    codigoarticulo = TextBox18.Text
                End If
                If Len(TextBox18.Text) > 13 Then
                    MessageBox.Show("El codigo digitado es invalido.", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                clase.consultar(sql1, "tbl")
                If clase.dt.Tables("tbl").Rows.Count > 0 Then
                    DataGridView2.Columns.Clear()
                    DataGridView2.DataSource = clase.dt.Tables("tbl")
                    columnas_datagridview2()
                    dtmovimientos.DataSource = Nothing
                    DataGridView3.RowCount = 0
                Else
                    MessageBox.Show("El codigo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    DataGridView2.DataSource = Nothing
                    DataGridView2.ColumnCount = 13
                    dtmovimientos.DataSource = Nothing
                    columnas_datagridview2()
                    Button9_Click(Nothing, Nothing)
                End If
            Case 1
                clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, SUM(inventario_bodega.inv_cantidad), articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN inventario_bodega ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) WHERE (articulos.ar_referencia LIKE '" & UCase(TextBox18.Text) & "%') GROUP BY articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre ORDER BY articulos.ar_referencia ASC", "tbl")
                If clase.dt.Tables("tbl").Rows.Count > 0 Then
                    DataGridView2.Columns.Clear()
                    DataGridView2.DataSource = clase.dt.Tables("tbl")
                    columnas_datagridview2()
                    dtmovimientos.DataSource = Nothing
                    DataGridView3.RowCount = 0
                Else
                    MessageBox.Show("La referencia digitada no fue encontrada.", "REFERENCIA NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    DataGridView2.DataSource = Nothing
                    DataGridView2.ColumnCount = 13
                    dtmovimientos.DataSource = Nothing
                    columnas_datagridview2()
                    Button9_Click(Nothing, Nothing)
                End If
            Case 2
                If (ComboBox1.Text.Trim = "") And (ComboBox7.Text.Trim = "") And (ComboBox6.Text.Trim = "") Then
                    MessageBox.Show("Debe especificar un criterio de busqueda.", "CRITERIO DE BUSQUEDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                Dim consql As String = ""
                Select Case ComboBox1.Text
                    Case Is = ""
                        If (ComboBox7.Text.Trim <> "") And (ComboBox6.Text.Trim = "") Then 'BUSCAR POR COLOR
                            Select Case ComboBox8.SelectedIndex
                                Case 0
                                    consql = " AND (inventario_bodega.inv_cantidad >0) "
                                Case 1
                                    consql = " AND (inventario_bodega.inv_cantidad =0) "
                                Case 2
                                    consql = ""
                            End Select
                            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, inventario_bodega.inv_cantidad, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1  ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2  ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN inventario_bodega ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) WHERE (articulos.ar_color =" & ComboBox7.SelectedValue & ") " & consql & "ORDER BY articulos.ar_referencia ASC", "resultados")
                            '   ComboBox6.Text = "SIN TALLA"
                        End If
                        If (ComboBox7.Text.Trim = "") And (ComboBox6.Text.Trim <> "") Then 'BUSCAR POR TALLA
                            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, inventario_bodega.inv_cantidad, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1  ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2  ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN inventario_bodega ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) WHERE (articulos.ar_talla =" & ComboBox6.SelectedValue & ") ORDER BY articulos.ar_referencia ASC", "resultados")
                            '  ComboBox7.Text = "SIN COLOR"
                        End If
                        If clase.dt.Tables("resultados").Rows.Count > 0 Then
                            DataGridView2.Columns.Clear()
                            DataGridView2.DataSource = clase.dt.Tables("resultados")
                            columnas_datagridview2()
                            dtmovimientos.DataSource = Nothing
                            DataGridView3.RowCount = 0
                        Else
                            MessageBox.Show("No se encontraron articulos en la linea de producto especificada.", "NO SE ENCONTRARON ARTICULOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            DataGridView2.DataSource = Nothing
                            DataGridView2.ColumnCount = 13
                            dtmovimientos.DataSource = Nothing
                            columnas_datagridview2()
                            Button9_Click(Nothing, Nothing)
                        End If

                    Case Is <> ""
                        Dim sub1, sub2, sub3, sub4, talla, color As String
                        If ComboBox2.Text <> "" Then
                            sub1 = " AND articulos.ar_sublinea1 = " & ComboBox2.SelectedValue
                        Else
                            sub1 = ""
                        End If
                        If ComboBox3.Text <> "" Then
                            sub2 = " AND articulos.ar_sublinea2 = " & ComboBox3.SelectedValue
                        Else
                            sub2 = ""
                        End If
                        If ComboBox4.Text <> "" Then
                            sub3 = " AND articulos.ar_sublinea3 = " & ComboBox4.SelectedValue
                        Else
                            sub3 = ""
                        End If
                        If ComboBox5.Text <> "" Then
                            sub4 = " AND articulos.ar_sublinea4 = " & ComboBox5.SelectedValue
                        Else
                            sub4 = ""
                        End If
                        If (ComboBox7.Text.Trim <> "") Then
                            color = " AND ar_color = " & ComboBox7.SelectedValue
                        Else
                            color = ""
                        End If
                        If (ComboBox6.Text.Trim <> "") Then
                            talla = " AND ar_talla = " & ComboBox6.SelectedValue
                        Else
                            talla = ""
                        End If
                        Select Case ComboBox8.SelectedIndex
                            Case 0
                                consql = " AND (inventario_bodega.inv_cantidad >0) "
                            Case 1
                                consql = " AND (inventario_bodega.inv_cantidad =0) "
                            Case 2
                                consql = ""
                        End Select
                        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, SUM(inventario_bodega.inv_cantidad), articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2  ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN inventario_bodega ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) WHERE (articulos.ar_linea = " & ComboBox1.SelectedValue & sub1 & sub2 & sub3 & sub4 & talla & color & ") " & consql & "ORDER BY articulos.ar_referencia ASC", "datos")
                        If clase.dt.Tables("datos").Rows.Count > 0 Then
                            DataGridView2.Columns.Clear()
                            dtmovimientos.DataSource = Nothing
                            DataGridView3.RowCount = 0
                            DataGridView2.DataSource = clase.dt.Tables("datos")
                            columnas_datagridview2()
                        Else
                            MessageBox.Show("No se encontraron articulos en la linea de producto especificada.", "NO SE ENCONTRARON ARTICULOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            DataGridView2.DataSource = Nothing
                            DataGridView2.ColumnCount = 13
                            columnas_datagridview2()
                            Button9_Click(Nothing, Nothing)
                            dtmovimientos.DataSource = Nothing
                        End If
                End Select
        End Select
        DataGridView1.RowCount = 0
        TabControl1.SelectTab(0)
        ajustar_cantidades_y_establecer_totales()
    End Sub

    Private Sub TextBox18_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox18.KeyPress
        If ComboBox9.SelectedIndex = 0 Then
            clase.validar_numeros(e)
        End If
    End Sub

    Private Sub ajustar_cantidades_y_establecer_totales()
        Dim x As Integer
        Dim cantidad As Integer = 0
        For x = 0 To DataGridView2.RowCount - 1
            If IsDBNull(DataGridView2.Item(8, x).Value) Then
                DataGridView2.Item(8, x).Value = 0
            Else
                cantidad = cantidad + DataGridView2.Item(8, x).Value
            End If
        Next
        DataGridView2.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        TextBox1.Text = cantidad
        TextBox2.Text = DataGridView2.RowCount
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        With DataGridView2
            '  clase.consultar1(buscar_codigos_a_apartir_de_la_referencia(.Item(0, .CurrentCell.RowIndex).Value, .Item(1, .CurrentCell.RowIndex).Value, .Item(8, .CurrentCell.RowIndex).Value, .Item(9, .CurrentCell.RowIndex).Value, .Item(10, .CurrentCell.RowIndex).Value, .Item(11, .CurrentCell.RowIndex).Value, .Item(12, .CurrentCell.RowIndex).Value), "data")
            clase.consultar("select ar_foto from articulos where ar_codigo = " & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value & "", "tablita")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        End With
        Dim sql As String
        If ComboBox9.SelectedIndex = 0 Then
            sql = "SELECT bodegas.bod_nombre, articulos_gondolas.gondola FROM bodegas INNER JOIN articulos_gondolas ON (bodegas.bod_codigo = articulos_gondolas.bodega) WHERE (articulos_gondolas.articulo =" & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value & ") ORDER BY articulos_gondolas.bodega, articulos_gondolas.gondola ASC"
        Else
            ' Dim x As Short
            Dim ind As Boolean = False
            sql = "SELECT bodegas.bod_nombre, articulos_gondolas.gondola FROM bodegas INNER JOIN articulos_gondolas ON (bodegas.bod_codigo = articulos_gondolas.bodega) WHERE (articulos_gondolas.articulo = " & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value & ")"
            'For x = 0 To clase.dt1.Tables("data").Rows.Count - 1
            '    If ind = False Then
            '        sql = sql & " (articulos_gondolas.articulo = " & clase.dt1.Tables("data").Rows(x)("ar_codigo") & ")"
            '        ind = True
            '    Else
            '        sql = sql & " OR (articulos_gondolas.articulo = " & clase.dt1.Tables("data").Rows(x)("ar_codigo") & ")"
            '    End If
            'Next
            'sql = sql & ") GROUP BY articulos_gondolas.bodega, articulos_gondolas.gondola ORDER BY articulos_gondolas.bodega, articulos_gondolas.gondola ASC"
        End If
        clase.consultar(sql, "datos1")
        If clase.dt.Tables("datos1").Rows.Count > 0 Then
            DataGridView1.RowCount = 0
            For x = 0 To clase.dt.Tables("datos1").Rows.Count - 1
                With DataGridView1
                    .RowCount = .RowCount + 1
                    .Item(0, x).Value = clase.dt.Tables("datos1").Rows(x)("bod_nombre")
                    .Item(1, x).Value = clase.dt.Tables("datos1").Rows(x)("gondola")
                End With

            Next
        Else
            DataGridView1.RowCount = 0
        End If
        clase.consultar("select* from inventario_bodega where inv_codigoart = " & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value, "listado")
        If clase.dt.Tables("listado").Rows.Count > 0 Then
            With DataGridView3
                .RowCount = 9
                .ColumnCount = 2
                .Item(0, 0).Value = "Inventario"
                .Item(0, 1).Value = "ENTRADAS"
                .Item(0, 2).Value = "Ajustes (+)"
                .Item(0, 3).Value = "Ingresados"
                .Item(0, 4).Value = "SALIDAS"
                .Item(0, 5).Value = "Transferencia"
                .Item(0, 6).Value = "Ajustes (-)"
                .Item(0, 7).Value = "Existencia Teorica"
                .Item(0, 8).Value = "Existencia Real"
                .Item(1, 0).Value = comprobar_nulidad_de_integer(clase.dt.Tables("listado").Rows(0)("inv_saldo_inicial"))
                .Item(1, 2).Value = comprobar_nulidad_de_integer(clase.dt.Tables("listado").Rows(0)("inv_ajustado_pos"))
                .Item(1, 3).Value = comprobar_nulidad_de_integer(clase.dt.Tables("listado").Rows(0)("inv_ingresado"))
                .Item(1, 5).Value = comprobar_nulidad_de_integer(clase.dt.Tables("listado").Rows(0)("inv_transferido"))
                .Item(1, 6).Value = comprobar_nulidad_de_integer(clase.dt.Tables("listado").Rows(0)("inv_ajustado_neg"))
                .Item(1, 7).Value = (.Item(1, 0).Value + (.Item(1, 2).Value + .Item(1, 3).Value)) - (.Item(1, 5).Value + .Item(1, 6).Value)
                .Item(1, 8).Value = comprobar_nulidad_de_integer(clase.dt.Tables("listado").Rows(0)("inv_cantidad"))
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
        Else
            dtmovimientos.Columns.Clear()
        End If
    End Sub

    Function fecha_ultimo_inventario() As Date
        clase.consultar("SELECT cabinventario.cabinv_fecha FROM informacion INNER JOIN cabinventario ON (informacion.ultimo_inventario = cabinventario.cabinv_codigo)", "fecha")
        Return clase.dt.Tables("fecha").Rows(0)("cabinv_fecha")
    End Function

    Function hora_ultimo_inventario() As Date
        clase.consultar("SELECT cabinventario.cabinv_hora FROM informacion INNER JOIN cabinventario ON (informacion.ultimo_inventario = cabinventario.cabinv_codigo)", "fecha")
        Return clase.dt.Tables("fecha").Rows(0)("cabinv_hora")
    End Function


    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        DataGridView2.DataSource = Nothing
        DataGridView2.ColumnCount = 13
        columnas_datagridview2()
        ComboBox9.SelectedIndex = 0
        PictureBox1.Image = Nothing
        TextBox18.Focus()
        TextBox1.Text = 0
        TextBox2.Text = 0
        codigoarticulo = vbEmpty
        DataGridView1.RowCount = 0
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim sql1 As String = ""
        Dim titulo As String = ""
        Dim columna As Short
        Select Case DataGridView3.CurrentRow.Index
            Case 2
                sql1 = "SELECT cabajuste.cabaj_codigo, cabajuste.cabaj_fecha, cabajuste.cabaj_hora, cabajuste.cabaj_operario, SUM(detajuste.detaj_cantidad) AS cant FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) WHERE (cabajuste.cabaj_procesado = TRUE AND cabajuste.cabaj_fecha BETWEEN '" & fecha_ultimo_inventario.ToString("yyyy-MM-dd") & "' AND '" & Now.ToString("yyyy-MM-dd") & "' AND detajuste.detaj_articulo =" & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value & " AND (" & codigos_ajustes_positivos() & ")) GROUP BY detajuste.detaj_articulo, detajuste.detaj_codigo_ajuste"
                titulo = "AJUSTES POSITIVOS"
                columna = 4
            Case 3
                sql1 = "SELECT cabentrada_bodega.cabent_codigo, cabentrada_bodega.cabent_fecha, cabentrada_bodega.cabent_hora, cabentrada_bodega.cabent_operario, SUM(detalle_entrada.detent_cantidad) AS cant FROM detalle_entrada INNER JOIN cabentrada_bodega ON (detalle_entrada.detent_cod_entrada = cabentrada_bodega.cabent_codigo) WHERE (cabentrada_bodega.cabent_estado = TRUE AND cabentrada_bodega.cabent_fecha BETWEEN '" & fecha_ultimo_inventario.ToString("yyyy-MM-dd") & "' AND '" & Now.ToString("yyyy-MM-dd") & "' AND detalle_entrada.detent_articulo =" & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value & ") GROUP BY detalle_entrada.detent_articulo, detalle_entrada.detent_cod_entrada"
                titulo = "ENTRADAS A BODEGA"
                columna = 4
            Case 5
                sql1 = "SELECT cabtransferencia.tr_numero, cabtransferencia.tr_fecha, cabtransferencia.tr_hora, tiendas.tienda, cabtransferencia.tr_operador, dettransferencia.dt_cantidad AS cant FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) WHERE (cabtransferencia.tr_fecha BETWEEN '" & fecha_ultimo_inventario.ToString("yyyy-MM-dd") & "' AND '" & Now.ToString("yyyy-MM-dd") & "' AND dettransferencia.dt_codarticulo =" & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value & " AND cabtransferencia.tr_bodega IS NOT NULL AND cabtransferencia.tr_estado = TRUE AND cabtransferencia.tr_finalizada = TRUE)"
                titulo = "TRANSFERENCIAS"
                columna = 5
            Case 6
                sql1 = "SELECT cabajuste.cabaj_codigo, cabajuste.cabaj_fecha, cabajuste.cabaj_hora, cabajuste.cabaj_operario, SUM(detajuste.detaj_cantidad) AS cant FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) WHERE (cabajuste.cabaj_procesado = TRUE AND cabajuste.cabaj_fecha BETWEEN '" & fecha_ultimo_inventario.ToString("yyyy-MM-dd") & "' AND '" & Now.ToString("yyyy-MM-dd") & "' AND detajuste.detaj_articulo =" & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value & " AND (" & codigos_ajustes_negativos() & ")) GROUP BY detajuste.detaj_articulo, detajuste.detaj_codigo_ajuste"
                titulo = "AJUSTES NEGATIVOS"
                columna = 4
        End Select
        If sql1 = "" Then
            dtmovimientos.DataSource = Nothing
            TextBox3.Text = "0"
            Label6.Text = ""
            Exit Sub
        End If
        Label6.Text = titulo
        clase.consultar2(sql1, "result")
        If clase.dt2.Tables("result").Rows.Count > 0 Then
            dtmovimientos.DataSource = clase.dt2.Tables("result")
            preparar_columnas_ajustes_pos()
            Dim x As Short
            Dim cantidad As Short = 0
            For x = 0 To dtmovimientos.RowCount - 1
                cantidad = cantidad + dtmovimientos.Item(columna, x).Value
            Next
            TextBox3.Text = System.Math.Abs(cantidad)
        Else
            dtmovimientos.DataSource = Nothing
            TextBox3.Text = "0"
        End If
    End Sub

    Function codigos_ajustes_negativos() As String
        clase.consultar2("select* from tipos_ajustes where tip_tipo = 'R'", "neg")
        If clase.dt2.Tables("neg").Rows.Count > 0 Then
            Dim a As Short
            Dim ind As Boolean = False
            Dim cadena As String = ""
            For a = 0 To clase.dt2.Tables("neg").Rows.Count - 1
                If ind = False Then
                    cadena = "cabajuste.cabaj_tipo_ajuste = " & clase.dt2.Tables("neg").Rows(a)("tip_codigo")
                    ind = True
                Else
                    cadena = cadena & " OR cabajuste.cabaj_tipo_ajuste = " & clase.dt2.Tables("neg").Rows(a)("tip_codigo")
                End If
            Next
            Return cadena
        Else
            Return ""
        End If
    End Function

    Function codigos_ajustes_positivos() As String
        clase.consultar2("select* from tipos_ajustes where tip_tipo = 'S'", "pos")
        If clase.dt2.Tables("pos").Rows.Count > 0 Then
            Dim a As Short
            Dim ind As Boolean = False
            Dim cadena As String = ""
            For a = 0 To clase.dt2.Tables("pos").Rows.Count - 1
                If ind = False Then
                    cadena = "cabajuste.cabaj_tipo_ajuste = " & clase.dt2.Tables("pos").Rows(a)("tip_codigo")
                    ind = True
                Else
                    cadena = cadena & " OR cabajuste.cabaj_tipo_ajuste = " & clase.dt2.Tables("pos").Rows(a)("tip_codigo")
                End If
            Next
            Return cadena
        Else
            Return ""
        End If
    End Function

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub preparar_columnas_ajustes_pos()
        Select Case DataGridView3.CurrentRow.Index
            Case 2
                With dtmovimientos
                    .Columns(0).HeaderText = "Codigo"
                    .Columns(1).HeaderText = "Fecha"
                    .Columns(2).HeaderText = "Hora"
                    .Columns(3).HeaderText = "Operario"
                    .Columns(4).HeaderText = "Cant"
                    .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(4).Width = 50
                    .Columns(3).Width = 150
                End With
            Case 3
                With dtmovimientos
                    .Columns(0).HeaderText = "Codigo"
                    .Columns(1).HeaderText = "Fecha"
                    .Columns(2).HeaderText = "Hora"
                    .Columns(3).HeaderText = "Operario"
                    .Columns(4).HeaderText = "Cant"
                    .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(4).Width = 50
                    .Columns(3).Width = 150
                End With
            Case 5
                With dtmovimientos
                    .Columns(0).HeaderText = "Codigo"
                    .Columns(1).HeaderText = "Fecha"
                    .Columns(2).HeaderText = "Hora"
                    .Columns(3).HeaderText = "Tienda"
                    .Columns(4).HeaderText = "Operario"
                    .Columns(5).HeaderText = "Cant"
                    .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(3).Width = 150
                    .Columns(4).Width = 150
                    .Columns(5).Width = 50
                End With
            Case 6
                With dtmovimientos
                    .Columns(0).HeaderText = "Codigo"
                    .Columns(1).HeaderText = "Fecha"
                    .Columns(2).HeaderText = "Hora"
                    .Columns(3).HeaderText = "Operario"
                    .Columns(4).HeaderText = "Cant"
                    .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .Columns(4).Width = 50
                    .Columns(3).Width = 150
                End With
        End Select
        dtmovimientos.Columns(0).Width = 60
        dtmovimientos.Columns(1).Width = 80
        dtmovimientos.Columns(2).Width = 60
    End Sub
End Class