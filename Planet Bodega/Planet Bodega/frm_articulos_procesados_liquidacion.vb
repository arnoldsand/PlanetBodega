Public Class frm_articulos_procesados_liquidacion
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_articulos_procesados_liquidacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        codiimportation = vbEmpty
        DataGridView1.ColumnCount = 11
        preparar_columnas()
    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button7_Click(sender As Object, e As EventArgs) Handles button7.Click
        '   frm_seleccionar_importacion.ShowDialog()
        '  frm_seleccionar_importacion.Dispose()
        TextBox1.Focus()
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 120
            .Columns(1).Width = 150
            .Columns(2).Width = 50
            .Columns(3).Width = 50
            .Columns(4).Width = 50
            .Columns(5).Width = 80
            .Columns(6).Width = 180
            .Columns(7).Width = 100
            .Columns(8).Width = 50
            .Columns(9).Width = 80
            .Columns(10).Width = 80
            .Columns(0).HeaderText = "Referencia"
            .Columns(0).DefaultCellStyle.Font = New Font(DataGridView1.Font.Name, DataGridView1.Font.Size, FontStyle.Bold)
            .Columns(1).HeaderText = "Descripcion"
            .Columns(2).HeaderText = "Cant"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).HeaderText = "Dañado"
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).HeaderText = "Cant Real"
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).HeaderText = "O. Produccion"
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).HeaderText = "Realizada"
            .Columns(7).HeaderText = "Fecha"
            .Columns(8).HeaderText = "Cant"
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).HeaderText = "Total"
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).HeaderText = "Faltante"
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox9, "Importación") = False Then Exit Sub
        filtrar_referencias(TextBox1.Text)
    End Sub

    Private Sub filtrar_referencias(txtref As String)
        clase.dt_global.Tables.Clear()
        clase.consultar_global("SELECT articulos.ar_referencia, articulos.ar_descripcion, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4, SUM(entradamercancia.com_unidades) AS sumacantidades FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) WHERE(entradamercancia.com_codigoimp = " & codiimportation & " AND articulos.ar_referencia like '" & txtref & "%') GROUP BY articulos.ar_referencia, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 ORDER BY articulos.ar_referencia ASC", "tabla")
        If clase.dt_global.Tables("tabla").Rows.Count > 0 Then
            Dim f, g As Integer
            g = 0
            DataGridView1.RowCount = 0
            Dim total As Integer = 0
            Dim dañado As Integer = 0
            For f = 0 To clase.dt_global.Tables("tabla").Rows.Count - 1
                With DataGridView1
                    .RowCount = .RowCount + 1
                    total = 0
                    .Item(0, g).Value = clase.dt_global.Tables("tabla").Rows(f)("ar_referencia")
                    .Item(1, g).Value = clase.dt_global.Tables("tabla").Rows(f)("ar_descripcion")
                    .Item(2, g).Value = clase.dt_global.Tables("tabla").Rows(f)("sumacantidades")
                    dañado = calcular_dañado_mercancia_no_almacenada_apartir_referencia(1, codiimportation, clase.dt_global.Tables("tabla")(f)("ar_referencia"), clase.dt_global.Tables("tabla")(f)("ar_descripcion"), clase.dt_global.Tables("tabla")(f)("ar_linea"), clase.dt_global.Tables("tabla")(f)("ar_sublinea1"), clase.dt_global.Tables("tabla")(f)("ar_sublinea2"), clase.dt_global.Tables("tabla")(f)("ar_sublinea3"), clase.dt_global.Tables("tabla")(f)("ar_sublinea4"))
                    .Item(3, g).Value = dañado
                    total = clase.dt_global.Tables("tabla").Rows(f)("sumacantidades")
                    .Item(4, g).Value = total - dañado
                    clase.consultar(buscar_codigos_a_apartir_de_la_referencia(clase.dt_global.Tables("tabla")(f)("ar_referencia"), clase.dt_global.Tables("tabla")(f)("ar_descripcion"), clase.dt_global.Tables("tabla")(f)("ar_linea"), clase.dt_global.Tables("tabla")(f)("ar_sublinea1"), clase.dt_global.Tables("tabla")(f)("ar_sublinea2"), clase.dt_global.Tables("tabla")(f)("ar_sublinea3"), clase.dt_global.Tables("tabla")(f)("ar_sublinea4")), "tabla1")
                    If clase.dt.Tables("tabla1").Rows.Count > 0 Then
                        'pendiente verificar si las referencias a mostrar pertenecen a ordenes de produccion recibidas o no
                        Dim sql As String = "SELECT ordenproduccion.op_fecha, ordenproduccion.op_codigo, ordenproduccion.op_realizadapor, SUM(deordenprod.do_unidades) AS cantidad FROM deordenprod INNER JOIN ordenproduccion ON (deordenprod.do_idcaborden = ordenproduccion.op_codigo) WHERE ((ordenproduccion.op_codigoimportacion =" & codiimportation & ") AND ("
                        Dim ind As Boolean = True
                        Dim x As Short
                        For x = 0 To clase.dt.Tables("tabla1").Rows.Count - 1
                            If ind = True Then
                                sql = sql & " deordenprod.do_codigo = " & clase.dt.Tables("tabla1")(x)("ar_codigo")
                                ind = False
                            Else
                                sql = sql & " OR deordenprod.do_codigo = " & clase.dt.Tables("tabla1")(x)("ar_codigo")
                            End If
                        Next
                        sql = sql & ")) GROUP BY ordenproduccion.op_codigo"
                        clase.consultar1(sql, "tabla2")
                        Dim cant As Integer
                        If clase.dt1.Tables("tabla2").Rows.Count > 0 Then
                            Dim fecha As Date
                            cant = 0
                            For x = 0 To clase.dt1.Tables("tabla2").Rows.Count - 1
                                If x <> 0 Then
                                    .RowCount = .RowCount + 1
                                End If
                                .Item(5, g).Value = clase.dt1.Tables("tabla2").Rows(x)("op_codigo")
                                .Item(6, g).Value = clase.dt1.Tables("tabla2").Rows(x)("op_realizadapor")
                                fecha = clase.dt1.Tables("tabla2").Rows(x)("op_fecha")
                                .Item(7, g).Value = fecha.ToString("dd/MM/yyyy")
                                .Item(8, g).Value = clase.dt1.Tables("tabla2").Rows(x)("cantidad")
                                cant = cant + clase.dt1.Tables("tabla2").Rows(x)("cantidad")
                                g = g + 1
                            Next
                            .Item(9, g - x).Value = cant
                            .Item(10, g - x).Value = total - cant - dañado
                        Else
                            cant = 0
                            g = g + 1
                            .Item(9, g - 1).Value = cant
                            .Item(10, g - 1).Value = total - cant - dañado
                        End If
                    Else
                        g = g + 1
                    End If
                End With
            Next
        Else
            DataGridView1.RowCount = 0
        End If
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        filtrar_referencias(TextBox1.Text)
    End Sub
End Class