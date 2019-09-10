Public Class frm_articulos_procesados_tiqueteo
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_articulos_procesados_tiqueteo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        codiimportation = vbEmpty
        With DataGridView1
            .ColumnCount = 12
            .Columns(0).Width = 120
            .Columns(1).Width = 150
            .Columns(2).Width = 50
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).Width = 60
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).Width = 50
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).Width = 50
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).Width = 80
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).Width = 180
            .Columns(8).Width = 100
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).Width = 50
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).Width = 80
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).Width = 80
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripcion"
            .Columns(2).HeaderText = "Cant"
            .Columns(3).HeaderText = "Restante"
            .Columns(4).HeaderText = "Dañado"
            .Columns(5).HeaderText = "Cant Real"
            .Columns(6).HeaderText = "Transferencia"
            .Columns(7).HeaderText = "Recibida Por"
            .Columns(8).HeaderText = "Fecha"
            .Columns(9).HeaderText = "Cant"
            .Columns(10).HeaderText = "Total"
            .Columns(11).HeaderText = "Faltante"

        End With
    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button7_Click(sender As Object, e As EventArgs) Handles button7.Click
        frm_seleccionar_importacion3.ShowDialog()
        frm_seleccionar_importacion3.Dispose()
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

   

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox9, "Importación") = False Then Exit Sub
        llenar_referencia(TextBox1.Text)
    End Sub

    Private Sub llenar_referencia(referenciabus As String)
        clase.dt_global.Tables.Clear()  'pendiente (linea de abajo) verificar si las referencias a mostrar pertenecen a ordenes de produccion recibidas o no
        clase.consultar_global("SELECT articulos.ar_referencia, articulos.ar_descripcion, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4, SUM(deordenprod.do_unidades) AS cantidad FROM deordenprod INNER JOIN ordenproduccion ON (deordenprod.do_idcaborden = ordenproduccion.op_codigo) INNER JOIN articulos ON (articulos.ar_codigo = deordenprod.do_codigo) WHERE (ordenproduccion.op_codigoimportacion =" & codiimportation & " AND articulos.ar_referencia like '" & referenciabus & "%') GROUP BY articulos.ar_referencia, articulos.ar_descripcion, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4", "tabla5")
        If clase.dt_global.Tables("tabla5").Rows.Count > 0 Then
            Dim restante, danado As Integer
            With DataGridView1
                Dim i, g As Integer
                g = 0
                DataGridView1.RowCount = 0
                Dim total As Integer = 0
                For i = 0 To clase.dt_global.Tables("tabla5").Rows.Count - 1
                    .RowCount = .RowCount + 1
                    .Item(0, g).Value = clase.dt_global.Tables("tabla5").Rows(i)("ar_referencia")
                    .Item(1, g).Value = clase.dt_global.Tables("tabla5").Rows(i)("ar_descripcion")
                    .Item(2, g).Value = clase.dt_global.Tables("tabla5").Rows(i)("cantidad")

                    restante = calcular_reserva_x_referencia(codiimportation, clase.dt_global.Tables("tabla5").Rows(i)("ar_referencia"), clase.dt_global.Tables("tabla5").Rows(i)("ar_descripcion"), clase.dt_global.Tables("tabla5").Rows(i)("ar_linea"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea1"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea2"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea3"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea4"))
                    .Item(3, g).Value = restante
                    danado = calcular_dañado_mercancia_no_almacenada_apartir_referencia(2, codiimportation, clase.dt_global.Tables("tabla5").Rows(i)("ar_referencia"), clase.dt_global.Tables("tabla5").Rows(i)("ar_descripcion"), clase.dt_global.Tables("tabla5").Rows(i)("ar_linea"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea1"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea2"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea3"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea4"))
                    .Item(4, g).Value = danado
                    .Item(5, g).Value = restante - danado
                    clase.consultar(buscar_codigos_a_apartir_de_la_referencia(clase.dt_global.Tables("tabla5").Rows(i)("ar_referencia"), clase.dt_global.Tables("tabla5").Rows(i)("ar_descripcion"), clase.dt_global.Tables("tabla5").Rows(i)("ar_linea"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea1"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea2"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea3"), clase.dt_global.Tables("tabla5").Rows(i)("ar_sublinea4")), "tabla10")
                    If clase.dt.Tables("tabla10").Rows.Count > 0 Then
                        Dim sql As String = "SELECT transferencia_bodega.trans_fecha, transferencia_bodega.trans_codigo, transferencia_bodega.trans_realizada_por, SUM(det_transferencia_bodega.detrans_cantidad) AS cantidad FROM det_transferencia_bodega INNER JOIN transferencia_bodega ON (det_transferencia_bodega.detrans_transferencia = transferencia_bodega.trans_codigo)  WHERE ((transferencia_bodega.trans_importacion =" & codiimportation & ") AND ("
                        Dim ind As Boolean = True
                        Dim x As Short
                        For x = 0 To clase.dt.Tables("tabla10").Rows.Count - 1
                            If ind = True Then
                                sql = sql & "(det_transferencia_bodega.detrans_articulo =" & clase.dt.Tables("tabla10")(x)("ar_codigo") & ")"
                                ind = False
                            Else
                                sql = sql & " OR (det_transferencia_bodega.detrans_articulo =" & clase.dt.Tables("tabla10")(x)("ar_codigo") & ")"
                            End If
                        Next
                        sql = sql & ")) GROUP BY det_transferencia_bodega.detrans_transferencia"

                        clase.consultar1(sql, "tabla11")
                        Dim cant As Integer
                        If clase.dt1.Tables("tabla11").Rows.Count > 0 Then

                            Dim fecha As Date
                            cant = 0
                            For x = 0 To clase.dt1.Tables("tabla11").Rows.Count - 1
                                If x <> 0 Then
                                    .RowCount = .RowCount + 1
                                End If
                                .Item(6, g).Value = clase.dt1.Tables("tabla11").Rows(x)("trans_codigo")
                                .Item(7, g).Value = clase.dt1.Tables("tabla11").Rows(x)("trans_realizada_por")
                                fecha = clase.dt1.Tables("tabla11").Rows(x)("trans_fecha")
                                .Item(8, g).Value = fecha.ToString("dd/MM/yyyy")
                                .Item(9, g).Value = clase.dt1.Tables("tabla11").Rows(x)("cantidad")
                                cant = cant + clase.dt1.Tables("tabla11").Rows(x)("cantidad")
                                g = g + 1
                            Next
                            .Item(10, g - x).Value = cant
                            .Item(11, g - x).Value = restante - cant - danado
                            '
                        Else
                            cant = 0
                            g = g + 1
                            .Item(10, g - 1).Value = cant
                            .Item(11, g - 1).Value = restante - cant - danado
                        End If
                    Else
                        g = g + 1
                    End If
                Next
            End With
        Else
            DataGridView1.RowCount = 0
        End If
    End Sub

    Function calcular_reserva_x_referencia(importacion As Short, ref As String, descri As String, linea As Short, sublinea1 As Short, sublinea2 As Object, sublinea3 As Object, sublinea4 As Object) As Integer
        clase.consultar(buscar_codigos_a_apartir_de_la_referencia(ref, descri, linea, sublinea1, sublinea2, sublinea3, sublinea4), "tabla6")
        If clase.dt.Tables("tabla6").Rows.Count > 0 Then
            Dim consql As String = "SELECT ordenproduccion.op_patrondist, deordenprod.do_idcaborden, SUM(deordenprod.do_unidades) AS cantidad FROM deordenprod INNER JOIN ordenproduccion ON (deordenprod.do_idcaborden = ordenproduccion.op_codigo) WHERE ((ordenproduccion.op_codigoimportacion =" & importacion & ") AND ("
            Dim x As Short
            Dim ind As Boolean = True
            For x = 0 To clase.dt.Tables("tabla6").Rows.Count - 1
                If ind = True Then
                    consql = consql & " (deordenprod.do_codigo =" & clase.dt.Tables("tabla6").Rows(x)("ar_codigo") & ")"
                    ind = False
                Else
                    consql = consql & " OR (deordenprod.do_codigo =" & clase.dt.Tables("tabla6").Rows(x)("ar_codigo") & ")"
                End If
            Next
            consql = consql & ")) GROUP BY deordenprod.do_idcaborden"

            clase.consultar1(consql, "tabla9")
            If clase.dt1.Tables("tabla9").Rows.Count > 0 Then
                Dim t As Short
                Dim cant_referencia As Integer = 0
                Dim calculo As Integer
                For t = 0 To clase.dt1.Tables("tabla9").Rows.Count - 1
                    calculo = clase.dt1.Tables("tabla9").Rows(t)("cantidad") - calcular_cantidad_total_patron(clase.dt1.Tables("tabla9").Rows(t)("op_patrondist"))
                    If calculo < 0 Then
                        calculo = 0
                    End If
                    cant_referencia = cant_referencia + calculo
                Next
                Return cant_referencia
            End If
        End If
    End Function

    Function calcular_cantidad_total_patron(patron As Short) As Integer
        clase.consultar("SELECT SUM(detpatrondist.dp_cantidad) AS Cantidad FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo) INNER JOIN tiendas ON (tiendas.id = detpatrondist.dp_tienda) WHERE (tiendas.estado =TRUE AND cabpatrondist.pat_codigo =" & patron & ")", "tabla8")
        If clase.dt.Tables("tabla8").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tabla8").Rows(0)("Cantidad")) Then
                Return 0
                Exit Function
            End If
            Return clase.dt.Tables("tabla8").Rows(0)("Cantidad")
        End If
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        llenar_referencia(TextBox1.Text)
    End Sub
End Class