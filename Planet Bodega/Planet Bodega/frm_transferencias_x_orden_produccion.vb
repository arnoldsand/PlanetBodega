Public Class frm_transferencias_x_orden_produccion
    Dim clase As New class_library
    Dim codigo_articulo As Long

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_transferencias_x_orden_produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT * FROM tiendas WHERE (estado =TRUE) ORDER BY tienda ASC", "tiendas")
        If clase.dt.Tables("tiendas").Rows.Count Then
            Dim a As Short
            ComboBox9.Items.Clear()
            ComboBox9.Items.Add("TODAS")
            For a = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
                ComboBox9.Items.Add(clase.dt.Tables("tiendas").Rows(a)("tienda"))
            Next
            ComboBox9.SelectedIndex = 0
        Else
            ComboBox9.Items.Clear()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '  If clase.validar_cajas_text(TextBox1, "Orden de Producción") = False Then Exit Sub
        If clase.validar_combobox(ComboBox9, "Tienda") = False Then Exit Sub
        llenar_grid(ComboBox9.Text)
    End Sub

    Private Sub llenar_grid(tiendas As String)
        Dim sql As String = ""
        If tiendas = "TODAS" Then
            sql = "select* from tiendas where estado = TRUE order by tienda ASC"
        Else
            sql = "select* from tiendas where tienda = '" & tiendas & "' order by tienda ASC"
        End If
        clase.consultar(sql, "tiendas")
        Dim a, c As Short
        Dim x As Short = 0 ' contador de filas
        Dim z As Short = 0
        Dim ind As Boolean = False
        DataGridView1.Rows.Clear()
        For a = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
            DataGridView1.RowCount = DataGridView1.RowCount + 1
            With DataGridView1
                .Item(0, x).Value = clase.dt.Tables("tiendas").Rows(a)("tienda")
                clase.consultar1("SELECT cabtransferencia.tr_numero, cabtransferencia.tr_operador, SUM(dettransferencia.dt_faltante) AS faltante, SUM(dettransferencia.dt_danado) AS danado, cabtransferencia.tr_revisada, cabtransferencia.tr_revisor FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) WHERE (cabtransferencia.tr_opproduccion =" & TextBox1.Text & " AND cabtransferencia.tr_destino =" & clase.dt.Tables("tiendas").Rows(a)("Id") & ") GROUP BY cabtransferencia.tr_numero", "trans-tiendas")
                If clase.dt1.Tables("trans-tiendas").Rows.Count > 0 Then
                    z = 0
                    For c = 0 To clase.dt1.Tables("trans-tiendas").Rows.Count - 1
                        If ind <> False Then
                            DataGridView1.RowCount = DataGridView1.RowCount + 1
                        End If
                        ind = True
                        DataGridView1.Item(1, c + x).Value = clase.dt1.Tables("trans-tiendas").Rows(c)("tr_numero")
                        DataGridView1.Item(2, c + x).Value = clase.dt1.Tables("trans-tiendas").Rows(c)("tr_operador")
                        DataGridView1.Item(3, c + x).Value = clase.dt1.Tables("trans-tiendas").Rows(c)("faltante")
                        DataGridView1.Item(4, c + x).Value = clase.dt1.Tables("trans-tiendas").Rows(c)("danado")
                        DataGridView1.Item(5, c + x).Value = clase.dt1.Tables("trans-tiendas").Rows(c)("tr_revisada")
                        DataGridView1.Item(6, c + x).Value = clase.dt1.Tables("trans-tiendas").Rows(c)("tr_revisor")
                        z = z + 1
                    Next
                    x = x + z
                    ind = False
                Else
                    x = x + 1
                End If
            End With
        Next
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox2, "Articulo") = False Then Exit Sub
        If Len(TextBox2.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox2.Text & "')", "tabla11")
            '  codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox5.Text)            esta linea seguramente debe ser quitada
        End If
        If Len(TextBox2.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox2.Text & ")", "tabla11")
            ' codigo_articulo = TextBox5.Text     esta linea seguramente debe ser quitada
        End If
        If Len(TextBox2.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Text = ""
            TextBox2.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            codigo_articulo = clase.dt.Tables("tabla11").Rows(0)("ar_codigo")
            clase.consultar1("SELECT acp_id, acp_operario, acp_cantidad + acp_danado AS recibido, acp_danado, acp_cantidad FROM det_asignacion_orden_produccion WHERE (aop_articulo =" & codigo_articulo & " AND acp_ordenproduccion =" & TextBox1.Text & " AND acp_fecha_fin IS NOT NULL AND acp_hora_fin IS NOT NULL)", "analisis-codigo")
            If clase.dt1.Tables("analisis-codigo").Rows.Count > 0 Then
                ' para continuar con esto debe hacer la prueba (simulacro) de la orden de produccio y el de asignacion de ordenes de trabajo
                DataGridView2.DataSource = clase.dt1.Tables("analisis-codigo")
                preparar_columnas2()
                TextBox3.Text = calcular_teorico_entregado()
                TextBox4.Text = calcular_teorico_en_digitacion()
                TextBox5.Text = calcular_dañado()
                TextBox6.Text = calcular_procesado_en_transferencia()
                TextBox7.Text = calcular_reserva()
            Else
                DataGridView2.DataSource = Nothing
                TextBox3.Text = 0
                TextBox4.Text = 0
                TextBox5.Text = 0
                TextBox6.Text = 0
                TextBox7.Text = 0
                MessageBox.Show("El articulo especificado no fue procesado en la orden de producción actual.", "ARTICULO NO PROCESADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox2.Text = ""
                TextBox2.Focus()
            End If
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Text = ""
            TextBox2.Focus()
        End If
    End Sub

    Function calcular_teorico_entregado() As Short
        clase.consultar1("SELECT SUM(acp_cantidad) AS neto FROM det_asignacion_orden_produccion WHERE (aop_articulo =" & codigo_articulo & " AND acp_fecha_fin IS NOT NULL AND acp_hora_fin IS NOT NULL AND acp_ordenproduccion =" & TextBox1.Text & ")", "teorico")
        If clase.dt1.Tables("teorico").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("teorico").Rows(0)("neto")) Then
                Return 0
            Else
                Return clase.dt1.Tables("teorico").Rows(0)("neto")
            End If
        Else
            Return 0
        End If
    End Function

    Function calcular_teorico_entregado_operario() As Short
        clase.consultar1("SELECT SUM(acp_cantidad) AS neto FROM det_asignacion_orden_produccion WHERE (aop_articulo =" & codigo_articulo & " AND acp_fecha_fin IS NOT NULL AND acp_hora_fin IS NOT NULL AND acp_ordenproduccion =" & TextBox1.Text & " AND acp_operario = " & DataGridView2.Item(1, DataGridView2.CurrentCell.RowIndex).Value & ")", "teorico")
        If clase.dt1.Tables("teorico").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("teorico").Rows(0)("neto")) Then
                Return 0
            Else
                Return clase.dt1.Tables("teorico").Rows(0)("neto")
            End If
        Else
            Return 0
        End If
    End Function

    Function calcular_dañado() As Short
        clase.consultar1("SELECT SUM(acp_danado) AS danado FROM det_asignacion_orden_produccion WHERE (acp_ordenproduccion =" & TextBox1.Text & " AND aop_articulo =" & codigo_articulo & ") GROUP BY acp_ordenproduccion", "danado")
        If clase.dt1.Tables("danado").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("danado").Rows(0)("danado")) Then
                Return 0
            Else
                Return clase.dt1.Tables("danado").Rows(0)("danado")
            End If
        Else
            Return 0
        End If
    End Function

    Function calcular_dañado_operario() As Short
        clase.consultar1("SELECT SUM(acp_danado) AS danado FROM det_asignacion_orden_produccion WHERE (acp_ordenproduccion =" & TextBox1.Text & " AND aop_articulo =" & codigo_articulo & " AND det_asignacion_orden_produccion.acp_operario = " & DataGridView2.Item(1, DataGridView2.CurrentCell.RowIndex).Value & ") GROUP BY acp_ordenproduccion", "danado")
        If clase.dt1.Tables("danado").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("danado").Rows(0)("danado")) Then
                Return 0
            Else
                Return clase.dt1.Tables("danado").Rows(0)("danado")
            End If
        Else
            Return 0
        End If
    End Function

    Function calcular_teorico_en_digitacion() As Short
        clase.consultar1("SELECT det_asignacion_tiendas.as_tienda FROM det_asignacion_tiendas INNER JOIN det_asignacion_orden_produccion ON (det_asignacion_tiendas.as_asigorden = det_asignacion_orden_produccion.acp_id) WHERE (det_asignacion_orden_produccion.aop_articulo =" & codigo_articulo & " AND det_asignacion_orden_produccion.acp_ordenproduccion =" & TextBox1.Text & ")", "digitacion")
        If clase.dt1.Tables("digitacion").Rows.Count > 0 Then
            Dim x As Short
            Dim contador As Short = 0
            For x = 0 To clase.dt1.Tables("digitacion").Rows.Count - 1
                clase.consultar("SELECT SUM(detpatrondist.dp_cantidad) AS suma FROM  detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo) INNER JOIN deordenprod  ON (deordenprod.do_patrondist = cabpatrondist.pat_codigo) WHERE (deordenprod.do_codigo =" & codigo_articulo & " AND deordenprod.do_idcaborden =" & TextBox1.Text & " AND detpatrondist.dp_tienda =" & clase.dt1.Tables("digitacion").Rows(x)("as_tienda") & ")", "detalleorden")
                If clase.dt.Tables("detalleorden").Rows.Count > 0 Then
                    If IsDBNull(clase.dt.Tables("detalleorden").Rows(0)("suma")) Then
                        contador = contador + 0
                    Else
                        contador = contador + clase.dt.Tables("detalleorden").Rows(0)("suma")
                    End If
                Else
                    contador = contador + 0
                End If
            Next
            Return contador
        Else
            Return 0
        End If
    End Function

    Function calcular_teorico_en_digitacion_operario() As Short
        clase.consultar1("SELECT det_asignacion_tiendas.as_tienda FROM det_asignacion_tiendas INNER JOIN det_asignacion_orden_produccion ON (det_asignacion_tiendas.as_asigorden = det_asignacion_orden_produccion.acp_id) WHERE (det_asignacion_orden_produccion.aop_articulo =" & codigo_articulo & " AND det_asignacion_orden_produccion.acp_ordenproduccion =" & TextBox1.Text & " AND det_asignacion_orden_produccion.acp_operario = " & DataGridView2.Item(1, DataGridView2.CurrentCell.RowIndex).Value & ")", "digitacion")
        If clase.dt1.Tables("digitacion").Rows.Count > 0 Then
            Dim x As Short
            Dim contador As Short = 0
            For x = 0 To clase.dt1.Tables("digitacion").Rows.Count - 1
                clase.consultar("SELECT SUM(detpatrondist.dp_cantidad) AS suma FROM  detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo) INNER JOIN deordenprod  ON (deordenprod.do_patrondist = cabpatrondist.pat_codigo) WHERE (deordenprod.do_codigo =" & codigo_articulo & " AND deordenprod.do_idcaborden =" & TextBox1.Text & " AND detpatrondist.dp_tienda =" & clase.dt1.Tables("digitacion").Rows(x)("as_tienda") & ")", "detalleorden")
                If clase.dt.Tables("detalleorden").Rows.Count > 0 Then
                    If IsDBNull(clase.dt.Tables("detalleorden").Rows(0)("suma")) Then
                        contador = contador + 0
                    Else
                        contador = contador + clase.dt.Tables("detalleorden").Rows(0)("suma")
                    End If
                Else
                    contador = contador + 0
                End If
            Next
            Return contador
        Else
            Return 0
        End If
    End Function

    Function calcular_procesado_en_transferencia() As Short
        clase.consultar1("SELECT SUM(dettransferencia.dt_cantidad) AS cant FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) WHERE (cabtransferencia.tr_opproduccion =" & TextBox1.Text & " AND dettransferencia.dt_codarticulo =" & codigo_articulo & " AND cabtransferencia.tr_estado = TRUE) GROUP BY cabtransferencia.tr_opproduccion", "procesado")
        If clase.dt1.Tables("procesado").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("procesado").Rows(0)("cant")) Then
                Return 0
            Else
                Return clase.dt1.Tables("procesado").Rows(0)("cant")
            End If
        Else
            Return 0
        End If
    End Function


    Private Sub preparar_columnas2()
        With DataGridView2
            .Columns(0).Visible = False
            .Columns(1).HeaderText = "Operario"
            .Columns(2).HeaderText = "Cantidad"
            .Columns(3).HeaderText = "Dañado"
            .Columns(4).HeaderText = "Entregado"
            .Columns(1).Width = 120
            .Columns(2).Width = 60
            .Columns(3).Width = 60
            .Columns(4).Width = 60
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If clase.validar_cajas_text(TextBox1, "Orden de Producción") = False Then Exit Sub
        clase.consultar("SELECT * FROM ordenproduccion WHERE (op_codigo =" & TextBox1.Text & ")", "ordenproduccion")
        If clase.dt.Tables("ordenproduccion").Rows.Count = 0 Then
            MessageBox.Show("La Orden de Producción digitada no existe.", "DATO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox1.Text = ""
            TextBox1.Focus()
            Exit Sub
        End If
        TextBox1.Enabled = False
        TabControl1.Enabled = True
        Button4.Enabled = False
    End Sub


    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        clase.consultar1("SELECT tiendas.Id, tiendas.tienda, det_asignacion_tiendas.as_cantidad  FROM det_asignacion_tiendas INNER JOIN tiendas ON (det_asignacion_tiendas.as_tienda = tiendas.id) WHERE (det_asignacion_tiendas.as_asigorden =" & DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value & ") ORDER BY tiendas.tienda ASC", "asignacion")
        If clase.dt1.Tables("asignacion").Rows.Count > 0 Then
            Dim a As Short
            DataGridView3.Rows.Clear()
            For a = 0 To clase.dt1.Tables("asignacion").Rows.Count - 1
                DataGridView3.RowCount = DataGridView3.RowCount + 1
                DataGridView3.Item(0, a).Value = clase.dt1.Tables("asignacion").Rows(a)("tienda")
                DataGridView3.Item(1, a).Value = clase.dt1.Tables("asignacion").Rows(a)("as_cantidad")
            Next
            TextBox12.Text = calcular_teorico_entregado_operario()
            TextBox11.Text = calcular_teorico_en_digitacion_operario()
            TextBox10.Text = calcular_dañado_operario()
            TextBox8.Text = calcular_reserva_operario()

        Else
            DataGridView3.Rows.Clear()
        End If
    End Sub

    Function calcular_reserva() As Short
        clase.consultar1("SELECT SUM(det_asignacion_tiendas.as_cantidad) AS cant FROM det_asignacion_tiendas INNER JOIN det_asignacion_orden_produccion  ON (det_asignacion_tiendas.as_asigorden = det_asignacion_orden_produccion.acp_id) WHERE (det_asignacion_orden_produccion.acp_ordenproduccion =" & TextBox1.Text & " AND det_asignacion_orden_produccion.aop_articulo =" & TextBox2.Text & " AND det_asignacion_tiendas.as_tienda =5000) GROUP BY det_asignacion_orden_produccion.aop_articulo", "reserva")
        If clase.dt1.Tables("reserva").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("reserva").Rows(0)("cant")) Then
                Return 0
            Else
                Return clase.dt1.Tables("reserva").Rows(0)("cant")
            End If
        Else
            Return 0
        End If
    End Function

    Function calcular_reserva_operario() As Short
        clase.consultar1("SELECT SUM(det_asignacion_tiendas.as_cantidad) AS cant FROM det_asignacion_tiendas INNER JOIN det_asignacion_orden_produccion  ON (det_asignacion_tiendas.as_asigorden = det_asignacion_orden_produccion.acp_id) WHERE (det_asignacion_orden_produccion.acp_ordenproduccion =" & TextBox1.Text & " AND det_asignacion_orden_produccion.aop_articulo =" & TextBox2.Text & " AND det_asignacion_tiendas.as_tienda =5000 AND det_asignacion_orden_produccion.acp_operario = " & DataGridView2.Item(1, DataGridView2.CurrentCell.RowIndex).Value & ") GROUP BY det_asignacion_orden_produccion.aop_articulo", "reserva")
        If clase.dt1.Tables("reserva").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("reserva").Rows(0)("cant")) Then
                Return 0
            Else
                Return clase.dt1.Tables("reserva").Rows(0)("cant")
            End If
        Else
            Return 0
        End If
    End Function

    
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub
End Class