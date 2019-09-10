Public Class frmAsignar
    Dim clase As New class_library
    Dim Validado As Boolean
    Dim Operario, Fin, Tienda, CantFinaliza, IdAsignacion As String
    Public Orden, Codigo, Descripcion, Referencia, Cant, CantAsig, Danado, SelDanado As String

    Private Sub txtOrden_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtOrden.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Public Sub LlenarGridOrden()
        grdOrden.DataSource = Nothing
        clase.consultar("SELECT ordenproduccion.op_codigo AS ORDEN, deordenprod.do_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, deordenprod.do_unidades AS CANT FROM deordenprod INNER JOIN articulos ON (deordenprod.do_codigo = articulos.ar_codigo) INNER JOIN ordenproduccion ON (deordenprod.do_idcaborden = ordenproduccion.op_codigo)WHERE (ordenproduccion.op_codigo ='" & txtOrden.Text & "');", "orden")
        If clase.dt.Tables("orden").Rows.Count = 0 Then
            MessageBox.Show("NO EXISTE ESTA ORDEN DE PRODUCCION, VUELVA A DIGITARLA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtOrden.Text = ""
            txtOrden.Focus()
            Validado = False
            Timer1.Enabled = False
            Exit Sub
        End If

        Validado = True
        'AGREGAMOS UNA COLUMNA AL DATASET CON EL NOMBRE "CANTASIG" PARA SABER LAS UNIDADES ASIGNADAS
        clase.dt.Tables("orden").Columns.Add("CANTASIG")
        clase.dt.Tables("orden").Columns.Add("DANADO")

        With grdOrden
            .DataSource = clase.dt.Tables("orden")
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .RowHeadersVisible = False

            .Columns(0).Visible = False
            .Columns(1).Width = 70
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).Width = 150
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).Width = 150
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(4).Width = 70
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).Width = 70
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).Width = 70
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        For i = 0 To grdOrden.RowCount - 1
            clase.consultar1("SELECT SUM(acp_cantidad) AS CANTASIG,SUM(acp_danado) AS DANADO FROM det_asignacion_orden_produccion WHERE (aop_articulo ='" & grdOrden(1, [i]).Value.ToString & "' AND acp_ordenproduccion ='" & txtOrden.Text & "');", "cantasig")
            If clase.dt1.Tables("cantasig").Rows.Count > 0 Then
                grdOrden(5, [i]).Value = clase.dt1.Tables("cantasig").Rows(0)("CANTASIG")
                grdOrden(6, [i]).Value = clase.dt1.Tables("cantasig").Rows(0)("DANADO")
                If grdOrden(4, [i]).Value = (Val(grdOrden(5, [i]).Value) + Val(grdOrden(6, [i]).Value)) Then
                    grdOrden.Rows(i).DefaultCellStyle.BackColor = Color.DarkSlateBlue
                    grdOrden.Rows(i).DefaultCellStyle.ForeColor = Color.White
                End If
            End If
        Next
        txtArticulo.Enabled = True
        btnBuscar.Enabled = True

        grdOperario.DataSource = Nothing
        LlenarTodos()
    End Sub

    Public Sub LlenarTodos()
        Dim horaini2, fechaini2, fechafin2, horafin2 As Date
        Dim diffecha2 As String
        Dim dif2 As TimeSpan
        grdOperario.DataSource = Nothing
        clase.consultar("SELECT det_asignacion_orden_produccion.acp_id AS ID, aop_articulo AS CODIGO, articulos.ar_referencia as REFERENCIA, articulos.ar_descripcion as DESCRIPCION, det_asignacion_orden_produccion.acp_operario AS OPERARIO, det_asignacion_orden_produccion.acp_cantidad AS CANT, det_asignacion_orden_produccion.acp_danado AS DANADO, CONCAT_WS(' // ',acp_fecha_inicio,acp_hora_inicio) AS INICIO, CONCAT_WS(' // ',acp_fecha_fin,acp_hora_fin) AS FIN FROM det_asignacion_orden_produccion INNER JOIN articulos ON (det_asignacion_orden_produccion.aop_articulo = articulos.ar_codigo) WHERE (det_asignacion_orden_produccion.acp_ordenproduccion ='" & txtOrden.Text & "');", "asignados2")
        clase.dt.Tables("asignados2").Columns.Add("TIEMPO")

        With grdOperario
            .DataSource = clase.dt.Tables("asignados2")
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .RowHeadersVisible = False
            .Columns(0).Visible = False

            .Columns(1).Width = 70
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).Width = 100
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).Width = 100
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            .Columns(4).Width = 70
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).Width = 50
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).Width = 70
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).Width = 165
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).Width = 165
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).Width = 160
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        End With
        For i = 0 To grdOperario.RowCount - 1
            clase.consultar1("SELECT acp_operario,acp_fecha_inicio AS FINI,acp_fecha_fin AS FFIN, acp_hora_inicio AS INICIO, acp_hora_fin AS FIN FROM det_asignacion_orden_produccion WHERE (acp_id='" & grdOperario(0, i).Value & "' and acp_operario='" & grdOperario(4, i).Value & "')", "hora2")
            If IsDBNull(clase.dt1.Tables("hora2").Rows(0)("FIN")) Then
                horaini2 = clase.dt1.Tables("hora2").Rows(0)("INICIO").ToString
                fechaini2 = clase.dt1.Tables("hora2").Rows(0)("FINI").ToString
                diffecha2 = DateDiff(DateInterval.Day, Now, fechaini2)
                horafin2 = clase.FormatoHora(TimeOfDay)
                dif2 = (horafin2 - horaini2)
                grdOperario(9, [i]).Value = diffecha2 + "//" + dif2.ToString
            Else
                fechaini2 = clase.dt1.Tables("hora2").Rows(0)("FINI").ToString
                fechafin2 = clase.dt1.Tables("hora2").Rows(0)("FFIN").ToString
                diffecha2 = DateDiff(DateInterval.Day, fechafin2, fechaini2)
                horaini2 = clase.dt1.Tables("hora2").Rows(0)("INICIO").ToString
                horafin2 = clase.dt1.Tables("hora2").Rows(0)("FIN").ToString
                dif2 = (horafin2 - horaini2)
                grdOperario(9, [i]).Value = diffecha2 + "//" + dif2.ToString
            End If
        Next
    End Sub

    Private Sub grdOrden_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdOrden.CellClick
        Dim Y As Integer = grdOrden.CurrentCell.RowIndex
        Orden = grdOrden(0, [Y]).Value.ToString
        Codigo = grdOrden(1, [Y]).Value.ToString
        Referencia = grdOrden(2, [Y]).Value.ToString
        Descripcion = grdOrden(3, [Y]).Value.ToString
        Cant = grdOrden(4, [Y]).Value.ToString
        SelDanado = grdOrden(6, [Y]).Value.ToString
        If grdOrden(5, [Y]).Value.ToString = "" Then
            CantAsig = "0"
        Else
            CantAsig = grdOrden(5, [Y]).Value.ToString
        End If
        Operario = ""
        Fin = ""
        CantFinaliza = ""
        IdAsignacion = ""
    End Sub

    Private Sub btnAsignar_Click(sender As Object, e As EventArgs) Handles btnAsignar.Click
        If Orden = "" Then
            MessageBox.Show("DEBE SELECCIONAR UN PRODUCTO PARA ASIGNAR", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        Else
            If Cant = (Val(CantAsig) + Val(SelDanado)) Then
                MessageBox.Show("ESTE PRODUCTO YA FUE REVISADO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            frmAsignarOperario.ShowDialog()
            frmAsignarOperario.Dispose()
        End If
    End Sub

    Private Sub btnFinalizar_Click(sender As Object, e As EventArgs) Handles btnFinalizar.Click
        Dim TotOrden, TotAsignado, TotDanado As String
        If Operario = "" Then
            MessageBox.Show("DEBE SELECCIONAR UN OPERARIO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If Fin <> "" Then
            MessageBox.Show("ESTA LABOR YA FUE FINALIZADA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If MessageBox.Show("DESEA FINALIZAR LA LABOR QUE SE ENCUENTRA REALIZANDO EL OPERARIO " & Operario & " EN LA ORDEN " & txtOrden.Text & " Y EL ARTICULO " & Codigo & "?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            frmDanado1.ShowDialog()
            frmDanado1.Dispose()
            'ACTUALIZAR LA TABLA DET_ASIGNACION_ORDEN_PRODUCCION.DANADO
            Dim TotalProducto = Val(CantFinaliza) - Val(Danado)
            clase.actualizar("UPDATE det_asignacion_orden_produccion SET acp_cantidad='" & TotalProducto & "',acp_danado='" & Danado & "' WHERE (acp_ordenproduccion='" & txtOrden.Text & "' AND aop_articulo='" & Codigo & "' AND acp_operario='" & Operario & "' AND acp_fecha_fin IS NULL)")

            'ASIGNAR LA FECHA DE FINALIZACION AL OPERARIO
            clase.actualizar("UPDATE det_asignacion_orden_produccion SET acp_fecha_fin='" & clase.FormatoFecha(Date.Today) & "', acp_hora_fin='" & clase.FormatoHora(TimeOfDay) & "' WHERE acp_id='" & IdAsignacion & "'")
            LlenarGridOrden()
            LlenarTodos()

            'GUARDAR DISTRIBUCION_ORDEN_DE_PRODUCCION
            clase.consultar("SELECT det_asignacion_tiendas.as_tienda AS TIENDA, det_asignacion_tiendas.as_cantidad AS CANTIDAD, det_asignacion_orden_produccion.aop_articulo as ARTICULO FROM det_asignacion_tiendas INNER JOIN det_asignacion_orden_produccion ON (det_asignacion_tiendas.as_asigorden = det_asignacion_orden_produccion.acp_id)WHERE (det_asignacion_orden_produccion.acp_id ='" & IdAsignacion & "');", "distribucion")
            For i = 0 To clase.dt.Tables("distribucion").Rows.Count - 1
                clase.agregar_registro("INSERT INTO distribucion_orden_de_produccion(ordendeproduccion,tienda,articulo,cantidad) VALUES('" & txtOrden.Text & "','" & clase.dt.Tables("distribucion").Rows(i)("TIENDA").ToString & "','" & clase.dt.Tables("distribucion").Rows(i)("ARTICULO").ToString & "','" & clase.dt.Tables("distribucion").Rows(i)("CANTIDAD").ToString & "')")
            Next

            'VERIFICAR SI ES EL ULTIMO PRODUCTO DE LA ORDEN PARA CAMBIAR ordenproduccion.op_procesado A TRUE
            clase.consultar("SELECT do_idcaborden, SUM(do_unidades) AS sumdeorden FROM deordenprod WHERE (do_idcaborden ='" & txtOrden.Text & "')GROUP BY do_idcaborden;", "sumdeorden")
            TotOrden = clase.dt.Tables("sumdeorden").Rows(0)("sumdeorden")
            clase.consultar("SELECT acp_ordenproduccion, SUM(acp_cantidad) AS sumdetasig,SUM(acp_danado) AS sumdanado FROM det_asignacion_orden_produccion WHERE (acp_ordenproduccion ='" & txtOrden.Text & "' AND acp_fecha_fin IS NOT NULL)GROUP BY acp_ordenproduccion;", "sumdetasig")
            TotAsignado = clase.dt.Tables("sumdetasig").Rows(0)("sumdetasig")
            TotDanado = clase.dt.Tables("sumdetasig").Rows(0)("sumdanado")

            If TotOrden = (Val(TotAsignado) + Val(TotDanado)) Then
                clase.actualizar2("UPDATE ordenproduccion SET op_procesado='1' WHERE op_codigo='" & txtOrden.Text & "'")
                LlenarGridOrden()
                MessageBox.Show("ESTE FUE EL ULTIMO ARTICULO DE ESTA ORDEN, LA ORDEN FUE FINALIZADA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub btnPreTransferencia_Click(sender As Object, e As EventArgs) Handles btnPreTransferencia.Click
        If Validado = False Then
            MessageBox.Show("DEBE DIGITAR UN NUMERO DE ORDEN VALIDA PARA VER LA PRETRANSFERENCIA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        frmPreTransferencia.ShowDialog()
        frmPreTransferencia.Dispose()
    End Sub

    Private Sub frmAsignar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Validado = False
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtOrden, "Orden de Producción") = False Then Exit Sub
        Timer1.Enabled = True
        LlenarGridOrden()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        If clase.validar_cajas_text(txtArticulo, "Codigo") = False Then Exit Sub
        clase.consultar("SELECT ordenproduccion.op_codigo AS ORDEN, deordenprod.do_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, deordenprod.do_unidades AS CANT FROM deordenprod INNER JOIN articulos ON (deordenprod.do_codigo = articulos.ar_codigo) INNER JOIN ordenproduccion ON (deordenprod.do_idcaborden = ordenproduccion.op_codigo)WHERE (ordenproduccion.op_codigo ='" & txtOrden.Text & "' AND (deordenprod.do_codigo='" & txtArticulo.Text & "') OR articulos.ar_codigobarras='" & txtArticulo.Text & "');", "busqueda")
        If clase.dt.Tables("busqueda").Rows.Count = 0 Then
            MessageBox.Show("No se encuentra articulo en esta orden de producción", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            For i = 0 To grdOrden.RowCount - 1
                If grdOrden(1, i).Value = clase.dt.Tables("busqueda").Rows(0)("CODIGO") Then
                    grdOrden.Rows(i).Selected = True
                    grdOrden.CurrentCell = grdOrden.Rows(i).Cells(1)
                End If
            Next
        End If
    End Sub

    Private Sub txtArticulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtArticulo.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub btnVer_Click(sender As Object, e As EventArgs)
        LlenarTodos()
    End Sub

    Private Sub grdOperario_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdOperario.CellClick
        Dim Y As Integer = grdOperario.CurrentCell.RowIndex
        IdAsignacion = grdOperario(0, [Y]).Value.ToString
        Operario = grdOperario(4, [Y]).Value.ToString
        Fin = grdOperario(8, [Y]).Value.ToString
        CantFinaliza = grdOperario(5, [Y]).Value.ToString
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LlenarTodos()
    End Sub
End Class
