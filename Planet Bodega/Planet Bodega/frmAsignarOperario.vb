Public Class frmAsignarOperario
    Dim clase As New class_library
    Dim Y As Integer
    Dim Tienda, Sel, SumPatron, SumAsigReserva, SumOrden, CantReserva, IdAsig As String
    Dim TablaReserva As New DataTable
    Dim Valida As Boolean = True

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtOperario.Text = ""
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        txtOperario.Text = txtOperario.Text & 1
    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        txtOperario.Text = txtOperario.Text & 2
    End Sub

    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        txtOperario.Text = txtOperario.Text & 3
    End Sub

    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        txtOperario.Text = txtOperario.Text & 4
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        txtOperario.Text = txtOperario.Text & 5
    End Sub

    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        txtOperario.Text = txtOperario.Text & 6
    End Sub

    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        txtOperario.Text = txtOperario.Text & 7
    End Sub

    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        txtOperario.Text = txtOperario.Text & 8
    End Sub

    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        txtOperario.Text = txtOperario.Text & 9
    End Sub

    Private Sub btn10_Click(sender As Object, e As EventArgs) Handles btn10.Click
        txtOperario.Text = txtOperario.Text & 0
    End Sub

    Private Sub frmAsignarOperario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtCodigo.Text = frmAsignar.Codigo
        txtReferencia.Text = frmAsignar.Referencia
        txtDescripcion.Text = frmAsignar.Descripcion
        clase.consultar("SELECT SUM(do_unidades) AS SUMORDEN FROM deordenprod WHERE (do_codigo ='" & frmAsignar.Codigo & "' AND do_idcaborden ='" & frmAsignar.Orden & "');", "sumorden")
        SumOrden = clase.dt.Tables("sumorden").Rows(0)("SUMORDEN")
        clase.consultar1("SELECT SUM(detpatrondist.dp_cantidad) AS SUMPATRON FROM deordenprod INNER JOIN detpatrondist ON (deordenprod.do_patrondist = detpatrondist.dp_idpatron)WHERE (deordenprod.do_idcaborden ='" & frmAsignar.Orden & "' AND deordenprod.do_codigo ='" & frmAsignar.Codigo & "');", "sumpatron")
        SumPatron = clase.dt1.Tables("sumpatron").Rows(0)("SUMPATRON")

        With TablaReserva
            .Columns.Add("TIENDA")
            .Columns.Add("CANT")
            .Columns.Add("OPERARIO")
        End With

        LlenarGridPatron()
    End Sub

    Private Sub btnAsignar_Click(sender As Object, e As EventArgs) Handles btnAsignar.Click
        If clase.validar_cajas_text(txtOperario, "OPERARIO") = False Then Exit Sub
        If clase.validar_cajas_text(txtCantidad, "CANTIDAD") = False Then Exit Sub
        If txtCantidad.Text = "0" Then
            MessageBox.Show("DEBE SELECCIONAR AL MENOS UNA TIENDA PARA ESTE OPERARIO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If txtCantidad.Text > (frmAsignar.Cant - (Val(frmAsignar.CantAsig) + Val(frmAsignar.SelDanado))) Then
            MessageBox.Show("LA CANTIDAD ASIGNADA NO PUEDE SER SUPERIOR A LA DISPONIBLE", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            LlenarGridPatron()
            txtCantidad.Text = ""
            txtCantidad.Focus()
            Exit Sub
        End If

        'ValidaOperario(txtOperario.Text)
        'If Valida = False Then
        '    MessageBox.Show("EL OPERARIO YA SE ENCUENTRA TRABAJANDO EN LA ORDEN DE PRODUCCION " & clase.dt.Tables("operario").Rows(0)("acp_ordenproduccion") & " CON EL ARTICULO " & clase.dt.Tables("operario").Rows(0)("aop_articulo") & "", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    txtOperario.Text = ""
        '    txtCantidad.Text = ""
        '    chkTodos.Checked = False
        '    LlenarGridPatron()
        '    Exit Sub
        'End If

        For i = 0 To grdPatron.RowCount - 1
            If IsDBNull(grdPatron(3, [i]).Value) Then
                Continue For
            End If
            If (grdPatron(3, [i]).Value = True) Then
                'VALIDAR QUE OTRO OPERARIO NO ESTE TRABAJANDO EN LA MISMA TIENDA
                clase.consultar("SELECT det_asignacion_tiendas.as_tienda AS TIENDA, det_asignacion_orden_produccion.acp_operario AS OPERARIO, det_asignacion_orden_produccion.aop_articulo AS ARTICULO, det_asignacion_orden_produccion.acp_fecha_inicio, det_asignacion_orden_produccion.acp_fecha_fin FROM det_asignacion_orden_produccion, det_asignacion_tiendas WHERE (det_asignacion_orden_produccion.acp_id=det_asignacion_tiendas.as_asigorden AND det_asignacion_tiendas.as_tienda ='" & grdPatron(0, [i]).Value.ToString & "' AND det_asignacion_orden_produccion.acp_fecha_inicio <>'' AND det_asignacion_orden_produccion.aop_articulo ='" & frmAsignar.Codigo & "' AND det_asignacion_orden_produccion.acp_fecha_fin IS NULL);", "tienda")
                If clase.dt.Tables("tienda").Rows.Count > 0 Then
                    MessageBox.Show("EL OPERARIO " & clase.dt.Tables("tienda").Rows(0)("OPERARIO") & " ESTA TRABAJANDO ESTE MISMO CODIGO PARA LA MISMA TIENDA " & clase.dt.Tables("tienda").Rows(0)("TIENDA") & "", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtOperario.Text = ""
                    txtCantidad.Text = ""
                    LlenarGridPatron()
                    Exit Sub
                End If
            End If
        Next

        If MessageBox.Show("DESEA ASIGNAR ESTE PRODUCTO AL OPERARIO?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            'GUARDAR LAS TIENDAS DE CADA OPERARIO
            Dim Total As String = "0"
            For i = 0 To grdPatron.RowCount - 1
                If IsDBNull(grdPatron(3, [i]).Value) Then
                    Continue For
                End If
                If (grdPatron(3, [i]).Value = True) Then
                    Total = (Val(grdPatron(2, [i]).Value.ToString) + Val(Total))
                End If
            Next
            clase.agregar_registro("INSERT INTO det_asignacion_orden_produccion(aop_articulo,acp_cantidad,acp_ordenproduccion,acp_operario,acp_fecha_inicio,acp_hora_inicio,acp_estado,acp_danado) VALUES('" & frmAsignar.Codigo & "','" & Total & "','" & frmAsignar.Orden & "','" & txtOperario.Text & "','" & clase.FormatoFecha(Date.Today) & "','" & clase.FormatoHora(TimeOfDay) & "','A','0')")

            'BUSCAR EL CONSECUTIVO PARA EL DETALLE
            clase.consultar("SELECT MAX(acp_id) as consecutivo FROM det_asignacion_orden_produccion", "consecutivo")
            Dim Consecutivo As String = clase.dt.Tables("consecutivo").Rows(0)("consecutivo")

            For i = 0 To grdPatron.RowCount - 1
                If IsDBNull(grdPatron(3, [i]).Value) Then
                    Continue For
                End If
                If (grdPatron(3, [i]).Value = True) Then
                    'GUARDAR LAS TIENDAS POR OPERARIO
                    clase.agregar_registro2("INSERT INTO det_asignacion_tiendas(as_asigorden,as_tienda,as_cantidad) VALUES('" & Consecutivo & "','" & grdPatron(0, [i]).Value.ToString & "','" & grdPatron(2, [i]).Value.ToString & "')")
                End If
            Next
            'FIN GUARDAR ASIGNACION TIENDA

            MessageBox.Show("ARTICULO ASIGNADO CORRECTAMENTE", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            frmAsignar.LlenarGridOrden()
            frmAsignar.LlenarTodos()
            frmAsignar.Enabled = True
            Me.Close()

        End If
    End Sub

    Private Sub btnAsigReserva_Click(sender As Object, e As EventArgs) Handles btnAsigReserva.Click
        ValidaOperario(txtOperario.Text)
        If Valida = False Then
            If (Val(txtCantidad.Text) <> Val(txtReserva.Text)) Then
                MessageBox.Show("DEBE SELECCIONAR TODA LA RESERVA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtReserva.Text = ""
                txtReserva.Focus()
                Exit Sub
            End If
        End If

        If txtOperario.Text = "" Then
            MessageBox.Show("DEBE SELECCIONAR UN OPERARIO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim Buscar() As DataRow
        Buscar = TablaReserva.Select("OPERARIO = '" & txtOperario.Text & "'")
        If Buscar.Length > 0 Then
            MessageBox.Show("ESTE OPERARIO YA SE ENCUENTRA TRABAJANDO EN LA RESERVA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtOperario.Text = ""
            txtReserva.Text = ""
            txtReserva.Focus()
            Exit Sub
        End If

        If txtReserva.Text <> "" Then
            If (Val(txtReserva.Text) <= Val(CantReserva)) Then
                SumAsigReserva = SumAsigReserva + Val(txtReserva.Text)
                If (Val(SumAsigReserva) > Val(CantReserva)) Then
                    MessageBox.Show("LA CANTIDAD ASIGNADA SUPERA LA RESERVA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SumAsigReserva = SumAsigReserva - Val(txtReserva.Text)
                    txtReserva.Text = ""
                    txtReserva.Focus()
                    Exit Sub
                End If

                Reserva(txtReserva.Text)
                txtCantidad.Text = Val((txtCantidad.Text) - Val(txtReserva.Text))
                txtReserva.Text = ""
                txtOperario.Text = ""
                txtOperario.Focus()
            End If
        End If
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If Val(txtCantidad.Text) > "0" Then
            MessageBox.Show("DEBE ASIGNAR TODA LA RESERVA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        'ASIGNAR LA RESERVA AL MISMO OPERARIO
        ValidaOperario(grdReserva(2, [Y]).Value.ToString)
        If Valida = False Then

            clase.consultar("SELECT * FROM det_asignacion_orden_produccion WHERE acp_id='" & IdAsig & "'", "asignado")
            Dim Total As String = (Val(grdReserva(1, [Y]).Value.ToString) + Val(clase.dt.Tables("asignado").Rows(0)("acp_cantidad")))
            'ACTUALIZAR DET_ASIGNACION_ORDEN_PRODUCCION
            clase.actualizar("UPDATE det_asignacion_orden_produccion SET acp_cantidad='" & Total & "' WHERE acp_id='" & IdAsig & "'")

            'INSERTAR EN DET_ASIGNACION_TIENDAS
            clase.agregar_registro2("INSERT INTO det_asignacion_tiendas(as_asigorden,as_tienda,as_cantidad) VALUES('" & IdAsig & "','5000','" & grdReserva(1, [Y]).Value.ToString & "')")

            MessageBox.Show("RESERVA GUARDADA CON EXITO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            frmAsignar.Enabled = True
            frmAsignar.LlenarGridOrden()
            frmAsignar.LlenarTodos()
            Me.Close()
        End If

        If MessageBox.Show("DESEA GUARDAR LA ASIGNACION DE LA RESERVA?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            For i = 0 To grdReserva.RowCount - 1
                clase.agregar_registro("INSERT INTO det_asignacion_orden_produccion(aop_articulo,acp_cantidad,acp_ordenproduccion,acp_operario,acp_fecha_inicio,acp_hora_inicio,acp_estado,acp_danado) VALUES('" & frmAsignar.Codigo & "','" & grdReserva(1, [i]).Value.ToString & "','" & frmAsignar.Orden & "','" & grdReserva(2, [i]).Value.ToString & "','" & clase.FormatoFecha(Date.Today) & "','" & clase.FormatoHora(TimeOfDay) & "','A','0')")

                clase.consultar("SELECT MAX(acp_id) as consecutivo FROM det_asignacion_orden_produccion", "consecutivo")
                Dim Consecutivo As String = clase.dt.Tables("consecutivo").Rows(0)("consecutivo")

                clase.agregar_registro2("INSERT INTO det_asignacion_tiendas(as_asigorden,as_tienda,as_cantidad) VALUES('" & Consecutivo & "','5000','" & grdReserva(1, [i]).Value.ToString & "')")
            Next
            MessageBox.Show("RESERVA GUARDADA CON EXITO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            frmAsignar.Enabled = True
            frmAsignar.LlenarGridOrden()
            frmAsignar.LlenarTodos()
            Me.Close()
        End If
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        '  frmAsignar.Enabled = True
        Me.Close()
    End Sub

    Public Sub LlenarGridPatron()
        grdPatron.DataSource = Nothing
        clase.consultar("SELECT tiendas.id AS ID, tiendas.tienda AS TIENDA, detpatrondist.dp_cantidad AS CANT FROM deordenprod INNER JOIN cabpatrondist ON (deordenprod.do_patrondist = cabpatrondist.pat_codigo) INNER JOIN detpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo) INNER JOIN tiendas ON (detpatrondist.dp_tienda = tiendas.id) WHERE (deordenprod.do_idcaborden ='" & frmAsignar.Orden & "' AND deordenprod.do_codigo ='" & frmAsignar.Codigo & "');", "patron")
        grdPatron.DataSource = clase.dt.Tables("patron")
        clase.dt.Tables("patron").Columns.Add("SEL", GetType(Boolean))

        'QUITAR LAS TIENDAS YA ASIGNADAS
        clase.consultar1("SELECT det_asignacion_orden_produccion.acp_ordenproduccion, det_asignacion_orden_produccion.aop_articulo, det_asignacion_tiendas.as_tienda AS TIENDA FROM det_asignacion_tiendas INNER JOIN det_asignacion_orden_produccion ON (det_asignacion_tiendas.as_asigorden = det_asignacion_orden_produccion.acp_id) WHERE (det_asignacion_orden_produccion.acp_ordenproduccion ='" & frmAsignar.Orden & "' AND det_asignacion_orden_produccion.aop_articulo ='" & frmAsignar.Codigo & "');", "asignados")
        CantReserva = Val(SumOrden) - Val(SumPatron)
        If CantReserva <> 0 Then
            clase.dt.Tables("patron").Rows.Add("5000", "RESERVA", CantReserva)
        End If

        For i = 0 To clase.dt1.Tables("asignados").Rows.Count - 1
            For a = 0 To grdPatron.RowCount - 1
                If grdPatron(0, [a]).Value.ToString = clase.dt1.Tables("asignados").Rows(i)("TIENDA").ToString Then
                    Me.grdPatron.CurrentCell = Nothing
                    Me.grdPatron.Rows(a).Visible = False
                End If
            Next
        Next

        grdPatron.Columns(0).Visible = False
        grdPatron.Columns(1).ReadOnly = True
        grdPatron.Columns(1).Width = 200
        grdPatron.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdPatron.Columns(2).Width = 70
        grdPatron.Columns(2).ReadOnly = True
        grdPatron.Columns(3).Width = 40
        grdPatron.RowHeadersWidth = 4
    End Sub

    Private Sub grdPatron_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdPatron.CellClick
        Dim A As Integer = grdPatron.CurrentCell.RowIndex
        Tienda = grdPatron(0, [A]).Value.ToString
        Sel = grdPatron(3, [A]).Value.ToString
        If grdPatron(0, [A]).Value.ToString = "5000" Then
            If MessageBox.Show("DESEA ASIGNAR LA RESERVA", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                grdPatron.Visible = False
                btnAsignar.Visible = False
                btnGuardar.Visible = True
                grdReserva.Visible = True
                btnDeshacer.Visible = True
                btnAsigReserva.Visible = True
                txtCantidad.Text = CantReserva
            End If
        End If
    End Sub

    Private Sub grdPatron_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles grdPatron.CurrentCellDirtyStateChanged
        If grdPatron.IsCurrentCellDirty Then
            grdPatron.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub grdPatron_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles grdPatron.CellValueChanged
        Try
            Dim A As Integer = grdPatron.CurrentCell.RowIndex
            If grdPatron.CurrentCell.Value = True Then
                txtCantidad.Text = Val(txtCantidad.Text) + Val(grdPatron(2, [A]).Value.ToString)
            ElseIf grdPatron.CurrentCell.Value = False Then
                txtCantidad.Text = Val(txtCantidad.Text) - Val(grdPatron(2, [A]).Value.ToString)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtReserva_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtReserva.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub chkTodos_CheckedChanged(sender As Object, e As EventArgs) Handles chkTodos.CheckedChanged
        Dim AsigReserva As Boolean = False
        If chkTodos.Checked = True Then
            If MessageBox.Show("DESEA ASIGNAR LA RESERVA", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
                AsigReserva = True
            End If
        End If

        txtCantidad.Text = ""
        grdPatron.CurrentCell = Nothing
        If chkTodos.Checked = True Then
            For i = 0 To grdPatron.RowCount - 1
                If grdPatron.Rows(i).Visible = True Then
                    If AsigReserva = True Then
                        grdPatron(3, [i]).Value = True
                        txtCantidad.Text = (Val(grdPatron(2, [i]).Value.ToString) + Val(txtCantidad.Text))
                    ElseIf AsigReserva = False Then
                        If grdPatron(1, [i]).Value <> "RESERVA" Then
                            grdPatron(3, [i]).Value = True
                            txtCantidad.Text = (Val(grdPatron(2, [i]).Value.ToString) + Val(txtCantidad.Text))
                        End If
                    End If
                End If
            Next
        Else
            LlenarGridPatron()
        End If
    End Sub

    Public Sub Reserva(Cantidad As String)
        Dim Asignacion As DataRow = TablaReserva.NewRow
        Asignacion("TIENDA") = "RESERVA"
        Asignacion("CANT") = Cantidad
        Asignacion("OPERARIO") = txtOperario.Text
        TablaReserva.Rows.Add(Asignacion)
        grdReserva.DataSource = TablaReserva
        grdReserva.RowHeadersWidth = 4
    End Sub

    Public Sub ValidaOperario(Operario As String)
        'CONSULTAR SI EL OPERARIO ESTA DISPONIBLE
        clase.consultar("SELECT * FROM det_asignacion_orden_produccion WHERE (acp_operario='" & Operario & "' AND acp_fecha_inicio<>'' AND acp_fecha_fin IS NULL)", "operario")
        If clase.dt.Tables("operario").Rows.Count > 0 Then
            IdAsig = clase.dt.Tables("operario").Rows(0)("acp_id")
            Valida = False
        End If
    End Sub

    Private Sub btnDeshacer_Click(sender As Object, e As EventArgs) Handles btnDeshacer.Click
        grdPatron.Visible = True
        btnAsignar.Visible = True
        btnGuardar.Visible = False
        grdReserva.Visible = False
        TablaReserva.Clear()
        btnAsigReserva.Visible = False
        btnDeshacer.Visible = False
        txtCantidad.Text = ""
        LlenarGridPatron()
        SumAsigReserva = 0
    End Sub

    Private Sub grdReserva_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdReserva.CellClick
        Y = grdPatron.CurrentCell.RowIndex
    End Sub

End Class