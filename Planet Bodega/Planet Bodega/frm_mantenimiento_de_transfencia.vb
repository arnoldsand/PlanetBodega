Public Class frm_mantenimiento_de_transfencia
    Dim clase As New class_library
   
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frm_mantenimiento_de_transfencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("select* from tiendas where estado = TRUE order by tienda asc", "tiendas")
        ComboBox9.Items.Clear()
        Dim x As Short
        ComboBox9.Items.Add("TODAS")
        For x = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
            ComboBox9.Items.Add(clase.dt.Tables("tiendas").Rows(x)("tienda"))
        Next
        'clase.llenar_combo(ComboBox9, "select* from tiendas order by tienda asc", "tienda", "id")
        ComboBox1.Text = "TODAS"
        DataGridView2.ColumnCount = 13
        preparar_columnas()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If clase.validar_combobox(ComboBox9, "Tiendas") = False Then Exit Sub
        LlenarListado()
    End Sub

    Private Sub LlenarListado()
        Dim sql As String = ""
        Dim fecha1 As Date = DateTimePicker1.Value
        Dim fecha2 As Date = DateTimePicker2.Value
        If ComboBox9.Text = "TODAS" Then
            Select Case ComboBox1.Text
                Case "Todas"
                    sql = "SELECT cabtransferencia.tr_numero, cabtransferencia.tr_codigolovepos, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado, cabtransferencia.tr_finalizada FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE (cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') GROUP BY dettransferencia.dt_trnumero"
                Case "Suspendidas"
                    sql = "SELECT cabtransferencia.tr_numero, cabtransferencia.tr_codigolovepos, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado, cabtransferencia.tr_finalizada FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_estado =False)) GROUP BY dettransferencia.dt_trnumero"
                Case "Cerradas"
                    sql = "SELECT cabtransferencia.tr_numero, cabtransferencia.tr_codigolovepos, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado, cabtransferencia.tr_finalizada FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_estado =True)) GROUP BY dettransferencia.dt_trnumero"
            End Select
        Else
            Select Case ComboBox1.Text
                Case "Todas"
                    sql = "SELECT cabtransferencia.tr_numero, cabtransferencia.tr_codigolovepos, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado, cabtransferencia.tr_finalizada FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_destino = " & codigo_tienda(ComboBox9.Text) & ")) GROUP BY dettransferencia.dt_trnumero"
                Case "Suspendidas"
                    sql = "SELECT cabtransferencia.tr_numero, cabtransferencia.tr_codigolovepos, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado, cabtransferencia.tr_finalizada FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_destino = " & codigo_tienda(ComboBox9.Text) & ") AND (cabtransferencia.tr_estado =False)) GROUP BY dettransferencia.dt_trnumero"
                Case "Cerradas"
                    sql = "SELECT cabtransferencia.tr_numero, cabtransferencia.tr_codigolovepos, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado, cabtransferencia.tr_finalizada FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_destino = " & codigo_tienda(ComboBox9.Text) & ") AND (cabtransferencia.tr_estado =True)) GROUP BY dettransferencia.dt_trnumero"
            End Select
        End If
        clase.consultar(sql, "tabla3")
        If clase.dt.Tables("tabla3").Rows.Count > 0 Then
            DataGridView2.Columns.Clear()
            DataGridView2.DataSource = clase.dt.Tables("tabla3")
            preparar_columnas()
        Else
            DataGridView2.DataSource = Nothing
            DataGridView2.ColumnCount = 13
            preparar_columnas()
            MessageBox.Show("No se encontraron transferencias con los criterios de busqueda especficados.", "NO SE ENCONTRARON TRANSFERENCIAS", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Function codigo_tienda(tienda As String) As Short
        clase.consultar("select* from tiendas where tienda = '" & tienda & "'", "result")
        Return clase.dt.Tables("result").Rows(0)("id")
    End Function

    Sub preparar_columnas()
        With DataGridView2
            .Columns(0).HeaderText = "Numero"
            .Columns(1).HeaderText = "Codigo LovePOS"
            .Columns(2).HeaderText = "Destino"
            .Columns(3).HeaderText = "Bodega"
            .Columns(4).HeaderText = "Fecha"
            .Columns(5).HeaderText = "Operador"
            .Columns(6).HeaderText = "Revisor"
            .Columns(7).HeaderText = "Precio Costo"
            .Columns(8).HeaderText = "Precio Venta1"
            .Columns(9).HeaderText = "Precio Venta2"
            .Columns(10).HeaderText = "Unidades"
            .Columns(11).HeaderText = "Estado"
            .Columns(12).HeaderText = "Finalizada"

            .Columns(0).Width = 70
            .Columns(1).Width = 120
            .Columns(2).Width = 120
            .Columns(3).Width = 130
            .Columns(4).Width = 80
            .Columns(5).Width = 150
            .Columns(6).Width = 150
            .Columns(7).Width = 100
            .Columns(8).Width = 100
            .Columns(9).Width = 100
            .Columns(10).Width = 60
            .Columns(11).Width = 60
            .Columns(12).Width = 60
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            Dim v As String = MessageBox.Show("¿Desea Imprimir la transferencia seleccionada?", "IMPRIMIR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                clase.consultar("SELECT tr_finalizada FROM cabtransferencia WHERE (tr_numero =" & DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value & ")", "transferencia")
                If IsDBNull(clase.dt.Tables("transferencia").Rows(0)("tr_finalizada")) Then
                    MessageBox.Show("No se puede imprimir una transferencia  que no se ha revisado.", "NO SE PUEDE IMPRIMIR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                Else
                    If clase.dt.Tables("transferencia").Rows(0)("tr_finalizada") = False Then
                        MessageBox.Show("No se puede imprimir una transferencia  que no se ha revisado.", "NO SE PUEDE IMPRIMIR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    imprimir_hoja_transferencia(DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value)
                End If
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Dim v As String = MessageBox.Show("¿Desea enviar el archivo de transferencia por correo?", "ENVIAR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            clase.consultar("SELECT tr_finalizada FROM cabtransferencia WHERE (tr_numero =" & DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value & ")", "transferencia")
            If IsDBNull(clase.dt.Tables("transferencia").Rows(0)("tr_finalizada")) Then
                MessageBox.Show("No se puede enviar una transferencia  que no se ha revisado.", "NO SE PUEDE ENVIAR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            Else
                If clase.dt.Tables("transferencia").Rows(0)("tr_finalizada") = False Then
                    MessageBox.Show("No se puede enviar una transferencia  que no se ha revisado.", "NO SE PUEDE ENVIAR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                construir_archivo(DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value)
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            frm_detalle_transferencia.ShowDialog()
            frm_detalle_transferencia.Dispose()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            Dim formulariodestino As New frmdestinotransferencias(DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value)
            formulariodestino.ShowDialog()
            formulariodestino.Dispose()
            LlenarListado()
        End If
    End Sub

    Private Sub btnAnular_Click(sender As Object, e As EventArgs) Handles btnAnular.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            Dim v As DialogResult = MessageBox.Show("¿Desea Anular la Transferencia Seleccionada?", "ANULAR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = DialogResult.Yes Then
                Dim FormularioAnulacion As ContrasenaAnulacion = New ContrasenaAnulacion(DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value)
                FormularioAnulacion.ShowDialog()
                FormularioAnulacion.Dispose()
            End If
            LlenarListado()
        End If
    End Sub
End Class