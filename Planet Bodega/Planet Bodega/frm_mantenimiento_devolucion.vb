Public Class frm_mantenimiento_devolucion
    Dim clase As New class_library

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frm_mantenimiento_de_transfencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("select* from tiendas order by tienda asc", "tiendas")
        ComboBox9.Items.Clear()
        Dim x As Short
        ComboBox9.Items.Add("TODAS")
        For x = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
            ComboBox9.Items.Add(clase.dt.Tables("tiendas").Rows(x)("tienda"))
        Next
        'clase.llenar_combo(ComboBox9, "select* from tiendas order by tienda asc", "tienda", "id")
        ComboBox1.Text = "TODAS"
        DataGridView2.ColumnCount = 11
        preparar_columnas()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If clase.validar_combobox(ComboBox9, "Tiendas") = False Then Exit Sub
        Dim sql As String = ""
        Dim fecha1 As Date = DateTimePicker1.Value
        Dim fecha2 As Date = DateTimePicker2.Value
        If ComboBox9.Text = "TODAS" Then
            Select Case ComboBox1.Text
                Case "Todas"
                    sql = "SELECT cabtransferencia.tr_numero, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado  FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE (cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "' AND (tr_opproduccion IS NULL) AND (tr_importacion IS NULL) AND (tr_operador IS NOT NULL)) GROUP BY dettransferencia.dt_trnumero"
                Case "Suspendidas"
                    sql = "SELECT cabtransferencia.tr_numero, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_estado =False) AND (tr_importacion IS NULL) AND(tr_opproduccion IS NULL) AND (tr_operador IS NOT NULL)) GROUP BY dettransferencia.dt_trnumero"
                Case "Cerradas"
                    sql = "SELECT cabtransferencia.tr_numero, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_estado =True) AND (tr_importacion IS NULL) AND (tr_opproduccion IS NULL) AND (tr_operador IS NOT NULL)) GROUP BY dettransferencia.dt_trnumero"
            End Select
        Else
            Select Case ComboBox1.Text
                Case "Todas"
                    sql = "SELECT cabtransferencia.tr_numero, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_destino = " & codigo_tienda(ComboBox9.Text) & ") AND (tr_importacion IS NULL) AND (tr_opproduccion IS NULL) AND (tr_operador IS NOT NULL)) GROUP BY dettransferencia.dt_trnumero"
                Case "Suspendidas"
                    sql = "SELECT cabtransferencia.tr_numero, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_destino = " & codigo_tienda(ComboBox9.Text) & ") AND (tr_importacion IS NULL) AND (cabtransferencia.tr_estado =False) AND (tr_opproduccion IS NULL) AND (tr_operador IS NOT NULL)) GROUP BY dettransferencia.dt_trnumero"
                Case "Cerradas"
                    sql = "SELECT cabtransferencia.tr_numero, tiendas.tienda, bodegas.bod_nombre, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS Costo, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS Venta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS Venta2, SUM(dt_cantidad),  cabtransferencia.tr_estado FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) LEFT JOIN bodegas ON (cabtransferencia.tr_bodega = bodegas.bod_codigo) WHERE ((cabtransferencia.tr_fecha BETWEEN '" & fecha1.ToString("yyyy-MM-dd") & "' AND '" & fecha2.ToString("yyyy-MM-dd") & "') AND (cabtransferencia.tr_destino = " & codigo_tienda(ComboBox9.Text) & ") AND (tr_importacion IS NULL) AND (cabtransferencia.tr_estado =True) AND (tr_opproduccion IS NULL) AND (tr_operador IS NOT NULL)) GROUP BY dettransferencia.dt_trnumero"
            End Select
        End If
        clase.consultar(sql, "tabla3")
        If clase.dt.Tables("tabla3").Rows.Count > 0 Then
            DataGridView2.Columns.Clear()
            DataGridView2.DataSource = clase.dt.Tables("tabla3")
            preparar_columnas()
        Else
            DataGridView2.DataSource = Nothing
            DataGridView2.ColumnCount = 11
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
            .Columns(1).HeaderText = "Destino"
            .Columns(2).HeaderText = "Bodega"
            .Columns(3).HeaderText = "Fecha"
            .Columns(4).HeaderText = "Operador"
            .Columns(5).HeaderText = "Revisor"
            .Columns(6).HeaderText = "Precio Costo"
            .Columns(7).HeaderText = "Precio Venta1"
            .Columns(8).HeaderText = "Precio Venta2"
            .Columns(9).HeaderText = "Unidades"
            .Columns(10).HeaderText = "Estado"
            .Columns(0).Width = 70
            .Columns(1).Width = 120
            .Columns(2).Width = 130
            .Columns(3).Width = 80
            .Columns(4).Width = 150
            .Columns(5).Width = 150
            .Columns(6).Width = 100
            .Columns(7).Width = 100
            .Columns(8).Width = 100
            .Columns(9).Width = 60
            .Columns(10).Width = 60
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            Dim v As String = MessageBox.Show("¿Desea Imprimir la transferencia seleccionada?", "IMPRIMIR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                imprimir_hoja_transferencia(DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value)
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        clase.consultar("SELECT tr_revisada FROM cabtransferencia WHERE (tr_numero =" & DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value & ")", "transferencia")

        construir_archivo(DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value)
        MessageBox.Show("Mensaje enviado con exito", "Planet Love", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            frm_detalle_devolucion.ShowDialog()
            frm_detalle_devolucion.Dispose()
        End If
    End Sub
End Class