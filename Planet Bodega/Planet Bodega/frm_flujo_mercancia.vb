Public Class frm_flujo_mercancia
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_flujo_mercancia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With dtgridcajas
            .ColumnCount = 5
            .Columns(0).HeaderText = "Movimiento"
            .Columns(1).HeaderText = "Fecha"
            .Columns(2).HeaderText = "Liquidador"
            .Columns(3).HeaderText = "Estado"
            .Columns(4).HeaderText = "Items Procesados"
            .Columns(0).Width = 80
            .Columns(1).Width = 100
            .Columns(2).Width = 250
            .Columns(3).Width = 80
            .Columns(4).Width = 80
        End With
    End Sub

    Private Sub llenar_lista(inicial As String, final As String)
        clase.dt_global.Tables.Clear()
        clase.consultar_global("SELECT cabsal_cod AS Movimiento, cabsal_fecha AS Fecha, cabsal_liquidador AS Liquidador, calsal_estado AS Estado FROM  cabsalidas_mercancia WHERE (cabsal_fecha BETWEEN '" & inicial & "' AND '" & final & "') ORDER BY cabsal_fecha ASC", "tabla")
        If clase.dt_global.Tables("tabla").Rows.Count > 0 Then
            Dim b As Short
            dtgridcajas.RowCount = 0
            With dtgridcajas
                For b = 0 To clase.dt_global.Tables("tabla").Rows.Count - 1
                    .RowCount = .RowCount + 1
                    .Item(0, b).Value = clase.dt_global.Tables("tabla").Rows(b)("Movimiento")
                    Dim fecha As Date = clase.dt_global.Tables("tabla").Rows(b)("Fecha")
                    .Item(1, b).Value = fecha.ToString("dd/MM/yyyy")
                    .Item(2, b).Value = clase.dt_global.Tables("tabla").Rows(b)("Liquidador")
                    .Item(3, b).Value = clase.dt_global.Tables("tabla").Rows(b)("Estado")
                    .Item(4, b).Value = buscar_items_procesados(clase.dt_global.Tables("tabla").Rows(b)("Movimiento"))
                Next
            End With
        Else
            dtgridcajas.RowCount = 0
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        llenar_lista(DateTimePicker1.Value.ToString("yyyy-MM-dd"), DateTimePicker2.Value.ToString("yyyy-MM-dd"))
    End Sub

    Function buscar_items_procesados(codmov As Integer) As Integer
        clase.consultar("SELECT COUNT(*) AS cantidad FROM detsalidas_mercancia WHERE (det_salidacodigo =" & codmov & " AND det_procesado = TRUE)", "restab")
        If clase.dt.Tables("restab").Rows.Count > 0 Then
            Return clase.dt.Tables("restab").Rows(0)("cantidad")
        End If
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If dtgridcajas.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dtgridcajas.CurrentRow.Index) Then
            clase.consultar("SELECT COUNT(*) AS cantidad FROM detsalidas_mercancia WHERE (det_salidacodigo =" & dtgridcajas.Item(0, dtgridcajas.CurrentRow.Index).Value & " AND det_procesado = TRUE)", "restab")
            If Val(clase.dt.Tables("restab").Rows(0)("cantidad")) > 0 Then
                MessageBox.Show("No se puede eliminar el movimiento porque ya hay items procesados.", "NO SE PUEDE ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim v As String = MessageBox.Show("¿Desea Borrar el movimiento seleccionado?", "BORRAR MOVIMIENTO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If v = 6 Then
                    clase.consultar_global("SELECT* FROM detsalidas_mercancia WHERE (det_salidacodigo =" & dtgridcajas.Item(0, dtgridcajas.CurrentRow.Index).Value & ")", "tbl")
                    If clase.dt_global.Tables("tbl").Rows.Count > 0 Then
                        Dim x As Short
                        For x = 0 To clase.dt_global.Tables("tbl").Rows.Count - 1
                            clase.consultar("SELECT detcab_cantidad FROM detalle_importacion_detcajas WHERE (detcab_coditem =" & clase.dt_global.Tables("tbl").Rows(x)("det_codref") & ")", "123")
                            Dim cantidad As Integer = clase.dt.Tables("123").Rows(0)("detcab_cantidad")
                            clase.actualizar("UPDATE detalle_importacion_detcajas SET detcab_cantidad = " & Val(cantidad + clase.dt_global.Tables("tbl").Rows(x)("det_cant")) & " WHERE detcab_coditem = " & clase.dt_global.Tables("tbl").Rows(x)("det_codref") & "")
                        Next x
                        clase.borradoautomatico("DELETE from cabsalidas_mercancia WHERE cabsal_cod =" & dtgridcajas.Item(0, dtgridcajas.CurrentRow.Index).Value & "")
                        llenar_lista(DateTimePicker1.Value.ToString("yyyy-MM-dd"), DateTimePicker2.Value.ToString("yyyy-MM-dd"))
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If dtgridcajas.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dtgridcajas.CurrentRow.Index) Then
            frm_ver_salida.ShowDialog()
            frm_ver_salida.Dispose()
        End If
    End Sub
End Class