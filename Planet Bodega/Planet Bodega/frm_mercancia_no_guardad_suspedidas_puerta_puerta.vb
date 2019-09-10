Public Class frm_mercancia_no_guardad_suspedidas_puerta_puerta
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_mercancia_no_guardad_suspedidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT cabmer_puertapuerta.cabpuerta_codigo, cabimportacion.imp_nombrefecha, cabmer_puertapuerta.cabpuerta_operario, SUM(detmer_puertapuerta.detpuerta_cantidad) Cant, cabmer_puertapuerta.cabpuerta_procesada, cabmer_puertapuerta.cabpuerta_importacion FROM detmer_puertapuerta INNER JOIN cabmer_puertapuerta ON (detmer_puertapuerta.detpuerta_codigopuerta = cabmer_puertapuerta.cabpuerta_codigo) INNER JOIN cabimportacion  ON (cabmer_puertapuerta.cabpuerta_importacion = cabimportacion.imp_codigo) GROUP BY cabmer_puertapuerta.cabpuerta_codigo ORDER BY cabmer_puertapuerta.cabpuerta_codigo DESC", "ingreso")

        'clase.consultar("SELECT cabmer_noimportada.cabmer_codigo, cabimportacion.imp_nombrefecha, proveedores.prv_codigoasignado, cabmer_noimportada.cabmer_operario, SUM(detmer_noimportada.detmer_cantidad) Cant, cabmer_noimportada.cabmer_procesada, cabmer_noimportada.cabmer_proveedor, cabmer_noimportada.cabmer_importacion FROM detmer_noimportada INNER JOIN cabmer_noimportada ON (detmer_noimportada.detmer_codigo_imp = cabmer_noimportada.cabmer_codigo) INNER JOIN cabimportacion  ON (cabmer_noimportada.cabmer_importacion = cabimportacion.imp_codigo) INNER JOIN proveedores  ON (cabmer_noimportada.cabmer_proveedor = proveedores.prv_codigo) GROUP BY cabmer_noimportada.cabmer_codigo ORDER BY cabmer_noimportada.cabmer_codigo DESC", "ingreso")
        If clase.dt.Tables("ingreso").Rows.Count > 0 Then
            grdArticulo.DataSource = clase.dt.Tables("ingreso")
            prepara_columnas()
        Else
            grdArticulo.DataSource = Nothing
            grdArticulo.ColumnCount = 8
            prepara_columnas()
        End If
    End Sub

    Private Sub prepara_columnas()
        With grdArticulo
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Visible = False
            .Columns(0).HeaderText = "Movimiento"
            .Columns(1).HeaderText = "Importación"
            .Columns(1).Width = 180
            '.Columns(2).HeaderText = "Proveedores"
            '.Columns(2).Width = 120
            .Columns(2).HeaderText = "Operario"
            .Columns(2).Width = 150
            .Columns(3).HeaderText = "Cant"
            .Columns(3).Width = 70
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).HeaderText = ""
            .Columns(4).Width = 20
            .Columns(5).Visible = False
            .Columns(6).Visible = False
        End With
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If grdArticulo.RowCount = 0 Then
            Exit Sub
        End If
        Try
            If IsNumeric(grdArticulo.CurrentRow.Index) Then
                frm_ingreso_puerta_puerta_puerta_puerta.txtImportacion.Text = grdArticulo.Item(1, grdArticulo.CurrentRow.Index).Value
                frm_ingreso_puerta_puerta_puerta_puerta.ConsImport = grdArticulo.Item(5, grdArticulo.CurrentRow.Index).Value
                frm_ingreso_puerta_puerta_puerta_puerta.txtOperario.Text = grdArticulo.Item(2, grdArticulo.CurrentRow.Index).Value
                frm_ingreso_puerta_puerta_puerta_puerta.llenar_grilla_mercancia(grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value)
                frm_ingreso_puerta_puerta_puerta_puerta.btnImportacion.Enabled = False
                frm_ingreso_puerta_puerta_puerta_puerta.txtOperario.Enabled = False
                frm_ingreso_puerta_puerta_puerta_puerta.txtArticulo.Enabled = True
                frm_ingreso_puerta_puerta_puerta_puerta.Button1.Enabled = False
                frm_ingreso_puerta_puerta_puerta_puerta.btnDeshacer.Enabled = True
                If grdArticulo.Item(4, grdArticulo.CurrentRow.Index).Value = True Then
                    frm_ingreso_puerta_puerta_puerta_puerta.btnGuardar.Enabled = False
                    frm_ingreso_puerta_puerta_puerta_puerta.btnDeshacer.Enabled = False
                    frm_ingreso_puerta_puerta_puerta_puerta.button5.Enabled = False
                    frm_ingreso_puerta_puerta_puerta_puerta.txtConsMov.Text = ""

                Else
                    frm_ingreso_puerta_puerta_puerta_puerta.btnGuardar.Enabled = True
                    frm_ingreso_puerta_puerta_puerta_puerta.btnDeshacer.Enabled = True
                    frm_ingreso_puerta_puerta_puerta_puerta.cargarconsecutivos() 'txtConsMov
                    frm_ingreso_puerta_puerta_puerta_puerta.codigo_guardado = grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value
                    frm_ingreso_puerta_puerta_puerta_puerta.button5.Enabled = True
                End If
                Me.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Debe seleccionar un movimiento, gracias.", "Plane Love", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
End Class