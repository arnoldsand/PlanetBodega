Public Class frm_mercancia_no_guardad_suspedidas
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_mercancia_no_guardad_suspedidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT cabmer_noimportada.cabmer_codigo, cabimportacion.imp_nombrefecha, cabmer_noimportada.cabmer_operario, SUM(detmer_noimportada.detmer_cantidad) Cant, cabmer_noimportada.cabmer_procesada, cabmer_noimportada.cabmer_importacion FROM detmer_noimportada INNER JOIN cabmer_noimportada ON (detmer_noimportada.detmer_codigo_imp = cabmer_noimportada.cabmer_codigo) INNER JOIN cabimportacion  ON (cabmer_noimportada.cabmer_importacion = cabimportacion.imp_codigo) GROUP BY cabmer_noimportada.cabmer_codigo ORDER BY cabmer_noimportada.cabmer_codigo DESC", "ingreso")
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
            .Columns(0).Visible = False
            .Columns(5).Visible = False
            ' .Columns(6).Visible = False
            .Columns(0).HeaderText = "Movimiento"
            .Columns(1).HeaderText = "Importación"
            '.Columns(2).HeaderText = "Proveedores"
            .Columns(2).HeaderText = "Operario"
            .Columns(3).HeaderText = "Cant"
            .Columns(4).HeaderText = ""
            ' .Columns(0).Width = 80
            .Columns(1).Width = 180
            .Columns(2).Width = 150
            .Columns(3).Width = 70
            .Columns(4).Width = 20
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If grdArticulo.RowCount = 0 Then
            Exit Sub
        End If
        Try
            If IsNumeric(grdArticulo.CurrentRow.Index) Then
                'frm_ingreso_no_importado.txtImportacion.Text = grdArticulo.Item(1, grdArticulo.CurrentRow.Index).Value
                'frm_ingreso_no_importado.ConsImport = grdArticulo.Item(5, grdArticulo.CurrentRow.Index).Value
                'frm_ingreso_no_importado.txtOperario.Text = grdArticulo.Item(2, grdArticulo.CurrentRow.Index).Value
                'frm_ingreso_no_importado.llenar_grilla_mercancia(grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value)
                'frm_ingreso_no_importado.btnImportacion.Enabled = False
                'frm_ingreso_no_importado.txtOperario.Enabled = False
                'frm_ingreso_no_importado.txtArticulo.Enabled = True
                'frm_ingreso_no_importado.Button1.Enabled = False
                'frm_ingreso_no_importado.btnDeshacer.Enabled = True
                If grdArticulo.Item(4, grdArticulo.CurrentRow.Index).Value = True Then
                    frm_ingreso_no_importado.btnGuardar.Enabled = False
                    '  frm_ingreso_no_importado.txtConsMov.Text = grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value
                    frm_ingreso_no_importado.button5.Enabled = False
                    '  frm_ingreso_no_importado.txtConsMov.Text = ""
                Else
                    frm_ingreso_no_importado.btnGuardar.Enabled = True
                    '   frm_ingreso_no_importado.cargarconsecutivos() 'txtConsMov
                    'frm_ingreso_no_importado.codigo_guardado = grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value
                    frm_ingreso_no_importado.button5.Enabled = True
                End If
                Me.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Debe seleccionar un movimiento, gracias.", "Plane Love", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
End Class