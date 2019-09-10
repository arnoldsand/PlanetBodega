Public Class frm_enviar_actualizacion
    Dim clase As New class_library
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frm_enviar_actualizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With dataGridView1
            .Columns(0).ReadOnly = False
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
        End With
        clase.consultar("SELECT tiendas.tienda, tiendas.ciudad, tiendas.email, empresas.nombre_empresa FROM empresas INNER JOIN tiendas ON (empresas.cod_empresa = tiendas.empresa) ORDER BY tiendas.tienda ASC", "tiendas")
        If clase.dt.Tables("tiendas").Rows.Count > 0 Then
            Dim a As Integer
            For a = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
                With dataGridView1
                    .RowCount = .RowCount + 1
                    .Item(0, a).Value = False
                    .Item(1, a).Value = clase.dt.Tables("tiendas").Rows(a)("tienda")
                    .Item(2, a).Value = clase.dt.Tables("tiendas").Rows(a)("ciudad")
                    .Item(3, a).Value = clase.dt.Tables("tiendas").Rows(a)("nombre_empresa")
                    .Item(4, a).Value = clase.dt.Tables("tiendas").Rows(a)("email")
                End With
            Next
        Else
            dataGridView1.Rows.Clear()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim c As Short
        Dim ind As Boolean = False
        Dim cont As Short = 0
        For c = 0 To dataGridView1.RowCount - 1
            If dataGridView1.Item(0, c).Value = True Then
                ind = True
                cont = cont + 1
            End If
        Next
        If ind = True Then
            Dim consecutivo As Integer = vbEmpty
            Select Case frm_actualizaciones_de_articulo.ComboBox6.SelectedIndex
                Case 0
                    consecutivo = frm_actualizaciones_de_articulo.consecutivo_actulizacion_marca(1)
                Case 1
                    consecutivo = frm_actualizaciones_de_articulo.consecutivo_actulizacion_marca(2)
            End Select
            ProgressBar1.Maximum = cont
            For c = 0 To dataGridView1.RowCount - 1
                If dataGridView1.Item(0, c).Value = True Then
                    enviar_actualizacion_x_correo(dataGridView1.Item(4, c).Value, "ACTUALIZACIÓN DE ARTICULO Y CONDICIÓN " & c, "auxiliarbodega@planetloveonline.com", "ARTICULO: " & consecutivo, "C:\Data\muznovar\MUZNOVAR.TXT")
                    ProgressBar1.Increment(1)
                End If
            Next
            MessageBox.Show("Las actualizaciones se han enviado satisfactoriamente.", "ACTUALIZACIONES ENVIADAS", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Select Case frm_actualizaciones_de_articulo.ComboBox6.SelectedIndex
                Case 0
                    clase.actualizar("UPDATE informacion set consecutivo_muznovar_planetlove = " & consecutivo + 1 & "")
                Case 1
                    clase.actualizar("UPDATE informacion set consecutivo_muznovar_tixy = " & consecutivo + 1 & "")
            End Select
            frm_actualizaciones_de_articulo.reinicar()
            Me.Close()

        Else
            MessageBox.Show("Debe seleccionar por lo menos un punto de venta para el envío de actualizaciones.", "ENVÍO DE ACTUALIZACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class