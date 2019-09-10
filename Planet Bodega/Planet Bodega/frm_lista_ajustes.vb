Public Class frm_lista_ajustes
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_lista_ajustes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bodegaalmacenaje = vbEmpty
        clase.consultar("SELECT tipos_ajustes.tip_nombre, bodegas.bod_nombre, cabajuste.cabaj_fecha, cabajuste.cabaj_operario, cabajuste.cabaj_codigo FROM cabajuste INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) INNER JOIN bodegas ON (cabajuste.cabaj_bodega = bodegas.bod_codigo) WHERE (cabajuste.cabaj_procesado =FALSE) ORDER BY cabajuste.cabaj_codigo DESC", "rest")
        If clase.dt.Tables("rest").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("rest")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 4
            preparar_columnas()
        End If
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 120
            .Columns(1).Width = 120
            .Columns(2).Width = 100
            .Columns(3).Width = 200
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).HeaderText = "Tipo Ajuste"
            .Columns(1).HeaderText = "Bodega"
            .Columns(2).HeaderText = "Fecha"
            .Columns(3).HeaderText = "Operario"
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            clase.consultar("SELECT cabaj_procesado FROM cabajuste WHERE (cabaj_codigo =" & DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value & ")", "verificacion")
            If clase.dt.Tables("verificacion").Rows(0)("cabaj_procesado") = True Then
                MessageBox.Show("No se puede modificar un ajuste que ya fue procesado.", "AJUSTE YA PROCESADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            frm_generar_ajuste_para_gondola2.TextBox5.Text = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
            Select Case tipo_ajuste(DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value)

                Case "A"
                    frm_generar_ajuste_para_gondola2.ComboBox1.SelectedIndex = -1
                    frm_generar_ajuste_para_gondola2.ComboBox1.Enabled = True
                Case "S"
                    frm_generar_ajuste_para_gondola2.ComboBox1.SelectedIndex = 0
                    frm_generar_ajuste_para_gondola2.ComboBox1.Enabled = False
                Case "R"
                    frm_generar_ajuste_para_gondola2.ComboBox1.SelectedIndex = 1
                    frm_generar_ajuste_para_gondola2.ComboBox1.Enabled = False
            End Select

            bodegaalmacenaje = bodega_ajuste(DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value)
            Me.Close()
        End If
    End Sub

    Function tipo_ajuste(ajuste As Short) As String
        clase.consultar("SELECT tipos_ajustes.tip_tipo FROM cabajuste INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE (cabajuste.cabaj_codigo =" & ajuste & ")", "tipo")
        If clase.dt.Tables("tipo").Rows.Count > 0 Then
            Return clase.dt.Tables("tipo").Rows(0)("tip_tipo")
        Else
            Return ""
        End If
    End Function

    Function bodega_ajuste(ajuste As Short) As Short
        clase.consultar("SELECT bodegas.bod_codigo FROM cabajuste INNER JOIN bodegas ON (cabajuste.cabaj_bodega = bodegas.bod_codigo) WHERE (cabajuste.cabaj_codigo =" & ajuste & ")", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("bod_codigo")
    End Function
End Class