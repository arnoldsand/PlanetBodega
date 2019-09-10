Public Class frm_ajustes_importacion
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ajustes_importacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bodegaalmacenaje = vbEmpty
        clase.consultar("SELECT cabajuste.cabaj_codigo, tipos_ajustes.tip_nombre, cabajuste.cabaj_fecha, cabajuste.cabaj_hora, cabajuste.cabaj_operario FROM cabajuste INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE (cabajuste.cabaj_procesado =FALSE AND tipos_ajustes.tip_tipo IS NOT NULL) ORDER BY cabajuste.cabaj_codigo DESC", "rest")
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
            .Columns(0).Width = 50
            .Columns(1).Width = 150
            .Columns(2).Width = 80
            .Columns(3).Width = 50
            .Columns(4).Width = 180
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Tipo Ajuste"
            .Columns(2).HeaderText = "Fecha"
            .Columns(3).HeaderText = "Hora"
            .Columns(4).HeaderText = "Operario"
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            clase.consultar("SELECT cabaj_procesado FROM cabajuste WHERE (cabaj_codigo =" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & ")", "verificacion")
            If clase.dt.Tables("verificacion").Rows(0)("cabaj_procesado") = True Then
                MessageBox.Show("No se puede modificar un ajuste que ya fue procesado.", "AJUSTE YA PROCESADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Select Case tipo_ajuste(DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value)
                Case "A"
                    MessageBox.Show("No se puede anexar la captura a un ajuste abierto, solo se admiten ajustes positivos o negativos.", "AJUSTE INCORRECTO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                Case "S"
                    tipoajuste = 1
                Case "R"
                    tipoajuste = -1
            End Select
            frm_entrada_de_mercancia_desde_hand_held.TextBox6.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
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

End Class