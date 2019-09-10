Public Class frm_elegir_linea_sublinea
    Dim clase As New class_library
    Dim ind_carga As Boolean
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.Rows.Count = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            With DataGridView1
                cadena_string = "select ar_codigo from articulos where ar_referencia = '" & UCase(frm_ajustes.TextBox1.Text) & "' AND ar_linea = " & .Item(0, .CurrentCell.RowIndex).Value & " AND ar_sublinea1 = " & .Item(1, .CurrentCell.RowIndex).Value & " " & generar_cadena_sql_apartir_del_combobox_de_coincidencias(verificar_nulidad_vacio(.Item(2, .CurrentCell.RowIndex).Value), verificar_nulidad_vacio(.Item(3, .CurrentCell.RowIndex).Value), verificar_nulidad_vacio(.Item(4, .CurrentCell.RowIndex).Value))
            End With
            Me.Close()
        End If
    End Sub

    Private Sub frm_elegir_linea_sublinea_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cadena_string = ""
        clase.consultar("select* from articulos where ar_referencia = '" & UCase(frm_ajustes.TextBox1.Text) & "'", "result")
        If clase.dt.Tables("result").Rows.Count > 0 Then

            clase.consultar("SELECT DISTINCT articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4, CONCAT(linea1.ln1_nombre, ', ', sublinea1.sl1_nombre) AS refart FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (sublinea2.sl2_sl1codigo = sublinea1.sl1_codigo) AND (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) AND (sublinea1.sl1_ln1codigo = linea1.ln1_codigo) LEFT JOIN sublinea3 ON (sublinea3.sl3_sl2codigo = sublinea2.sl2_codigo) AND (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (sublinea4.sl4_sl3codigo = sublinea3.sl3_codigo) AND (articulos.ar_sublinea4 = sublinea4.sl4_codigo) WHERE (articulos.ar_referencia = '" & UCase(frm_ajustes.TextBox1.Text) & "')", "tableresult11")
            If clase.dt.Tables("tableresult11").Rows.Count > 0 Then
                DataGridView1.DataSource = clase.dt.Tables("tableresult11")
                With DataGridView1
                    .Columns(0).Width = 1
                    .Columns(1).Width = 1
                    .Columns(2).Width = 1
                    .Columns(3).Width = 1
                    .Columns(4).Width = 1
                    .Columns(5).Width = 400
                    .Columns(5).HeaderText = "Articulo"
                End With
            Else
                DataGridView1.DataSource = Nothing
            End If
        End If
    End Sub
End Class