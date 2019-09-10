Public Class frm_codigos_asociados_inventario
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_codigos_asociados_inventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With frm_inventario.DataGridView1
            clase.dt_global.Tables.Clear()
            clase.consultar_global(buscar_codigos_a_apartir_de_la_referencia(.Item(1, .CurrentRow.Index).Value, .Item(2, .CurrentRow.Index).Value, .Item(6, .CurrentRow.Index).Value, .Item(7, .CurrentRow.Index).Value, .Item(8, .CurrentRow.Index).Value, .Item(9, .CurrentRow.Index).Value, .Item(10, .CurrentRow.Index).Value), "tabla4")

            If clase.dt_global.Tables("tabla4").Rows.Count > 0 Then
                Dim x As Short
                For x = 0 To clase.dt_global.Tables("tabla4").Rows.Count - 1
                    dataGridView1.RowCount = dataGridView1.RowCount + 1
                    dataGridView1.Item(0, x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_codigo")
                    dataGridView1.Item(1, x).Value = buscar_talla_color(clase.dt_global.Tables("tabla4").Rows(x)("ar_codigo"))
                    dataGridView1.Item(2, x).Value = existencia_codigo_en_gondola(clase.dt_global.Tables("tabla4").Rows(x)("ar_codigo"), .Item(11, .CurrentRow.Index).Value)
                Next
            End If
        End With
    End Sub

    Function buscar_talla_color(codigoarticulo As Integer) As String
        clase.consultar("SELECT tallas.nombretalla, colores.colornombre FROM tallas INNER JOIN articulos ON (tallas.codigo_talla = articulos.ar_talla) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) WHERE (articulos.ar_codigo =" & codigoarticulo & ")", "tabla5")
        If clase.dt.Tables("tabla5").Rows.Count > 0 Then
            Return clase.dt.Tables("tabla5").Rows(0)("nombretalla") & " - " & clase.dt.Tables("tabla5").Rows(0)("colornombre")
        Else
            Return ""
        End If
    End Function

    Function existencia_codigo_en_gondola(codigo As Integer, gondola As String) As Integer
        clase.consultar("SELECT inv_cantidad FROM inventario_bodega WHERE (inv_gondola ='" & gondola & "' AND inv_codigoart =" & codigo & ")", "tabla6")
        If clase.dt.Tables("tabla6").Rows.Count > 0 Then
            Return clase.dt.Tables("tabla6").Rows(0)("inv_cantidad")
        Else
            Return 0
        End If
    End Function
End Class