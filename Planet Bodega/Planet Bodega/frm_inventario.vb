Public Class frm_inventario
    Dim clase As New class_library
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Dim arreglobloque(26, 2) As String

    Dim ind_carga As Boolean
    Private Sub frm_inventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ind_carga = False
        DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        If ind_carga = False Then
            clase.llenar_combo(ComboBox6, "select* from bodegas order by bod_codigo asc", "bod_nombre", "bod_codigo")
        End If
        ind_carga = True
        Dim a As Short
        For a = 0 To 25
            arreglobloque(a, 0) = bloque(a)
            arreglobloque(a, 1) = a + 1
        Next
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_combobox(ComboBox6, "Bodega") = False Then Exit Sub
        If (textBox2.Text.Trim <> "") And (ComboBox1.Text.Trim = "") Then
            If verificar_existencia_gondola(UCase(textBox2.Text.Trim), ComboBox6.SelectedValue) = False Then
                MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                textBox2.Text = ""
                textBox2.Focus()
                Exit Sub
            Else
                DataGridView1.RowCount = 0
                llenar_gondola(textBox2.Text, ComboBox6.SelectedValue)
            End If
        End If
        If (ComboBox1.Text.Trim <> "") And (textBox2.Text.Trim = "") Then
            Dim gondolasarray(0) As String
            'llenar_array_con_gondola_existentes(gondolasarray, ComboBox6.SelectedValue, ComboBox1.Text)
            clase.consultar("SELECT dtbod_cant_gondola FROM detbodega WHERE (dtbod_bloque =" & convertir_letra_en_numero(ComboBox1.Text) & " AND dtbod_bodega =" & ComboBox6.SelectedValue & ")", "tabla7")
            If clase.dt.Tables("tabla7").Rows.Count > 0 Then
                ReDim gondolasarray(clase.dt.Tables("tabla7").Rows(0)("dtbod_cant_gondola"))
                Dim c As Short
                For c = 0 To clase.dt.Tables("tabla7").Rows(0)("dtbod_cant_gondola") - 1
                    gondolasarray(c) = ComboBox1.Text & c + 1
                Next
            Else
                ReDim gondolasarray(0)
            End If
            '________________________
            If gondolasarray.Length = 0 Then
                MessageBox.Show("No hay bloques en esta bodega.", "NO SE ENCONTRARON BLOQUES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            DataGridView1.RowCount = 0
            For c = 0 To clase.dt.Tables("tabla7").Rows(0)("dtbod_cant_gondola") - 1
                llenar_gondola(gondolasarray(c), ComboBox6.SelectedValue)
            Next
        End If
        If (ComboBox1.Text.Trim = "") And (textBox2.Text.Trim = "") Then

            Dim cont As Integer = -1
            clase.consultar("SELECT SUM(dtbod_cant_gondola) AS cantidadgondolas FROM detbodega WHERE (dtbod_bodega =" & ComboBox6.SelectedValue & ")", "tabla8")
            If clase.dt.Tables("tabla8").Rows.Count > 0 Then
                If IsDBNull(clase.dt.Tables("tabla8").Rows(0)("cantidadgondolas")) Then
                    MessageBox.Show("No hay góndolas en esta bodega.", "NO SE ENCONTRARON GÓNDOLAS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                Dim gondolasarray(clase.dt.Tables("tabla8").Rows(0)("cantidadgondolas")) As String
                clase.dt_global.Tables.Clear()
                clase.consultar_global("SELECT dtbod_bloque FROM detbodega WHERE (dtbod_bodega =" & ComboBox6.SelectedValue & ")", "tabla9")
                If clase.dt_global.Tables("tabla9").Rows.Count > 0 Then
                    Dim a As Integer
                    For a = 0 To clase.dt_global.Tables("tabla9").Rows.Count - 1
                        clase.consultar("SELECT dtbod_cant_gondola FROM detbodega WHERE (dtbod_bloque =" & clase.dt_global.Tables("tabla9").Rows(a)("dtbod_bloque") & " AND dtbod_bodega =" & ComboBox6.SelectedValue & ")", "tabla7")
                        If clase.dt.Tables("tabla7").Rows.Count > 0 Then
                            Dim c As Short
                            For c = 0 To clase.dt.Tables("tabla7").Rows(0)("dtbod_cant_gondola") - 1
                                cont += 1
                                gondolasarray(cont) = bloque(clase.dt_global.Tables("tabla9").Rows(a)("dtbod_bloque") - 1) & c + 1

                            Next
                        End If
                    Next
                    DataGridView1.RowCount = 0
                    For c = 0 To gondolasarray.Length - 1
                        llenar_gondola(gondolasarray(c), ComboBox6.SelectedValue)
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub llenar_gondola(gondo As String, bod As Short)
        Dim g As Integer = DataGridView1.RowCount
        clase.dt_global.Tables.Clear()
        DataGridView1.RowCount = DataGridView1.RowCount + 1
        DataGridView1.Item(0, g).Value = UCase(gondo)
        clase.consultar_global("SELECT articulos.ar_referencia, articulos.ar_descripcion, SUM(inventario_bodega.inv_cantidad) AS cantidades, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4 FROM inventario_bodega INNER JOIN articulos ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) WHERE (inventario_bodega.inv_gondola ='" & UCase(gondo) & "' AND inventario_bodega.inv_bodega = " & bod & ") GROUP BY articulos.ar_referencia, articulos.ar_descripcion, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4", "tabla4")
        If clase.dt_global.Tables("tabla4").Rows.Count > 0 Then
            Dim x As Integer
            Dim cant As Integer = 0
            ' DataGridView1.RowCount = 0
            With DataGridView1
                For x = 0 To clase.dt_global.Tables("tabla4").Rows.Count - 1
                    If x <> 0 Then
                        .RowCount = .RowCount + 1
                    End If
                    .Item(1, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_referencia")
                    .Item(2, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_descripcion")
                    .Item(4, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("cantidades")
                    .Item(6, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_linea")
                    .Item(7, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_sublinea1")
                    .Item(8, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_sublinea2")
                    .Item(9, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_sublinea3")
                    .Item(10, g + x).Value = clase.dt_global.Tables("tabla4").Rows(x)("ar_sublinea4")
                    .Item(11, g + x).Value = UCase(gondo)
                    cant = cant + clase.dt_global.Tables("tabla4").Rows(x)("cantidades")
                Next
                .Item(5, .RowCount - x).Value = cant
            End With
        Else
            '    DataGridView1.RowCount = 0
        End If
    End Sub

    Function verificar_existencia_gondola(gondo As String, bodega As Short) As Boolean
        Dim gondola(cant_gondola_bodega(bodega)) As String
        clase.consultar("SELECT detbodega.dtbod_bloque, detbodega.dtbod_cant_gondola FROM detbodega INNER JOIN bodegas ON (detbodega.dtbod_bodega = bodegas.bod_codigo) WHERE (bodegas.bod_codigo =" & bodega & ")", "tabla1")
        If clase.dt.Tables("tabla1").Rows.Count > 0 Then
            Dim a, b As Short
            Dim cont As Integer = -1
            For a = 0 To clase.dt.Tables("tabla1").Rows.Count - 1
                For b = 1 To clase.dt.Tables("tabla1").Rows(a)("dtbod_cant_gondola")
                    cont += 1
                    gondola(cont) = bloque(clase.dt.Tables("tabla1").Rows(a)("dtbod_bloque") - 1) & b
                Next
            Next
            Dim i As Integer
            For i = 0 To cont
                If gondo = gondola(i) Then
                    Return True
                    Exit Function
                End If
            Next
            Return False
        End If
    End Function

    Function cant_gondola_bodega(bodega As Short) As Integer
        clase.consultar("SELECT SUM(detbodega.dtbod_cant_gondola) AS totalgondola FROM detbodega INNER JOIN bodegas ON (detbodega.dtbod_bodega = bodegas.bod_codigo) WHERE (bodegas.bod_codigo =" & bodega & ")", "tablatotal")
        If clase.dt.Tables("tablatotal").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tablatotal").Rows(0)("totalgondola")) Then
                Return 0
                Exit Function
            End If
            Return clase.dt.Tables("tablatotal").Rows(0)("totalgondola")
        End If
    End Function

    Private Sub ComboBox6_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedValueChanged
        If ind_carga = True Then
            clase.consultar("select* from detbodega where dtbod_bodega = " & ComboBox6.SelectedValue & " order by dtbod_bloque asc", "tabla3")
            If clase.dt.Tables("tabla3").Rows.Count > 0 Then
                ComboBox1.Items.Clear()
                ComboBox1.Items.Add("")
                Dim d As Short
                For d = 0 To clase.dt.Tables("tabla3").Rows.Count - 1
                    ComboBox1.Items.Add(bloque(clase.dt.Tables("tabla3").Rows(d)("dtbod_bloque") - 1))
                Next
            Else
                ComboBox1.Items.Clear()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 3 Then
            If DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value = "" Then
                Exit Sub
            End If
            frm_codigos_asociados_inventario.ShowDialog()
            frm_codigos_asociados_inventario.Dispose()
        End If
    End Sub

    Function convertir_letra_en_numero(letra As String) As Short
        Dim b As Short
        For b = 0 To 25
            If arreglobloque(b, 0) = letra Then
                Return arreglobloque(b, 1)
            End If
        Next
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_ajustar_gondola.ShowDialog()
        frm_ajustar_gondola.Dispose()
    End Sub

    Sub llenar_array_con_gondola_existentes(arreglogondolas() As String, bodega As Short, bloq As String)
       
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub
End Class

