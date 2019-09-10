Public Class frm_imprimir_codigos_barras_gondolas
    Dim clase As New class_library
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Dim ind_carga As Boolean
    Dim arreglobloque(26, 2) As String

    Private Sub frm_imprimir_codigos_barras_gondolas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ind_carga = False
        clase.llenar_combo(ComboBox1, "select* from bodegas order by bod_codigo ASC", "bod_nombre", "bod_codigo")
        ind_carga = True
        Dim a As Short
        For a = 0 To 25
            arreglobloque(a, 0) = bloque(a)
            arreglobloque(a, 1) = a + 1
        Next
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        textBox2.Enabled = True
        textBox2.Focus()
        ComboBox2.Enabled = False
        ComboBox2.Text = ""
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        textBox2.Enabled = False
        textBox2.Text = ""
        ComboBox2.Enabled = True
        ComboBox2.Focus()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        If ind_carga = True Then
            clase.consultar("SELECT dtbod_bloque FROM detbodega WHERE (dtbod_bodega =" & ComboBox1.SelectedValue & ") ORDER BY dtbod_bloque ASC", "tabla")
            ComboBox2.Items.Clear()
            If clase.dt.Tables("tabla").Rows.Count > 0 Then
                Dim i As Integer
                For i = 0 To clase.dt.Tables("tabla").Rows.Count - 1
                    ComboBox2.Items.Add(bloque(i))
                Next
            End If
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

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_combobox(ComboBox1, "Bodega") = False Then Exit Sub
        If RadioButton1.Checked = True Then
            If clase.validar_cajas_text(textBox2, "Góndola") = False Then Exit Sub
            If verificar_existencia_gondola(UCase(textBox2.Text), ComboBox1.SelectedValue) = False Then
                MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                textBox2.Text = ""
                textBox2.Focus()
                Exit Sub
            End If
            clase.agregar_registro("insert into codigo_de_bara_gondolas (codigobarra) values ('" & "*" & abreviatura_bodega(ComboBox1.SelectedValue) & UCase(textBox2.Text) & "*" & "')")

        Else
            If clase.validar_combobox(ComboBox2, "Bloque") = False Then Exit Sub
            clase.consultar("SELECT dtbod_cant_gondola FROM detbodega WHERE (dtbod_bloque =" & convertir_letra_en_numero(ComboBox2.Text) & " AND dtbod_bodega =" & ComboBox1.SelectedValue & ")", "tabla3")
            If IsDBNull(clase.dt.Tables("tabla3").Rows(0)("dtbod_cant_gondola")) = False Then
                Dim a As Integer
                For a = 1 To clase.dt.Tables("tabla3").Rows(0)("dtbod_cant_gondola")
                    clase.agregar_registro("insert into codigo_de_bara_gondolas (codigobarra) values ('" & "*" & abreviatura_bodega(ComboBox1.SelectedValue) & ComboBox2.Text & a & "*" & "')")
                Next
            Else
                MessageBox.Show("Bloque no definido", "BLOQUE SIN GONDOLAS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
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

    Function abreviatura_bodega(codbod As Integer) As String
        clase.consultar("SELECT bod_abreviatura FROM bodegas WHERE (bod_codigo =" & codbod & ")", "tbl")
        Return clase.dt.Tables("tbl").Rows(0)("bod_abreviatura")
    End Function
End Class