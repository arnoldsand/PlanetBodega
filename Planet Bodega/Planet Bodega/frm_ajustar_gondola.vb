Public Class frm_ajustar_gondola
    Dim clase As New class_library
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
       
            Me.Close()

    End Sub


    Private Sub frm_ajustar_gondola_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox6, "SELECT bod_codigo, bod_nombre FROM bodegas ORDER BY bod_codigo ASC", "bod_nombre", "bod_codigo")
    End Sub

  
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_combobox(ComboBox6, "Bodega Almacenaje") = False Then Exit Sub
        If clase.validar_cajas_text(textBox2, "Góndola") = False Then Exit Sub
        If verificar_existencia_gondola(UCase(textBox2.Text.Trim), ComboBox6.SelectedValue) = False Then
            MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
            Exit Sub
        End If
        clase.consultar("SELECT* FROM inventario_bodega WHERE (inv_gondola ='" & UCase(textBox2.Text) & "')", "tabla10")
        If clase.dt.Tables("tabla10").Rows.Count > 0 Then
            Dim v As String = MessageBox.Show("La góndola especificada ya fue capturada, cuando guarde el ajuste se borrará la captura hecha previamente ¿Desea Continuar?", "GÓNDOLA YA CAPTURADA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                TextBox4.Enabled = True
                TextBox4.Focus()
                TextBox1.Text = 0
                TextBox3.Text = 0
                ComboBox6.Enabled = False
                textBox2.Enabled = False
                Button1.Enabled = False
                Button3.Enabled = True
                Button4.Enabled = True
                Button6.Enabled = True
                Me.AcceptButton = Button3
                'clase.borradoautomatico("delete from inventario_bodega where inv_gondola = '" & UCase(textBox2.Text) & "' and inv_bodega = " & ComboBox6.SelectedValue & "")
            End If
            If v = 7 Then
                textBox2.Text = ""
                textBox2.Focus()
                Exit Sub
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox4.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox4.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox4.Text)
        End If
        If Len(TextBox4.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox4.Text & ")", "tabla11")
            codigo_articulo = TextBox4.Text
        End If
        If Len(TextBox4.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            Dim z As Short

            For z = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Item(0, z).Value = codigo_articulo Then
                    Dim cantid As Integer = DataGridView1.Item(5, z).Value
                    DataGridView1.Item(5, z).Value = cantid + 1
                    TextBox4.Text = ""
                    TextBox4.Focus()
                    cantidad_unidades()
                    Exit Sub
                End If
            Next
            DataGridView1.RowCount = DataGridView1.RowCount + 1
            With DataGridView1
                .Item(0, .RowCount - 1).Value = codigo_articulo
                .Item(1, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_referencia")
                .Item(2, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_descripcion")
                .Item(3, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("colornombre")
                .Item(4, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("nombretalla")
                .Item(5, .RowCount - 1).Value = 1
            End With
            cantidad_unidades()
            TextBox4.Text = ""
            TextBox4.Focus()
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox4.Focus()
        End If

    End Sub

    Private Sub cantidad_unidades()
        Dim i As Integer
        Dim cant_filas As Integer = 0
        For i = 0 To DataGridView1.RowCount - 1
            cant_filas = cant_filas + DataGridView1.Item(5, i).Value
        Next
        TextBox1.Text = cant_filas
        TextBox3.Text = DataGridView1.RowCount
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        restablecer()
    End Sub

    Private Sub restablecer()
        ComboBox6.SelectedIndex = -1
        ComboBox6.Enabled = True
        textBox2.Text = ""
        textBox2.Enabled = True
        textBox2.Focus()
        TextBox1.Text = 0
        TextBox3.Text = 0
        TextBox4.Enabled = False
        TextBox4.Text = ""
        DataGridView1.RowCount = 0
        Button1.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Me.AcceptButton = Button1
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DataGridView1.RowCount = 0 Then
            MessageBox.Show("Debe agregar por los menos un item para guardar el ajuste", "AGREGAR ITEMS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim v As String = MessageBox.Show("¿Desea guardar el ajuste hecho en este momento?", "GUARDAR AJUSTE", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            clase.borradoautomatico("delete from inventario_bodega where inv_gondola = '" & UCase(textBox2.Text) & "' and inv_bodega = " & ComboBox6.SelectedValue & "")
            Dim i As Short
            For i = 0 To DataGridView1.RowCount - 1
                With DataGridView1
                    clase.agregar_registro("INSERT INTO `inventario_bodega`(`inv_codigoart`,`inv_bodega`,`inv_gondola`,`inv_cantidad`) VALUES ('" & .Item(0, i).Value & "','" & ComboBox6.SelectedValue & "','" & UCase(textBox2.Text) & "','" & .Item(5, i).Value & "')")
                End With
            Next
            restablecer()
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
End Class