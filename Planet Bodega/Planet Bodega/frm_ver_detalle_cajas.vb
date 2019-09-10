Public Class frm_ver_detalle_cajas
    Dim clase As New class_library
    Dim codigoinicial As String
    Dim fila As Short
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ver_detalle_cajas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox1, "select* from unidadesmedida order by nombremedida asc", "nombremedida", "nombremedida")
        fila = frm_tabulacion_importacion.dtgridcajas.CurrentRow.Index
        clase.consultar("select* from detalle_importacion_cabcajas where det_caja = '" & frm_tabulacion_importacion.dtgridcajas.Item(0, fila).Value & "'", "tablita")
        If clase.dt.Tables("tablita").Rows.Count > 0 Then
            textBox2.Text = clase.dt.Tables("tablita").Rows(0)("det_caja")
            TextBox7.Text = clase.dt.Tables("tablita").Rows(0)("det_codigoproveedor")
            TextBox8.Text = clase.dt.Tables("tablita").Rows(0)("det_hoja")
            TextBox9.Text = clase.dt.Tables("tablita").Rows(0)("det_peso")
            codigoinicial = clase.dt.Tables("tablita").Rows(0)("det_caja")
        End If
        llenar_grilla_datos(frm_tabulacion_importacion.dtgridcajas.Item(0, fila).Value)
    End Sub

    Private Sub llenar_grilla_datos(codigocaja As String)
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, detalle_importacion_detcajas.detcab_marca, detalle_importacion_detcajas.detcab_composicion, detalle_importacion_detcajas.detcab_cantidad, unidadesmedida.nombremedida, detalle_importacion_detcajas.detcab_coditem, detalle_importacion_detcajas.detcab_subpartida FROM unidadesmedida INNER JOIN detalle_importacion_detcajas ON (unidadesmedida.nombremedida = detalle_importacion_detcajas.detcab_unimedida) WHERE (detalle_importacion_detcajas.detcab_codigocaja ='" & codigocaja & "')", "tbl123")
        If clase.dt.Tables("tbl123").Rows.Count > 0 Then
            prepara_columnas_grilla()
            Dim d As Short
            DataGridView1.RowCount = 0
            For d = 0 To clase.dt.Tables("tbl123").Rows.Count - 1
                With DataGridView1
                    .RowCount = .RowCount + 1
                    .Item(0, d).Value = clase.dt.Tables("tbl123").Rows(d)("detcab_referencia")
                    .Item(1, d).Value = clase.dt.Tables("tbl123").Rows(d)("detcab_descripcion")
                    .Item(2, d).Value = clase.dt.Tables("tbl123").Rows(d)("detcab_marca")
                    .Item(3, d).Value = clase.dt.Tables("tbl123").Rows(d)("detcab_composicion")
                    .Item(4, d).Value = clase.dt.Tables("tbl123").Rows(d)("detcab_subpartida")
                    .Item(5, d).Value = clase.dt.Tables("tbl123").Rows(d)("detcab_cantidad")
                    .Item(6, d).Value = clase.dt.Tables("tbl123").Rows(d)("nombremedida")
                    .Item(7, d).Value = clase.dt.Tables("tbl123").Rows(d)("detcab_coditem")
                End With
            Next
        Else
            DataGridView1.RowCount = 0
        End If
    End Sub

    Private Sub prepara_columnas_grilla()
        With DataGridView1
            .ColumnCount = 8
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Marca"
            .Columns(3).HeaderText = "Composición"
            .Columns(4).HeaderText = "Subpartida"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "Medida"
            .Columns(0).Width = 130
            .Columns(1).Width = 221
            .Columns(3).Width = 280
            .Columns(4).Width = 100
            .Columns(5).Width = 50
            .Columns(6).Width = 80
        End With
    End Sub

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As EventArgs) Handles TextBox7.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_proveedores2.ShowDialog()
        frm_proveedores2.Dispose()
    End Sub

    Private Sub textBox2_LostFocus(sender As Object, e As EventArgs) Handles textBox2.LostFocus
        textBox2.BackColor = Color.White
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        textBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox8_LostFocus(sender As Object, e As EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub textBox8_GotFocus(sender As Object, e As EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox9_LostFocus(sender As Object, e As EventArgs) Handles TextBox9.LostFocus
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub textBox9_GotFocus(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        TextBox9.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub textBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub textBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        TextBox3.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub textBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        TextBox4.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.enter(e)
    End Sub

    Private Sub textBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub textBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        TextBox5.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub textBox6_LostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        TextBox6.BackColor = Color.White
    End Sub

    Private Sub textBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        TextBox6.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        'restablecer()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If clase.validar_cajas_text(TextBox1, "Referencia") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox3, "Descripción") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox4, "Marca") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox5, "Composición") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox6, "Cantidad") = False Then Exit Sub
        If clase.validar_combobox(ComboBox1, "Unidad de Medida") = False Then Exit Sub
        If Val(TextBox6.Text) = 0 Then
            MessageBox.Show("La cantidad a digitar deber ser superior a 0", "ESPECIFICAR CANTIDAD", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox6.Focus()
            Exit Sub
        End If
        agregar_item_a_tabla(DataGridView1, TextBox1.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, ComboBox1.Text, TextBox10.Text)
    End Sub

    Private Sub agregar_item_a_tabla(dtgrid As DataGridView, ref As String, decri As String, marca As String, compo As String, cant As Integer, medida As String, subpartida As String)
        Dim c As Short
        Dim v As String
        Dim valcant As Integer
        Dim filas As Short = dtgrid.RowCount
        For c = 0 To filas - 1
            If UCase(TextBox1.Text) = dtgrid.Item(0, c).Value Then
                valcant = Val(dtgrid.Item(4, c).Value)
                v = MessageBox.Show("La Referencia que intenta agregar ya existe en el listado, ¿Desea añadir las cantidades al listado ya agregado?", "REFERENCIA YA EXISTE", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If v = "6" Then
                    cant = cant + valcant
                    dtgrid.Item(4, c).Value = cant
                    limpiar_despues_de_insercion()
                    Exit Sub
                End If
                If v = "7" Then
                    limpiar_despues_de_insercion()
                    Exit Sub
                End If
            End If
        Next
        dtgrid.RowCount = filas + 1
        With dtgrid
            .Item(0, dtgrid.RowCount - 1).Value = UCase(ref)
            .Item(1, dtgrid.RowCount - 1).Value = UCase(decri)
            .Item(2, dtgrid.RowCount - 1).Value = UCase(marca)
            .Item(3, dtgrid.RowCount - 1).Value = UCase(compo)
            .Item(4, dtgrid.RowCount - 1).Value = TextBox10.Text
            .Item(5, dtgrid.RowCount - 1).Value = cant
            .Item(6, dtgrid.RowCount - 1).Value = medida
        End With
        limpiar_despues_de_insercion()
    End Sub

    Private Sub limpiar_despues_de_insercion()
        TextBox1.Text = ""
        TextBox1.Focus()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox10.Text = ""
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        clase.consultar("SELECT det_fecharecepcion FROM detalle_importacion_cabcajas WHERE (det_caja ='" & textBox2.Text & "')", "caja")
        If IsDBNull(clase.dt.Tables("caja").Rows(0)("det_fecharecepcion")) = False Then
            MessageBox.Show("No se  puede eliminar articulos de una caja que ya fue recibida.", "CAJA YA RECIBIDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If DataGridView1.Item(7, DataGridView1.CurrentRow.Index).Value <> vbEmpty Then
            clase.borradoautomatico("delete from detalle_importacion_detcajas where detcab_coditem = " & DataGridView1.Item(7, DataGridView1.CurrentRow.Index).Value & "")
        End If
        DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'restringir la actualizacion cuando ya la caja se ha recibido
        clase.consultar("select* from detalle_importacion_cabcajas where det_caja = '" & codigoinicial & "'", "result")
        If clase.dt.Tables("result").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("result").Rows(0)("det_fecharecepcion")) = False Then
                MessageBox.Show("No se puede modificar una caja que ya fue recibida.", "IMPOSIBLE MODIFICAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If
        'verificar que el codigo nuevo para la caja no exista  ya
        If codigoinicial <> textBox2.Text Then
            clase.consultar("select* from detalle_importacion_cabcajas where det_caja = '" & textBox2.Text & "'", "result")
            If clase.dt.Tables("result").Rows.Count > 0 Then
                MessageBox.Show("El número de caja que digitó ya fue asignado a otra caja antes, por favor digite otro y vuelva a intentarlo.", "NUMERO DE CAJA YA EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                textBox2.Focus()
                textBox2.Text = ""
                Exit Sub
            End If
        End If
        If DataGridView1.RowCount = 0 Then
            MessageBox.Show("La caja debe tener por los menos un item agregado.", "AGREGAR ITEMS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Focus()
            Exit Sub
        End If
        Dim v As String = MessageBox.Show("¿Desea actualizar el detalle de Caja en este momento?", "ACTUALIZAR CAJA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            clase.actualizar("UPDATE `detalle_importacion_cabcajas` SET `det_caja`='" & textBox2.Text & "',`det_codigoproveedor`='" & TextBox7.Text & "',`det_peso`='" & TextBox9.Text & "',`det_hoja`='" & TextBox8.Text & "' WHERE `det_caja`='" & codigoinicial & "'")
            Dim i As Short
            Dim sql As String = ""
            Dim ind1 As Boolean = False
            Dim subpartida As String
            For i = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Item(4, i).Value.ToString() = "" Then
                    '  MsgBox("hola")
                    subpartida = "NULL"
                Else
                    ' MsgBox("hola1")
                    subpartida = "'" & DataGridView1.Item(4, i).Value & "'"
                End If
                If DataGridView1.Item(7, i).Value = vbEmpty Then
                    clase.agregar_registro("INSERT INTO `detalle_importacion_detcajas`(`detcab_codigocaja`,`detcab_referencia`,`detcab_descripcion`,`detcab_marca`,`detcab_composicion`,`detcab_cantidad`,`detcab_unimedida`,`detcab_subpartida`) VALUES ('" & textBox2.Text & "', '" & DataGridView1.Item(0, i).Value & "', '" & DataGridView1.Item(1, i).Value & "', '" & DataGridView1.Item(2, i).Value & "', '" & DataGridView1.Item(3, i).Value & "', '" & DataGridView1.Item(5, i).Value & "', '" & DataGridView1.Item(6, i).Value & "'," & subpartida & ")")
                Else
                    clase.actualizar("UPDATE detalle_importacion_detcajas SET detcab_referencia = '" & DataGridView1.Item(0, i).Value & "', detcab_descripcion = '" & DataGridView1.Item(1, i).Value & "', detcab_marca = '" & DataGridView1.Item(2, i).Value & "', detcab_composicion = '" & DataGridView1.Item(3, i).Value & "', detcab_cantidad = " & DataGridView1.Item(5, i).Value & ", detcab_unimedida = '" & DataGridView1.Item(6, i).Value & "', detcab_subpartida = " & subpartida & " WHERE detcab_coditem = " & DataGridView1.Item(7, i).Value & "")
                End If
            Next
            llenar_grilla_datos(frm_tabulacion_importacion.dtgridcajas.Item(0, fila).Value)
        End If
    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As EventArgs) Handles TextBox10.GotFocus
        TextBox10.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox10_LostFocus(sender As Object, e As EventArgs) Handles TextBox10.LostFocus
        TextBox10.BackColor = Color.White
    End Sub
End Class