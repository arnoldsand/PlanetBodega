Public Class frm_salidas_mercancia
    Dim clase As New class_library
    Dim consecutivo As Integer
    Dim ind_carga As Boolean
    Dim cantidad As Integer

    Private Sub frm_salidas_mercancia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        calcular_consecutivo()
        DataGridView1.ColumnCount = 7
        ' preparar_columnas()
    End Sub

    Private Sub calcular_consecutivo()
        clase.consultar("select max(cabsal_cod) as maximo from cabsalidas_mercancia", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tabla").Rows(0)("maximo")) Then
                consecutivo = 1
            Else
                consecutivo = clase.dt.Tables("tabla").Rows(0)("maximo") + 1
            End If
            TextBox3.Text = consecutivo
        Else
            consecutivo = 1
            TextBox3.Text = consecutivo
        End If
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs)
        clase.validar_numeros(e)
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If clase.validar_cajas_text(TextBox1, "Compañia") = False Then Exit Sub
        'If clase.validar_cajas_text(TextBox5, "Liquidador") = False Then Exit Sub
        frm_agregar_item_salida.ShowDialog()
        frm_agregar_item_salida.Dispose()
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Marca"
            .Columns(3).HeaderText = "Composición"
            .Columns(4).HeaderText = "Cant"
            .Columns(5).HeaderText = "Medida"
            .Columns(0).Width = 120
            .Columns(1).Width = 200
            .Columns(2).Width = 100
            .Columns(3).Width = 200
            .Columns(4).Width = 50
            .Columns(5).Width = 100
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs)
        '   MsgBox(ComboBox1.SelectedValue)
        If ind_carga = False Then
            clase.dt_global.Tables.Clear()
            '  clase.consultar_global("select* from detalle_importacion_detcajas where detcab_coditem = " & ComboBox1.SelectedValue & "", "table123")
            '  If clase.dt.Tables("table123").Rows.Count > 0 Then
            'cantidad = clase.dt_global.Tables("table123").Rows(0)("detcab_cantidad")
            '
        End If
        '   End If
    End Sub

    Function calcular_cant_referencias_agregadas(ref As String) As Integer
        Dim a As Short
        For a = 0 To DataGridView1.RowCount - 1
            If ref = DataGridView1.Item(0, a).Value Then
                Return DataGridView1.Item(3, a).Value
                Exit For
            End If
        Next
        Return 0
    End Function

    Function determinar_existencia_de_referencia(ref As String) As Boolean
        Dim a As Short
        For a = 0 To DataGridView1.RowCount - 1
            If ref = DataGridView1.Item(0, a).Value Then
                Return True
                Exit For
            End If
        Next
        Return False
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        restablecer()
    End Sub

    Private Sub restablecer()
        TextBox5.Text = ""
        Button1.Enabled = True
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.RowCount = 0 Then
            MessageBox.Show("Debe agregar por lo menos un item para registrar la salida de bodega.", "AGREGAR ITEMS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If clase.validar_cajas_text(TextBox5, "Liquidador") = False Then Exit Sub
        Dim v As String = MessageBox.Show("¿Desea registar la salida de bodega generada en este momento?", "GENERAR SALIDA DE BODEGA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            Dim fecha As String = Now.ToString("yyyy-MM-dd")
            calcular_consecutivo()
            clase.agregar_registro("INSERT INTO `cabsalidas_mercancia`(`cabsal_cod`,`cabsal_fecha`,`calsal_estado`,`cabsal_liquidador`,`cabsal_proveedor`) VALUES ( '" & TextBox3.Text & "','" & fecha & "', True, '" & UCase(TextBox5.Text) & "',NULL)")
            Dim i As Short
            For i = 0 To DataGridView1.RowCount - 1
                clase.agregar_registro("INSERT INTO `detsalidas_mercancia`(`det_salidacodigo`,`det_codref`,`det_cant`,`det_procesado`) VALUES ('" & TextBox3.Text & "','" & DataGridView1.Item(0, i).Value & "','" & DataGridView1.Item(4, i).Value & "',False)")
                clase.actualizar("UPDATE detalle_importacion_detcajas SET detcab_cantidad = '0' WHERE detcab_coditem = " & DataGridView1.Item(0, i).Value & "")
            Next
            restablecer()
            calcular_consecutivo()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        frm_listado_proveedores7.ShowDialog()
        frm_listado_proveedores7.Dispose()
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            DataGridView1.Rows.Remove(DataGridView1.Rows(DataGridView1.CurrentRow.Index))
        End If
    End Sub
End Class