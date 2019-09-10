Public Class frm_ver_detalle_factura
    Dim clase As New class_library
    Dim cdfactura As Short
    Dim z As Short
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_listado_proveedores6.ShowDialog()
        frm_listado_proveedores6.Dispose()
    End Sub

    Private Sub frm_ver_detalle_factura_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cdfactura = frm_facturas_tabuladas.dtgridcajas.Item(0, frm_facturas_tabuladas.dtgridcajas.CurrentCell.RowIndex).Value
        clase.consultar("SELECT cabfacturas.cabfact_importacion, cabfacturas.cabfect_funcionario, cabfacturas.cabfact_numero, proveedores.prv_codigo FROM cabfacturas INNER JOIN proveedores  ON (cabfacturas.cabfact_proveedor = proveedores.prv_codigo) WHERE (cabfacturas.cabfact_codigo =" & cdfactura & ")", "facturas")
        textBox2.Text = clase.dt.Tables("facturas").Rows(0)("cabfact_importacion")
        TextBox1.Text = clase.dt.Tables("facturas").Rows(0)("prv_codigo")
        TextBox4.Text = verificar_nulidad_vacio(clase.dt.Tables("facturas").Rows(0)("cabfact_numero"))
        TextBox3.Text = verificar_nulidad_vacio(clase.dt.Tables("facturas").Rows(0)("cabfect_funcionario"))
        llenar_grilla()
        clase.llenar_combo(ComboBox1, "select* from unidadesmedida order by nombremedida asc", "nombremedida", "nombremedida")
        ComboBox1.Text = "Docena(s)"
    End Sub


    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub llenar_grilla()
        clase.consultar("SELECT detfact_referencia, detfact_descripcion, detfact_cant, detfact_unimedida FROM detfacturas WHERE (detfact_factura =" & cdfactura & ")", "llenar")
        If clase.dt.Tables("llenar").Rows.Count > 0 Then
            Dim p As Short
            dtgridcajas.Rows.Clear()
            dtgridcajas.ColumnCount = 4
            preparar_columnas()
            For p = 0 To clase.dt.Tables("llenar").Rows.Count - 1
                dtgridcajas.RowCount = dtgridcajas.RowCount + 1
                dtgridcajas.Item(0, p).Value = clase.dt.Tables("llenar").Rows(p)("detfact_referencia")
                dtgridcajas.Item(1, p).Value = clase.dt.Tables("llenar").Rows(p)("detfact_descripcion")
                dtgridcajas.Item(2, p).Value = clase.dt.Tables("llenar").Rows(p)("detfact_cant")
                dtgridcajas.Item(3, p).Value = clase.dt.Tables("llenar").Rows(p)("detfact_unimedida")
                z = z + 1
            Next
        Else
            dtgridcajas.Rows.Clear()
            dtgridcajas.ColumnCount = 4
            preparar_columnas()
        End If
    End Sub

    Private Sub preparar_columnas()
        With dtgridcajas
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Cant"
            .Columns(3).HeaderText = "Medida"
            .Columns(0).Width = 200
            .Columns(1).Width = 250
            .Columns(2).Width = 50
            .Columns(3).Width = 120
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If clase.validar_cajas_text(TextBox8, "Referencia") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox5, "Descripción") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox7, "Cant") = False Then Exit Sub
        dtgridcajas.RowCount = dtgridcajas.RowCount + 1
        dtgridcajas.Item(0, z).Value = UCase(TextBox8.Text)
        dtgridcajas.Item(1, z).Value = UCase(TextBox5.Text)
        dtgridcajas.Item(2, z).Value = UCase(TextBox7.Text)
        dtgridcajas.Item(3, z).Value = ComboBox1.Text
        z = z + 1
        TextBox8.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Focus()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If dtgridcajas.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dtgridcajas.CurrentRow.Index) Then
            dtgridcajas.Rows.Remove(dtgridcajas.Rows(dtgridcajas.CurrentRow.Index))
            z = z - 1
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If dtgridcajas.Rows.Count = 0 Then
            Exit Sub
        End If
        clase.actualizar("UPDATE cabfacturas SET cabfact_proveedor = " & TextBox1.Text & ", cabfact_numero = '" & UCase(TextBox4.Text) & "', cabfect_funcionario = '" & UCase(TextBox3.Text) & "' WHERE cabfact_codigo = " & cdfactura & "")
        clase.borradoautomatico("delete from detfacturas WHERE detfact_factura = " & cdfactura & "")
        Dim i As Short
        For i = 0 To dtgridcajas.Rows.Count - 1
            clase.agregar_registro("INSERT INTO `detfacturas`(`detfact_codigo`,`detfact_factura`,`detfact_referencia`,`detfact_descripcion`,`detfact_cant`,`detfact_unimedida`) VALUES ( NULL,'" & cdfactura & "','" & dtgridcajas.Item(0, i).Value & "','" & dtgridcajas.Item(1, i).Value & "','" & dtgridcajas.Item(2, i).Value & "','" & dtgridcajas.Item(3, i).Value & "')")
        Next
        Me.Close()
    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.enter(e)
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox8_LostFocus(sender As Object, e As EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        TextBox5.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub
End Class