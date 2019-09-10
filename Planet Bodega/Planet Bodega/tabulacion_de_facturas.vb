Public Class tabulacion_de_facturas
    Dim clase As New class_library

    Dim a As Short
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    
    Private Sub tabulacion_de_facturas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        a = 0
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_listado_importacion5.ShowDialog()
        frm_listado_importacion5.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_listado_proveedores5.ShowDialog()
        frm_listado_proveedores5.Dispose()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Importación") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Proveedor") = False Then Exit Sub
        GroupBox2.Enabled = True
        TextBox8.Focus()
        Button8.Enabled = True
        Button9.Enabled = True
        Button11.Enabled = True
        Button2.Enabled = False
        clase.llenar_combo(ComboBox1, "select* from unidadesmedida order by nombremedida asc", "nombremedida", "nombremedida")
        ComboBox1.Text = "Docena(s)"
    End Sub

    Function consecutivo_cabfactura() As Integer
        clase.consultar("SELECT MAX(cabfact_codigo) AS maximo FROM cabfacturas", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tabla").Rows(0)("maximo")) Then
                Return 1
            Else
                Return clase.dt.Tables("tabla").Rows(0)("maximo") + 1
            End If
        End If
    End Function


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

    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        TextBox4.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        TextBox3.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub Button2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Button2.KeyPress
        clase.enter(e)
    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As EventArgs) Handles ComboBox1.GotFocus
        Me.AcceptButton = Button8
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If clase.validar_cajas_text(TextBox8, "Referencia") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox5, "Descripción") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox7, "Cant") = False Then Exit Sub
        dtgridcajas.RowCount = dtgridcajas.RowCount + 1
        dtgridcajas.Item(0, a).Value = a + 1
        dtgridcajas.Item(1, a).Value = UCase(TextBox8.Text)
        dtgridcajas.Item(2, a).Value = UCase(TextBox5.Text)
        dtgridcajas.Item(3, a).Value = UCase(TextBox7.Text)
        dtgridcajas.Item(4, a).Value = ComboBox1.Text
        a = a + 1
        TextBox8.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Focus()
    End Sub

    Private Sub ComboBox1_LostFocus(sender As Object, e As EventArgs) Handles ComboBox1.LostFocus
        Me.AcceptButton = Nothing
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If dtgridcajas.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dtgridcajas.CurrentRow.Index) Then
            Dim v As String = MessageBox.Show("¿Desea guardar los datos de esta factura?", "GUARDAR DATOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                Dim fech As Date = Now
                Dim codigocabfactura As Integer = consecutivo_cabfactura()
                clase.agregar_registro("INSERT INTO `cabfacturas`(`cabfact_codigo`,`cabfact_numero`,`cabfact_proveedor`,`cabfact_importacion`,`cabfact_fecha`,`cabfect_funcionario`) VALUES ( '" & codigocabfactura & "'," & comprobar_nulidad(UCase(TextBox4.Text)) & ",'" & TextBox1.Text & "','" & textBox2.Text & "','" & fech.ToString("yyyy-MM-dd") & "'," & comprobar_nulidad(UCase(TextBox3.Text)) & ")")
                Dim x As Short
                For x = 0 To dtgridcajas.RowCount - 1
                    clase.agregar_registro("INSERT INTO `detfacturas`(`detfact_factura`,`detfact_referencia`,`detfact_descripcion`,`detfact_cant`,`detfact_unimedida`) VALUES ('" & codigocabfactura & "','" & dtgridcajas.Item(1, x).Value & "','" & dtgridcajas.Item(2, x).Value & "','" & dtgridcajas.Item(3, x).Value & "','" & dtgridcajas.Item(4, x).Value & "')")
                Next
                dtgridcajas.Rows.Clear()
                GroupBox2.Enabled = False
                TextBox3.Text = ""
                TextBox4.Text = ""
                textBox2.Text = ""
                TextBox1.Text = ""
                Button8.Enabled = False
                Button9.Enabled = False
                Button11.Enabled = False
                Button2.Enabled = True
                a = 0
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If dtgridcajas.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dtgridcajas.CurrentRow.Index) Then
            dtgridcajas.Rows.Remove(dtgridcajas.Rows(dtgridcajas.CurrentRow.Index))
            a = a - 1
        End If
    End Sub
End Class