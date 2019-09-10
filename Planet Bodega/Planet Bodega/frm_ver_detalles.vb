Public Class frm_ver_detalles
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        '   clase.validar_numeros(e)
      
    End Sub



    Private Sub frm_ver_detalles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("select* from registros_de_importaciones where articulo = " & registros_importacion_articulo.dtgridcajas.Item(0, registros_importacion_articulo.dtgridcajas.CurrentRow.Index).Value & "", "busqueda")
        If clase.dt.Tables("busqueda").Rows.Count > 0 Then
            textBox2.Text = clase.dt.Tables("busqueda").Rows(0)("registro")
            TextBox1.Text = clase.dt.Tables("busqueda").Rows(0)("fecha")
        Else
            textBox2.Text = ""
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Registros de Importación") = False Then Exit Sub
        If Not IsDate(TextBox1.Text) Then
            MessageBox.Show("La fecha escrita es invalida. Por favor escribala en este formato dd/mm/yyyy.", "FECHA INVALIDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim fecha1 As Date = TextBox1.Text
        clase.consultar("select* from registros_de_importaciones where articulo = " & registros_importacion_articulo.dtgridcajas.Item(0, registros_importacion_articulo.dtgridcajas.CurrentRow.Index).Value & "", "busqueda1")
        If clase.dt.Tables("busqueda1").Rows.Count > 0 Then
            clase.actualizar("update registros_de_importaciones set registro = '" & textBox2.Text & "', fecha = '" & fecha1.ToString("yyyy-MM-dd") & "' where articulo = " & registros_importacion_articulo.dtgridcajas.Item(0, registros_importacion_articulo.dtgridcajas.CurrentRow.Index).Value & " ")
        Else
            clase.agregar_registro("insert into registros_de_importaciones (articulo, registro, fecha) values ('" & registros_importacion_articulo.dtgridcajas.Item(0, registros_importacion_articulo.dtgridcajas.CurrentRow.Index).Value & "', '" & textBox2.Text & "', '" & fecha1.ToString("yyyy-MM-dd") & "')")
        End If
        registros_importacion_articulo.llenar_grilla()
        Me.Close()
    End Sub
End Class