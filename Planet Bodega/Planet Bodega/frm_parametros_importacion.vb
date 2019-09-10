Public Class frm_parametros_importacion
    Dim clase As New class_library
    Dim encontrado As Boolean
    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        Me.Close()
    End Sub

    Private Sub frm_parametros_importacion_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        clase.consultar("Select* from cabimportacion where imp_codigo = " & cod_importacion & "", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            TextBox2.Text = clase.dt.Tables("tabla")(0)("imp_codigo")
            TextBox1.Text = clase.dt.Tables("tabla")(0)("imp_nombrefecha")
        End If
        Dim i As Short
        Dim campos As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
        clase.consultar("select* from imp_parametros where codigo_importacion = " & cod_importacion & "", "tabla2")
        If clase.dt.Tables("tabla2").Rows.Count > 0 Then
            encontrado = True
            For i = 0 To 25
                ComboBox1.Items.Add(campos(i))
                If clase.dt.Tables("tabla2")(0)("referencia") = campos(i) Then
                    ComboBox1.Text = campos(i)
                End If
                ComboBox2.Items.Add(campos(i))
                If clase.dt.Tables("tabla2")(0)("descripcion") = campos(i) Then
                    ComboBox2.Text = campos(i)
                End If
                ComboBox3.Items.Add(campos(i))
                If clase.dt.Tables("tabla2")(0)("marca") = campos(i) Then
                    ComboBox3.Text = campos(i)
                End If
                ComboBox4.Items.Add(campos(i))
                If clase.dt.Tables("tabla2")(0)("unidades") = campos(i) Then
                    ComboBox4.Text = campos(i)
                End If
            Next
        Else
            encontrado = False
            For i = 0 To 25
                ComboBox1.Items.Add(campos(i))
                ComboBox2.Items.Add(campos(i))
                ComboBox3.Items.Add(campos(i))
                ComboBox4.Items.Add(campos(i))
            Next
        End If
        


    End Sub

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        If ComboBox1.Text = "" Then
            MessageBox.Show("Debe seleccionar  algo en Referencia. Pulse aceptar para volverlo a intentar.", "SELECCIONAR VALOR", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox1.Focus()
            Exit Sub
        End If
        If ComboBox2.Text = "" Then
            MessageBox.Show("Debe seleccionar  algo en Descripción. Pulse aceptar para volverlo a intentar.", "SELECCIONAR VALOR", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox2.Focus()
            Exit Sub
        End If
        If ComboBox3.Text = "" Then
            MessageBox.Show("Debe seleccionar  algo en Marca. Pulse aceptar para volverlo a intentar.", "SELECCIONAR VALOR", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox3.Focus()
            Exit Sub
        End If
        If ComboBox4.Text = "" Then
            MessageBox.Show("Debe seleccionar  algo en Unidades. Pulse aceptar para volverlo a intentar.", "SELECCIONAR VALOR", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox4.Focus()
            Exit Sub
        End If
        If encontrado = False Then
            clase.agregar_registro("INSERT INTO imp_parametros (codigo_importacion, referencia, descripcion, marca, unidades) Values(" & cod_importacion & ", '" & ComboBox1.Text & "', '" & ComboBox2.Text & "', '" & ComboBox3.Text & "', '" & ComboBox4.Text & "')")
        Else
            clase.actualizar("UPDATE imp_parametros set referencia = '" & ComboBox1.Text & "', descripcion = '" & ComboBox2.Text & "', marca = '" & ComboBox3.Text & "', unidades = '" & ComboBox4.Text & "' where codigo_importacion = " & cod_importacion & "")
        End If
        Me.Close()
    End Sub
End Class