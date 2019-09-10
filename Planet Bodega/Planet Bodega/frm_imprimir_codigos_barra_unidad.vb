Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class frm_imprimir_codigos_barra_unidad
    Dim clase As New class_library
    Dim ean13 As New clase_ean13_codigo_de_barar

    Private Sub frm_imprimir_codigos_barra_unidad_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With dataGridView1
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = False
            .Columns(4).ReadOnly = False
        End With
        Dim impresora As Printer
        ComboBox6.Items.Clear()
        For Each impresora In Printers
            ComboBox6.Items.Add(impresora.DeviceName)
        Next
        ComboBox6.Text = ultima_impresora_usada()
    End Sub

    Function ultima_impresora_usada() As String
        clase.consultar("select* from informacion", "impresora")
        If IsDBNull(clase.dt.Tables("impresora").Rows(0)("ultima_impresora_usada")) Then
            Return ""
        Else
            Return clase.dt.Tables("impresora").Rows(0)("ultima_impresora_usada")
        End If
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Articulo") = False Then Exit Sub
        If (Val(TextBox1.Text) = 0) And (Val(TextBox3.Text) = 0) Then
            MessageBox.Show("Debe digitar cantidades validas. Pulse aceptar para volverlo a intentar.", "CAMPOS INCOMPLETOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Focus()
            Exit Sub
        End If
        Dim codigo_articulo As Long
        If Len(textBox2.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & textBox2.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(textBox2.Text)
        End If
        If Len(textBox2.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & textBox2.Text & ")", "tabla11")
            codigo_articulo = textBox2.Text
        End If
        If Len(textBox2.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            Dim a As Short
            Dim cant As Short
            Dim ind As Boolean = False
            For a = 0 To dataGridView1.RowCount - 1
                If codigo_articulo = dataGridView1.Item(0, a).Value Then
                    cant = Val(dataGridView1.Item(3, a).Value)
                    dataGridView1.Item(3, a).Value = cant + Val(TextBox1.Text)
                    cant = Val(dataGridView1.Item(4, a).Value)
                    dataGridView1.Item(4, a).Value = cant + Val(TextBox3.Text)
                    ind = True
                    Exit For
                End If
            Next
            If ind = False Then
                dataGridView1.RowCount = dataGridView1.RowCount + 1
                With dataGridView1
                    .Item(0, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_codigo")
                    .Item(1, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_referencia")
                    .Item(2, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_descripcion")
                    .Item(3, .RowCount - 1).Value = Val(TextBox1.Text)
                    .Item(4, .RowCount - 1).Value = Val(TextBox3.Text)
                End With
            End If
            textBox2.Text = ""
            textBox2.Focus()
            TextBox1.Text = ""
            TextBox3.Text = ""
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dataGridView1.EditingControlShowing
        ' obtener indice de la columna   
        Dim columna As Integer = dataGridView1.CurrentCell.ColumnIndex
        ' comprobar si la celda en edición corresponde a la columna 12
        If (columna = 3 Or columna = 4) Then
            ' referencia a la celda   
            Dim validar As TextBox = CType(e.Control, TextBox)
            ' agregar el controlador de eventos para el KeyPress   
            AddHandler validar.KeyPress, AddressOf validar_Keypress
        End If
    End Sub

    Private Sub validar_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        ' Obtener caracter   
        Dim caracter As Char = e.KeyChar
        ' comprobar si el caracter es un número o el retroceso   
        If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
            'Me.Text = e.KeyChar   
            e.KeyChar = Chr(0)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim b As Short
        Dim ind = False
        For b = 0 To dataGridView1.RowCount - 1
            If (Val(dataGridView1.Item(3, b).Value) = 0) And (Val(dataGridView1.Item(4, b).Value) = 0) Then
                MessageBox.Show("No ha especificado algunas cantidades. Por favor revise y vuelva a intentarlo.", "ESPECIFICAR CANTIDADES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ind = True
                Exit For
            End If
        Next
        If ind = False Then
            clase.borradoautomatico("delete from tabla_codigos_barra_ean13")
            Dim z As Short
            For z = 0 To dataGridView1.RowCount - 1
                ean13.imprimir_codigos_de_barra_33_x_25(dataGridView1.Item(0, z).Value, dataGridView1.Item(5, z).Value, ComboBox6.Text)
            Next
            Button1_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        If dataGridView1.RowCount > 0 And Len(textBox2.Text) = 0 Then
            Me.AcceptButton = Button4
        End If
    End Sub

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dataGridView1.Rows.Clear()
        textBox2.Focus()
        textBox2.Text = ""
        TextBox1.Text = ""
        TextBox3.Text = ""
        Me.AcceptButton = Nothing
    End Sub

    Private Sub textBox2_KeyUp(sender As Object, e As KeyEventArgs) Handles textBox2.KeyUp
        If Len(textBox2.Text) > 0 Then
            Me.AcceptButton = Nothing
        Else
            Me.AcceptButton = Button4
        End If
    End Sub

    Private Sub textBox2_LostFocus(sender As Object, e As EventArgs) Handles textBox2.LostFocus
        Me.AcceptButton = Nothing
    End Sub
End Class


