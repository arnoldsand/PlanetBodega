Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class frm_impresion_codigo_barra_x_unidad_19_x_15
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
        If (Val(TextBox1.Text) = 0) Then
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
                    cant = Val(dataGridView1.Item(5, a).Value)
                    dataGridView1.Item(5, a).Value = cant + Val(TextBox1.Text)
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
                    .Item(3, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("colornombre")
                    .Item(4, .RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("nombretalla")
                    .Item(5, .RowCount - 1).Value = Val(TextBox1.Text)
                    '   .Item(6, .RowCount - 1).Value = Val(TextBox3.Text)
                End With
                EstablecerFoto(clase.dt.Tables("tabla11").Rows(0)("ar_codigo"))
            End If
            textBox2.Text = ""
            textBox2.Focus()
            TextBox1.Text = ""
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
        End If
    End Sub

    Sub EstablecerFoto(IdArticulo As Integer)
        Try
            clase.consultar("select ar_foto from articulos where ar_codigo = " & IdArticulo & "", "tablita")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        Catch
            PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            SetImage(PictureBox1)
        End Try
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
        If (ComboBox1.Text = "") Or (ComboBox6.Text = "") Then
            MessageBox.Show("Debe seleccionar la impresora y la resolución para imprimir tiquetes", "ESTABLECER IMPRESORA Y RESOLUCIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim b As Short
        Dim ind = False
        For b = 0 To dataGridView1.RowCount - 1
            If Val(dataGridView1.Item(5, b).Value) = 0 Then
                MessageBox.Show("No ha especificado algunas cantidades. Por favor revise y vuelva a intentarlo.", "ESPECIFICAR CANTIDADES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ind = True
                Exit For
            End If
        Next
        If ind = False Then
            clase.borradoautomatico("delete from tabla_codigos_barra_ean13")
            Dim z As Short
            Select Case ComboBox1.Text
                Case "203 dpi"
                    For z = 0 To dataGridView1.RowCount - 1
                        ean13.imprimir_codigos_de_barra_19_x_15_203dpi(dataGridView1.Item(0, z).Value, dataGridView1.Item(5, z).Value, ComboBox6.Text)
                    Next
                Case "300 dpi"
                    For z = 0 To dataGridView1.RowCount - 1
                        ean13.imprimir_codigos_de_barra_19_x_15_300dpi(dataGridView1.Item(0, z).Value, dataGridView1.Item(5, z).Value, ComboBox6.Text)
                    Next
            End Select

            Button1_Click(Nothing, Nothing)
            PictureBox1.Image = Nothing
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

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs)
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dataGridView1.Rows.Clear()
        textBox2.Focus()
        textBox2.Text = ""
        TextBox1.Text = ""
        Me.AcceptButton = Nothing
        PictureBox1.Image = Nothing
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

    Private Sub dataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridView1.CellClick
        If dataGridView1.Rows.Count = 0 Then Exit Sub
        EstablecerFoto(dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value)
    End Sub
End Class