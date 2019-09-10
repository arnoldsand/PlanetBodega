Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class frm_codigo_de_barras_unidad_29_x_15
    Dim clase As New class_library
    Dim printer As New Printer
    Dim ean13 As New clase_ean13_codigo_de_barar

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(textBox2, "Movimiento") = False Then Exit Sub
        clase.consultar("SELECT DISTINCT detalle_importacion_cabcajas.det_codigoimportacion, cabsalidas_mercancia.cabsal_liquidador FROM detsalidas_mercancia INNER JOIN cabsalidas_mercancia ON (detsalidas_mercancia.det_salidacodigo = cabsalidas_mercancia.cabsal_cod) INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (cabsalidas_mercancia.cabsal_cod =" & textBox2.Text & ")", "mov")
        If clase.dt.Tables("mov").Rows.Count > 0 Then
            TextBox3.Text = clase.dt.Tables("mov").Rows(0)("cabsal_liquidador")
            '"SELECT detsalidas_mercancia.det_codigo, articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, linea1.ln1_nombre, detsalidas_mercancia.det_cant, detalle_importacion_detcajas.detcab_unimedida FROM detsalidas_mercancia INNER JOIN asociaciones_codigos ON (detsalidas_mercancia.det_codref = asociaciones_codigos.asc_precodref) INNER JOIN articulos ON (asociaciones_codigos.asc_postcodart = articulos.ar_codigo) INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) AND (asociaciones_codigos.asc_precodref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) WHERE (detsalidas_mercancia.det_salidacodigo = " & textBox2.Text & ") ORDER BY articulos.ar_codigo ASC"
            clase.consultar("SELECT detsalidas_mercancia.det_codigo, articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, linea1.ln1_nombre, entradamercancia.com_unidades FROM detsalidas_mercancia INNER JOIN asociaciones_codigos ON (detsalidas_mercancia.det_codref = asociaciones_codigos.asc_precodref) INNER JOIN articulos ON (asociaciones_codigos.asc_postcodart = articulos.ar_codigo) INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) INNER JOIN entradamercancia ON (entradamercancia.com_codigoart = articulos.ar_codigo) WHERE (detsalidas_mercancia.det_salidacodigo =" & textBox2.Text & " AND entradamercancia.com_codigoimp =" & clase.dt.Tables("mov").Rows(0)("det_codigoimportacion") & ") ORDER BY articulos.ar_codigo ASC", "articulos")
            Dim a As Short
            dataGridView1.RowCount = 0
            Dim cant1 As Short = 0
            For a = 0 To clase.dt.Tables("articulos").Rows.Count - 1
                dataGridView1.RowCount = dataGridView1.RowCount + 1
                dataGridView1.Item(0, a).Value = False
                dataGridView1.Item(1, a).Value = clase.dt.Tables("articulos").Rows(a)("ar_codigo")
                dataGridView1.Item(2, a).Value = clase.dt.Tables("articulos").Rows(a)("ar_referencia")
                dataGridView1.Item(3, a).Value = clase.dt.Tables("articulos").Rows(a)("ar_descripcion")
                dataGridView1.Item(4, a).Value = clase.dt.Tables("articulos").Rows(a)("colornombre")
                dataGridView1.Item(5, a).Value = clase.dt.Tables("articulos").Rows(a)("nombretalla")
                dataGridView1.Item(6, a).Value = clase.dt.Tables("articulos").Rows(a)("ln1_nombre")
                cant1 = clase.dt.Tables("articulos").Rows(a)("com_unidades")
                dataGridView1.Item(7, a).Value = cant1
                dataGridView1.Item(8, a).Value = cant1
                dataGridView1.Item(9, a).Value = 0
            Next
            CheckBox1.Checked = True
        Else
            MessageBox.Show("El número del movimiento especificado no fue encontrado. Vuelva a intentarlo.", "MOVIMIENTO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
            TextBox3.Text = ""
            dataGridView1.RowCount = 0
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frm_impresion_codigos_de_barra_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With dataGridView1
            .Columns(0).ReadOnly = False
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = True
            .Columns(7).ReadOnly = True
            .Columns(8).ReadOnly = False
            .Columns(9).ReadOnly = False
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

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If dataGridView1.RowCount = 0 Then
            CheckBox1.Checked = False
            Exit Sub
        End If
        Dim c As Short
        Select Case CheckBox1.Checked
            Case False
                For c = 0 To dataGridView1.RowCount - 1
                    dataGridView1.Item(0, c).Value = False
                Next
            Case True
                For c = 0 To dataGridView1.RowCount - 1
                    dataGridView1.Item(0, c).Value = True
                Next
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        frm_asignar_cantidades_multiples.ShowDialog()
        frm_asignar_cantidades_multiples.Dispose()
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        Me.AcceptButton = Button2
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If ComboBox6.Text = "" Then
            MessageBox.Show("Debe seleccionar una impresora.", "SELECCIONAR IMPRESORA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If ComboBox1.Text = "" Then
            MessageBox.Show("Debe seleccionar una resolución.", "SELECCIONAR RESOLUCIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim b As Short
        Dim ind = False
        'rellenar espacios vacios con ceros(0)
        'For b = 0 To dataGridView1.RowCount - 1
        '    If dataGridView1.Item(6, b).Value = "" Then
        '        dataGridView1.Item(6, b).Value = "0"
        '    End If
        '    If dataGridView1.Item(7, b).Value = "" Then
        '        dataGridView1.Item(7, b).Value = "0"
        '    End If
        'Next
        For b = 0 To dataGridView1.RowCount - 1
            If dataGridView1.Item(0, b).Value = True Then
                If (Val(dataGridView1.Item(7, b).Value) = 0) Or (Val(dataGridView1.Item(8, b).Value) = 0) Then
                    MessageBox.Show("No ha especificado algunas cantidades. Por favor revise y vuelva a intentarlo.", "ESPECIFICAR CANTIDADES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ind = True
                    Exit For
                End If
            End If
        Next
        If ind = False Then
            Dim z As Short
            'TIQUETE DE  PLANET LOVE
            For z = 0 To dataGridView1.RowCount - 1
                If dataGridView1.Item(0, z).Value = True Then
                    Select Case ComboBox1.Text
                        Case "203 dpi"
                            ean13.imprimir_codigos_de_barra_19_x_15_203dpi(dataGridView1.Item(1, z).Value, dataGridView1.Item(8, z).Value, ComboBox6.Text)
                        Case "300 dpi"
                            ean13.imprimir_codigos_de_barra_19_x_15_300dpi(dataGridView1.Item(1, z).Value, dataGridView1.Item(8, z).Value, ComboBox6.Text)
                    End Select


                End If
            Next
            clase.actualizar("UPDATE informacion SET ultima_impresora_usada = '" & ComboBox6.Text & "'")
        End If
    End Sub



    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dataGridView1.EditingControlShowing
        ' obtener indice de la columna   
        Dim columna As Integer = dataGridView1.CurrentCell.ColumnIndex
        ' comprobar si la celda en edición corresponde a la columna 12
        If (columna = 6 Or columna = 7) Then
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

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.validar_numeros(e)
    End Sub


End Class