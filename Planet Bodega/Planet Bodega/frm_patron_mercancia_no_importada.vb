Public Class frm_patron_mercancia_no_importada
    Dim clase As New class_library
    Dim columna As New System.Windows.Forms.DataGridViewCheckBoxColumn
    Dim inddeubicacioncolumna As Boolean
    Dim indguardado As Boolean = False
    Dim codigoaeditar As Short

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_orden_de_produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_grid1(TextBox3.Text)
        inddeubicacioncolumna = False
        If ind_formulario_patron_distribucion = True Then
            btnAceptar.Enabled = False
        End If
    End Sub

    Private Sub llenar_grid1(parametro As String)
        clase.consultar("SELECT cabpatrondist.pat_codigo, cabpatrondist.pat_nombre, cabpatrondist.pat_fecha, SUM(detpatrondist.dp_cantidad), cabpatrondist.pat_codigo FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo) WHERE cabpatrondist.pat_nombre like '" & UCase(parametro) & "%' GROUP BY cabpatrondist.pat_codigo ORDER BY cabpatrondist.pat_nombre ASC, cabpatrondist.pat_fecha DESC", "patron")
        If clase.dt.Tables("patron").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("patron")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 5
            preparar_columnas()
        End If
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 50
            .Columns(1).Width = 150
            .Columns(2).Width = 80
            .Columns(3).Width = 50
            .Columns(4).Visible = False
            .Columns(0).HeaderText = "No"
            .Columns(1).HeaderText = "Nombre"
            .Columns(2).HeaderText = "Fecha"
            .Columns(3).HeaderText = "Cant"
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub DataGridView2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit
        calcular_total()
    End Sub

    Sub calcular_total()
        Dim a As Short
        Dim sum As Short = 0
        For a = 0 To DataGridView2.RowCount - 1
            sum = sum + Val(DataGridView2.Item(2, a).Value)
        Next
        TextBox2.Text = sum
    End Sub

    Private Sub DataGridView2_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView2.EditingControlShowing
        ' obtener indice de la columna   
        Dim columna As Integer = DataGridView2.CurrentCell.ColumnIndex
        ' comprobar si la celda en edición corresponde a la columna 12
        If (columna = 2) Then
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Select Case indguardado
            Case True
                If clase.validar_cajas_text(TextBox1, "Nombre Patrón") = False Then Exit Sub
                Dim c As Short
                Dim ind As Boolean = False
                For c = 0 To DataGridView2.RowCount - 1
                    If DataGridView2.Item(2, c).Value <> vbEmpty Then
                        ind = True
                    End If
                Next
                If ind = True Then
                    Dim consecutivo As Integer = hallar_consecutivo()
                    clase.agregar_registro("INSERT INTO `cabpatrondist`(`pat_codigo`,`pat_nombre`,`pat_fecha`) VALUES ( '" & consecutivo & "','" & UCase(TextBox1.Text) & "','" & Now.ToString("yyyy-MM-dd") & "')")
                    For c = 0 To DataGridView2.RowCount - 1
                        If DataGridView2.Item(2, c).Value <> vbEmpty Then
                            clase.agregar_registro("INSERT INTO `detpatrondist`(`dp_idpatron`,`dp_tienda`,`dp_cantidad`,`dp_codigo`) VALUES ( '" & consecutivo & "','" & DataGridView2.Item(0, c).Value & "','" & DataGridView2.Item(2, c).Value & "',NULL)")
                        End If
                    Next
                    estado_formulario_despues_de_guardar()
                Else
                    MessageBox.Show("Debe establecer por lo menos una cantidad a un punto de venta para generar un patrón de distribución.", "GENERAR PATRÓN DE DISTRIBUCIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    DataGridView2.Focus()
                End If
            Case False
                clase.actualizar("UPDATE cabpatrondist set pat_nombre = '" & UCase(TextBox1.Text) & "' Where pat_codigo = " & codigoaeditar & "")
                Dim a As Short
                For a = 0 To DataGridView2.RowCount - 1
                    If Val(DataGridView2.Item(2, a).Value) <> 0 Then
                        clase.consultar("select* from detpatrondist where dp_idpatron = " & codigoaeditar & " and dp_tienda = " & DataGridView2.Item(0, a).Value & "", "busquedapatron")
                        If clase.dt.Tables("busquedapatron").Rows.Count > 0 Then
                            clase.actualizar("update detpatrondist set dp_cantidad = " & DataGridView2.Item(2, a).Value & " where dp_idpatron = " & codigoaeditar & " and dp_tienda = " & DataGridView2.Item(0, a).Value & "")
                        Else
                            clase.agregar_registro("insert into detpatrondist (dp_idpatron, dp_tienda, dp_cantidad) values ('" & codigoaeditar & "', '" & DataGridView2.Item(0, a).Value & "', '" & DataGridView2.Item(2, a).Value & "')")
                        End If
                    End If
                Next
                estado_formulario_despues_de_guardar()
        End Select

    End Sub

    Private Sub estado_formulario_despues_de_guardar()
        DataGridView2.Rows.Clear()
        TextBox1.Text = ""
        llenar_grid1(TextBox3.Text)
        DataGridView1.Enabled = True
        Button3.Enabled = False
        Button6.Enabled = False
        Button2.Enabled = True
        TextBox2.Text = 0
        Button1.Enabled = False
        Button4.Enabled = True
        TextBox3.Enabled = True
    End Sub

    Function hallar_consecutivo() As Integer
        clase.consultar("SELECT MAX(pat_codigo) AS maximo FROM cabpatrondist", "consecutivo")
        If clase.dt.Tables("consecutivo").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("consecutivo").Rows(0)("maximo")) Then
                Return 1
            Else
                Return clase.dt.Tables("consecutivo").Rows(0)("maximo") + 1
            End If
        Else
            Return 1
        End If
    End Function

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        clase.consultar("SELECT tiendas.id, tiendas.tienda, detpatrondist.dp_cantidad FROM tiendas INNER JOIN detpatrondist ON (tiendas.id = detpatrondist.dp_tienda) WHERE (detpatrondist.dp_idpatron =" & DataGridView1.Item(4, DataGridView1.CurrentCell.RowIndex).Value & ") ORDER BY tiendas.tienda ASC", "patrondist")
        If clase.dt.Tables("patrondist").Rows.Count > 0 Then
            DataGridView2.Columns.Clear()
            DataGridView2.DataSource = clase.dt.Tables("patrondist")
            preparar_columnas2()
            'columna.Width = 30
            'DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.columna})
            ' DataGridView2.Columns(2).Width = 50
            inddeubicacioncolumna = True
            CheckBox1.Visible = False
            DataGridView2.Columns(2).ReadOnly = True
            calcular_total()
            TextBox1.Text = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
        Else
            CheckBox1.Visible = False
            DataGridView2.DataSource = Nothing
            DataGridView2.ColumnCount = 3
            preparar_columnas2()
            TextBox1.Text = ""

        End If
        Dim Y As Integer = DataGridView1.CurrentCell.RowIndex
        frm_ingreso_no_importado.Patron = DataGridView1(0, [Y]).Value.ToString
        frm_ingreso_no_importado.NombrePatron = DataGridView1(1, [Y]).Value.ToString
    End Sub

    Private Sub preparar_columnas2()
        With DataGridView2
            .Columns(0).Visible = False
            .Columns(1).Width = 200
            .Columns(2).Width = 50
            .Columns(1).HeaderText = "Tienda"
            .Columns(1).ReadOnly = True
            .Columns(2).HeaderText = "Cant Unds"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox1.Text = ""
        DataGridView2.DataSource = Nothing
        DataGridView2.ColumnCount = 3
        preparar_columnas2()
        DataGridView2.Rows.Clear()
        llenar_grid1(TextBox3.Text)
        preparar_columnas()
        Button2.Enabled = True
        Button3.Enabled = False
        Button6.Enabled = False
        DataGridView1.Enabled = True
        CheckBox1.Visible = False
        btnAceptar.Enabled = True
        Button4.Enabled = True
        TextBox3.Enabled = True
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If frm_ingreso_no_importado.Patron = "" Then
            MessageBox.Show("DEBE SELECCIONAR UN PATRO DE DISTRIBUCION", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        frm_ingreso_no_importado.indicador_presionador_de_boton = True
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        indguardado = True
        CheckBox1.Visible = True
        TextBox3.Text = ""
        TextBox1.Text = ""
        Button1.Enabled = True
        DataGridView1.Enabled = False
        DataGridView1.DataSource = Nothing
        DataGridView1.ColumnCount = 5
        preparar_columnas()
        DataGridView1.Rows.Clear()
        TextBox1.Focus()
        TextBox2.Text = "0"
        Button2.Enabled = False
        Button3.Enabled = True
        Button4.Enabled = False
        Button6.Enabled = True
        DataGridView2.DataSource = Nothing
        DataGridView2.Columns.Clear()
        DataGridView2.ColumnCount = 3
        preparar_columnas2()
        columna.Width = 30
        DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewCheckBoxColumn() {Me.columna})
        DataGridView2.Rows.Clear()
        clase.consultar("SELECT * FROM tiendas WHERE (estado =TRUE) ORDER BY tienda ASC", "tiendas")
        If clase.dt.Tables("tiendas").Rows.Count Then
            'columna.Width = 30
            'DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewCheckBoxColumn() {Me.columna})
            Dim a As Integer
            For a = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
                With DataGridView2
                    .RowCount = .RowCount + 1
                    .Item(0, a).Value = clase.dt.Tables("tiendas").Rows(a)("Id")
                    .Item(1, a).Value = clase.dt.Tables("tiendas").Rows(a)("tienda")
                End With
            Next
            ' ubicarcheck()
        End If
    End Sub

    'Private Sub ubicarcheck()
    '    If inddeubicacioncolumna = True Then

    '        CheckBox1.Location = New System.Drawing.Point(376, 101)
    '    Else

    '        CheckBox1.Location = New System.Drawing.Point(626, 101)
    '    End If
    'End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        Dim a As Short
        Select Case CheckBox1.Checked
            Case True
                For a = 0 To DataGridView2.RowCount - 1
                    DataGridView2.Item(3, a).Value = True
                Next
            Case False
                For a = 0 To DataGridView2.RowCount - 1
                    DataGridView2.Item(3, a).Value = False
                Next
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        asignar_cant_patron_dist.ShowDialog()
        asignar_cant_patron_dist.Dispose()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            DataGridView1.Enabled = False
            Button2.Enabled = False
            Button4.Enabled = False
            btnAceptar.Enabled = False
            Button3.Enabled = True
            Button6.Enabled = True
            codigoaeditar = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value
            TextBox1.Text = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
            Button1.Enabled = True
            TextBox3.Enabled = False
            CheckBox1.Visible = True
            indguardado = False
            DataGridView2.DataSource = Nothing
            DataGridView2.ColumnCount = 3
            preparar_columnas2()
            clase.consultar("select* from tiendas where estado = TRUE order by tienda ASC", "tiendas")
            If clase.dt.Tables("tiendas").Rows.Count > 0 Then
                Dim a As Short
                For a = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
                    DataGridView2.RowCount = DataGridView2.RowCount + 1
                    DataGridView2.Item(0, a).Value = clase.dt.Tables("tiendas").Rows(a)("id")
                    DataGridView2.Item(1, a).Value = clase.dt.Tables("tiendas").Rows(a)("tienda")
                    DataGridView2.Item(2, a).Value = cantidad_del_patron(clase.dt.Tables("tiendas").Rows(a)("id"), codigoaeditar)
                Next
                columna.Width = 30
                DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.columna})
                DataGridView2.Columns(2).ReadOnly = False
            End If


        End If
    End Sub

    Function cantidad_del_patron(tienda As Short, patron As Short) As Short
        clase.consultar1("SELECT dp_cantidad FROM detpatrondist WHERE (dp_tienda =" & tienda & " AND dp_idpatron =" & patron & ")", "cant")
        If clase.dt1.Tables("cant").Rows.Count > 0 Then
            Return clase.dt1.Tables("cant").Rows(0)("dp_cantidad")
        Else
            Return vbEmpty
        End If
    End Function

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        llenar_grid1(TextBox3.Text)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class