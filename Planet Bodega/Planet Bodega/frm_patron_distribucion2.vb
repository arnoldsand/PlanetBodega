Public Class frm_patron_distribucion2
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_orden_de_produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_grid1()
    End Sub

    Private Sub llenar_grid1()
        clase.consultar("SELECT cabpatrondist.pat_codigo, cabpatrondist.pat_nombre, cabpatrondist.pat_fecha, SUM(detpatrondist.dp_cantidad), cabpatrondist.pat_codigo FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo) GROUP BY cabpatrondist.pat_codigo ORDER BY cabpatrondist.pat_nombre ASC, cabpatrondist.pat_fecha DESC", "patron")
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
            DataGridView2.Rows.Clear()
            TextBox1.Text = ""
            llenar_grid1()
            DataGridView1.Enabled = True
            Button3.Enabled = False
            Button6.Enabled = False
            Button2.Enabled = True
        Else
            MessageBox.Show("Debe establecer por lo menos una cantidad a un punto de venta para generar un patrón de distribución.", "GENERAR PATRÓN DE DISTRIBUCIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DataGridView2.Focus()
        End If
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
        Else
            DataGridView2.DataSource = Nothing
            DataGridView2.ColumnCount = 3
            preparar_columnas2()
        End If
        Dim Y As Integer = DataGridView1.CurrentCell.RowIndex
        frmIngresoNoImportados.Patron = DataGridView1(0, [Y]).Value.ToString
        frmIngresoNoImportados.NombrePatron = DataGridView1(1, [Y]).Value.ToString
    End Sub

    Private Sub preparar_columnas2()
        With DataGridView2
            .Columns(0).Visible = False
            .Columns(1).Width = 200
            .Columns(2).Width = 50
            .Columns(1).HeaderText = "Tienda"
            .Columns(2).HeaderText = "Cant Unds"
        End With
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox1.Text = ""
        DataGridView2.DataSource = Nothing
        DataGridView2.ColumnCount = 3
        preparar_columnas2()
        DataGridView2.Rows.Clear()
        llenar_grid1()
        preparar_columnas()
        Button2.Enabled = True
        Button3.Enabled = False
        Button6.Enabled = False
        DataGridView1.Enabled = True
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If frmIngresoNoImportados.Patron = "" Then
            MessageBox.Show("DEBE SELECCIONAR UN PATRO DE DISTRIBUCION", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Enabled = False
        DataGridView1.DataSource = Nothing
        DataGridView1.ColumnCount = 5
        preparar_columnas()
        DataGridView1.Rows.Clear()
        TextBox1.Focus()
        Button2.Enabled = False
        Button3.Enabled = True
        Button6.Enabled = True
        DataGridView2.DataSource = Nothing

        DataGridView2.ColumnCount = 3
        preparar_columnas2()
        DataGridView2.Rows.Clear()
        clase.consultar("SELECT * FROM tiendas WHERE (estado =TRUE) ORDER BY tienda ASC", "tiendas")
        If clase.dt.Tables("tiendas").Rows.Count Then

            Dim a As Integer
            For a = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
                With DataGridView2
                    .RowCount = .RowCount + 1
                    .Item(0, a).Value = clase.dt.Tables("tiendas").Rows(a)("Id")
                    .Item(1, a).Value = clase.dt.Tables("tiendas").Rows(a)("tienda")
                End With
            Next
        End If
    End Sub
End Class
