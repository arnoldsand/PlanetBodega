Public Class frm_articulos
    Dim clase As class_library = New class_library
    Private Sub frm_articulos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With ComboBox1
            .Items.Add("Codigo")
            .Items.Add("Descripción")
            .Items.Add("Referencia")
            .Text = "Codigo"
        End With
        With dataGridView1
            .RowCount = 0
            .RowHeadersWidth = 4
        End With
        llenar_datagrid(textBox2.Text, "")
    End Sub


    Sub llenar_datagrid(ByVal txt As String, ByVal tipobusquedas As String)

        Dim complebus As String = ""
        Dim consulta As String = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, cabclasificacion.Clasificacion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2, articulos.ar_precioanterior, articulos.ar_precioanterior2, articulos.ar_precioanttixy1, articulos.ar_precioanttixy2, articulos.ar_fechaingreso, articulos.ar_tasadesc, articulos.ar_creadopor, articulos.ar_fecdescontinua, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4  FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (sublinea2.sl2_sl1codigo = sublinea1.sl1_codigo) AND (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (sublinea3.sl3_sl2codigo = sublinea2.sl2_codigo) AND (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) AND (sublinea1.sl1_ln1codigo = linea1.ln1_codigo) LEFT JOIN sublinea4 ON (sublinea4.sl4_sl3codigo = sublinea3.sl3_codigo) AND (articulos.ar_sublinea4 = sublinea4.sl4_codigo) INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) LEFT JOIN cabclasificacion ON (articulos.ar_IdClasificacion = articulos.ar_IdClasificacion)"
        'Dim agrupamiento As String = " GROUP BY articulos.ar_referencia, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre"
        Select Case tipobusquedas
            Case "Codigo"
                If txt <> "" Then
                    If Len(textBox2.Text) = 13 Then
                        complebus = consulta & " WHERE (articulos.ar_codigobarras = '" & txt & "')"
                    Else
                        complebus = consulta & " WHERE (articulos.ar_codigo = " & txt & ")"
                    End If
                Else
                    complebus = consulta
                End If
            Case "Descripción"
                complebus = consulta & " WHERE articulos.ar_descripcion like '" & txt & "%' ORDER BY articulos.ar_descripcion asc"
            Case "Referencia"
                complebus = consulta & " WHERE articulos.ar_referencia like '" & txt & "%' ORDER BY articulos.ar_referencia asc"
            Case ""
                complebus = consulta & " ORDER BY articulos.ar_fechaingreso DESC"
        End Select

        clase.consultar(complebus, "tblres")
        If clase.dt.Tables("tblres").Rows.Count > 0 Then

            dataGridView1.DataSource = clase.dt.Tables("tblres")
        Else

            dataGridView1.DataSource = Nothing
        End If
    End Sub


    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.GotFocus
        textBox2.BackColor = System.Drawing.SystemColors.GradientActiveCaption
    End Sub

    Private Sub textBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textBox2.KeyPress
        If ComboBox1.Text = "Codigo" Then
            clase.validar_numeros(e)
        End If
    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.LostFocus
        textBox2.BackColor = Color.White
    End Sub

    Private Sub textBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox2.TextChanged
        llenar_datagrid(textBox2.Text, ComboBox1.Text)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        frm_crear_articulo.ShowDialog()
        frm_crear_articulo.Dispose()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If IsNumeric(dataGridView1.CurrentCell.ColumnIndex) And IsNumeric(dataGridView1.CurrentCell.RowIndex) Then
            frm_ver_articulo.ShowDialog()
            frm_ver_articulo.Dispose()
        End If
    End Sub

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub
End Class