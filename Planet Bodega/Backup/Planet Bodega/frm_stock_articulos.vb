Public Class frm_stock_articulos
    Dim clase As New class_library
    Public codigos As Integer()
    Public codigosdebarra As String()
    Public coloresvector As Short()
    Public tallasvector As Short()


    Private Sub frm_stock_articulos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Referencia"
        

    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox9_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox10_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox10.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox11_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox11.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox12_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox12.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox13_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox13.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox14_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox14.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox15_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox15.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox17_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox17.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub textBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textBox2.KeyPress
        If ComboBox1.Text = "Codigo" Then
            clase.validar_numeros(e)
        End If
    End Sub


    Private Sub ComboBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        If ComboBox1.Text = "Codigo" Then
            Button1.Visible = False
            Button2.Visible = False
            Button6.Visible = False
            Button7.Visible = False
            TextBox1.Visible = True
            TextBox3.Visible = True
            TextBox15.Visible = True
            TextBox14.Visible = True
        End If
        If ComboBox1.Text = "Referencia" Then
            Button1.Visible = True
            Button2.Visible = True
            Button6.Visible = True
            Button7.Visible = True
            TextBox1.Visible = False
            TextBox3.Visible = False
            TextBox15.Visible = False
            TextBox14.Visible = False
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Select Case ComboBox1.Text
            Case "Codigo"
                If clase.validar_cajas_text(textBox2, "Codigo") = False Then Exit Sub
                If Len(textBox2.Text) = 13 Then
                    clase.consultar("SELECT articulos.ar_codigo, articulos.ar_codigobarras, articulos.ar_descripcion, articulos.ar_referencia, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, articulos.ar_precio1, articulos.ar_precio2, articulos.ar_costo, articulos.ar_activo, articulos.ar_fecdescontinua, colores.colornombre, tallas.nombretalla FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) LEFT JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) LEFT JOIN colores ON (articulos.ar_color = colores.cod_color) WHERE articulos.ar_codigobarras = '" & textBox2.Text & "'", "tbla")
                Else
                    clase.consultar("SELECT articulos.ar_codigo, articulos.ar_codigobarras, articulos.ar_descripcion, articulos.ar_referencia, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, articulos.ar_precio1, articulos.ar_precio2, articulos.ar_costo, articulos.ar_activo, articulos.ar_fecdescontinua, colores.colornombre, tallas.nombretalla FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (articulos.ar_sublinea4 = sublinea4.sl4_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) LEFT JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) LEFT JOIN colores ON (articulos.ar_color = colores.cod_color) WHERE articulos.ar_codigo = " & textBox2.Text & "", "tbla")
                End If
                If clase.dt.Tables("tbla").Rows.Count > 0 Then
                    TextBox1.Text = clase.dt.Tables("tbla").Rows(0)("ar_codigo")
                    TextBox3.Text = clase.dt.Tables("tbla").Rows(0)("ar_codigobarras")
                    TextBox5.Text = clase.dt.Tables("tbla").Rows(0)("ar_referencia")
                    TextBox4.Text = clase.dt.Tables("tbla").Rows(0)("ar_descripcion")
                    TextBox15.Text = clase.dt.Tables("tbla").Rows(0)("colornombre")
                    TextBox6.Text = clase.dt.Tables("tbla").Rows(0)("ar_precio1")
                    TextBox7.Text = clase.dt.Tables("tbla").Rows(0)("ar_precio2")
                    TextBox8.Text = clase.dt.Tables("tbla").Rows(0)("ar_costo")
                    TextBox13.Text = clase.dt.Tables("tbla").Rows(0)("ln1_nombre")
                    TextBox14.Text = clase.dt.Tables("tbla").Rows(0)("nombretalla")
                    TextBox12.Text = clase.dt.Tables("tbla").Rows(0)("sl1_nombre")
                    If IsDBNull(clase.dt.Tables("tbla").Rows(0)("sl2_nombre")) Then
                        TextBox11.Text = ""
                    Else
                        TextBox11.Text = clase.dt.Tables("tbla").Rows(0)("sl2_nombre")
                    End If
                    If IsDBNull(clase.dt.Tables("tbla").Rows(0)("sl3_nombre")) Then
                        TextBox10.Text = ""
                    Else
                        TextBox10.Text = clase.dt.Tables("tbla").Rows(0)("sl3_nombre")
                    End If
                    If IsDBNull(clase.dt.Tables("tbla").Rows(0)("sl4_nombre")) Then
                        TextBox9.Text = ""
                    Else
                        TextBox9.Text = clase.dt.Tables("tbla").Rows(0)("sl4_nombre")
                    End If
                    TextBox17.Text = clase.dt.Tables("tbla").Rows(0)("ar_activo")
                Else
                    MessageBox.Show("No se encontró ningun articulo con coincida con el codigo escrito. Pulse aceptar para volverlo a intentar.", "ARTICULO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    limpiar()
                End If
            Case "Referencia"

                If clase.validar_cajas_text(textBox2, "Referencia") = False Then Exit Sub
                clase.consultar("select* from articulos where ar_codigo = " & textBox2.Text & "", "table1")
                If clase.dt.Tables("table1").Rows.Count > 0 Then
                    'verificar si dos o mas articulos diferentes tienen la misma referencia y mostrar un menu para seleccionar con cual se va a trabajar
                    'llenar los arrays de carateristicas
                Else
                    MessageBox.Show("No se encontró ningun articulo con coincida con el referencia escrita. Pulse aceptar para volverlo a intentar.", "ARTICULO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    limpiar()
                End If
        End Select
    End Sub

    Private Sub llenar_colores(ByVal ref As String)
        clase.consultar("SELECT DISTINCT ar_color FROM articulos WHERE  ar_referencia = '" & ref & "'", "tbl1")
        If clase.dt.Tables("tbl1").Rows.Count > 0 Then
            Dim d As Short
            ReDim coloresvector(clase.dt.Tables("tbl1").Rows.Count)
            For d = 0 To clase.dt.Tables("tbl1").Rows.Count - 1
                coloresvector(d) = clase.dt.Tables("tbl1").Rows(d)("ar_color")
            Next
        Else
            ReDim coloresvector(0)
        End If
    End Sub

    Private Sub llenar_codigosdebarras(ByVal ref As String)
        clase.consultar("SELECT DISTINCT ar_codigobarras FROM articulos WHERE  ar_referencia = '" & ref & "'", "tbl1")
        If clase.dt.Tables("tbl1").Rows.Count > 0 Then
            Dim d As Short
            ReDim codigosdebarra(clase.dt.Tables("tbl1").Rows.Count)
            For d = 0 To clase.dt.Tables("tbl1").Rows.Count - 1
                codigosdebarra(d) = clase.dt.Tables("tbl1").Rows(d)("ar_codigobarras")
            Next
        Else
            ReDim codigosdebarra(0)
        End If
    End Sub

    Private Sub llenar_tallas(ByVal ref As String)
        clase.consultar("SELECT DISTINCT ar_talla FROM articulos WHERE  ar_referencia = '" & ref & "'", "tbl1")
        If clase.dt.Tables("tbl1").Rows.Count > 0 Then
            Dim d As Short
            ReDim tallasvector(clase.dt.Tables("tbl1").Rows.Count)
            For d = 0 To clase.dt.Tables("tbl1").Rows.Count - 1
                tallasvector(d) = clase.dt.Tables("tbl1").Rows(d)("ar_talla")
            Next
        Else
            ReDim tallasvector(0)
        End If
    End Sub

    Private Sub llenar_codigos(ByVal ref As String)
        clase.consultar("SELECT DISTINCT ar_codigo FROM articulos WHERE  ar_referencia = '" & ref & "'", "tbl1")
        If clase.dt.Tables("tbl1").Rows.Count > 0 Then
            Dim d As Short
            ReDim codigos(clase.dt.Tables("tbl1").Rows.Count)
            For d = 0 To clase.dt.Tables("tbl1").Rows.Count - 1
                codigos(d) = clase.dt.Tables("tbl1").Rows(d)("ar_codigo")
            Next
        Else
            ReDim codigos(0)
        End If
    End Sub


    Private Sub limpiar()
        For Each a As Control In Me.GroupBox1.Controls
            If TypeOf a Is TextBox Then
                a.Text = ""
            End If
        Next
        textBox2.Text = ""
        textBox2.Focus()
        DataGridView1.Rows.Clear()
        DataGridView1.RowCount = 1
    End Sub


End Class