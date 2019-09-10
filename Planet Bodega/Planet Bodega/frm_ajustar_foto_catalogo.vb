Public Class frm_ajustar_foto_catalogo
    Dim column1 As New System.Windows.Forms.DataGridViewCheckBoxColumn
    Dim column2 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim column3 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim column4 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim column5 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim column6 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim column7 As New System.Windows.Forms.DataGridViewImageColumn
    Dim clase As New class_library
    Dim c, d As Integer
    Dim pag As Integer
    Dim sql As String


    Private Sub frm_ajustar_foto_catalogo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox9.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        d = 0
        DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.column1, Me.column2, Me.column3, Me.column4, Me.column5, Me.column6, Me.column7})
        clase.llenar_combo(ComboBox1, "select* from linea1 order by ln1_nombre ASC", "ln1_nombre", "ln1_codigo")
        With DataGridView1
            .Columns(1).HeaderText = "Codigo"
            .Columns(2).HeaderText = "Referencia"
            .Columns(3).HeaderText = "Descripción"
            .Columns(4).HeaderText = "Línea"
            .Columns(5).HeaderText = "Sublinea"
            .Columns(6).HeaderText = "Foto"
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Width = 25
            .Columns(1).Width = 80
            .Columns(2).Width = 120
            .Columns(3).Width = 120
            .Columns(4).Width = 120
            .Columns(5).Width = 120
            .Columns(6).Width = 150
            .Columns(0).ReadOnly = False
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = True

        End With
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        Select Case ComboBox9.SelectedIndex
            Case 0
                ComboBox1.Visible = True
                TextBox18.Visible = False
                TextBox18.Text = ""
                ComboBox1.SelectedIndex = -1
                ComboBox1.Focus()
                Label2.Text = "Línea"
                DataGridView1.RowCount = 0
                ComboBox2.Enabled = True
                ComboBox2.SelectedIndex = 0
            Case 1
                ComboBox1.Visible = False
                TextBox18.Visible = True
                TextBox18.Text = ""
                ComboBox1.SelectedIndex = -1
                TextBox18.Focus()
                Label2.Text = "Codigo"
                DataGridView1.RowCount = 0
                ComboBox2.SelectedIndex = -1
                ComboBox2.Enabled = False
        End Select
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Select Case ComboBox9.SelectedIndex
            Case 0
                If clase.validar_combobox(ComboBox1, "Línea") = False Then Exit Sub

                Select Case ComboBox2.SelectedIndex
                    Case 0
                        sql = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, articulos.ar_foto FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) WHERE (articulos.ar_linea =" & ComboBox1.SelectedValue & " and articulos.ar_activo = TRUE) ORDER BY linea1.ln1_nombre, sublinea1.sl1_nombre, articulos.ar_referencia ASC"
                    Case 1
                        sql = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, articulos.ar_foto FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) WHERE (articulos.ar_linea =" & ComboBox1.SelectedValue & " and articulos.ar_activo = FALSE) ORDER BY linea1.ln1_nombre, sublinea1.sl1_nombre, articulos.ar_referencia ASC"
                End Select
                llenar_listado(sql, 0)
            Case 1
                If clase.validar_cajas_text(TextBox18, "Codigo") = False Then Exit Sub
                sql = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, articulos.ar_foto FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) WHERE (articulos.ar_codigo =" & TextBox18.Text & ") ORDER BY linea1.ln1_nombre, sublinea1.sl1_nombre, articulos.ar_referencia ASC"
                llenar_listado(sql, 0)
                clase.consultar("select ar_activo from articulos where ar_codigo = " & TextBox18.Text & "", "bus")
                Select Case clase.dt.Tables("bus").Rows(0)("ar_activo")
                    Case False
                        ComboBox2.Text = "Inactivos"
                    Case True
                        ComboBox2.Text = "Activos"
                End Select
        End Select
    End Sub


    Private Sub llenar_listado(sql2 As String, ini As Integer)
       clase.consultar(sql2, "tabla1")
        If clase.dt.Tables("tabla1").Rows.Count > 0 Then
            Dim i As Integer
            DataGridView1.RowCount = 0
            For i = 0 To clase.dt.Tables("tabla1").Rows.Count - 1
                With clase.dt.Tables("tabla1")
                    DataGridView1.RowCount = DataGridView1.RowCount + 1
                    DataGridView1.Item(0, i).Value = obtener_estado_catalogo(.Rows(i)("ar_codigo"))
                    DataGridView1.Item(1, i).Value = .Rows(i)("ar_codigo")
                    DataGridView1.Item(2, i).Value = .Rows(i)("ar_referencia")
                    DataGridView1.Item(3, i).Value = .Rows(i)("ar_descripcion")
                    DataGridView1.Item(4, i).Value = .Rows(i)("ln1_nombre")
                    DataGridView1.Item(5, i).Value = .Rows(i)("sl1_nombre")
                    If System.IO.File.Exists(.Rows(i)("ar_foto")) Then
                        DataGridView1.Item(6, i).Value = Image.FromFile(.Rows(i)("ar_foto"))
                        'DataGridView1.Item(6, i).Value = "C:\Users\Arnold\Pictures\" & .Rows(i)("ar_codigo") & ".jpg"

                    Else
                        DataGridView1.Item(6, i).Value = Global.WindowsApplication1.My.Resources.Resources.sinfoto
                    End If
                    SetImage1(DataGridView1.Item(6, i))
                    DataGridView1.Rows(i).Height = 120
                End With
            Next
        Else

            DataGridView1.RowCount = 0
        End If
    End Sub

    Private Sub TextBox18_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox18.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim v As String = ""
        Select Case ComboBox2.SelectedIndex
            Case 0
                v = MessageBox.Show("¿Desea Deshabilitar los articulos seleccionados?", "DESHABILITAR ARTICULOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            Case 1
                v = MessageBox.Show("¿Desea Habilitar los articulos seleccionados?", "HABILITAR ARTICULOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        End Select
        If v = 6 Then
            Dim a As Integer
            Dim cont As Integer = 0
            For a = 0 To DataGridView1.RowCount - 1
                Application.DoEvents()
                If DataGridView1.Item(0, a).Value = True Then
                    cont = cont + 1
                    clase.consultar("select ar_activo from articulos where ar_codigo = " & DataGridView1.Item(1, a).Value & "", "busqueda")
                    Dim fecha As Date = Now
                    Select Case clase.dt.Tables("busqueda").Rows(0)("ar_activo")
                        Case False
                            clase.actualizar("update articulos set ar_activo = TRUE, ar_fecdescontinua = 'NULL' where ar_codigo = " & DataGridView1.Item(1, a).Value & "")
                        Case True
                            clase.actualizar("update articulos set ar_activo = FALSE, ar_fecdescontinua = '" & fecha.ToString("yyyy-MM-dd") & "' where ar_codigo = " & DataGridView1.Item(1, a).Value & "")
                            clase.borradoautomatico("delete from articulos_para_catalogo where codigo = " & DataGridView1.Item(1, a).Value & "")
                    End Select
                End If
            Next
            If cont = 0 Then
                MessageBox.Show("Debe seleccionar por los menos un item para habilitar-deshabilitar. Pulse aceptar para volverlo a intentar.", "SELECCIONAR ITEM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                llenar_listado(sql, 0)
            End If
        End If
    End Sub

    Function obtener_estado_catalogo(codigo As Integer) As Boolean
        clase.consultar1("select* from articulos_para_catalogo where codigo = " & codigo & "", "tabla3")
        If clase.dt1.Tables("tabla3").Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim v As String = MessageBox.Show("¿Desea guardar los cambios para el catalogo?", "MODIFICAR CATALOGO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            Dim a As Integer
            Dim cont As Integer = 0
            For a = 0 To DataGridView1.RowCount - 1
                Application.DoEvents()
                If DataGridView1.Item(0, a).Value = True Then
                    cont = cont + 1
                    establecer_estado(DataGridView1.Item(1, a).Value, True)
                Else
                    establecer_estado(DataGridView1.Item(1, a).Value, False)
                End If
            Next
            llenar_listado(sql, 0)
        End If
    End Sub

    Private Sub establecer_estado(codigo As Integer, estado As Boolean)
        clase.consultar("select* from articulos_para_catalogo where codigo = " & codigo & "", "tabl")
        If clase.dt.Tables("tabl").Rows.Count > 0 Then
            Select Case estado
                Case False
                    clase.borradoautomatico("delete from articulos_para_catalogo where codigo = " & codigo & "")

            End Select
        Else
            Select Case estado
                Case True
                    clase.agregar_registro("insert into articulos_para_catalogo (codigo) values ('" & codigo & "') ")
            End Select
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Select Case ComboBox2.SelectedIndex
            Case 0
                Button9.Enabled = True
                DataGridView1.RowCount = 0
            Case 1
                Button9.Enabled = False
                DataGridView1.RowCount = 0
        End Select
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        llenar_listado(sql, d - 1)
    End Sub
End Class
