Public Class frm_generar_secciones
    Dim clase As New class_library
    Dim ind As Boolean

    Private Sub preparar_columnas()
        With dataGridView1
            .Columns(0).HeaderText = "Sección"
            .Columns(1).HeaderText = "Bodega"
            .Columns(2).HeaderText = "Góndola"
            .Columns(3).HeaderText = "Operario"
            .Columns(4).HeaderText = "Procesado"
            .Columns(0).Width = 100
            .Columns(1).Width = 100
            .Columns(2).Width = 80
            .Columns(3).Width = 150
            .Columns(4).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = False
            .Columns(4).ReadOnly = True
        End With
    End Sub

    Private Sub frm_generar_secciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ind = False
        clase.llenar_combo(ComboBox1, "select* from bodegas order by bod_codigo asc", "bod_nombre", "bod_codigo")
        ComboBox2.Text = "No Capturadas"
        ComboBox3.Text = "Conteo 1"
        clase.consultar("SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = FALSE) ORDER BY secciones_inventario.secc_codigo ASC;", "todos")
        If clase.dt.Tables("todos").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("todos")
            preparar_columnas()
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 5
            preparar_columnas()
        End If
        ind = True
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ind = True Then
            TextBox1.Text = ""
            TextBox18.Text = ""
            Dim sql As String = ""
            Select Case ComboBox2.Text
                Case "No Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = FALSE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = TRUE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Todas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & ") ORDER BY secciones_inventario.secc_codigo ASC;"
            End Select
            clase.consultar(sql, "combo")
            If clase.dt.Tables("combo").Rows.Count > 0 Then
                dataGridView1.Columns.Clear()
                dataGridView1.DataSource = clase.dt.Tables("combo")
                preparar_columnas()
            Else
                dataGridView1.DataSource = Nothing
                dataGridView1.ColumnCount = 5
                preparar_columnas()
            End If
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If ind = True Then
            ComboBox1.SelectedIndex = -1
            TextBox18.Text = ""
            TextBox1.Text = ""
            ComboBox1.Enabled = False
            TextBox18.Enabled = False
            TextBox1.Enabled = False
            Dim sql As String = ""
            Select ComboBox2.Text
                Case "No Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = FALSE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = TRUE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Todas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & ") ORDER BY secciones_inventario.secc_codigo ASC;"
            End Select
            clase.consultar(sql, "todos")
            If clase.dt.Tables("todos").Rows.Count > 0 Then
                dataGridView1.Columns.Clear()
                dataGridView1.DataSource = clase.dt.Tables("todos")
                preparar_columnas()
            Else
                dataGridView1.DataSource = Nothing
                dataGridView1.ColumnCount = 5
                preparar_columnas()
            End If
        End If
        ind = True
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        ind = False
        ComboBox1.SelectedIndex = -1
        TextBox18.Text = ""
        TextBox1.Text = ""
        ComboBox1.Enabled = True
        TextBox18.Enabled = True
        TextBox1.Enabled = True
        dataGridView1.DataSource = Nothing
        dataGridView1.ColumnCount = 5
        preparar_columnas()
        ind = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsNumeric(dataGridView1.CurrentCell.RowIndex) Then
            If dataGridView1.Item(4, dataGridView1.CurrentCell.RowIndex).Value = True Then
                MessageBox.Show("La sección seleccionada ya fue finalizada.", "SECCIÓN YA FINALIZADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                clase.actualizar("UPDATE secciones_inventario SET secc_procesado = TRUE WHERE secc_codigo = " & dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value & " AND secc_inventario = " & frm_modulo_inventario.TextBox1.Text & "")
                MessageBox.Show("La sección fue finalizada exitosamente.", "SECCIÓN FINALIZADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                procedimiento_combo2()
            End If
        End If
    End Sub
    
    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        If ind = True Then
           procedimiento_textbox()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If ind = True Then
            procedimiento_textbox()
        End If
    End Sub

    Private Sub procedimiento_textbox()
        Dim sql As String = ""
        Select Case ComboBox2.Text
            Case "No Capturadas"
                sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_gondola like '" & UCase(TextBox18.Text) & "%' AND secciones_inventario.secc_operario like '" & UCase(TextBox1.Text) & "%' AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = FALSE) ORDER BY secciones_inventario.secc_codigo ASC;"
            Case "Capturadas"
                sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_gondola like '" & UCase(TextBox18.Text) & "%' AND secciones_inventario.secc_operario like '" & UCase(TextBox1.Text) & "%' AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = TRUE) ORDER BY secciones_inventario.secc_codigo ASC;"
            Case "Todas"
                sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_gondola like '" & UCase(TextBox18.Text) & "%' AND secciones_inventario.secc_operario like '" & UCase(TextBox1.Text) & "%' AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & ") ORDER BY secciones_inventario.secc_codigo ASC;"
        End Select
        clase.consultar(sql, "combo")
        If clase.dt.Tables("combo").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("combo")
            preparar_columnas()
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 5
            preparar_columnas()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ind = True Then
            procedimiento_combo2()
        End If
    End Sub

    Private Sub procedimiento_combo2()
        If RadioButton1.Checked = True Then
            Dim sql As String = ""
            Select Case ComboBox2.Text
                Case "No Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = FALSE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = TRUE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Todas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & ") ORDER BY secciones_inventario.secc_codigo ASC;"
            End Select
            clase.consultar(sql, "combo")
            If clase.dt.Tables("combo").Rows.Count > 0 Then
                dataGridView1.Columns.Clear()
                dataGridView1.DataSource = clase.dt.Tables("combo")
                preparar_columnas()
            Else
                dataGridView1.DataSource = Nothing
                dataGridView1.ColumnCount = 5
                preparar_columnas()
            End If
        End If
        If RadioButton2.Checked = True Then
            Dim sql As String = ""
            Select Case ComboBox2.Text
                Case "No Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_gondola like '" & UCase(TextBox18.Text) & "%' AND secciones_inventario.secc_operario like '" & UCase(TextBox1.Text) & "%' AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = FALSE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Capturadas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_gondola like '" & UCase(TextBox18.Text) & "%' AND secciones_inventario.secc_operario like '" & UCase(TextBox1.Text) & "%' AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & " AND secciones_inventario.secc_procesado = TRUE) ORDER BY secciones_inventario.secc_codigo ASC;"
                Case "Todas"
                    sql = "SELECT secciones_inventario.secc_codigo, bodegas.bod_abreviatura, secciones_inventario.secc_gondola, secciones_inventario.secc_operario, secciones_inventario.secc_procesado FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_inventario =" & frm_modulo_inventario.TextBox1.Text & " AND secciones_inventario.secc_bodega = " & ComboBox1.SelectedValue & " AND secciones_inventario.secc_gondola like '" & UCase(TextBox18.Text) & "%' AND secciones_inventario.secc_operario like '" & UCase(TextBox1.Text) & "%' AND secciones_inventario.secc_conteo = " & ComboBox3.SelectedIndex + 1 & ") ORDER BY secciones_inventario.secc_codigo ASC;"
            End Select
            clase.consultar(sql, "combo")
            If clase.dt.Tables("combo").Rows.Count > 0 Then
                dataGridView1.Columns.Clear()
                dataGridView1.DataSource = clase.dt.Tables("combo")
                preparar_columnas()
            Else
                dataGridView1.DataSource = Nothing
                dataGridView1.ColumnCount = 5
                preparar_columnas()
            End If
        End If
    End Sub

    Private Sub dataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridView1.CellEndEdit
        If dataGridView1.Item(3, dataGridView1.CurrentCell.RowIndex).Value <> "" Then
            clase.actualizar("UPDATE secciones_inventario SET secc_operario = '" & UCase(dataGridView1.Item(3, dataGridView1.CurrentCell.RowIndex).Value) & "' WHERE secc_codigo = " & dataGridView1.Item(0, dataGridView1.CurrentCell.RowIndex).Value & "")
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ind = True Then
            procedimiento_combo2()
        End If
    End Sub
End Class