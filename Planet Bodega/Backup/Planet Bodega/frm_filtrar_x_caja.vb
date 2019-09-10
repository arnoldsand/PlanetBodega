Public Class frm_filtrar_x_caja
    Dim clase As New class_library
    Dim campos As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Dim listado(26) As String

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        Me.Close()
    End Sub

    Private Sub frm_filtrar_x_caja_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        For i = 0 To frm_crear_articulos_lotes.dataGridView1.Columns.Count - 1
            ComboBox1.Items.Add(frm_crear_articulos_lotes.dataGridView1.Columns(i).HeaderText)
        Next
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "" Then
            Exit Sub
        End If
        Dim parametro As String = convertir_encabezado_comumna_enletra(ComboBox1.Text)
        Dim sql As String = "Select distinct " & parametro & " from detalleimportacion where codigo_importacion = " & cod_importacion & " order by " & parametro & " asc"
        clase.consultar(sql, "tabla6")
        If clase.dt.Tables("tabla6").Rows.Count > 0 Then
            ListBox1.DataSource = clase.dt.Tables("tabla6")
            ListBox1.DisplayMember = parametro
        End If
    End Sub

    Function convertir_encabezado_comumna_enletra(ByVal txt As String) As String
        If ind_cargador_lote = True Then
            Dim a As Short
            convertir_encabezado_comumna_enletra = ""
            For a = 0 To frm_crear_articulos_lotes.dataGridView1.Columns.Count - 1
                listado(a) = frm_crear_articulos_lotes.dataGridView1.Columns(a).HeaderText
            Next
            a = 0
            Do While a <= frm_crear_articulos_lotes.dataGridView1.Columns.Count - 1
                If listado(a) = txt Then
                    convertir_encabezado_comumna_enletra = campos(a)
                    Exit Do
                End If
                a += 1
            Loop
        Else
            convertir_encabezado_comumna_enletra = txt
        End If
    End Function

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        If ComboBox1.Text = "" Then
            Exit Sub
        End If
        Dim sql As String = "select A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z from detalleimportacion where codigo_importacion = " & cod_importacion & " and " & convertir_encabezado_comumna_enletra(ComboBox1.Text) & " = '" & ListBox1.Text & "'"
        clase.consultar(sql, "tabal3")
        If clase.dt.Tables("tabal3").Rows.Count > 0 Then
            frm_crear_articulos_lotes.dataGridView1.DataSource = clase.dt.Tables("tabal3")
        Else
            frm_crear_articulos_lotes.dataGridView1.DataSource = Nothing
        End If
    End Sub
End Class