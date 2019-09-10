Public Class frm_fotos_desde_hand_held
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        inicializar_cadena_access()
        conexion.Open()
        Dim pedido As Integer = 0
        Dim cm1 As New OleDbCommand("select* from detalle_entrada where codigoarticulo <> 'N'", conexion)
        Dim c1 As New OleDbDataAdapter
        c1.SelectCommand = cm1
        Dim dt4 As New DataSet
        c1.Fill(dt4, "transferencia")
        conexion.Close()
        If dt4.Tables("transferencia").Rows.Count > 0 Then
            Dim v As String = MessageBox.Show("Se importarán " & dt4.Tables("transferencia").Rows.Count & " articulos ¿Desea continuar?", "IMPORTAR ARTICULOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                pedido = dt4.Tables("transferencia").Rows(0)("pedido")
                TextBox1.Text = pedido
                TextBox2.Text = dt4.Tables("transferencia").Rows.Count
                Button4.Enabled = False
                Button1.Enabled = True
                Button2.Enabled = True
                clase.borradoautomatico("delete from tabla_temporal_maquinas where pedido = " & pedido & "")
                Dim a As Integer
                For a = 0 To dt4.Tables("transferencia").Rows.Count - 1
                    If verificar_existencia_articulo(dt4.Tables("transferencia").Rows(a)("codigoarticulo")) = True Then
                        clase.agregar_registro("insert into tabla_temporal_maquinas (pedido, codigo) values ('" & pedido & "', '" & dt4.Tables("transferencia").Rows(a)("codigoarticulo") & "')")
                    End If
                Next
                clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre FROM articulos INNER JOIN tabla_temporal_maquinas ON (articulos.ar_codigobarras = tabla_temporal_maquinas.codigo) INNER JOIN linea1  ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) WHERE (tabla_temporal_maquinas.pedido = " & pedido & ")", "tabla")
                If clase.dt.Tables("tabla").Rows.Count > 0 Then
                    DataGridView1.Columns.Clear()
                    DataGridView1.DataSource = clase.dt.Tables("tabla")
                    preparar_columnas()
                Else
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Columns.Clear()
                    preparar_columnas()
                End If
                Dim cm3 As New OleDbCommand("delete from detalle_entrada")
                conexion.Open()
                cm3.Connection = conexion
                cm3.ExecuteNonQuery()
                conexion.Close()
            End If
        Else
            MessageBox.Show("No se encontraron datos para importar.", "IMPOSIBLE IMPORTAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Linea"
            .Columns(4).HeaderText = "Sublinea"
            .Columns(0).Width = 100
            .Columns(1).Width = 150
            .Columns(2).Width = 180
            .Columns(3).Width = 120
            .Columns(4).Width = 120
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub frm_fotos_desde_hand_held_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnCount = 5
        preparar_columnas()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            clase.consultar("select ar_foto from articulos where ar_codigo = " & .Item(0, .CurrentCell.RowIndex).Value & "", "tablita")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.ColumnCount = 5
        preparar_columnas()
        TextBox1.Text = ""
        TextBox2.Text = ""
        PictureBox1.Image = Nothing
        Button4.Enabled = True
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 5
            preparar_columnas()
            TextBox1.Text = ""
            TextBox2.Text = ""
            PictureBox1.Image = Nothing
            MessageBox.Show("Articulos procesados satisfactoriamente.", "OPERACIÓN EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Button4.Enabled = True
            Button1.Enabled = False
            Button2.Enabled = False
        End If
    End Sub
End Class