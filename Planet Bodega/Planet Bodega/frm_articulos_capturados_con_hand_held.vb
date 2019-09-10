Public Class frm_articulos_capturados_con_hand_held
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_articulos_capturados_con_hand_held_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT pedido, COUNT(codigo) FROM tabla_temporal_maquinas GROUP BY pedido ORDER BY pedido ASC", "busqueda")
        If clase.dt.Tables("busqueda").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("busqueda")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 2
            preparar_columnas()
        End If
    End Sub

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Pedido"
            .Columns(1).HeaderText = "Cant"
            .Columns(0).Width = 120
            .Columns(1).Width = 120
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim v As String = MessageBox.Show("¿Desea agregar los articulos capturados al catálogo?", "AGREGAR ARTICULOS AL HAND HELD", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            clase.borradoautomatico("delete from articulos_para_catalogo")
            clase.consultar("SELECT DISTINCT codigo FROM tabla_temporal_maquinas", "tabla1")
            Dim i As Integer
            For i = 0 To clase.dt.Tables("tabla1").Rows.Count - 1
                clase.agregar_registro("insert into articulos_para_catalogo (codigo) values ('" & convertir_codigobarra_a_codigo_normal(clase.dt.Tables("tabla1").Rows(i)("codigo")) & "')")
            Next
            clase.borradoautomatico("delete from tabla_temporal_maquinas")
            MessageBox.Show("Articulos subidos satisfactoriamente.", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            DataGridView1.DataSource = Nothing
            Button2.Enabled = False
        End If
    End Sub
End Class