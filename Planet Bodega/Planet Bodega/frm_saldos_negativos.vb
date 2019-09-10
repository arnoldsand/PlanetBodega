Public Class frm_saldos_negativos
    Dim clase As New class_library
    Dim ind As Boolean
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
        ind = False
    End Sub

    Private Sub frm_saldos_negativos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ind = True
        dataGridView1.ColumnCount = 7
        preparar_columnas()
    End Sub


    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs)
        If ind = True Then
            
        End If
    End Sub

    Private Sub preparar_columnas()
        With dataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"

            .Columns(5).HeaderText = "Cant"
            .Columns(0).Width = 80
            .Columns(1).Width = 100
            .Columns(2).Width = 200
            .Columns(3).Width = 80
            .Columns(4).Width = 80

            .Columns(5).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

   

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, inventario_bodega.inv_cantidad FROM inventario_bodega INNER JOIN articulos ON (inventario_bodega.inv_codigoart = articulos.ar_codigo) INNER JOIN colores  ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas  ON (tallas.codigo_talla = articulos.ar_talla) WHERE (inventario_bodega.inv_cantidad <0) ORDER BY articulos.ar_referencia ASC, articulos.ar_descripcion ASC", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("tabla")
            preparar_columnas()
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 7
            preparar_columnas()
        End If
    End Sub
End Class