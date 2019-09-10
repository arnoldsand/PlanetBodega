Public Class frm_registro_dian_referencia
    Dim clase As New class_library
    Private WithEvents bs As New BindingSource
    Dim c5 As New MySqlDataAdapter
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_listado_importacion7.ShowDialog()
        frm_listado_importacion7.Dispose()
    End Sub

    Private Sub frm_registro_dian_referencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnCount = 9
        preparar_columnas()
    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox9, "Importación") = False Then Exit Sub
        llenar_grilla(TextBox1.Text)
    End Sub

    Sub llenar_grilla(ref As String)
        clase.Conectar()
        Dim cm5 As New MySqlCommand("SELECT detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, detalle_importacion_detcajas.detcab_composicion, detalle_importacion_detcajas.detcab_marca, proveedores.prv_codigoasignado, detalle_importacion_detcajas.detcab_cantidad, detalle_importacion_detcajas.detcab_unimedida, detalle_importacion_detcajas.detcab_registrodian, detalle_importacion_detcajas.detcab_fecharegistrodian, detalle_importacion_detcajas.detcab_coditem FROM  detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas  ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores  ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & TextBox9.Text & " AND detalle_importacion_detcajas.detcab_referencia LIKE '" & ref & "%' ) ORDER BY detalle_importacion_detcajas.detcab_referencia ASC, detalle_importacion_detcajas.detcab_descripcion ASC", clase.conex)
        c5.SelectCommand = cm5
        Dim dt5 As New DataTable
        c5.Fill(dt5)
        clase.desconectar()
        If dt5.Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = dt5
            preparar_columnas()
        Else

            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 9
            preparar_columnas()
        End If
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 120
            .Columns(1).Width = 120
            .Columns(2).Width = 150
            .Columns(3).Width = 100
            .Columns(4).Width = 120
            .Columns(5).Width = 50
            .Columns(6).Width = 80
            .Columns(7).Width = 180
            .Columns(8).Width = 90
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Composición"
            .Columns(3).HeaderText = "Marca"
            .Columns(4).HeaderText = "Proveedor"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "Medida"
            .Columns(7).HeaderText = "Registro DIAN"
            .Columns(8).HeaderText = "Fecha"
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        llenar_grilla(TextBox1.Text)
    End Sub

    'Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
    '    If Not bs.DataSource Is Nothing Then
    '        c5.Update(CType(bs.DataSource, DataTable))
    '        llenar_grilla(TextBox1.Text)
    '    End If
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            frm_dian_registro_dian.ShowDialog()
            frm_dian_registro_dian.Dispose()
        End If
    End Sub
End Class
