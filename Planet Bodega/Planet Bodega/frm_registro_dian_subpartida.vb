Public Class frm_registro_dian_subpartida
    Dim clase As New class_library
    Dim Subpartida, Fecha As String
    Dim Fbase As Date

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtSubPartida, "SubPartida") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Importación") = False Then Exit Sub
        LlenarGrid()
        GroupBox1.Enabled = True
        txtSubPartida.Text = ""
    End Sub

    Public Sub LlenarGrid()
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, detalle_importacion_detcajas.detcab_composicion, detalle_importacion_detcajas.detcab_marca, proveedores.prv_codigoasignado, detalle_importacion_detcajas.detcab_cantidad, detalle_importacion_detcajas.detcab_unimedida, detalle_importacion_detcajas.detcab_registrodian, detalle_importacion_detcajas.detcab_fecharegistrodian,detalle_importacion_detcajas.detcab_coditem FROM  detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas  ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores  ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detalle_importacion_detcajas.detcab_subpartida ='" & txtSubPartida.Text & "' AND detalle_importacion_cabcajas.det_codigoimportacion = " & TextBox1.Text & ")", "consulta")

        With grdRegistro
            .DataSource = clase.dt.Tables("consulta")
            .Columns(0).Width = 120
            .Columns(1).Width = 120
            .Columns(2).Width = 150
            .Columns(3).Width = 100
            .Columns(4).Width = 120
            .Columns(5).Width = 50
            .Columns(6).Width = 80
            .Columns(7).Width = 180
            .Columns(8).Width = 100
            .Columns(9).Visible = False
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Composición"
            .Columns(3).HeaderText = "Marca"
            .Columns(4).HeaderText = "Proveedor"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "Medida"
            .Columns(7).HeaderText = "Registro Dian"
            .Columns(8).HeaderText = "Fecha Registro"
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = True
            .Columns(7).ReadOnly = True
        End With
    End Sub

    Private Sub btnAsignar_Click(sender As Object, e As EventArgs) Handles btnAsignar.Click
        If clase.validar_cajas_text(txtRegistro, "Registro Dian") = False Then Exit Sub
        For i = 0 To grdRegistro.RowCount - 1
            grdRegistro(7, i).Value = txtRegistro.Text
            grdRegistro(8, i).Value = dtpFechaRegistro.Value.Date.ToString("yyyy-MM-dd")
        Next
        btnGuardar.Enabled = True
        Subpartida = txtSubPartida.Text
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Fecha = Fbase.ToString("yyyy-MM-dd")
        For i = 0 To grdRegistro.RowCount - 1
            If IsDBNull(grdRegistro(7, i).Value) Then
                MessageBox.Show("Debe asignar un registro de importación.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            clase.actualizar("UPDATE detalle_importacion_detcajas SET detcab_registrodian='" & grdRegistro(7, i).Value.ToString & "', detcab_fecharegistrodian='" & Fecha & "' WHERE detcab_coditem='" & grdRegistro(9, i).Value.ToString & "'")
        Next
        MessageBox.Show("Registro asignado con éxito.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Limpiar()
    End Sub

    Private Sub btnDeshacer_Click(sender As Object, e As EventArgs) Handles btnDeshacer.Click
        Limpiar()
    End Sub

    Public Sub Limpiar()
        grdRegistro.DataSource = Nothing
        btnGuardar.Enabled = False
        txtRegistro.Text = ""
        txtSubPartida.Text = ""
        txtSubPartida.Focus()
        GroupBox1.Enabled = False
    End Sub

    Private Sub dtpFechaRegistro_ValueChanged(sender As Object, e As EventArgs) Handles dtpFechaRegistro.ValueChanged
        Fbase = dtpFechaRegistro.Value.Date
    End Sub

    Private Sub txtSubPartida_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSubPartida.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub txtRegistro_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtRegistro.KeyPress
        clase.enter(e)
    End Sub

    Private Sub frm_registro_dian_subpartida_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_listado_importacion10.ShowDialog()
        frm_listado_importacion10.Dispose()
    End Sub
End Class