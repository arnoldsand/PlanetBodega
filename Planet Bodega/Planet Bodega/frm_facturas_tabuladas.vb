Public Class frm_facturas_tabuladas
    Dim clase As New class_library
    Private Sub frm_facturas_tabuladas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
       
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_listado_importacion6.ShowDialog()
        frm_listado_importacion6.Dispose()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        llenar_grilla()
    End Sub

    Sub llenar_grilla()
        clase.consultar("SELECT cabfacturas.cabfact_codigo, proveedores.prv_codigoasignado, proveedores.prv_ciudad, COUNT(detfacturas.detfact_codigo), cabfacturas.cabfact_fecha, cabfacturas.cabfect_funcionario FROM detfacturas INNER JOIN cabfacturas ON (detfacturas.detfact_factura = cabfacturas.cabfact_codigo) INNER JOIN proveedores ON (cabfacturas.cabfact_proveedor = proveedores.prv_codigo) INNER JOIN cabimportacion  ON (cabfacturas.cabfact_importacion = cabimportacion.imp_codigo) WHERE (cabfacturas.cabfact_importacion =" & textBox2.Text & ") GROUP BY cabfacturas.cabfact_codigo", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            dtgridcajas.DataSource = clase.dt.Tables("tabla")
            preparar_columnas()
        Else
            dtgridcajas.DataSource = Nothing
            'preparar_columnas()
        End If
    End Sub

    Sub preparar_columnas()
        With dtgridcajas
            .Columns(0).Width = 100
            .Columns(1).Width = 150
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).HeaderText = "Número Factura"
            .Columns(1).HeaderText = "Proveedor"
            .Columns(2).HeaderText = "Ciudad"
            .Columns(3).HeaderText = "Items"
            .Columns(4).HeaderText = "Fecha"
            .Columns(5).HeaderText = "Funcionario"
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If dtgridcajas.Rows.Count = 0 Then
            Exit Sub
        End If
        frm_ver_detalle_factura.ShowDialog()
        frm_ver_detalle_factura.Dispose()
    End Sub
End Class