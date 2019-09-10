Public Class frm_ver_salida
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ver_salida_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim parametro As Integer = frm_flujo_mercancia.dtgridcajas.Item(0, frm_flujo_mercancia.dtgridcajas.CurrentRow.Index).Value
        clase.consultar("select* from cabsalidas_mercancia where cabsal_cod = " & parametro & "", "rest")
        TextBox1.Text = parametro
        TextBox2.Text = clase.dt.Tables("rest").Rows(0)("cabsal_fecha")
        TextBox3.Text = clase.dt.Tables("rest").Rows(0)("cabsal_liquidador") & ""
        ' MsgBox("SELECT detalle_importacion_detcajas.detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_descripcion AS Descripcion, detsalidas_mercancia.det_cant AS Cantidad, detsalidas_mercancia.det_procesado AS Procesado FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detsalidas_mercancia.det_salidacodigo =" & parametro & ")")
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_descripcion AS Descripcion, detsalidas_mercancia.det_cant AS Cantidad, detsalidas_mercancia.det_procesado AS Procesado FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detsalidas_mercancia.det_salidacodigo =" & parametro & ")", "tabla123")
        If clase.dt.Tables("tabla123").Rows.Count > 0 Then
            dtgridcajas.DataSource = clase.dt.Tables("tabla123")
            preparar_columnas()
        Else
            dtgridcajas.DataSource = Nothing
            dtgridcajas.ColumnCount = 4
            preparar_columnas()
        End If
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub preparar_columnas()
        With dtgridcajas
            .Columns(0).Width = 150
            .Columns(1).Width = 200
            .Columns(2).Width = 80
            .Columns(3).Width = 80
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripcion"
            .Columns(2).HeaderText = "Cantidad"
            .Columns(3).HeaderText = "Procesado"
        End With
    End Sub
End Class