Public Class frm_detalle_devolucion
    Dim clase As New class_library
    Private Sub frm_detalle_transferencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT tiendas.tienda, cabtransferencia.tr_operador, SUM(dt_costo * dt_cantidad) AS Costo, SUM(dt_venta1 * dt_cantidad) AS Precio1, SUM(dt_venta2 * dt_cantidad) AS Precio2 FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas  ON (cabtransferencia.tr_destino = tiendas.id) WHERE (cabtransferencia.tr_numero =" & frm_mantenimiento_devolucion.DataGridView2.Item(0, frm_mantenimiento_devolucion.DataGridView2.CurrentRow.Index).Value & ") GROUP BY dettransferencia.dt_trnumero", "tabla5")
        TextBox1.Text = frm_mantenimiento_devolucion.DataGridView2.Item(0, frm_mantenimiento_devolucion.DataGridView2.CurrentRow.Index).Value
        TextBox5.Text = clase.dt.Tables("tabla5").Rows(0)("tienda")
        TextBox18.Text = clase.dt.Tables("tabla5").Rows(0)("tr_operador")
        TextBox2.Text = FormatCurrency(clase.dt.Tables("tabla5").Rows(0)("Costo"), 0)
        TextBox3.Text = FormatCurrency(clase.dt.Tables("tabla5").Rows(0)("Precio1"), 0)
        TextBox4.Text = FormatCurrency(clase.dt.Tables("tabla5").Rows(0)("Precio2"), 0)
        clase.consultar("SELECT dettransferencia.dt_codarticulo AS Codigo, articulos.ar_referencia AS Referencia, articulos.ar_descripcion AS Descripcion, dettransferencia.dt_cantidad AS Cant, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS PrecioCosto, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS PrecioVenta1, FORMAT(SUM(dt_venta2 * dt_cantidad), 'Currency') AS PrecioVenta1 FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) WHERE (dettransferencia.dt_trnumero =" & frm_mantenimiento_devolucion.DataGridView2.Item(0, frm_mantenimiento_devolucion.DataGridView2.CurrentRow.Index).Value & ") GROUP BY dettransferencia.dt_codarticulo", "tabla6")
        If clase.dt.Tables("tabla6").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("tabla6")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 70
            .Columns(1).Width = 120
            .Columns(2).Width = 160
            .Columns(3).Width = 50
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox18_GotFocus(sender As Object, e As EventArgs) Handles TextBox18.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        clase.consultar("SELECT tr_revisada FROM cabtransferencia WHERE (tr_numero =" & TextBox1.Text & ")", "transferencia")
        construir_archivo(TextBox1.Text)
        MessageBox.Show("Mensaje enviado con exito", "Planet Love", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            Dim v As String = MessageBox.Show("¿Desea Imprimir la transferencia seleccionada?", "IMPRIMIR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                imprimir_hoja_transferencia(TextBox1.Text)
            End If
        End If
    End Sub
End Class