Public Class registros_importacion_articulo
    Dim clase As New class_library
    Dim bloqueado As Boolean

  
    Private Sub registros_importacion_articulo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        textBox2.Focus()
        dtgridcajas.ColumnCount = 5
        preparar_columnas()
    End Sub

    Private Sub preparar_columnas()
        With dtgridcajas
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Registro"
            .Columns(4).HeaderText = "Fecha"
            .Columns(0).Width = 80
            .Columns(1).Width = 120
            .Columns(2).Width = 180
            .Columns(3).Width = 200
            .Columns(4).Width = 100
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedIndex
            Case 0
                Label2.Text = "Codigo:"
                Button6.Visible = False
                bloqueado = False
                textBox2.Text = ""
                textBox2.Focus()
            Case 1
                Label2.Text = "Importación:"
                Button6.Visible = True
                bloqueado = True
                textBox2.Text = ""
            Case 2
                Label2.Text = "Referencia:"
                Button6.Visible = False
                bloqueado = False
                textBox2.Focus()
                textBox2.Text = ""
        End Select

    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        llenar_grilla()
    End Sub

    Sub llenar_grilla()
        Dim criterio As String = ""
        Select Case ComboBox1.SelectedIndex
            Case 0
                If clase.validar_cajas_text(textBox2, "Codigo Articulo") = False Then Exit Sub
                criterio = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, registros_de_importaciones.registro, registros_de_importaciones.fecha FROM registros_de_importaciones RIGHT JOIN articulos ON (registros_de_importaciones.articulo = articulos.ar_codigo) WHERE (articulos.ar_codigo =" & textBox2.Text & ") ORDER BY articulos.ar_referencia, articulos.ar_descripcion ASC"
            Case 1
                If clase.validar_cajas_text(textBox2, "Importación") = False Then Exit Sub
                criterio = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, registros_de_importaciones.registro, registros_de_importaciones.fecha FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN asociaciones_codigos  ON (asociaciones_codigos.asc_precodref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN articulos ON (asociaciones_codigos.asc_postcodart = articulos.ar_codigo) INNER JOIN registros_de_importaciones ON (registros_de_importaciones.articulo = articulos.ar_codigo) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & textBox2.Text & ") ORDER BY articulos.ar_referencia, articulos.ar_descripcion ASC"
            Case 2
                If clase.validar_cajas_text(textBox2, "Referencia") = False Then Exit Sub
                criterio = "SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, registros_de_importaciones.registro, registros_de_importaciones.fecha FROM registros_de_importaciones RIGHT JOIN articulos ON (registros_de_importaciones.articulo = articulos.ar_codigo) WHERE (articulos.ar_referencia like '" & textBox2.Text & "%') ORDER BY articulos.ar_referencia, articulos.ar_descripcion ASC"
        End Select
        clase.consultar(criterio, "registros")
        If clase.dt.Tables("registros").Rows.Count > 0 Then
            dtgridcajas.Columns.Clear()
            dtgridcajas.DataSource = clase.dt.Tables("registros")
            preparar_columnas()
        Else
            MessageBox.Show("Los criterios proporcionados no generaron resultados.", "SIN RESULTADOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtgridcajas.DataSource = Nothing
            dtgridcajas.ColumnCount = 5
            preparar_columnas()
        End If
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        If bloqueado = True Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If

    End Sub

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        If ComboBox1.SelectedIndex = 0 Then
            clase.validar_numeros(e)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_listado_importacion9.ShowDialog()
        frm_listado_importacion9.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If dtgridcajas.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dtgridcajas.CurrentCell.RowIndex) Then
            frm_ver_detalles.ShowDialog()
            frm_ver_detalles.Dispose()
        End If
    End Sub
End Class