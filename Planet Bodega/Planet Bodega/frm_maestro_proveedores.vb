Public Class frm_maestro_proveedores
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_maestro_proveedores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_proveedores("")
    End Sub

    Sub llenar_proveedores(parametro As String)
        clase.consultar("SELECT prv_codigoasignado, prv_nombre, prv_ciudad, prv_contacto, prv_telefono1, prv_telefono2, prv_email1, prv_email2, prv_direccion, prv_web, prv_codigo FROM proveedores WHERE (prv_nombre LIKE '" & parametro & "%') ORDER BY prv_nombre ASC", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            datagridreferencia.Columns.Clear()
            datagridreferencia.DataSource = clase.dt.Tables("tabla")
            preparar_columnas()
        Else
            datagridreferencia.DataSource = Nothing
        End If
    End Sub

    Private Sub textBox2_TextChanged(sender As Object, e As EventArgs) Handles textBox2.TextChanged
        llenar_proveedores(textBox2.Text)
    End Sub

    Sub preparar_columnas()
        With datagridreferencia
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Compañia"
            .Columns(2).HeaderText = "Ciudad"
            .Columns(3).HeaderText = "Contacto"
            .Columns(4).HeaderText = "Telefeno 1"
            .Columns(5).HeaderText = "Telefono 2"
            .Columns(6).HeaderText = "E-mail 1"
            .Columns(7).HeaderText = "E-mail 2"
            .Columns(8).HeaderText = "Direccion"
            .Columns(9).HeaderText = "Pagina Web"
            .Columns(0).Width = 80
            .Columns(1).Width = 150
            .Columns(2).Width = 100
            .Columns(3).Width = 150
            .Columns(4).Width = 100
            .Columns(5).Width = 100
            .Columns(6).Width = 160
            .Columns(7).Width = 160
            .Columns(8).Width = 130
            .Columns(9).Width = 130
            .Columns(10).Width = 1
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm_nuevo_proveedor.ShowDialog()
        frm_nuevo_proveedor.Dispose()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If datagridreferencia.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(datagridreferencia.CurrentRow.Index) Then
            frm_ver_proveedor.ShowDialog()
            frm_ver_proveedor.Dispose()
        End If
    End Sub
End Class