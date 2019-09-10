Public Class frm_tienda
    Dim clase As New class_library
    Private Sub frm_tienda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_tienda(devolver_valor(ComboBox2.Text))
        ComboBox2.SelectedIndex = 0
    End Sub

    Sub llenar_tienda(actividad As Boolean)
        clase.consultar("SELECT tiendas.tienda, tiendas.ciudad, empresas.nombre_empresa, tiendas.email, tiendas.telefono1, tiendas.telefono2, tiendas.direccion, tiendas.administrador, tiendas.local , tiendas.id FROM empresas INNER JOIN tiendas ON (empresas.cod_empresa = tiendas.empresa) WHERE (tiendas.estado =" & actividad & ") ORDER BY tiendas.tienda ASC", "resultados")
        If clase.dt.Tables("resultados").Rows.Count > 0 Then
            datagridreferencia.Columns.Clear()
            datagridreferencia.DataSource = clase.dt.Tables("resultados")
            preparar_columnas()
        Else
            datagridreferencia.DataSource = Nothing
            datagridreferencia.ColumnCount = 10
            preparar_columnas()
        End If
    End Sub

    Private Sub preparar_columnas()
        With datagridreferencia
            .Columns(0).Width = 150
            .Columns(1).Width = 100
            .Columns(2).Width = 150
            .Columns(3).Width = 200
            .Columns(4).Width = 120
            .Columns(5).Width = 120
            .Columns(6).Width = 150
            .Columns(7).Width = 180
            .Columns(8).Width = 50
            .Columns(9).Width = 1
            .Columns(0).HeaderText = "Tienda"
            .Columns(1).HeaderText = "Ciudad"
            .Columns(2).HeaderText = "Empresa"
            .Columns(3).HeaderText = "E-mail"
            .Columns(4).HeaderText = "Telefono 1"
            .Columns(5).HeaderText = "Teléfono 2"
            .Columns(6).HeaderText = "Dirección"
            .Columns(7).HeaderText = "Administrador"
            .Columns(8).HeaderText = "Local"
        End With
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm_nueva_tienda.ShowDialog()
        frm_nueva_tienda.Dispose()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If IsNumeric(datagridreferencia.CurrentRow.Index) Then
            frm_ver_tienda.ShowDialog()
            frm_ver_tienda.Dispose()
        End If
    End Sub

   
    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        llenar_tienda(devolver_valor(ComboBox2.Text))
    End Sub

    Function devolver_valor(para As String) As Boolean
        If ComboBox2.Text = "Activa" Then
            Return True
        End If
        If ComboBox2.Text = "Inactiva" Then
            Return False
        End If
    End Function

End Class