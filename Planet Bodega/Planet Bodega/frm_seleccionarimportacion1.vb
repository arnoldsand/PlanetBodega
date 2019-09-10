Public Class frm_seleccionarimportacion1
    Dim clase As New class_library
    Public CodImport, NomImport As String

    Private Sub frm_seleccionarimportacion_Load(sender As Object, e As EventArgs) Handles Me.Load
        llenar_combo()
    End Sub

    Sub llenar_combo()
        Dim c1 As New MySqlDataAdapter("select imp_nombrefecha as NOMBRE_IMPORTACION, imp_codigo from cabimportacion order by imp_fecha desc LIMIT 0, 10", clase.conex)
        Dim dt As New DataTable
        clase.Conectar()
        c1.SelectCommand.Connection = clase.conex
        c1.Fill(dt)
        clase.desconectar()
        If dt.Rows.Count > 0 Then
            grdImportacion.DataSource = dt
            With grdImportacion
                .Columns(0).Width = 350
                .Columns(1).Visible = False
            End With
        Else
            grdImportacion.DataSource = Nothing
        End If
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If grdImportacion.RowCount = 0 Then
            Exit Sub
        End If
        frmNuevaOrden.txtImportacion.Text = NomImport
        Me.Close()
    End Sub

    Private Sub grdImportacion_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdImportacion.CellClick
        'OBTENEMOS EL CAMPO SELECCIONADO EN EL DATAGRIDVIEW PARA REALIZAR EL PROCESO DE ANULACION
        Dim y As Integer = grdImportacion.CurrentCell.RowIndex
        NomImport = grdImportacion(0, [y]).Value.ToString()
        CodImport = grdImportacion(1, [y]).Value.ToString()
    End Sub
End Class