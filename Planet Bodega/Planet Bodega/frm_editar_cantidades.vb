Public Class frm_editar_cantidades
    Dim clase As New class_library

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_editar_cantidades_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, entradamercancia.com_unidades FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) WHERE (articulos.ar_codigo =" & frmInforme.grdImportacion.Item(0, frmInforme.grdImportacion.CurrentRow.Index).Value & " AND entradamercancia.com_codigoimp =" & frmInforme.ConsImport & ")", "entrada")
        If clase.dt.Tables("entrada").Rows.Count > 0 Then
            TextBox1.Text = clase.dt.Tables("entrada").Rows(0)("ar_codigo")
            TextBox2.Text = clase.dt.Tables("entrada").Rows(0)("ar_referencia")
            TextBox3.Text = clase.dt.Tables("entrada").Rows(0)("ar_descripcion")
            TextBox4.Text = clase.dt.Tables("entrada").Rows(0)("com_unidades")
        End If
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox4, "Cantidad") = False Then Exit Sub
        frm_password_liquidacion.ShowDialog()
        If frm_password_liquidacion.ValidaPassword1 = pass_liquidacion() Then
            clase.actualizar("UPDATE entradamercancia set com_unidades = " & TextBox4.Text & " where com_codigoart = " & frmInforme.grdImportacion.Item(0, frmInforme.grdImportacion.CurrentRow.Index).Value & " and com_codigoimp = " & frmInforme.ConsImport & "")
            MessageBox.Show("Las cantidades se actualizaron satisfactoriamente.", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information)
            frm_password_liquidacion.Dispose()

            Me.Close()
        Else
            MessageBox.Show("Contraseña incorrecta.", "CONTRASEÑA INVALIDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            frm_password_liquidacion.Dispose()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub

    Function pass_liquidacion() As String
        clase.consultar1("select password_liquidacion from informacion", "tabla")
        Return clase.dt1.Tables("tabla").Rows(0)("password_liquidacion")
    End Function
End Class