Public Class frm_diferencias_transferencia
    Dim clase As New class_library
    Public ValidaPassword As String
    Dim PasswordRevision As String

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        frm_password_revision.txtPassword.Text = ""
        frm_password_revision.ShowDialog()

        If frm_password_revision.ValidaPassword = PasswordRevision Then
            'ACTUALIZAR EL DATATABLE TABLADIFERENCIA CON LAS CORRECCIONES
            Dim ActDiferencia() As DataRow
            For i = 0 To grdDiferencia.RowCount - 1
                ActDiferencia = frmRevision.TablaDiferencia.Select("rev_articulo='" & grdDiferencia(1, i).Value.ToString & "'")
                ActDiferencia(0)("rev_faltante") = (Val(grdDiferencia(6, i).Value.ToString) - Val(grdDiferencia(5, i).Value.ToString))
            Next

            Me.Close()
        Else
            MessageBox.Show("Constraseña invalida. Intente nuevamente.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub frm_diferencias_transferencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT * FROM informacion", "password")
        PasswordRevision = clase.dt.Tables("password").Rows(0)("password_revision")

        With grdDiferencia
            .RowHeadersVisible = False
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .Columns(0).Visible = False
            .Columns(0).ReadOnly = True
            .Columns(1).Width = 70
            .Columns(1).HeaderText = "Codigo"
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).ReadOnly = True
            .Columns(2).Visible = False
            .Columns(2).ReadOnly = True
            .Columns(3).Width = 200
            .Columns(3).ReadOnly = True
            .Columns(4).Width = 200
            .Columns(4).ReadOnly = True
            .Columns(5).Width = 70
            .Columns(5).HeaderText = "Revisado"
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).ReadOnly = False
            .Columns(6).Width = 70
            .Columns(6).HeaderText = "Esperado"
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).ReadOnly = True
            .Columns(6).Visible = False
        End With
    End Sub

    Private Sub grdDiferencia_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles grdDiferencia.EditingControlShowing
        Dim validar As TextBox = CType(e.Control, TextBox)
        AddHandler validar.KeyPress, AddressOf validar_Keypress
    End Sub

    Private Sub validar_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim columna As Integer = grdDiferencia.CurrentCell.ColumnIndex
        If columna = 5 Then
            Dim caracter As Char = e.KeyChar
            If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
                Me.Text = e.KeyChar
                e.KeyChar = Chr(0)
            End If
        End If
    End Sub

    Private Sub btnRegresar_Click(sender As Object, e As EventArgs) Handles btnRegresar.Click
        frmRevision.Retorno = True
        frmRevision.TablaDiferencia.Rows.Clear()
        Me.Close()
    End Sub
End Class