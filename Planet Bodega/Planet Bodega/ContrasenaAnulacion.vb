Public Class ContrasenaAnulacion
    Dim clase As New class_library
    Private _NroTransferencias As Integer

    Public Sub New(IdTransferencia As Integer)

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()
        _NroTransferencias = IdTransferencia
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If clase.validar_cajas_text(TextBox1, "Contraseña") = False Then Exit Sub
        clase.consultar("SELECT* FROM informacion", "clave")
        If clase.dt.Tables("clave").Rows.Count > 0 Then
            If TextBox1.Text = clase.dt.Tables("clave").Rows(0)("clave_edicion_transferencias") Then
                clase.consultar("select* from cabtransferencia where tr_numero = " & _NroTransferencias & "", "bus")
                If clase.dt.Tables("bus").Rows(0)("tr_revisada") = True Then
                    Dim CodigoLovePOS As String = clase.dt.Tables("bus").Rows(0)("tr_codigolovepos")
                    clase.ConsultarSQLServer("SELECT* FROM [dbo].[CabMovimientoInventario] WHERE IdEmpresa = 100 AND IdTienda = 99 AND IdBodega = 1 AND IdTipoMovimiento = 'TRS' AND IdCodMovimiento = '" & CodigoLovePOS & "' ", "sqlserver")
                    If clase.dtSql.Tables("sqlserver").Rows.Count > 0 Then
                        If clase.dtSql.Tables("sqlserver").Rows(0)("Anulado") = True Then
                            MessageBox.Show("No se puede anular una transferencia que ya fue anulada con anterioridad.", "TRANSFERENCIA YA ANULADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        If clase.dtSql.Tables("sqlserver").Rows(0)("Recibido") = True Then
                            MessageBox.Show("No se puede anular una transferencia que ya fue recibida.", "TRANSFERENCIA YA RECIBIDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        AnularMovimientoDeInventario(100, 99, 1, "TRS", CodigoLovePOS)
                        'clase.
                    End If
                End If
                clase.agregar_registro("DELETE FROM cabtransferencia WHERE tr_numero = " & _NroTransferencias & "")
                MessageBox.Show("La Transferencia se ha Anulado Satisfactoriamente.", "TRANSFERENCIA ANULADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Close()
            Else
                MessageBox.Show("La contraseña escrita es incorrecta. Vuelva a intentarlo.", "CONSTRASEÑA ERRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Text = ""
                TextBox1.Focus()
            End If
        End If
    End Sub
End Class