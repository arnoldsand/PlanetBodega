Public Class frmdestinotransferencias
    Dim clase As New class_library
    Private _NroTransferencias As Integer
    Dim tiendaanterior As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Sub New(NroTransferencia As Integer)

        ' Llamada necesaria para el diseñador.
        InitializeComponent()
        _NroTransferencias = NroTransferencia
        
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub

    Private Sub frmdestinotransferencias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox9, "SELECT* FROM tiendas WHERE estado = TRUE ORDER BY tienda ASC", "tienda", "id")
        clase.consultar1("SELECT cabtransferencia.tr_destino, tiendas.tienda FROM cabtransferencia INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) WHERE cabtransferencia.tr_numero = " & _NroTransferencias & "", "trans")
        tiendaanterior = clase.dt1.Tables("trans").Rows(0)("tienda")
        ComboBox9.SelectedValue = clase.dt1.Tables("trans").Rows(0)("tr_destino")
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If clase.validar_combobox(ComboBox9, "Destino") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Contraseña") = False Then Exit Sub

        Dim CodLovePOS As String = ""
        clase.consultar("SELECT* FROM informacion", "clave")
        If clase.dt.Tables("clave").Rows.Count > 0 Then
            If TextBox1.Text = clase.dt.Tables("clave").Rows(0)("clave_edicion_transferencias") Then
                clase.consultar("SELECT* FROM cabtransferencia WHERE tr_numero = " & _NroTransferencias & "", "cabtrans")
                If clase.dt.Tables("cabtrans").Rows(0)("tr_finalizada") = False Then
                    MessageBox.Show("Debe finalizar la transferencia antes de modificar el destino de esta.", "FINALIZAR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                If clase.dt.Tables("cabtrans").Rows(0)("tr_destino") = ComboBox9.SelectedValue Then
                    MessageBox.Show("Para cambiar el destino de la transferencia debe seleccionar uno diferente al original.", "SELECCIONAR OTRO DESTINO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                CodLovePOS = clase.dt.Tables("cabtrans").Rows(0)("tr_codigolovepos")
                    If (IsDBNull(clase.dt.Tables("cabtrans").Rows(0)("tr_codigolovepos")) = False) Then
                        clase.ConsultarSQLServer("SELECT* FROM [dbo].[CabMovimientoInventario] WHERE IdEmpresa = 100 AND IdTienda = 99 AND IdBodega = 1 AND IdTipoMovimiento = 'TRS' AND IdCodMovimiento = '" & CodLovePOS & "'", "cab")
                        If (clase.dtSql.Tables("cab").Rows(0)("Recibido") = True) Then
                            MessageBox.Show("No se puede cambiar el destino de una transferencia que ya fue recibida.", "NO SE PUEDE CAMBIAR DESTINO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    End If
                    Dim parametro As String
                    If IsDBNull(clase.dt.Tables("cabtrans").Rows(0)("tr_auditoria")) Then
                        parametro = ""
                    Else
                        parametro = clase.dt.Tables("cabtrans").Rows(0)("tr_auditoria")
                    End If
                    parametro = "CAMBIO DE DESTINO: DESTINO ANTERIOR : " & tiendaanterior & " DESTINO NUEVO: " & ComboBox9.Text & " // " & parametro
                    If clase.dt.Tables("cabtrans").Rows(0)("tr_finalizada") = True Then
                    Dim CodigoLovePOS As String = CalcularCodigoTransferenciaSaliente(GetCodigoEmpresaByTienda(ComboBox9.SelectedValue), ComboBox9.SelectedValue)
                    If CodigoLovePOS <> "" Then
                            clase.actualizar("UPDATE cabtransferencia SET tr_destino = " & ComboBox9.SelectedValue & ", tr_auditoria = '" & parametro & "', tr_codigolovepos = '" & CodigoLovePOS & "' WHERE tr_numero = " & _NroTransferencias & "")
                        End If
                        AnularMovimientoDeInventario(100, 99, 1, "TRS", CodLovePOS)
                        GuardarMovimientoInventarioEnLovePOS(_NroTransferencias, CodigoLovePOS)
                    ActualizarConsetuvosDocumentosInventarios(GetCodigoEmpresaByTienda(ComboBox9.SelectedValue), ComboBox9.SelectedValue, CodigoLovePOS)
                Else
                        clase.actualizar("UPDATE cabtransferencia SET tr_destino = " & ComboBox9.SelectedValue & ", tr_auditoria = '" & parametro & "' WHERE tr_numero = " & _NroTransferencias & "")
                    End If
                    'If CalcularCodigoTransferencia(ComboBox9.SelectedValue) <> "" Then
                    '    ActualizarConsetuvosDocumentosInventarios(ComboBox9.SelectedValue)
                    'End If
                    MessageBox.Show("El destino de la transferencia se ha cambiado con exito.", "OPERACIÓN SATISFACTORIA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Close()
                Else
                    MessageBox.Show("La contraseña escrita es incorrecta. Vuelva a intentarlo.", "CONSTRASEÑA ERRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Text = ""
                TextBox1.Focus()
            End If
        End If
    End Sub


End Class