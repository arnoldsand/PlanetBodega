Public Class frm_edicion_cantidad
    Dim clase As New class_library
    Private _NroTransferencia As Integer
    Private _Articulo As Integer
    Private _Cantidad As Integer

    Sub New(NroTransferencia As Integer, Articulo As Integer, Cantidad As Integer)

        ' Llamada necesaria para el diseñador.
        InitializeComponent()
        _NroTransferencia = NroTransferencia
        _Articulo = Articulo
        _Cantidad = Cantidad
        TextBox1.Text = Cantidad
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(TextBox1, "Cantidad") = False Then Exit Sub
        If clase.validar_cajas_text(txtcontrasena, "Contraseña") = False Then Exit Sub
        If TextBox1.Text.Length > 3 Then
            MessageBox.Show("No puede especificar un valor de mas de tres digitos.", "VALIDACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Text = ""
            TextBox1.Focus()
        Else
            clase.consultar("SELECT* FROM informacion", "clave")
            If clase.dt.Tables("clave").Rows.Count > 0 Then
                If txtcontrasena.Text = clase.dt.Tables("clave").Rows(0)("clave_edicion_transferencias") Then
                    clase.consultar("SELECT dt.*, cb.tr_codigolovepos, cb.tr_revisada FROM dettransferencia dt INNER JOIN cabtransferencia cb ON (dt.dt_trnumero = cb.tr_numero) WHERE dt.dt_trnumero = " & _NroTransferencia & " AND dt.dt_codarticulo = " & _Articulo & "", "cabtrans")
                    Dim parametro As String
                    If IsDBNull(clase.dt.Tables("cabtrans").Rows(0)("dt_auditoria")) Then
                        parametro = ""
                    Else
                        parametro = clase.dt.Tables("cabtrans").Rows(0)("dt_auditoria")
                    End If
                    parametro = "CAMBIO DE CANTIDAD: CANTIDAD ANTERIOR : " & _Cantidad & " CANTIDAD NUEVA: " & TextBox1.Text & " // " & parametro
                    clase.actualizar("UPDATE dettransferencia SET dt_cantidad = " & TextBox1.Text & ", dt_auditoria = '" & parametro & "' WHERE dt_trnumero = " & _NroTransferencia & " AND dt_codarticulo = " & _Articulo & "")
                    If clase.dt.Tables("cabtrans").Rows(0)("tr_revisada") = True Then
                        Dim CodigoLovePOS As String = clase.dt.Tables("cabtrans").Rows(0)("tr_codigolovepos")
                        clase.ConsultarSQLServer("SELECT* FROM [dbo].[CabMovimientoInventario] WHERE IdEmpresa = 100 AND IdTienda = 99 AND IdBodega = 1 AND IdTipoMovimiento = 'TRS' AND IdCodMovimiento = '" & CodigoLovePOS & "'", "sqlbus")
                        If clase.dtSql.Tables("sqlbus").Rows(0)("Recibido") = True Then
                            MessageBox.Show("No se Puede Modificar una Transferencias que ya fue Recibida.", "TRANSFERENCIA YA RECIBIDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        Dim IdEmpresaDestino As Integer = clase.dtSql.Tables("sqlbus").Rows(0)("IdEmpresaDestino")
                        Dim IdTiendaDestino As Integer = clase.dtSql.Tables("sqlbus").Rows(0)("IdTiendaDestino")
                        Dim IdBodegaDestino As Integer = clase.dtSql.Tables("sqlbus").Rows(0)("IdBodegaDestino")
                        Dim Diferencia As Integer = _Cantidad - TextBox1.Text
                        clase.ActualizarSQLServer("UPDATE [dbo].[DetMovimientoInventario] SET Cantidad = " & TextBox1.Text & ", FechaModif = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "', EquipoModif = '" & Environment.MachineName & "' WHERE IdEmpresa = 100 AND IdTienda = 99 AND IdBodega = 1 AND IdTipoMovimiento = 'TRS' AND IdCodMovimiento = '" & CodigoLovePOS & "' AND IdArticulo = " & _Articulo & "")
                        ConsolidarArticulo(100, 99, 1, _Articulo, Diferencia, Diferencia * clase.dt.Tables("cabtrans").Rows(0)("dt_costo"), Diferencia * clase.dt.Tables("cabtrans").Rows(0)("dt_costo2"))
                        clase.ActualizarSQLServer("UPDATE [dbo].[DetMovimientoInventario] SET Cantidad = " & TextBox1.Text & ", FechaModif = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "', EquipoModif = '" & Environment.MachineName & "' WHERE IdEmpresa = " & IdEmpresaDestino & " AND IdTienda = " & IdTiendaDestino & " AND IdBodega = " & IdBodegaDestino & " AND IdTipoMovimiento = 'TRR' AND IdCodMovimiento = '" & CodigoLovePOS & "' AND IdArticulo = " & _Articulo & "")
                    End If
                        MessageBox.Show("La cantidad se ha actualizado satisfactoriamente.", "CANTIDAD ACTUALIZADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Close()
                Else
                    MessageBox.Show("La contraseña escrita es incorrecta. Vuelva a intentarlo.", "CONSTRASEÑA ERRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtcontrasena.Text = ""
                    txtcontrasena.Focus()
                End If
            End If

        End If
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.validar_numeros(e)
    End Sub
End Class