Public Class frmDevolucionesDeTiendas
    Dim clase As New class_library
    Dim IdTienda As Integer
    Dim fechaOperativa As Date
    Dim Operador As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtCodigoLovePOS, "Codigo Transferencia LovePOS") = False Then Exit Sub
        Dim sql As String = "SELECT  CabMovimientoInventario.FechaOperativa, CabMovimientoInventario.IdEmpresa, CabMovimientoInventario.IdTienda, CabMovimientoInventario.IdBodega,  CabMovimientoInventario.IdCodMovimiento, Tienda.Nombre, CabMovimientoInventario.FechaOperativa, SUM(DetMovimientoInventario.PrecioCosto * DetMovimientoInventario.Cantidad) AS TotalCosto, SUM(DetMovimientoInventario.PrecioVenta * DetMovimientoInventario.Cantidad) AS TotalVenta, COunt(DetMovimientoInventario.IdArticulo) as Referencias, Sum(DetMovimientoInventario.Cantidad) As Cantidad, CabMovimientoInventario.Operario " &
                            "From            CabMovimientoInventario INNER JOIN DetMovimientoInventario ON CabMovimientoInventario.IdEmpresa = DetMovimientoInventario.IdEmpresa AND CabMovimientoInventario.IdTienda = DetMovimientoInventario.IdTienda AND " &
                             "CabMovimientoInventario.IdBodega = DetMovimientoInventario.IdBodega AND CabMovimientoInventario.IdTipoMovimiento = DetMovimientoInventario.IdTipoMovimiento AND  CabMovimientoInventario.IdCodMovimiento = DetMovimientoInventario.IdCodMovimiento INNER JOIN " &
                             "Tienda ON CabMovimientoInventario.IdEmpresa = Tienda.IdEmpresa AND CabMovimientoInventario.IdTienda = Tienda.IdTienda WHERE        (CabMovimientoInventario.IdTipoMovimiento = 'TRS') AND (CabMovimientoInventario.IdEmpresaDestino = 100) AND (CabMovimientoInventario.IdTiendaDestino = 99) AND (CabMovimientoInventario.Finalizado = 1) AND " &
                              "(CabMovimientoInventario.Revisada = 1) AND (CabMovimientoInventario.Recibido = 0) GROUP BY CabMovimientoInventario.FechaOperativa, CabMovimientoInventario.IdEmpresa, CabMovimientoInventario.IdTienda, CabMovimientoInventario.IdBodega,  CabMovimientoInventario.IdCodMovimiento, Tienda.Nombre, CabMovimientoInventario.FechaOperativa,  CabMovimientoInventario.Operario HAVING        (CabMovimientoInventario.IdCodMovimiento = '" & txtCodigoLovePOS.Text.ToUpper().Trim() & "')"
        clase.ConsultarSQLServer(sql, "lovepos")
        If clase.dtSql.Tables("lovepos").Rows.Count > 0 Then
            txtDestino.Text = clase.dtSql.Tables("lovepos").Rows(0)("Nombre")
            txtRef.Text = clase.dtSql.Tables("lovepos").Rows(0)("Referencias")
            txtUnds.Text = clase.dtSql.Tables("lovepos").Rows(0)("Cantidad")
            txtTotalCosto.Text = clase.dtSql.Tables("lovepos").Rows(0)("TotalCosto")
            txtTotalVenta.Text = clase.dtSql.Tables("lovepos").Rows(0)("TotalVenta")
            txtOperario.Text = clase.dtSql.Tables("lovepos").Rows(0)("Operario")
            Dim IdEmpresa As Integer = clase.dtSql.Tables("lovepos").Rows(0)("IdEmpresa")
            IdTienda = clase.dtSql.Tables("lovepos").Rows(0)("IdTienda")
            Dim IdBodega As Integer = clase.dtSql.Tables("lovepos").Rows(0)("IdBodega")
            fechaOperativa = clase.dtSql.Tables("lovepos").Rows(0)("FechaOperativa")
            Operador = clase.dtSql.Tables("lovepos").Rows(0)("Operario")
            Dim sql1 As String = "SELECT        Articulo.IdArticulo AS Codigo, Articulo.Referencia, Articulo.DescripcionCorta AS Descripcion,  DetMovimientoInventario.Cantidad AS Cant, DetMovimientoInventario.PrecioCosto AS Costo, DetMovimientoInventario.PrecioVenta AS Venta " &
                                  "FROM            DetMovimientoInventario INNER JOIN Articulo ON DetMovimientoInventario.IdArticulo = Articulo.IdArticulo WHERE (DetMovimientoInventario.IdCodMovimiento = '" & txtCodigoLovePOS.Text.ToUpper().Trim() & "') AND (DetMovimientoInventario.IdTipoMovimiento = 'TRS') AND (DetMovimientoInventario.IdEmpresa = " & IdEmpresa & ") AND (DetMovimientoInventario.IdTienda = " & IdTienda & ") AND " &
                                  "(DetMovimientoInventario.IdBodega = " & IdBodega & ")"
            clase.ConsultarSQLServer(sql1, "det")
            If clase.dtSql.Tables("det").Rows.Count > 0 Then
                dtgDetalle.DataSource = clase.dtSql.Tables("det")
                PrepararColumnas()
            Else
                dtgDetalle = Nothing
                PrepararColumnas()
            End If
            btnGuardar.Enabled = True
            btnAceptar.Enabled = False
            txtCodigoLovePOS.Enabled = False
        Else
            MessageBox.Show("El codigo de bodega escrito no existe, no ha sido finalizado o ya fue recibido en bodega.", "NO HAY RESULTADOS DE ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub PrepararColumnas()
        With dtgDetalle
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Width = 80
            .Columns(1).Width = 120
            .Columns(2).Width = 180
            .Columns(3).Width = 50
            .Columns(4).Width = 120
            .Columns(5).Width = 120
            .Columns(4).DefaultCellStyle.Format = "C"
            .Columns(5).DefaultCellStyle.Format = "C"
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = False
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
        End With
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If dtgDetalle.Rows.Count > 0 Then
            Dim a As Integer
            Dim NumeroTransferencia As Integer = consecutivo_transferencia()
            clase.agregar_registro("INSERT INTO `cabtransferencia`(`tr_numero`,`tr_destino`,`tr_origen`,`tr_bodega`,`tr_fecha`,`tr_estado`,`tr_operador`,`tr_codigolovepos`) VALUES ( '" & NumeroTransferencia & "','" & IdTienda & "',NULL,'4','" & fechaOperativa.ToString("yyyy-MM-dd") & "',TRUE,'" & Operador & "','" & txtCodigoLovePOS.Text.ToUpper().Trim() & "')")
            For a = 0 To dtgDetalle.RowCount - 1
                clase.agregar_registro("INSERT INTO `dettransferencia`(`dt_trnumero`,`dt_gondola`,`dt_codarticulo`,`dt_cantidad`,`dt_danado`,`dt_costo`,`dt_venta1`,`dt_venta2`) VALUES ('" & NumeroTransferencia & "',NULL,'" & dtgDetalle.Item(0, a).Value & "','" & dtgDetalle.Item(3, a).Value * -1 & "','0','" & Str(dtgDetalle.Item(0, a).Value) & "','" & Str(dtgDetalle.Item(4, a).Value) & "','" & Str(dtgDetalle.Item(5, a).Value) & "')")
            Next
            Dim fechaRecibo As Date = Now
            clase.AgregarSQLServer("UPDATE [dbo].[CabMovimientoInventario] SET [Recibido] = '1', [FechaRecibo] = CONVERT(DATETIME, SYSDATETIME(), 102) WHERE (CabMovimientoInventario.IdTipoMovimiento = 'TRS') AND (CabMovimientoInventario.IdEmpresaDestino = 100) AND (CabMovimientoInventario.IdTiendaDestino = 99) AND (CabMovimientoInventario.Finalizado = 1) AND (CabMovimientoInventario.Revisada = 1) AND (CabMovimientoInventario.Recibido = 0) AND (CabMovimientoInventario.IdCodMovimiento = '" & txtCodigoLovePOS.Text.ToUpper().Trim() & "')")
            MessageBox.Show("La transfencia se ha guardado satisfactoriamente.", "OPERACION FINALIZADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dtgDetalle.DataSource = Nothing
            PrepararColumnas()
            txtCodigoLovePOS.Text = ""
            txtDestino.Text = ""
            txtOperario.Text = ""
            txtRef.Text = ""
            txtUnds.Text = ""
            txtTotalCosto.Text = ""
            txtTotalVenta.Text = ""
            btnGuardar.Enabled = False
            btnAceptar.Enabled = True
        End If
    End Sub

    Function consecutivo_transferencia() As Long
        clase.consultar("SELECT MAX(tr_numero) AS maximo FROM cabtransferencia", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tbl").Rows(0)("maximo")) Then
                Return 1
            End If
            Return clase.dt.Tables("tbl").Rows(0)("maximo") + 1
        End If
    End Function




    Private Sub Validar(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        'Validar Fila seleccionada 
        Dim drwFila As DataGridViewCell = dtgDetalle.CurrentCell()
        If drwFila.ColumnIndex = 3 Then
            'Si son digitos o si es la tecla borrar
            If Char.IsDigit(e.KeyChar) Or (Asc(e.KeyChar) = 8) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub dtgDetalle_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dtgDetalle.EditingControlShowing
        AddHandler e.Control.KeyPress, AddressOf Validar
    End Sub
End Class