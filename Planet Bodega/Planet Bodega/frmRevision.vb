Public Class frmRevision
    Dim clase As New class_library
    Dim Transferencia, Registros As Integer
    Public TablaAleatoria, TablaTransferencia, TablaRevisado, TablaDiferencia As New DataTable
    Public CantidadRev As Integer
    Public Retorno As Boolean
    Dim idtransferencia As Boolean
    Dim PasswordRevision, Codigo As String

    Public Sub CargarAleatorioDetalle()
        clase.consultar("SELECT porcentaje_ref_rev FROM informacion", "aleatorio")
        Dim Porcentaje As String = clase.dt.Tables("aleatorio").Rows(0)("porcentaje_ref_rev")
        clase.consultar1("SELECT dettransferencia.dt_codarticulo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, dettransferencia.dt_trnumero,dettransferencia.dt_cantidad AS CANTIDAD FROM dettransferencia INNER JOIN articulos ON (dt_codarticulo = articulos.ar_codigo) WHERE dettransferencia.dt_trnumero ='" & txtTransferencia.Text & "';", "articulo")
        TablaTransferencia = clase.dt1.Tables("articulo")
        TablaTransferencia.Columns.Add("Agregado")

        Registros = clase.dt1.Tables("articulo").Rows.Count
        Dim limite As String = ((Registros * Porcentaje) / 100)
        clase.consultar("SELECT dettransferencia.dt_codarticulo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, dettransferencia.dt_cantidad AS CANTIDAD, dettransferencia.dt_costo AS COSTO, dettransferencia.dt_venta1 AS VENTA1, dettransferencia.dt_venta2 AS VENTA2, dettransferencia.dt_trnumero FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo)WHERE (dettransferencia.dt_trnumero ='" & txtTransferencia.Text & "') ORDER BY RAND() LIMIT " & CInt(limite) & ";", "detalle")
        TablaAleatoria = clase.dt.Tables("detalle")
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub txtTransferencia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTransferencia.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub txtTransferencia_LostFocus(sender As Object, e As EventArgs) Handles txtTransferencia.LostFocus
        If txtTransferencia.Text <> "" Then

            clase.consultar("SELECT * FROM cabtransferencia WHERE (tr_numero =" & txtTransferencia.Text & ")", "transferencia")

            If clase.dt.Tables("transferencia").Rows.Count = 0 Then

                MessageBox.Show("TRANSFERENCIA NO EXISTE", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtTransferencia.Text = ""
                txtTransferencia.Focus()
                Exit Sub
            End If

            clase.consultar1("SELECT * FROM cabtransferencia WHERE (tr_numero='" & txtTransferencia.Text & "' AND tr_revisada='1' )", "finalizada")
            If clase.dt1.Tables("finalizada").Rows.Count > 0 Then
                MessageBox.Show("ESTA TRANSFERENCIA YA FUE REVISADA", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtTransferencia.Text = ""
                txtTransferencia.Focus()
                Exit Sub
            End If

            clase.borradoautomatico("DELETE FROM faltante_transferencia WHERE rev_transferencia=" & txtTransferencia.Text & "")

            If IsDBNull(clase.dt.Tables("transferencia").Rows(0)("tr_destino")) = False Then
                txtDestino.Text = destino_x_numero_transferencia(clase.dt.Tables("transferencia").Rows(0)("tr_destino"))
            End If
            If IsDBNull(clase.dt.Tables("transferencia").Rows(0)("tr_operador")) = False Then
                txtOperario.Text = clase.dt.Tables("transferencia").Rows(0)("tr_operador")
            End If
            If IsDBNull(clase.dt.Tables("transferencia").Rows(0)("tr_bodega")) Then
                idtransferencia = False
            Else
                idtransferencia = True
            End If
            txtFecha.Text = clase.dt.Tables("transferencia").Rows(0)("tr_fecha")
            txtTransferencia.Enabled = False
            Transferencia = clase.dt.Tables("transferencia").Rows(0)("tr_numero")

            'GENERAR EL ALEATORIO DE REFERENCIAS
            CargarAleatorioDetalle()
            txtRevisor.Enabled = True
            txtRevisor.Focus()
        End If
    End Sub

    Private Function destino_x_numero_transferencia(codi As Short) As String
        clase.consultar2("select* from tiendas where id = " & codi & "", "tienda")
        Return clase.dt2.Tables("tienda").Rows(0)("tienda")
    End Function

    Private Sub txtRevisor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtRevisor.KeyPress
        clase.enter(e)
    End Sub

    Private Sub txtRevisor_LostFocus(sender As Object, e As EventArgs) Handles txtRevisor.LostFocus
        If txtRevisor.Text <> "" Then
            txtArticulo.Enabled = True
            txtRevisor.Enabled = False
            txtArticulo.Focus()
        End If
    End Sub

    Private Sub txtArticulo_Enter(sender As Object, e As EventArgs) Handles txtArticulo.Enter

    End Sub

    Private Sub txtArticulo_GotFocus(sender As Object, e As EventArgs) Handles txtArticulo.GotFocus
        Me.AcceptButton = btnAceptar
    End Sub

    Private Sub txtArticulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtArticulo.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Public Sub Limpiar()
        TablaRevisado.Rows.Clear()
        TablaAleatoria.Rows.Clear()
        TablaTransferencia.Rows.Clear()
        TablaDiferencia.Rows.Clear()
        txtTransferencia.Enabled = True
        txtTransferencia.Focus()
        txtArticulo.Enabled = False
        txtRevisor.Enabled = False
        txtDestino.Text = ""
        txtTransferencia.Text = ""
        txtFecha.Text = ""
        txtOperario.Text = ""
        txtRevisor.Text = ""
        clase.Limpiar_Cajas(Me)
        Retorno = False
    End Sub

    Private Sub frmTransferencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT * FROM informacion", "password")
        PasswordRevision = clase.dt.Tables("password").Rows(0)("password_revision")
        grdRevision.RowHeadersVisible = False
        TablaRevisado.Columns.Add("CODIGO")
        TablaRevisado.Columns.Add("REFERENCIA")
        TablaRevisado.Columns.Add("DESCRIPCION")
        TablaRevisado.Columns.Add("CANTIDAD")
        TablaDiferencia.Columns.Add("rev_transferencia")
        TablaDiferencia.Columns.Add("rev_articulo")
        TablaDiferencia.Columns.Add("rev_faltante")
        TablaDiferencia.Columns.Add("Referencia")
        TablaDiferencia.Columns.Add("Descripcion")
        TablaDiferencia.Columns.Add("total_producto")
        TablaDiferencia.Columns.Add("esperado")
        txtTransferencia.Focus()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Limpiar()
    End Sub

    Private Sub txtFecha_GotFocus(sender As Object, e As EventArgs) Handles txtFecha.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtDestino_GotFocus(sender As Object, e As EventArgs) Handles txtDestino.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtOperario_GotFocus(sender As Object, e As EventArgs) Handles txtOperario.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtReferencia_GotFocus(sender As Object, e As EventArgs) Handles txtReferencia.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtDescripcion_GotFocus(sender As Object, e As EventArgs) Handles txtDescripcion.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtTotal_GotFocus(sender As Object, e As EventArgs) Handles txtTotal.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtRevisadas_GotFocus(sender As Object, e As EventArgs) Handles txtRevisadas.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtArticulo_LostFocus(sender As Object, e As EventArgs) Handles txtArticulo.LostFocus
        Me.AcceptButton = Nothing
    End Sub

    Private Sub btnFinalizar_Click(sender As Object, e As EventArgs) Handles btnFinalizar.Click
        Dim Diferencia As Integer = 0
        Dim revcodigo As String = "1"
        'COMPARAR TABLATRANSFERENCIA & TABLAREVISADO PARA OBTENER LOS SOBRANTES O FALTANTES
        For i = 0 To TablaTransferencia.Rows.Count - 1
            For j = 0 To TablaRevisado.Rows.Count - 1
                If TablaTransferencia.Rows(i)("CODIGO") = TablaRevisado.Rows(j)("CODIGO") Then

                    If TablaTransferencia.Rows(i)("CANTIDAD") <> TablaRevisado.Rows(j)("CANTIDAD") Then
                        Diferencia = Val(TablaTransferencia.Rows(i)("CANTIDAD")) - Val(TablaRevisado.Rows(j)("CANTIDAD"))

                        Dim AgregarDiferencia As DataRow = TablaDiferencia.NewRow
                        AgregarDiferencia("rev_transferencia") = txtTransferencia.Text
                        AgregarDiferencia("rev_articulo") = TablaRevisado.Rows(j)("CODIGO")
                        AgregarDiferencia("rev_faltante") = Diferencia
                        AgregarDiferencia("total_producto") = TablaRevisado.Rows(j)("CANTIDAD")
                        AgregarDiferencia("esperado") = TablaTransferencia.Rows(i)("CANTIDAD")
                        AgregarDiferencia("Referencia") = TablaTransferencia.Rows(i)("REFERENCIA")
                        AgregarDiferencia("Descripcion") = TablaTransferencia.Rows(i)("DESCRIPCION")

                        TablaDiferencia.Rows.Add(AgregarDiferencia)
                    End If
                Else
                    If revcodigo <> TablaTransferencia.Rows(i)("CODIGO") Then
                        revcodigo = TablaTransferencia.Rows(i)("CODIGO")
                        Dim Buscar2() As DataRow
                        Buscar2 = TablaRevisado.Select("CODIGO = '" & TablaTransferencia.Rows(i)("CODIGO") & "'")
                        If Buscar2.Length = 0 Then
                            If MessageBox.Show("Falta este articulo por revisar " & TablaTransferencia.Rows(i)("CODIGO") & " - " & TablaTransferencia.Rows(i)("REFERENCIA") & " - " & TablaTransferencia.Rows(i)("DESCRIPCION") & ". Desea volver al formulario y agregarlo?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                                TablaDiferencia.Rows.Clear()
                                Exit Sub
                            Else
                                'BORRAR ARTICULO DE LA TABLA DETTRANSFERENCIA
                                clase.borradoautomatico("DELETE FROM dettransferencia WHERE dt_trnumero='" & txtTransferencia.Text & "' AND dt_codarticulo='" & TablaTransferencia.Rows(i)("CODIGO") & "'")

                                'CARGAR ARTICULO A LA TABLA FALTANTETRANSFERENCIA                  (1)
                                clase.agregar_registro("INSERT INTO faltante_transferencia(rev_transferencia,rev_articulo,rev_faltante) VALUES('" & txtTransferencia.Text & "','" & TablaTransferencia.Rows(i)("CODIGO") & "','" & TablaTransferencia.Rows(i)("CANTIDAD") & "')")

                                'GUARDAR ARTICULO EN LA TABLA ELIMINADOS_TRANSFERENCIA
                                clase.agregar_registro("INSERT INTO eliminados_transferencia(trelim_idtransferencia,trelim_articulo,trelim_cantidad) VALUES('" & txtTransferencia.Text & "','" & TablaTransferencia.Rows(i)("CODIGO") & "','" & TablaTransferencia.Rows(i)("CANTIDAD") & "')")
                                Continue For
                            End If
                        End If
                    End If
                End If
            Next
        Next

        If TablaDiferencia.Rows.Count > 0 Then
            'MOSTRAR DIFERENCIAS Y SOLICITAR AUTORIZACION PARA REALIZAR LOS AJUSTES
            frm_diferencias_transferencia.grdDiferencia.DataSource = TablaDiferencia
            frm_diferencias_transferencia.ShowDialog()
            frm_diferencias_transferencia.Dispose()

            If Retorno = True Then
                Exit Sub
            End If

            For i = 0 To TablaDiferencia.Rows.Count - 1
                'ACTUALIZAR TABLA DETTRANSFERENCIA
                clase.actualizar("UPDATE dettransferencia SET dt_cantidad='" & TablaDiferencia.Rows(i)("total_producto") & "' WHERE dt_trnumero='" & txtTransferencia.Text & "' AND dt_codarticulo='" & TablaDiferencia.Rows(i)("rev_articulo") & "'")

                'AGREGAR DIFERENCIA A LA TABLA FALTANTE_TRANSFERENCIA                     (2)
                clase.agregar_registro("INSERT INTO faltante_transferencia(rev_transferencia,rev_articulo,rev_faltante) VALUES('" & txtTransferencia.Text & "','" & TablaDiferencia.Rows(i)("rev_articulo") & "','" & TablaDiferencia.Rows(i)("rev_faltante") & "')")
            Next
        End If
        If idtransferencia = True Then
            clase.consultar1("SELECT * FROM faltante_transferencia WHERE (rev_transferencia =" & txtTransferencia.Text & ")", "revision")
            If clase.dt1.Tables("revision").Rows.Count > 0 Then
                Dim z As Short
                For z = 0 To clase.dt1.Tables("revision").Rows.Count - 1
                    sumar_al_inventario(clase.dt1.Tables("revision").Rows(z)("rev_articulo"), clase.dt1.Tables("revision").Rows(z)("rev_faltante"))
                Next
            End If
        End If
        'ACTUALIZAR EL CAMPO TR_REVISADA Y TR_REVISOR EN LA TABLA CABTRANSFERENCIA A "SI"
        clase.actualizar("UPDATE cabtransferencia SET tr_revisada='1', tr_revisor='" & UCase(txtRevisor.Text) & "' WHERE tr_numero='" & txtTransferencia.Text & "'")
        'IMPRIMIR LA TRANSFERENCIA
        If MessageBox.Show("DESEA ENVIAR EL ARCHIVO DE TRANSFERENCIA POR EMAIL AL DESTINO?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
            construir_archivo(txtTransferencia.Text)
        End If
        imprimir_hoja_transferencia(txtTransferencia.Text)
        'LIMPIAR LOS DATASET Y LOS TEXT BOX
        MessageBox.Show("La transferencia fue revisada exitosamente", "TRANSFERENCIA REVISADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Limpiar()
        Exit Sub
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        CantidadRev = 0
        Dim ArtRevisado As DataRow = TablaRevisado.NewRow
        Dim Nuevo As Boolean = True
        Dim NuevoArticulo As Boolean = False
        If txtArticulo.Text <> "" Then
            'VALIDAR QUE AL ARTICULO PERTENEZCA AL DETALLE DE LA TRASFERENCIA
            clase.consultar("SELECT articulos.ar_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, dettransferencia.dt_trnumero FROM dettransferencia INNER JOIN articulos ON (dt_codarticulo = articulos.ar_codigo) WHERE ((dettransferencia.dt_codarticulo ='" & txtArticulo.Text & "' OR articulos.ar_codigobarras='" & txtArticulo.Text & "') AND dettransferencia.dt_trnumero ='" & txtTransferencia.Text & "');", "articulo")
            If clase.dt.Tables("articulo").Rows.Count = 0 Then

                If MessageBox.Show("El articulo no se encuentra registrado en esta transferencia, deseea agregarlo?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    frm_password_revision.txtPassword.Text = ""
                    frm_password_revision.ShowDialog()
                    If frm_password_revision.ValidaPassword = PasswordRevision Then
                        NuevoArticulo = True
                    Else
                        MessageBox.Show("Contraseña invalida.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtArticulo.Text = ""
                        txtArticulo.Focus()
                        Exit Sub
                    End If
                    frm_password_revision.Dispose()
                Else
                    txtArticulo.Text = ""
                    txtArticulo.Focus()
                    Exit Sub
                End If
            End If

            If NuevoArticulo = True Then
                clase.consultar1("SELECT ar_codigo AS CODIGO,ar_referencia AS REFERENCIA,ar_descripcion AS DESCRIPCION,ar_costo AS dt_costo,ar_precio1 AS dt_venta1,ar_precio2 AS dt_venta2 FROM articulos WHERE ar_codigo='" & txtArticulo.Text & "' or ar_codigobarras='" & txtArticulo.Text & "'", "articulo")
                If clase.dt1.Tables("articulo").Rows.Count > 0 Then
                    txtArticulo.Text = clase.dt1.Tables("articulo").Rows(0)("CODIGO")
                    txtDescripcion.Text = clase.dt1.Tables("articulo").Rows(0)("DESCRIPCION")
                    txtReferencia.Text = clase.dt1.Tables("articulo").Rows(0)("REFERENCIA")

                    frmCantidad.ShowDialog()

                    'AGREGAR EL REGISTRO A LA TABLA DETTRANSFERENCIA CON EL NUEVO PRODUCTO
                    clase.agregar_registro("INSERT INTO dettransferencia(dt_trnumero,dt_codarticulo,dt_cantidad,dt_danado,dt_costo,dt_venta1,dt_venta2) VALUES('" & txtTransferencia.Text & "','" & clase.dt1.Tables("articulo").Rows(0)("CODIGO") & "','" & CantidadRev & "','0','" & clase.dt1.Tables("articulo").Rows(0)("dt_costo") & "','" & clase.dt1.Tables("articulo").Rows(0)("dt_venta1") & "','" & clase.dt1.Tables("articulo").Rows(0)("dt_venta2") & "')")

                    'GUARDAR EN LA TABLA FALTANTETRANSFERENCIA COMO -                                 (3)
                    clase.agregar_registro("INSERT INTO faltante_transferencia(rev_transferencia,rev_articulo,rev_faltante) VALUES('" & txtTransferencia.Text & "','" & clase.dt1.Tables("articulo").Rows(0)("CODIGO") & "','" & -1 * CantidadRev & "')")

                    'AGREGAR PRODUCTO AL DATATABLE TABLATRANSFERENCIA
                    Dim NuevaFilaArticulo As DataRow = TablaTransferencia.NewRow
                    NuevaFilaArticulo("CODIGO") = clase.dt1.Tables("articulo").Rows(0)("CODIGO")
                    NuevaFilaArticulo("DESCRIPCION") = clase.dt1.Tables("articulo").Rows(0)("DESCRIPCION")
                    NuevaFilaArticulo("REFERENCIA") = clase.dt1.Tables("articulo").Rows(0)("REFERENCIA")
                    NuevaFilaArticulo("dt_trnumero") = txtTransferencia.Text
                    NuevaFilaArticulo("CANTIDAD") = CantidadRev
                    NuevaFilaArticulo("Agregado") = "SI"
                    TablaTransferencia.Rows.Add(NuevaFilaArticulo)
                Else
                    MessageBox.Show("Articulo no existe.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtArticulo.Text = ""
                    txtArticulo.Focus()
                    Exit Sub
                End If
            Else
                txtArticulo.Text = clase.dt.Tables("articulo").Rows(0)("CODIGO")
                txtDescripcion.Text = clase.dt.Tables("articulo").Rows(0)("DESCRIPCION")
                txtReferencia.Text = clase.dt.Tables("articulo").Rows(0)("REFERENCIA")
                frmCantidad.ShowDialog()
            End If

            'VALIDAR SI EL ARTICULO YA SE ENCUENTRA Y EN DADO CASO AGREGARLE LAS UNIDADES
            For i = 0 To TablaRevisado.Rows.Count - 1
                If TablaRevisado.Rows(i)("CODIGO") = txtArticulo.Text Then
                    TablaRevisado.Rows(i)("CANTIDAD") = Val(TablaRevisado.Rows(i)("CANTIDAD")) + CantidadRev
                    Nuevo = False
                End If
            Next

            'GUARDAR EN UN DATASET LO QUE SE VA REVISANDO PARA SU POSTERIOR VERIFICACION
            If Nuevo = True Then
                ArtRevisado("CODIGO") = txtArticulo.Text
                ArtRevisado("REFERENCIA") = txtReferencia.Text
                ArtRevisado("DESCRIPCION") = txtDescripcion.Text
                ArtRevisado("CANTIDAD") = CantidadRev
                TablaRevisado.Rows.InsertAt(ArtRevisado, 0)
            End If
            grdRevision.DataSource = TablaRevisado
            grdRevision.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            grdRevision.Columns(1).Width = 250

            grdRevision.Rows(0).Selected = True
            grdRevision.CurrentCell = grdRevision.Rows(0).Cells(1)

            txtTotal.Text = Val(txtTotal.Text) + 1
            txtArticulo.Text = ""
            txtArticulo.Focus()
        End If
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If Codigo = "" Then
            MessageBox.Show("Seleccione el código a eliminar.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If MessageBox.Show("Desea eliminar el codigo " & Codigo & " ?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                frm_password_revision.txtPassword.Text = ""
                frm_password_revision.ShowDialog()

                If frm_password_revision.ValidaPassword = PasswordRevision Then
                    'Borrar del datagrid
                    For a = 0 To grdRevision.Rows.Count - 1
                        If Codigo = grdRevision(0, a).Value Then
                            txtTotal.Text = Val(txtTotal.Text) - 1
                            txtRevisadas.Text = Val(txtRevisadas.Text) - Val(grdRevision(3, a).Value)
                            grdRevision.Rows.Remove(grdRevision.Rows(a))
                            Exit For
                        End If
                    Next

                    For a = 0 To TablaTransferencia.Rows.Count - 1
                        If Codigo = TablaTransferencia.Rows(a)("CODIGO") And IsDBNull(TablaTransferencia.Rows(a)("Agregado")) = False Then
                            clase.borradoautomatico("DELETE FROM dettransferencia WHERE dt_trnumero='" & txtTransferencia.Text & "' AND dt_codarticulo='" & Codigo & "'")

                            clase.borradoautomatico("DELETE FROM faltante_transferencia WHERE rev_transferencia='" & txtTransferencia.Text & "' AND rev_articulo='" & Codigo & "'")
                            Exit For
                        End If
                    Next
                Else
                    MessageBox.Show("Contraseña invalida.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtArticulo.Text = ""
                    txtArticulo.Focus()
                    Exit Sub
                End If
            End If
        End If

    End Sub

    Private Sub grdRevision_Click(sender As Object, e As EventArgs) Handles grdRevision.Click
        Codigo = grdRevision(0, grdRevision.CurrentCell.RowIndex).Value
    End Sub

    Private Sub txtTransferencia_TextChanged(sender As Object, e As EventArgs) Handles txtTransferencia.TextChanged

    End Sub

    Private Sub txtArticulo_TextChanged(sender As Object, e As EventArgs) Handles txtArticulo.TextChanged

    End Sub
End Class
