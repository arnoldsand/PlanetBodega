Imports System.Windows.Forms
Public Class mdi_principal
    Dim clase As New class_library
    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click, NewToolStripButton.Click
        frm_crear_articulos_lotes.MdiParent = Me
        frm_crear_articulos_lotes.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripButton.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: agregue código aquí para abrir el archivo.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: agregue código aquí para guardar el contenido actual del formulario en un archivo.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        frm_articulos.MdiParent = Me
        frm_articulos.Show()

    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Utilice My.Computer.Clipboard.GetText() o My.Computer.Clipboard.GetData para recuperar la información del Portapapeles.
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Cierre todos los formularios secundarios del principal.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub MenuStrip_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip.ItemClicked

    End Sub

    Private Sub MDIParent1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'frm_loginacceso.ShowDialog()
        'frm_loginacceso.Dispose()
        '  SetImage(PictureBox1)      

    End Sub

    Private Sub StocksDeArticulosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StocksDeArticulosToolStripMenuItem.Click
        frm_stock_articulos.MdiParent = Me
        frm_stock_articulos.Show()
    End Sub

    Private Sub LineasYSublineasDeArticuloToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LineasYSublineasDeArticuloToolStripMenuItem.Click
        frm_lineas_sublineas.MdiParent = Me
        frm_lineas_sublineas.Show()
    End Sub

    Private Sub ImportacionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportacionToolStripMenuItem.Click

    End Sub

    Private Sub ImprimirCodigosDeBarraParaCajasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImprimirCodigosDeBarraParaCajasToolStripMenuItem.Click
        Codigos_de_Barra_para_Preinspección.ShowDialog()
        Codigos_de_Barra_para_Preinspección.Dispose()
    End Sub

    Private Sub TabulaciónDeImportaciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TabulaciónDeImportaciónToolStripMenuItem.Click
        frm_tabulacion_importacion.MdiParent = Me
        frm_tabulacion_importacion.Show()
    End Sub

    Private Sub RecibirCajasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RecibirCajasToolStripMenuItem.Click
        frm_recibo_cajas.ShowDialog()
        frm_recibo_cajas.Dispose()
    End Sub


    Private Sub StockBodegaMercancíaMuertaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StockBodegaMercancíaMuertaToolStripMenuItem.Click
        frm_stock_bodega_muerta.ShowDialog()
        frm_stock_bodega_muerta.Dispose()
    End Sub

    Private Sub FlijoDeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FlijoDeToolStripMenuItem.Click
        frm_flujo_mercancia.ShowDialog()
        frm_flujo_mercancia.Dispose()
    End Sub

    Private Sub PuntosDeVentaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PuntosDeVentaToolStripMenuItem.Click
        frm_tienda.MdiParent = Me
        frm_tienda.Show()
    End Sub

    Private Sub ProveedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProveedoresToolStripMenuItem.Click
        frm_maestro_proveedores.MdiParent = Me
        frm_maestro_proveedores.Show()
    End Sub

    Private Sub TallasYColoresToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frm_tallas_colores.MdiParent = Me
        frm_tallas_colores.Show()
    End Sub



    Private Sub RelacionDeArticulosProcesadosEnLiquidaciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RelacionDeArticulosProcesadosEnLiquidaciónToolStripMenuItem.Click
        frm_articulos_procesados_liquidacion.MdiParent = Me
        frm_articulos_procesados_liquidacion.Show()
    End Sub

    Private Sub RelaciónDeArticulosProcesadosEnTiqueteoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RelaciónDeArticulosProcesadosEnTiqueteoToolStripMenuItem.Click
        frm_articulos_procesados_tiqueteo.MdiParent = Me
        frm_articulos_procesados_tiqueteo.Show()
    End Sub

    Private Sub InventarioToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles InventarioToolStripMenuItem1.Click
        frm_almacenaje_articulos.MdiParent = Me
        frm_almacenaje_articulos.Show()
    End Sub

    Private Sub AjustesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AjustesToolStripMenuItem.Click
        frm_ajustes.MdiParent = Me
        frm_ajustes.Show()
    End Sub

    Private Sub SaldosNegativosXBodegaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaldosNegativosXBodegaToolStripMenuItem.Click
        frm_saldos_negativos.MdiParent = Me
        frm_saldos_negativos.Show()
    End Sub

    Private Sub TranferenciasDesdeBodegaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TranferenciasDesdeBodegaToolStripMenuItem.Click
        frm_transferencias_desde_bodega.MdiParent = Me
        frm_transferencias_desde_bodega.Show()
    End Sub

    Private Sub TransferenciasDesdeDigitaciónToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TransferenciaDesdeDigitaciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransferenciaDesdeDigitaciónToolStripMenuItem.Click
        frmTransferencia.MdiParent = Me
        frmTransferencia.Show()
    End Sub

    Private Sub AjustarGóndolaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AjustarGóndolaToolStripMenuItem.Click
        frm_generar_ajuste_para_gondola2.ShowDialog()
        frm_generar_ajuste_para_gondola2.Dispose()
    End Sub

    Private Sub MantenimientoDeTransferenciasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MantenimientoDeTransferenciasToolStripMenuItem.Click
        frm_mantenimiento_de_transfencia.MdiParent = Me
        frm_mantenimiento_de_transfencia.Show()
    End Sub

    Private Sub TransferenciaDesdeHandHeldToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransferenciaDesdeHandHeldToolStripMenuItem.Click
        frm_salida_de_mercancia_desde_hand_held.MdiParent = Me
        frm_salida_de_mercancia_desde_hand_held.Show()
        'frm_salida_de_mercancia_desde_hand_held.ShowDialog()
        ' frm_salida_de_mercancia_desde_hand_held.Dispose()
    End Sub

    Private Sub MantenimientoDePedidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MantenimientoDePedidosToolStripMenuItem.Click
        frm_mantenimiento_pedidos.MdiParent = Me
        frm_mantenimiento_pedidos.Show()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        frm_entrada_de_mercancia_desde_hand_held.ShowDialog()
        frm_entrada_de_mercancia_desde_hand_held.Dispose()
    End Sub

    Private Sub CatalogoDeFotosToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ConfiguaraciónManualToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConfiguaraciónManualToolStripMenuItem.Click
        frm_ajustar_foto_catalogo.MdiParent = Me
        frm_ajustar_foto_catalogo.Show()
    End Sub

    Private Sub ConfiguraciónDesdeHandHeldToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConfiguraciónDesdeHandHeldToolStripMenuItem.Click
        frm_fotos_desde_hand_held.ShowDialog()
        frm_fotos_desde_hand_held.Dispose()
    End Sub

    Private Sub SubirAlCatalogoArticulosCapturadosConHandHeldToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SubirAlCatalogoArticulosCapturadosConHandHeldToolStripMenuItem.Click
        frm_articulos_capturados_con_hand_held.ShowDialog()
        frm_articulos_capturados_con_hand_held.Dispose()
    End Sub

    Private Sub ImpresionDeCodigosDeBarraParaGóndolasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frm_imprimir_codigos_barras_gondolas.ShowDialog()
        frm_imprimir_codigos_barras_gondolas.Dispose()
    End Sub

    Private Sub TabulaciónDeFacturasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TabulaciónDeFacturasToolStripMenuItem.Click
        tabulacion_de_facturas.ShowDialog()
        tabulacion_de_facturas.Dispose()
    End Sub

    Private Sub FacturasTabuladasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FacturasTabuladasToolStripMenuItem.Click
        frm_facturas_tabuladas.ShowDialog()
        frm_facturas_tabuladas.Dispose()
    End Sub

    Private Sub ReporteDeDañadoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frm_reporte_de_danado.ShowDialog()
        frm_reporte_de_danado.Dispose()
    End Sub



    Private Sub PorUnidadYCantidadToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub GenerarActualizacionesDeArticulosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerarActualizacionesDeArticulosToolStripMenuItem.Click
        frm_actualizaciones_de_articulo.ShowDialog()
        frm_actualizaciones_de_articulo.Dispose()
    End Sub

    Private Sub PatroneDeDistribuciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PatroneDeDistribuciónToolStripMenuItem.Click
        ind_formulario_patron_distribucion = True
        'frm_patron_distribucion.ShowDialog()
        ' frm_patron_distribucion.Dispose()
        frm_patron_mercancia_no_importada.ShowDialog()
        frm_patron_mercancia_no_importada.Dispose()
        ind_formulario_patron_distribucion = False
    End Sub

    Private Sub OrdenesDeProducciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrdenesDeProducciónToolStripMenuItem.Click
        frmOrdenProduccion.ShowDialog()
        frmOrdenProduccion.Dispose()
    End Sub

    Private Sub TiqueteoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TiqueteoToolStripMenuItem.Click
        frmAsignar.ShowDialog()
        frmAsignar.Dispose()
    End Sub

    Private Sub RevisiónDeTransferenciaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RevisiónDeTransferenciaToolStripMenuItem.Click
        frm_revision.ShowDialog()
        frm_revision.Dispose()
    End Sub

    Private Sub IngresoManualDeMercancíaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IngresoManualDeMercancíaToolStripMenuItem.Click
        frm_ingreso_no_importado.ShowDialog()
        frm_ingreso_no_importado.Dispose()
    End Sub

    Private Sub DevolucionesABodegaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DevolucionesABodegaToolStripMenuItem.Click
        frmDevolucionesDeTiendas.ShowDialog()
        frmDevolucionesDeTiendas.Dispose()
        ' frmTransferenciaDevoluciones.ShowDialog()
        ' frmTransferenciaDevoluciones.Dispose()
    End Sub

    Private Sub MantenimientoDeDevolucionesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MantenimientoDeDevolucionesToolStripMenuItem.Click
        frm_mantenimiento_devolucion.ShowDialog()
        frm_mantenimiento_devolucion.Dispose()
    End Sub

    Private Sub ArticulosIngresadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ArticulosIngresadosToolStripMenuItem.Click
        frmInforme.ShowDialog()
        frmInforme.Dispose()
    End Sub

    Private Sub ImpresionDePatronesDeDistribuciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImpresionDePatronesDeDistribuciónToolStripMenuItem.Click
        frmInformePatron.ShowDialog()
        frmInformePatron.Dispose()
    End Sub

    Private Sub ConsultaDeFotoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frmConsulta_x_foto.ShowDialog()
        frmConsulta_x_foto.Dispose()
    End Sub

    Private Sub IngresoDeMercanciaPuertaAPuertaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frm_ingreso_puerta_puerta_puerta_puerta.ShowDialog()
        frm_ingreso_puerta_puerta_puerta_puerta.Dispose()
    End Sub

    Private Sub CargarPreInspeccionDesdeExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CargarPreInspeccionDesdeExcelToolStripMenuItem.Click
        frmPreInspeccion.ShowDialog()
        frmPreInspeccion.Dispose()
    End Sub

    Private Sub RegistroDIANPorSubpartidaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistroDIANPorSubpartidaToolStripMenuItem.Click
        frm_registro_dian_subpartida.ShowDialog()
        frm_registro_dian_subpartida.Dispose()
    End Sub

    Private Sub RegistroDeImportaciónPorArticuloToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistroDeImportaciónPorArticuloToolStripMenuItem.Click
        registros_importacion_articulo.ShowDialog()
        registros_importacion_articulo.Dispose()
    End Sub

    Private Sub CodificacionDeArticulosNuevosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CodificacionDeArticulosNuevosToolStripMenuItem.Click
        frm_planillas_ripley_codigos_nuevos.ShowDialog()
        frm_planillas_ripley_codigos_nuevos.Dispose()
    End Sub

    Private Sub VariacionesDePreciosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VariacionesDePreciosToolStripMenuItem.Click
        frm_variaciones_de_precios_ripley.ShowDialog()
        frm_variaciones_de_precios_ripley.Dispose()
    End Sub

    Private Sub SubirMaestroDeCodigoDeRipleyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SubirMaestroDeCodigoDeRipleyToolStripMenuItem.Click
        frm_subir_codigos_ripley.ShowDialog()
        frm_subir_codigos_ripley.Dispose()
    End Sub

    Private Sub ConsultaDeCatalogoDeFotosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultaDeCatalogoDeFotosToolStripMenuItem.Click
        frmConsulta_x_foto.ShowDialog()
        frmConsulta_x_foto.Dispose()
    End Sub

    Private Sub AnexionAImportaciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnexionAImportaciónToolStripMenuItem.Click
        frm_almacenaje_x_hand_held.ShowDialog()
        frm_almacenaje_x_hand_held.Dispose()
    End Sub

    Private Sub SeguimientoAInventarioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeguimientoAInventarioToolStripMenuItem.Click
        frm_seguimiento_inventario.ShowDialog()
        frm_seguimiento_inventario.Dispose()
    End Sub

    Private Sub PorMovimientoToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        frm_impresion_codigos_de_barra.ShowDialog()
        frm_impresion_codigos_de_barra.Dispose()
    End Sub

    Private Sub PorUnidadaYCantidadToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frm_imprimir_codigos_barra_unidad.ShowDialog()
        frm_imprimir_codigos_barra_unidad.Dispose()
    End Sub

    Private Sub PorMovimientoToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        frm_codigo_de_barras_unidad_29_x_15.ShowDialog()
        frm_codigo_de_barras_unidad_29_x_15.Dispose()
    End Sub

    Private Sub PorUnidadYCantidadToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        frm_impresion_codigo_barra_x_unidad_19_x_15.ShowDialog()
        frm_impresion_codigo_barra_x_unidad_19_x_15.Dispose()
    End Sub

    Private Sub FinalizarUnaTransferenciaSinRevisarToolStripMenuItem_Click(sender As Object, e As EventArgs)
        frm_finalizar_sin_revisar.ShowDialog()
        frm_finalizar_sin_revisar.Dispose()
    End Sub

    Private Sub InventariarBodegaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InventariarBodegaToolStripMenuItem.Click
        frm_modulo_inventario.ShowDialog()
        frm_modulo_inventario.Dispose()
    End Sub

    Private Sub ToolStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip.ItemClicked

    End Sub

    Private Sub MantenimientoDePrepedidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MantenimientoDePrepedidosToolStripMenuItem.Click
        frm_mantenimiento_prepedidos.ShowDialog()
        frm_mantenimiento_prepedidos.Dispose()
    End Sub

    Private Sub EditorDeRecibosDeMercancíaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditorDeRecibosDeMercancíaToolStripMenuItem.Click
        frm_mantenimiento_recibo_bodega.ShowDialog()
        frm_mantenimiento_recibo_bodega.Dispose()
    End Sub

    Private Sub MovimientoDeMercancíaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MovimientoDeMercancíaToolStripMenuItem.Click
        frm_movimiento_de_mercancia.ShowDialog()
        'If ind_formulario = False Then
        frm_movimiento_de_mercancia.Dispose()

        'Else
        '    frm_movimiento_de_mercancia.Dispose()
        '    frm_movimiento_de_mercancia.ShowDialog()
        '    ind_formulario = False
        'End If
    End Sub

    Private Sub FinalizarTransferenciaSinRevisarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FinalizarTransferenciaSinRevisarToolStripMenuItem.Click
        clase.consultar("select* from informacion", "inf")
        If clase.dt.Tables("inf").Rows(0)("habilitar_finalizar_sin_revisar") Then
            frm_finalizar_sin_revisar.ShowDialog()
            frm_finalizar_sin_revisar.Dispose()
        Else
            MessageBox.Show("La finalización de transferencías sin revisar esta deshabilitado. No se puede utilizar en este momento.", "FINALIZAR SIN REVISAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub MigraciónDeArticulosYActualizacionesDePrecioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MigraciónDeArticulosYActualizacionesDePrecioToolStripMenuItem.Click
        frmMigracionArticulos.ShowDialog()
        frmMigracionArticulos.Dispose()
    End Sub

    Private Sub VerificarArticulosSinFotoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VerificarArticulosSinFotoToolStripMenuItem.Click
        ' frmVerificarExistenciaFotosArticulos.ShowDialog()
        '  frmVerificarExistenciaFotosArticulos.Dispose()
    End Sub

    Private Sub FacturasImportacionesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FacturasImportacionesToolStripMenuItem.Click
        frmFotosImportacion.ShowDialog()
        frmFotosImportacion.Dispose()
    End Sub

    Private Sub GenerarAjusteEnGóndolaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerarAjusteEnGóndolaToolStripMenuItem.Click

    End Sub

    Private Sub ImpresionDeCodigosDeBaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImpresionDeCodigosDeBaToolStripMenuItem.Click
        frm_impresion_codigo_barra_x_unidad_19_x_15.ShowDialog()
        frm_impresion_codigo_barra_x_unidad_19_x_15.Dispose()
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        frmClasificaciones.ShowDialog()
        frmClasificaciones.Dispose()
    End Sub

    Private Sub SalirToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem1.Click
        Me.Close()
    End Sub
End Class
