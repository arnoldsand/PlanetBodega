Public Class frm_revision
    Dim clase As New class_library
    Dim ind As Boolean = False 'esta variable la utilizo para distinguir las pulsaciones del boton aceptar (cuando es presionado para validar la transferenica y cuando es para validar el revisor)
    Dim cont_inconsistencia As Short
    Dim ref_faltantes As Short
    Dim descontar_inventario As Boolean  'esta varible me indica si la transferencia proviene de la bodega de almacenaje y por ende si debo descontar los saldos
    Dim indicadorcerrado As Boolean
    Dim CodigoTiendaDestino As Integer
    Dim CodigoBodegaOrigen As Integer
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub frm_revision_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If indicadorcerrado = True Then
            Dim f As String = MessageBox.Show("¿Esta seguro que desea cerrar el formulario, se perderán los datos de la revisión actual?", "CERRAR FORMULARIO", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If f = 7 Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub frm_revision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        textBox2.Focus()
        cont_inconsistencia = 0
        ref_faltantes = 0
        indicadorcerrado = False
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(textBox2, "Número Transferencia") = False Then Exit Sub
        If ind = False Then
            clase.consultar("SELECT cabtransferencia.tr_destino, tiendas.tienda, cabtransferencia.tr_fecha, cabtransferencia.tr_bodega, cabtransferencia.tr_estado, cabtransferencia.tr_operador, cabtransferencia.tr_revisada FROM cabtransferencia INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) WHERE (cabtransferencia.tr_numero =" & textBox2.Text & ")", "transferencia")
            If clase.dt.Tables("transferencia").Rows.Count > 0 Then
                If (clase.dt.Tables("transferencia").Rows(0)("tr_revisada") = True) Or (clase.dt.Tables("transferencia").Rows(0)("tr_estado") = False) Then
                    MessageBox.Show("La trasferencia ya fue revisada o no ha sido finalizada, verifique y vuelva a intentarlo", "NO SE PUEDE REVISAR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    textBox2.Text = ""
                    textBox2.Focus()
                Else
                    CodigoTiendaDestino = clase.dt.Tables("transferencia").Rows(0)("tr_destino")
                    If IsDBNull(clase.dt.Tables("transferencia").Rows(0)("tr_bodega")) Then
                        CodigoBodegaOrigen = 0
                    Else
                        CodigoBodegaOrigen = clase.dt.Tables("transferencia").Rows(0)("tr_bodega")
                    End If
                    cont_inconsistencia = 0
                    textBox2.Enabled = False
                    txtDestino.Text = clase.dt.Tables("transferencia").Rows(0)("tienda")
                    Dim fechatrans As Date = clase.dt.Tables("transferencia").Rows(0)("tr_fecha")
                    txtFecha.Text = fechatrans.ToString("dd/MM/yyyy")
                    txtOperario.Text = clase.dt.Tables("transferencia").Rows(0)("tr_operador")
                    txtRevisor.Enabled = True
                    txtRevisor.Focus()
                    ref_faltantes = referencias_transferencias()
                    txtfaltantes.Text = ref_faltantes
                    txtinconsistencias.Text = cont_inconsistencia
                    btnagregar.Enabled = True
                    btnLimpiar.Enabled = True
                    btnEliminar.Enabled = True
                    btnFinalizar.Enabled = True
                    If IsDBNull(clase.dt.Tables("transferencia").Rows(0)("tr_bodega")) Then  'aqui compruebo el valor del campo tr_bodega que me indicará si descuento o no del inventario lo que hubiese en la tabla faltante_transferencia
                        descontar_inventario = False
                        clase.borradoautomatico("delete from faltante_transferencia where rev_transferencia = " & textBox2.Text)
                    Else
                        descontar_inventario = True
                        clase.consultar1("select* from faltante_transferencia where rev_transferencia = " & textBox2.Text, "faltante") 'esto que se hace aquí es para cargar al inventario lo que estuviese en la tabla faltante_transferencia, esto solo pasará si entre la descarga del formulario frm_articulos_faltantes_x_revision y la carga de frm_diferencias_revision se va la luz o es descargado el formulario frm_revision y solo estará lo que se hubiese eliminado de las transferencias (lo que no se hubiese incluido)
                        Dim z As Short
                        For z = 0 To clase.dt1.Tables("faltante").Rows.Count - 1
                            sumar_al_inventario(clase.dt1.Tables("faltante").Rows(z)("rev_articulo"), clase.dt1.Tables("faltante").Rows(z)("rev_faltante"))
                        Next
                        clase.borradoautomatico("delete from faltante_transferencia where rev_transferencia = " & textBox2.Text) 'borro para llenarlo con las nuevas diferencias (en la nueva revision)
                    End If
                    ind = True  'esta variable la utilizo para controlar la funcion del boton btnaceptar
                    indicadorcerrado = True 'esta variable me indica si debo preguntar para cerra el formulario cuando estoy en plena revisión
                End If
            Else
                MessageBox.Show("No se encontró ninguna transferencia que coincida con el codigo escrito.", "TRANSFERENCIA NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            txtRevisor.Text = UCase(txtRevisor.Text)

            txtArticulo.Enabled = True
            btnAceptar.Enabled = False
            Me.AcceptButton = btnagregar
            txtArticulo.Focus()
            txtRevisor.Enabled = False
        End If
    End Sub

    Private Sub textBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textBox2.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub txtDestino_GotFocus(sender As Object, e As EventArgs) Handles txtDestino.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtFecha_GotFocus(sender As Object, e As EventArgs) Handles txtFecha.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtOperario_GotFocus(sender As Object, e As EventArgs) Handles txtOperario.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtDescripcion_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtReferencia_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtTotal_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtRevisadas_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnagregar.Click
        If clase.validar_cajas_text(txtArticulo, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(txtArticulo.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & txtArticulo.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(txtArticulo.Text)         '   esta linea seguramente debe ser quitada
        End If
        If Len(txtArticulo.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & txtArticulo.Text & ")", "tabla11")
            codigo_articulo = txtArticulo.Text   '  esta linea seguramente debe ser quitada
        End If
        If Len(txtArticulo.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtArticulo.Text = ""
            txtArticulo.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            clase.consultar1("SELECT articulos.ar_referencia, articulos.ar_descripcion, dettransferencia.dt_cantidad FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) WHERE (cabtransferencia.tr_estado =TRUE AND cabtransferencia.tr_revisada =FALSE AND cabtransferencia.tr_numero =" & textBox2.Text & " AND dettransferencia.dt_codarticulo =" & codigo_articulo & ")", "bus-trans")
            If clase.dt1.Tables("bus-trans").Rows.Count > 0 Then
                txtReferencia.Text = clase.dt1.Tables("bus-trans").Rows(0)("ar_referencia")
                txtDescripcion.Text = clase.dt1.Tables("bus-trans").Rows(0)("ar_descripcion")
                colocar_foto(codigo_articulo)
                frm_cantidad_revision.ShowDialog()
                If frm_cantidad_revision.FormularioCerrado = False Then
                    Dim ind As Boolean = False
                    Dim x As Short
                    Dim cant As Integer = 0
                    For x = 0 To grdRevision.RowCount - 1
                        If grdRevision.Item(0, x).Value = codigo_articulo Then
                            grdRevision.Item(3, x).Value = grdRevision.Item(3, x).Value + frm_cantidad_revision.cantidadrevisada
                            grdRevision.CurrentCell = grdRevision.Item(0, x)
                            If (grdRevision.Item(3, x).Value) <> obtener_cant_total_transferencia(textBox2.Text, codigo_articulo) Then
                                grdRevision.Rows(x).DefaultCellStyle.BackColor = Color.AliceBlue
                                If grdRevision.Item(4, x).Value <> "Inconsistencia" Then  'aqui me aseguro de que cuando ya la fila en el campo observaciones diga inconsistencia no aumentar el contador cont_inconsistencia
                                    cont_inconsistencia += 1
                                    txtinconsistencias.Text = cont_inconsistencia
                                End If
                                grdRevision.Item(4, x).Value = "Inconsistencia"
                            Else
                                If grdRevision.Item(4, x).Value = "Inconsistencia" Then  'aqui me aseguro que cuando haya una inconsistencia sea corregida el contador cont_inconsistencia disminuyo en una unidad
                                    cont_inconsistencia -= 1
                                    txtinconsistencias.Text = cont_inconsistencia
                                End If
                                grdRevision.Item(4, x).Value = "OK"
                                grdRevision.Rows(grdRevision.RowCount - 1).DefaultCellStyle.BackColor = Color.White
                            End If
                            txtArticulo.Text = ""
                            txtArticulo.Focus()
                            ind = True
                            Exit For
                        End If
                    Next
                    If ind = False Then
                        grdRevision.RowCount = grdRevision.RowCount + 1
                        grdRevision.Item(0, grdRevision.RowCount - 1).Value = codigo_articulo
                        grdRevision.Item(1, grdRevision.RowCount - 1).Value = clase.dt1.Tables("bus-trans").Rows(0)("ar_referencia")
                        grdRevision.Item(2, grdRevision.RowCount - 1).Value = clase.dt1.Tables("bus-trans").Rows(0)("ar_descripcion")
                        grdRevision.Item(3, grdRevision.RowCount - 1).Value = frm_cantidad_revision.cantidadrevisada
                        grdRevision.CurrentCell = grdRevision.Item(0, grdRevision.RowCount - 1)
                        If (grdRevision.Item(3, grdRevision.RowCount - 1).Value) <> obtener_cant_total_transferencia(textBox2.Text, codigo_articulo) Then
                            grdRevision.Item(4, grdRevision.RowCount - 1).Value = "Inconsistencia"
                            grdRevision.Rows(grdRevision.RowCount - 1).DefaultCellStyle.BackColor = Color.AliceBlue
                            cont_inconsistencia += 1
                            txtinconsistencias.Text = cont_inconsistencia
                        Else
                            grdRevision.Item(4, grdRevision.RowCount - 1).Value = "OK"
                            grdRevision.Rows(grdRevision.RowCount - 1).DefaultCellStyle.BackColor = Color.White
                        End If
                        ref_faltantes -= 1
                        txtfaltantes.Text = ref_faltantes
                    End If
                End If
                frm_cantidad_revision.Dispose()
                    txtArticulo.Text = ""
                    txtArticulo.Focus()
                Else
                    clase.consultar2("select ar_referencia, ar_descripcion from articulos where ar_codigo = " & codigo_articulo, "art")

                txtReferencia.Text = clase.dt2.Tables("art").Rows(0)("ar_referencia")
                txtDescripcion.Text = clase.dt2.Tables("art").Rows(0)("ar_descripcion")
                Dim v As String = MessageBox.Show("Este articulo no pertenece a la transferencia actual. Verifique el codigo. Si es correcto entonces ¿Desea Agregarlo a la transferencia?", "ARTICULO NO PERTENECE A TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
                If v = 6 Then
                    frm_contrasena_revision.ShowDialog()
                    If frm_contrasena_revision.validado = True Then
                        frm_cantidad_revision.ShowDialog()
                        clase.consultar2("select ar_referencia, ar_descripcion, ar_precio1, ar_precio2 from articulos where ar_codigo = " & codigo_articulo & "", "articulo")
                        clase.agregar_registro("INSERT INTO dettransferencia(dt_trnumero,dt_codarticulo,dt_cantidad,dt_danado,dt_costo,dt_venta1,dt_venta2,dt_costo2) VALUES('" & textBox2.Text & "','" & codigo_articulo & "','" & frm_cantidad_revision.cantidadrevisada & "','0','" & Str(RecuperarCostoFiscal(codigo_articulo)) & "','" & clase.dt2.Tables("articulo").Rows(0)("ar_precio1") & "','" & clase.dt2.Tables("articulo").Rows(0)("ar_precio2") & "','" & Str(RecuperarCostoReal(codigo_articulo)) & "')")
                        '    clase.agregar_registro()
                        grdRevision.RowCount = grdRevision.RowCount + 1
                        grdRevision.Item(0, grdRevision.RowCount - 1).Value = codigo_articulo
                        grdRevision.Item(1, grdRevision.RowCount - 1).Value = clase.dt2.Tables("articulo").Rows(0)("ar_referencia")
                        grdRevision.Item(2, grdRevision.RowCount - 1).Value = clase.dt2.Tables("articulo").Rows(0)("ar_descripcion")
                        grdRevision.Item(3, grdRevision.RowCount - 1).Value = frm_cantidad_revision.cantidadrevisada
                        grdRevision.Item(4, grdRevision.RowCount - 1).Value = "Agregado"
                        grdRevision.CurrentCell = grdRevision.Item(0, grdRevision.RowCount - 1)
                        txtArticulo.Text = ""
                        txtArticulo.Focus()
                        frm_cantidad_revision.Dispose()
                    End If
                    frm_contrasena_revision.Dispose()
                End If
            End If
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtArticulo.Text = ""
            txtArticulo.Focus()
        End If
    End Sub

    Private Sub txtArticulo_KeyPress(sender As Object, e As KeyPressEventArgs)
        clase.validar_numeros(e)
    End Sub

    Public Function obtener_cant_total_transferencia(transferencia As Long, articulo As Long) As Short
        clase.consultar1("select dt_cantidad from dettransferencia where dt_trnumero = " & transferencia & " and dt_codarticulo = " & articulo & "", "total")
        If clase.dt1.Tables("total").Rows.Count > 0 Then
            Return clase.dt1.Tables("total").Rows(0)("dt_cantidad")
        Else
            Return 0
        End If
    End Function

    Function referencias_transferencias() As Short
        clase.consultar2("SELECT COUNT(dt_cantidad) AS cant FROM dettransferencia WHERE (dt_trnumero =" & textBox2.Text & ")", "transferencia")
        Return clase.dt2.Tables("transferencia").Rows(0)("cant")
    End Function
   
    Private Sub txtinconsistencias_GotFocus(sender As Object, e As EventArgs) Handles txtinconsistencias.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtfaltantes_GotFocus(sender As Object, e As EventArgs) Handles txtfaltantes.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub colocar_foto(articulo As Long)
        Try
            clase.consultar("select ar_foto from articulos where ar_codigo = " & articulo & "", "tablita")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        Catch
            PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            SetImage(PictureBox1)
        End Try
    End Sub

    Private Sub grdRevision_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdRevision.CellClick
        Try
            clase.consultar("select ar_foto, ar_referencia, ar_descripcion from articulos where ar_codigo = " & grdRevision.Item(0, grdRevision.CurrentCell.RowIndex).Value & "", "tablita")
            txtReferencia.Text = clase.dt.Tables("tablita").Rows(0)("ar_referencia")
            txtDescripcion.Text = clase.dt.Tables("tablita").Rows(0)("ar_descripcion")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        Catch
            PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            SetImage(PictureBox1)
        End Try
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If grdRevision.RowCount = 0 Then Exit Sub
        If IsNumeric(grdRevision.CurrentCell.RowIndex) Then
            Dim v As String = MessageBox.Show("¿Esta seguro que desea eliminar el articulo seleccionado?", "ELIMINAR ARTICULO", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If v = 6 Then
                frm_contrasena_revision.ShowDialog()
                If frm_contrasena_revision.validado = True Then
                    Select Case grdRevision.Item(4, grdRevision.CurrentCell.RowIndex).Value
                        Case "OK"  'cuando es "Agregado" no debe hacer nada con respecto a los sumadores
                            ref_faltantes += 1
                        Case "Inconsistencia"
                            cont_inconsistencia -= 1
                            ref_faltantes += 1
                    End Select
                    txtinconsistencias.Text = cont_inconsistencia
                    txtfaltantes.Text = ref_faltantes
                    grdRevision.Rows.Remove(grdRevision.Rows(grdRevision.CurrentCell.RowIndex))
                End If
                frm_contrasena_revision.Dispose()
                txtArticulo.Focus()
            End If
        End If
    End Sub

    Private Sub txtReferencia_GotFocus1(sender As Object, e As EventArgs) Handles txtReferencia.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtDescripcion_GotFocus1(sender As Object, e As EventArgs) Handles txtDescripcion.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Dim v As String = MessageBox.Show("¿Desea deshacer la revisión actual de esta transferencia?", "DESHACER REVISIÓN", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If v = 6 Then
            reiniciar()
        End If
    End Sub

    Private Sub reiniciar()
        cont_inconsistencia = 0
        ref_faltantes = 0
        txtinconsistencias.Text = ""
        txtfaltantes.Text = ""
        grdRevision.Rows.Clear()
        PictureBox1.Image = Nothing
        txtReferencia.Text = ""
        txtDescripcion.Text = ""
        txtArticulo.Text = ""
        txtArticulo.Enabled = False
        txtDestino.Text = ""
        txtRevisor.Text = ""
        txtRevisor.Enabled = False
        txtFecha.Text = ""
        txtOperario.Text = ""
        textBox2.Text = ""
        textBox2.Enabled = True
        textBox2.Focus()
        btnAceptar.Enabled = True
        Me.AcceptButton = btnAceptar
        btnagregar.Enabled = False
        btnLimpiar.Enabled = False
        btnEliminar.Enabled = False
        btnFinalizar.Enabled = False
        ind = False
        descontar_inventario = vbEmpty
        indicadorcerrado = False
    End Sub


    Private Sub btnFinalizar_Click(sender As Object, e As EventArgs) Handles btnFinalizar.Click
        If grdRevision.RowCount = 0 Then Exit Sub
        Dim v As String = MessageBox.Show("¿Desea Finalizar la revisión de la transferencia en este momento?", "FINALIZAR REVISIÓN", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If v = 6 Then
            'verificar articulos de la transferencia que faltan por ser revisados
            clase.consultar1("select* from dettransferencia where dt_trnumero = " & textBox2.Text, "detalle")
            If clase.dt1.Tables("detalle").Rows.Count > 0 Then
                Dim x As Short
                Dim z As Short
                Dim ind As Boolean
                For x = 0 To clase.dt1.Tables("detalle").Rows.Count - 1
                    ind = False
                    For z = 0 To grdRevision.RowCount - 1
                        If grdRevision.Item(0, z).Value = clase.dt1.Tables("detalle").Rows(x)("dt_codarticulo") Then
                            ind = True
                            Exit For
                        End If
                    Next
                    If ind = False Then
                        Exit For
                    End If
                Next
                If ind = False Then
                    Dim indicador As Boolean = False
                    frm_articulos_faltantes_x_revision.ShowDialog()
                    indicador = frm_articulos_faltantes_x_revision.indicador_no_agregado
                    frm_articulos_faltantes_x_revision.Dispose()
                    If indicador = True Then
                        Exit Sub 'esto sucede cuando se presiona el boton "SI" en el formulario frm_articulos_faltantes_x_revision descargado anteriormente
                    End If
                End If
                Dim x1 As Short
                Dim ind1 As Boolean = False
                For x1 = 0 To grdRevision.RowCount - 1
                    If (grdRevision.Item(4, x1).Value = "Inconsistencia") Or (grdRevision.Item(4, x1).Value = "Agregado") Then
                        ind1 = True
                        Exit For
                    End If
                Next
                If ind1 = True Then
                    Dim validado As Boolean
                    frm_diferencias_revision.ShowDialog()
                    validado = frm_diferencias_revision.validado
                    frm_diferencias_revision.Dispose()
                    If validado = False Then
                        Exit Sub
                    End If
                End If
                If descontar_inventario = True Then 'evaluo esta variable, si es verdadero significa que viene de bodega y por ende se deben descontar los saldos
                    clase.consultar1("select* from faltante_transferencia where rev_transferencia = " & textBox2.Text, "faltante")
                    Dim z3 As Short
                    For z3 = 0 To clase.dt1.Tables("faltante").Rows.Count - 1
                        sumar_al_inventario(clase.dt1.Tables("faltante").Rows(z3)("rev_articulo"), clase.dt1.Tables("faltante").Rows(z3)("rev_faltante"))
                    Next
                End If
                SubirTransferenciaLovePos(txtRevisor.Text, textBox2.Text, CodigoTiendaDestino)
                ' imprimir_hoja_transferencia(textBox2.Text)
                'Dim v2 As String = MessageBox.Show("¿Desea enviar el archivo de la transferencia por correo electrónico?", "ENVIAR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
                'If v2 = 6 Then
                '    construir_archivo(textBox2.Text)
                'End If
                MessageBox.Show("La revisión de la transferencia se completó satisfactoriamente.", "TRANSFERENCIA REVISADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                reiniciar()
            End If
        End If
    End Sub

   

End Class