Public Class frm_ingreso_no_importado
    Dim clase As New class_library
    Dim Codigo As String
    Public i As Integer
    Public Patron, NombrePatron, ConsImport, NombreProveedor As String
    Public IdProveedor As Integer
    Dim indicador As Boolean
    Dim ind_guardado As Boolean
    Public codigo_guardado As Short
    Public indicador_presionador_de_boton As Boolean  'este varible me indica cuando el formulario patron de distribucion se descarga si lo hace por accion del boton aceptar
    Dim filaActual As Integer
    Dim MovCreado As Boolean


    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    'Public Sub cargarconsecutivos()
    '    clase.consultar("SELECT MAX(op_codigo) AS CODIGO FROM ordenproduccion", "consecutivo")
    '    If IsDBNull(clase.dt.Tables("consecutivo").Rows(0)("CODIGO")) Then
    '        txtConsOrden.Text = "1"
    '    Else
    '        txtConsOrden.Text = clase.dt.Tables("consecutivo").Rows(0)("CODIGO") + 1
    '    End If
    '    clase.consultar("SELECT MAX(cabsal_cod) AS ConsCabSal FROM cabsalidas_mercancia", "conscabsal")
    '    If IsDBNull(clase.dt.Tables("conscabsal").Rows(0)("ConsCabSal")) Then
    '        txtConsMov.Text = "1"
    '    Else
    '        txtConsMov.Text = clase.dt.Tables("conscabsal").Rows(0)("ConsCabSal") + 1
    '    End If
    'End Sub

    Private Sub frmIngresoNoImportados_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '     PrepararColumnas()
        '   cargarconsecutivos()
        indicador = False
        ind_guardado = False
        indicador_presionador_de_boton = False
    End Sub

    'Public Sub PrepararColumnas()
    '    'PREPARAMOS LAS COLUMNAS PARA RECIBIR LAS INFORMACION DE LOS ARTICULOS
    '    With grdArticulo
    '        .Columns.Add("0", "Codigo")
    '        .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '        .Columns(0).ReadOnly = True
    '        .Columns(0).Width = 80
    '        .Columns.Add("1", "Referencia")
    '        .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    '        .Columns(1).ReadOnly = True
    '        .Columns(1).Width = 200
    '        .Columns.Add("2", "Descripción")
    '        .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    '        .Columns(2).Width = 200
    '        .Columns(2).DefaultCellStyle.WrapMode = DataGridViewTriState.True
    '        .Columns(2).ReadOnly = True
    '        .Columns.Add("3", "Cant")
    '        .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '        .Columns(3).ReadOnly = False
    '        .Columns(3).Width = 50
    '        .Columns.Add("4", "Costo Unitario")
    '        .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '        .Columns(4).ReadOnly = False
    '        .Columns(4).Visible = True
    '        .Columns(4).Width = 80
    '        .Columns(4).DefaultCellStyle = DataGridViewCellStyle2
    '        .Columns.Add("5", "IdProv")
    '        .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '        .Columns(5).ReadOnly = True
    '        .Columns(5).Visible = False
    '        .Columns.Add("6", "Prov")
    '        .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '        .Columns(6).ReadOnly = True
    '        .Columns(6).Width = 70
    '        .Columns(6).Visible = False
    '        .RowHeadersWidth = 4
    '    End With
    'End Sub

    Sub guardar_datos()
        clase.consultar("SELECT a.ar_codigo AS CODIGO, a.ar_referencia AS REFERENCIA, a.ar_descripcion AS DESCRIPCION, a.ar_precio1 AS PRECIO_1, a.ar_precio2 AS PRECIO_2, c.colornombre, t.nombretalla FROM articulos a INNER JOIN colores c ON (c.cod_color = a.ar_color) INNER JOIN tallas t ON (a.ar_talla = t.codigo_talla) WHERE a.ar_codigobarras='" & txtArticulo.Text & "' OR a.ar_codigo='" & txtArticulo.Text & "'", "articulo")
        If clase.dt.Tables("articulo").Rows.Count > 0 Then
            If ind_guardado = False And (MovCreado = False) Then
                clase.agregar_registro("INSERT INTO `cabmer_noimportada`(`cabmer_codigo`,`cabmer_importacion`,`cabmer_proveedor`,`cabmer_operario`,`cabmer_procesada`) VALUES ( NULL,'" & ConsImport & "','" & IdProveedor & "','" & UCase(txtOperario.Text) & "',False)")
                codigo_guardado = calcular_consecutivo_cabmer()
                ind_guardado = True
            End If
            clase.consultar1("SELECT dt.detmer_cantidad, dt.detmer_costounitario FROM detmer_noimportada dt INNER JOIN cabmer_noimportada cb ON (dt.detmer_codigo_imp = cb.cabmer_codigo) WHERE (cb.cabmer_importacion =" & ConsImport & " AND cb.cabmer_proveedor =" & IdProveedor & " AND dt.detmer_articulo = " & txtArticulo.Text & ")", "art")
            If clase.dt1.Tables("art").Rows.Count > 0 Then
                Dim costo As Double = clase.dt1.Tables("art").Rows(0)("detmer_costounitario")
                Dim cant As Integer
                frm_cantidad_revision.ShowDialog()
                If frm_cantidad_revision.FormularioCerrado = False Then

                    If IsDBNull(clase.dt1.Tables("art").Rows(0)("detmer_cantidad")) Then
                        cant = 0
                    Else
                        cant = clase.dt1.Tables("art").Rows(0)("detmer_cantidad")
                    End If
                    Dim v As DialogResult = MessageBox.Show("Se encontraron " & cant & "  unidades de este articulos que ya fueron ingresados con anterioridad en esta entrada de mercancia con este mismo proveedor. La cantidad " & frm_cantidad_revision.cantidadrevisada & " que ha especificado se sumará a la ya guardada con anterioridad para un total de " & cant + frm_cantidad_revision.cantidadrevisada & " unidades. ¿Desea Continuar?", "ARTICULOS YA INGRESADOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    If v = DialogResult.Yes Then
                        If frm_cantidad_revision.cantidadrevisada > 0 Then
                            clase.actualizar("UPDATE detmer_noimportada SET detmer_cantidad = " & frm_cantidad_revision.cantidadrevisada + cant & " WHERE (detmer_codigo_imp =" & codigo_guardado & " AND detmer_articulo =" & clase.dt.Tables("articulo").Rows(0)("CODIGO") & ")")
                        End If
                        Dim a As Integer
                        Dim enc As Boolean = False
                        For a = 0 To dtgDetalleArticulos.RowCount - 1
                            If txtArticulo.Text = dtgDetalleArticulos(0, a).Value.ToString Then
                                dtgDetalleArticulos.Rows(a).Cells(3).Selected = True
                                dtgDetalleArticulos.CurrentCell = dtgDetalleArticulos.Rows(a).Cells(3)
                                If MovCreado Then
                                    dtgDetalleArticulos.Item(3, a).Value = cant + frm_cantidad_revision.cantidadrevisada
                                Else
                                    dtgDetalleArticulos.Item(3, a).Value = cant + frm_cantidad_revision.cantidadrevisada 'dtgDetalleArticulos.Item(3, a).Value
                                End If
                                txtArticulo.Text = ""
                                txtArticulo.Focus()
                                enc = True
                                Exit For
                            End If
                        Next
                        If enc = False Then
                            AgregarItem(clase.dt.Tables("articulo").Rows(0)("CODIGO"), clase.dt.Tables("articulo").Rows(0)("REFERENCIA"), clase.dt.Tables("articulo").Rows(0)("DESCRIPCION"), clase.dt.Tables("articulo").Rows(0)("colornombre"), clase.dt.Tables("articulo").Rows(0)("nombretalla"), frm_cantidad_revision.cantidadrevisada, IdProveedor)
                            dtgDetalleArticulos.Rows(dtgDetalleArticulos.RowCount - 1).Cells(3).Selected = True
                            dtgDetalleArticulos.CurrentCell = dtgDetalleArticulos.Rows(dtgDetalleArticulos.RowCount - 1).Cells(3)
                            ' If MovCreado Then
                            dtgDetalleArticulos.Item(3, dtgDetalleArticulos.RowCount - 1).Value = cant + frm_cantidad_revision.cantidadrevisada
                            dtgDetalleArticulos.Item(4, dtgDetalleArticulos.RowCount - 1).Value = costo.ToString("C0")
                            ' Else
                            ' dtgDetalleArticulos.Item(3, dtgDetalleArticulos.RowCount - 1).Value = cant + frm_cantidad_revision.cantidadrevisada

                            txtArticulo.Text = ""
                            txtArticulo.Focus()
                            'End If
                        End If
                    Else
                        txtArticulo.Text = ""
                        txtArticulo.Focus()
                    End If
                Else
                    txtArticulo.Text = ""
                    txtArticulo.Focus()
                End If
                frm_cantidad_revision.Dispose()
            Else
                frm_cantidad_revision.ShowDialog()
                If frm_cantidad_revision.FormularioCerrado = False Then
                    If frm_cantidad_revision.cantidadrevisada > 0 Then
                        clase.agregar_registro("INSERT INTO `detmer_noimportada`(`detmer_codigo`,`detmer_codigo_imp`,`detmer_articulo`,`detmer_cantidad`,`detmer_patron`,`detmer_proveedor`) VALUES ( NULL,'" & codigo_guardado & "','" & clase.dt.Tables("articulo").Rows(0)("CODIGO") & "','" & frm_cantidad_revision.cantidadrevisada & "',NULL,'" & IdProveedor & "')")
                        AgregarItem(clase.dt.Tables("articulo").Rows(0)("CODIGO"), clase.dt.Tables("articulo").Rows(0)("REFERENCIA"), clase.dt.Tables("articulo").Rows(0)("DESCRIPCION"), clase.dt.Tables("articulo").Rows(0)("colornombre"), clase.dt.Tables("articulo").Rows(0)("nombretalla"), frm_cantidad_revision.cantidadrevisada, IdProveedor)
                    End If
                Else
                    txtArticulo.Text = ""
                    txtArticulo.Focus()
                End If
                frm_cantidad_revision.Dispose()
            End If
        Else
            MessageBox.Show("El articulo con el codigo digitado no existe.", "ARTICULO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtArticulo.Enabled = True
            txtArticulo.Text = ""
            txtArticulo.Focus()
        End If
    End Sub

    Private Sub AgregarItem(Codigo As Integer, Referencia As String, Descripcion As String, Color As String, Talla As String, Cantidad As Integer, IdProveedor As Integer)
        i = dtgDetalleArticulos.Rows.Add()
        dtgDetalleArticulos(0, i).Value = Codigo
        dtgDetalleArticulos(1, i).Value = Referencia
        dtgDetalleArticulos(2, i).Value = Descripcion
        dtgDetalleArticulos(3, i).Value = Color
        dtgDetalleArticulos(4, i).Value = Talla
        dtgDetalleArticulos(5, i).Value = Cantidad
        dtgDetalleArticulos(7, i).Value = IdProveedor
        dtgDetalleArticulos.CurrentCell = dtgDetalleArticulos.Rows(i).Cells(6)
        txtArticulo.Text = ""
        txtArticulo.Enabled = False
        btnConsulta.Enabled = False
        EstablecerFoto(Codigo)
    End Sub

    Function calcular_consecutivo_cabmer() As Short
        clase.consultar2("SELECT MAX(cabmer_codigo) AS consecutivo FROM cabmer_noimportada", "maximo")
        If IsDBNull(clase.dt2.Tables("maximo").Rows(0)("consecutivo")) Then
            Return 1
        Else
            Return clase.dt2.Tables("maximo").Rows(0)("consecutivo")
        End If
    End Function

    Private Sub txtArticulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtArticulo.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub txtArticulo_LostFocus(sender As Object, e As EventArgs) Handles txtArticulo.LostFocus
        'VALIDAMOS SI LOS DATOS DEL ENCABEZADO ESTAN BIEN DIGITADOS
        If indicador = False Then
            If clase.validar_cajas_text(txtOperario, "NOMBRE DE OPERARIO") And clase.validar_cajas_text(txtImportacion, "IMPORTACION") Then
                btnDeshacer.Enabled = True
                btnGuardar.Enabled = True
                If txtArticulo.Text <> "" Then
                    'Agregar()

                    guardar_datos()
                End If
            End If
        End If
    End Sub

    'Private Sub grdArticulo_CellClick(sender As Object, e As DataGridViewCellEventArgs)

    'End Sub

    'Private Sub grdArticulo_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)

    'End Sub

    Private Sub validar_Keypress(
      ByVal sender As Object,
      ByVal e As System.Windows.Forms.KeyPressEventArgs)
        'obtener indice de la columna  
        Dim columna As Integer = dtgDetalleArticulos.CurrentCell.ColumnIndex
        'comprobar si la celda en edición corresponde a la columna 5
        If columna = 5 Then
            'Obtener caracter
            Dim caracter As Char = e.KeyChar
            'comprobar si el caracter es un número o el retroceso  
            If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
                Me.Text = e.KeyChar
                e.KeyChar = Chr(0)
            End If
        End If
        If columna = 6 Then
            'Obtener caracter
            Dim caracter As Char = e.KeyChar
            'comprobar si el caracter es un número o el retroceso  
            If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
                Me.Text = e.KeyChar
                e.KeyChar = Chr(0)
            End If
        End If
    End Sub

    Private Sub btnImportacion_Click(sender As Object, e As EventArgs) Handles btnImportacion.Click
        frm_importacion_mercancia_no_importada.ShowDialog()
        frm_importacion_mercancia_no_importada.Dispose()
    End Sub

    Public Sub Limpiar()
        indicador = True
        ' GroupBox1.Enabled = True
        txtImportacion.Text = ""
        txtOperario.Text = ""
        dtgDetalleArticulos.Rows.Clear()
        btnDeshacer.Enabled = False
        btnGuardar.Enabled = False
        btnImportacion.Enabled = True
        btnImportacion.Focus()
        Button1.Enabled = True
        txtArticulo.Enabled = False
        btnConsulta.Enabled = False
        txtOperario.Enabled = True
        indicador = False
        button5.Enabled = False
        '   cargarconsecutivos()
        ind_guardado = False
        txtProveedor.Text = ""
        PictureBox1.Image = Nothing
    End Sub

    Private Sub btnDeshacer_Click(sender As Object, e As EventArgs) Handles btnDeshacer.Click
        Limpiar()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If dtgDetalleArticulos.RowCount = 0 Then
            Exit Sub
        End If
        clase.consultar("SELECT* FROM detmer_noimportada WHERE (detmer_codigo_imp =" & codigo_guardado & ")", "detmer")
        If clase.dt.Tables("detmer").Rows.Count > 0 Then
            Dim x As Short
            Dim indvalidad As Boolean = False
            For x = 0 To clase.dt.Tables("detmer").Rows.Count - 1
                If IsDBNull(clase.dt.Tables("detmer").Rows(x)("detmer_cantidad")) Then
                    MessageBox.Show("Debe establecer las cantidades para cada item agregado.", "DATOS INCOMPLETOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    dtgDetalleArticulos.CurrentCell = dtgDetalleArticulos.Item(5, x)
                    indvalidad = True
                    Exit For
                End If
            Next
            If indvalidad = True Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If


        Dim a As Integer
        '1. ENTRADAMERCANCIA
        For a = 0 To dtgDetalleArticulos.RowCount - 1
            clase.consultar("SELECT entradamercancia.* FROM entradamercancia WHERE (com_codigoart =" & dtgDetalleArticulos.Item(0, a).Value & " AND com_codigoimp =" & ConsImport & ")", "entradamercancia")
            If clase.dt.Tables("entradamercancia").Rows.Count > 0 Then
                Dim cant As Integer = clase.dt.Tables("entradamercancia").Rows(0)("com_unidades")
                clase.actualizar("update entradamercancia set com_unidades = " & cant + dtgDetalleArticulos(5, a).Value & " WHERE (com_codigoart =" & dtgDetalleArticulos.Item(0, a).Value & " AND com_codigoimp =" & ConsImport & ")")
            Else
                clase.agregar_registro("INSERT INTO entradamercancia(com_codigoimp,com_codigoart,com_unidades,com_danado,com_operario) VALUES('" & ConsImport & "','" & dtgDetalleArticulos(0, a).Value.ToString & "','" & dtgDetalleArticulos(5, a).Value.ToString & "','0','" & txtOperario.Text & "')")
            End If

        Next
        '8. DETALLE_PROVEEDORES ARTICULO
        For d = 0 To dtgDetalleArticulos.RowCount - 1
            clase.consultar("SELECT * FROM detalle_proveedores_articulos WHERE codigo_articulo='" & dtgDetalleArticulos(0, d).Value & "'", "prov")
            If clase.dt.Tables("prov").Rows.Count = 0 Then
                clase.agregar_registro("INSERT INTO detalle_proveedores_articulos(codigo_articulo,codigo_proveedor) VALUES('" & dtgDetalleArticulos(0, d).Value & "','" & dtgDetalleArticulos(7, d).Value & "')")
            ElseIf dtgDetalleArticulos(5, d).Value <> clase.dt.Tables("prov").Rows(0)("codigo_proveedor") Then
                clase.agregar_registro("INSERT INTO detalle_proveedores_articulos(codigo_articulo,codigo_proveedor) VALUES('" & dtgDetalleArticulos(0, d).Value & "','" & dtgDetalleArticulos(7, d).Value & "')")
            End If
        Next


        clase.actualizar("UPDATE cabmer_noimportada set cabmer_procesada = TRUE WHERE cabmer_codigo = " & codigo_guardado & "")
        ind_guardado = False
        codigo_guardado = vbEmpty
        MigrarArticulosLineasProveedores()
        CrearMovimientoCompraLovePOSNacional(ConsImport, IdProveedor, txtOperario.Text.ToUpper)
        MessageBox.Show("Datos guardados satisfactoriamente.", "Operación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Limpiar()

    End Sub

    Private Sub btnConsulta_Click(sender As Object, e As EventArgs) Handles btnConsulta.Click
        frm_busqueda_articulos.ShowDialog()
        frm_busqueda_articulos.Dispose()
        txtArticulo.Focus()
        If txtArticulo.Text <> "" Then
            SendKeys.Send("{ENTER}")
        End If
    End Sub

    Private Sub btnProveedor_Click_1(sender As Object, e As EventArgs)
        frm_proveedores_mercancia_no_importada.ShowDialog()
        frm_proveedores_mercancia_no_importada.Dispose()
    End Sub

    Private Sub txtConsMov_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtImportacion_GotFocus(sender As Object, e As EventArgs) Handles txtImportacion.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(txtImportacion, "Importación") = False Then Exit Sub
        If clase.validar_cajas_text(txtProveedor, "Proveedor") = False Then Exit Sub
        If clase.validar_cajas_text(txtOperario, "Operario") = False Then Exit Sub
        clase.consultar1("SELECT* FROM cabmer_noimportada WHERE cabmer_proveedor = " & IdProveedor & " AND cabmer_importacion = " & ConsImport & "", "busqueda-mov")
        If clase.dt1.Tables("busqueda-mov").Rows.Count > 0 Then
            'MessageBox.Show("Ya existe un ingreso de mercancía con este proveedor e importación. ", "RELACIÓN DE PRODUCTO YA INGRESADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Exit Sub
            MovCreado = True
            codigo_guardado = clase.dt1.Tables("busqueda-mov").Rows(0)("cabmer_codigo")
        Else
            MovCreado = False
        End If
        btnGuardar.Enabled = True
        btnDeshacer.Enabled = True
        Button1.Enabled = False
        txtArticulo.Enabled = True
        txtArticulo.Focus()
        btnImportacion.Enabled = False
        btnConsulta.Enabled = True
        txtOperario.Enabled = False
        button5.Enabled = True
    End Sub


    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        If dtgDetalleArticulos.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dtgDetalleArticulos.CurrentCell.RowIndex) Then
            clase.borradoautomatico("delete from detmer_noimportada WHERE detmer_codigo_imp = " & codigo_guardado & " AND detmer_articulo = " & dtgDetalleArticulos.Item(0, dtgDetalleArticulos.CurrentRow.Index).Value & " ")
            dtgDetalleArticulos.Rows.Remove(dtgDetalleArticulos.Rows(dtgDetalleArticulos.CurrentCell.RowIndex))

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frm_proveedores_mercancia_no_importada.ShowDialog()
        frm_proveedores_mercancia_no_importada.Dispose()
        txtProveedor.Text = NombreProveedor
    End Sub



    'Private Sub grdArticulo_KeyDown(sender As Object, e As KeyEventArgs) Handles grdArticulo.KeyDown

    'End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        frm_mercancia_no_guardad_suspedidas.ShowDialog()
        frm_mercancia_no_guardad_suspedidas.Dispose()
    End Sub

    'Sub llenar_grilla_mercancia(codigo As Short) 'esta funcion la utilizo para cargar los movimientos suspendidos, 
    '    clase.consultar("SELECT detmer_noimportada.detmer_articulo, articulos.ar_referencia, articulos.ar_descripcion, detmer_noimportada.detmer_cantidad, cabpatrondist.pat_codigo, cabpatrondist.pat_nombre,proveedores.prv_codigo,proveedores.prv_codigoasignado FROM detmer_noimportada INNER JOIN cabmer_noimportada ON (detmer_noimportada.detmer_codigo_imp = cabmer_noimportada.cabmer_codigo) INNER JOIN articulos ON (detmer_noimportada.detmer_articulo = articulos.ar_codigo) INNER JOIN proveedores ON (detmer_noimportada.detmer_proveedor = proveedores.prv_codigo) LEFT JOIN cabpatrondist ON (detmer_noimportada.detmer_patron = cabpatrondist.pat_codigo) WHERE (cabmer_noimportada.cabmer_codigo =" & codigo & ")", "mov")
    '    If clase.dt.Tables("mov").Rows.Count > 0 Then
    '        grdArticulo.Rows.Clear()
    '        Dim c As Short
    '        For c = 0 To clase.dt.Tables("mov").Rows.Count - 1
    '            grdArticulo.RowCount = grdArticulo.RowCount + 1
    '            grdArticulo.Item(0, c).Value = clase.dt.Tables("mov").Rows(c)("detmer_articulo")
    '            grdArticulo.Item(1, c).Value = clase.dt.Tables("mov").Rows(c)("ar_referencia")
    '            grdArticulo.Item(2, c).Value = clase.dt.Tables("mov").Rows(c)("ar_descripcion")
    '            If IsDBNull(clase.dt.Tables("mov").Rows(c)("detmer_cantidad")) Then
    '                grdArticulo.Item(3, c).Value = ""
    '            Else
    '                grdArticulo.Item(3, c).Value = clase.dt.Tables("mov").Rows(c)("detmer_cantidad")
    '            End If
    '            If IsDBNull(clase.dt.Tables("mov").Rows(c)("pat_codigo")) Then
    '                grdArticulo.Item(4, c).Value = ""
    '                grdArticulo.Item(5, c).Value = ""
    '            Else
    '                grdArticulo.Item(4, c).Value = clase.dt.Tables("mov").Rows(c)("pat_codigo")
    '                grdArticulo.Item(5, c).Value = clase.dt.Tables("mov").Rows(c)("pat_nombre")
    '            End If
    '            If IsDBNull(clase.dt.Tables("mov").Rows(c)("prv_codigo")) Then
    '                grdArticulo.Item(6, c).Value = ""
    '                grdArticulo.Item(7, c).Value = ""
    '            Else
    '                grdArticulo.Item(6, c).Value = clase.dt.Tables("mov").Rows(c)("prv_codigo")
    '                grdArticulo.Item(7, c).Value = clase.dt.Tables("mov").Rows(c)("prv_codigoasignado")
    '            End If
    '        Next
    '    Else
    '        grdArticulo.Rows.Clear()
    '    End If
    'End Sub

    Private Sub txtConsOrden_GotFocus1(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub





    'Private Sub grdArticulo_KeyPress(sender As Object, e As KeyPressEventArgs)
    '    Dim p As Short = grdArticulo.CurrentRow.Index
    '    If grdArticulo(3, p).Value = Nothing Then
    '        MessageBox.Show("DEBES LLENAR LA CANTIDAD DE ESTE PRODUCTO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        Exit Sub
    '    End If
    '    If Asc(e.KeyChar) = 13 Then
    '        'ELEGIR EL PROVEEDOR
    '        'frm_proveedores_mercancia_no_importada.ShowDialog()
    '        'frm_proveedores_mercancia_no_importada.Dispose()

    '        '  grdArticulo(6, p).Value = IdProveedor
    '        '  grdArticulo(7, p).Value = CodProveedor

    '        'ELEGIR EL PATRON DE DISTRIBUCION
    '        '   frm_patron_mercancia_no_importada.ShowDialog()
    '        '  frm_patron_mercancia_no_importada.Dispose()
    '        'VALIDAR SI LA CANTIDAD CUMPLE CON LAS REQUERIDAS POR EL PATRON
    '        If indicador_presionador_de_boton = True Then
    '            '   clase.consultar("SELECT cabpatrondist.pat_codigo AS PATRON, cabpatrondist.pat_nombre AS NOMBRE, SUM(detpatrondist.dp_cantidad) AS SUMPATRON FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo)WHERE (cabpatrondist.pat_codigo ='" & Patron & "') GROUP BY cabpatrondist.pat_codigo;", "patron")
    '            '  If grdArticulo(3, p).Value >= clase.dt.Tables("patron").Rows(0)("SUMPATRON") Then
    '            '  grdArticulo(4, p).Value = Patron
    '            ' grdArticulo(5, p).Value = clase.dt.Tables("patron").Rows(0)("NOMBRE")
    '            '   clase.actualizar("UPDATE detmer_noimportada SET detmer_proveedor=" & grdArticulo.Item(6, p).Value & " WHERE detmer_codigo_imp = " & codigo_guardado & " AND detmer_articulo = " & grdArticulo.Item(0, p).Value & "")
    '            'Else
    '            '    clase.actualizar("UPDATE detmer_noimportada SET detmer_cantidad = NULL WHERE detmer_codigo_imp = " & codigo_guardado & " AND detmer_articulo = " & grdArticulo.Item(0, p).Value & "")
    '            '    MessageBox.Show("LA CANTIDAD NO CUMPLE CON LAS REQUERIDAS PARA SURTIR LAS TIENDAS", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            '    grdArticulo.CurrentCell = grdArticulo.Rows(p).Cells(3)
    '            '    grdArticulo(3, p).Value = ""
    '            '    Exit Sub
    '            'End If
    '            'FIN ELECCION DE PATRON
    '            btnConsulta.Enabled = True
    '            txtArticulo.Enabled = True
    '            txtArticulo.Focus()
    '            indicador_presionador_de_boton = False
    '        End If
    '    End If
    'End Sub

    Private Sub dtgDetalleArticulos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles dtgDetalleArticulos.KeyPress
        Dim p As Short = dtgDetalleArticulos.CurrentCell.ColumnIndex
        '   If indicador_presionador_de_boton = True Then
        Select Case p
            Case 5
                If dtgDetalleArticulos.Item(5, dtgDetalleArticulos.CurrentRow.Index).Value = Nothing Then
                    MessageBox.Show("Debe especificar la cantidad de este producto.", "ESPECIFICAR CANTIDAD", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
                dtgDetalleArticulos.CurrentCell = dtgDetalleArticulos.Item(6, dtgDetalleArticulos.CurrentCell.RowIndex)
            Case 6
                If dtgDetalleArticulos.Item(6, dtgDetalleArticulos.CurrentRow.Index).Value = Nothing Then
                    MessageBox.Show("Debe especificar el costo unitario de este producto.", "ESPECIFICAR COSTO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
                txtArticulo.Enabled = True
                btnConsulta.Enabled = True
                txtArticulo.Focus()
            End Select
        ' End If
    End Sub

    Private Sub dtgDetalleArticulos_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dtgDetalleArticulos.EditingControlShowing
        'referencia a la celda  
        Dim validar As TextBox = CType(e.Control, TextBox)
        'agregar el controlador de eventos para el KeyPress  
        AddHandler validar.KeyPress, AddressOf validar_Keypress
    End Sub

    Private Sub dtgDetalleArticulos_KeyDown(sender As Object, e As KeyEventArgs) Handles dtgDetalleArticulos.KeyDown
        If e.KeyCode = Keys.Return Then
            e.Handled = True
        End If
    End Sub


    Private Sub dtgDetalleArticulos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtgDetalleArticulos.CellClick
        If dtgDetalleArticulos.RowCount > 0 Then
            Dim Y As Integer = dtgDetalleArticulos.CurrentCell.RowIndex
            Codigo = dtgDetalleArticulos(0, [Y]).Value.ToString
            EstablecerFoto(Codigo)
        End If
    End Sub

    Private Sub dtgDetalleArticulos_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dtgDetalleArticulos.CellEndEdit
        Dim col As Short = dtgDetalleArticulos.CurrentCell.ColumnIndex
        If col = 5 Then

            clase.actualizar("UPDATE detmer_noimportada SET detmer_cantidad = " & dtgDetalleArticulos.Item(5, dtgDetalleArticulos.CurrentRow.Index).Value & " WHERE detmer_codigo_imp = " & codigo_guardado & " AND detmer_articulo = " & dtgDetalleArticulos.Item(0, dtgDetalleArticulos.CurrentRow.Index).Value & "")
        End If
        If col = 6 Then
            clase.actualizar("UPDATE detmer_noimportada SET detmer_costounitario = " & Str(dtgDetalleArticulos.Item(6, dtgDetalleArticulos.CurrentRow.Index).Value) & " WHERE detmer_codigo_imp = " & codigo_guardado & " AND detmer_articulo = " & dtgDetalleArticulos.Item(0, dtgDetalleArticulos.CurrentRow.Index).Value & "")
            Dim Costo As Double = dtgDetalleArticulos.Item(6, dtgDetalleArticulos.CurrentRow.Index).Value
            dtgDetalleArticulos.Item(6, dtgDetalleArticulos.CurrentRow.Index).Value = Costo.ToString("C0")
        End If
    End Sub

    Private Sub dtgDetalleArticulos_RowLeave(sender As Object, e As DataGridViewCellEventArgs) Handles dtgDetalleArticulos.RowLeave

    End Sub

    Private Sub txtArticulo_TextChanged(sender As Object, e As EventArgs) Handles txtArticulo.TextChanged

    End Sub

    Private Sub ValidarFila(sender As Object, data As DataGridViewCellCancelEventArgs) Handles dtgDetalleArticulos.RowValidating
        If dtgDetalleArticulos.RowCount > 0 Then
            filaActual = dtgDetalleArticulos.CurrentCell.RowIndex
            If dtgDetalleArticulos.Item(5, filaActual).Value = Nothing Then
                MessageBox.Show("Debe escribir algo en el campo cantidad.", "ESPECIFICAR CANTIDAD", MessageBoxButtons.OK, MessageBoxIcon.Error)
                data.Cancel = True
                Exit Sub
            End If
            If dtgDetalleArticulos.Item(6, filaActual).Value = Nothing Then
                If dtgDetalleArticulos.Item(6, filaActual).Value = Nothing Then
                    MessageBox.Show("Debe escribir algo en el campo costo unitario.", "ESPECIFICAR COSTO UNITARIO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    data.Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Sub EstablecerFoto(IdArticulo As Integer)
        Try
            clase.consultar("select ar_foto from articulos where ar_codigo = " & IdArticulo & "", "tablita")
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

    Private Sub txtProveedor_GotFocus(sender As Object, e As EventArgs) Handles txtProveedor.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub
End Class