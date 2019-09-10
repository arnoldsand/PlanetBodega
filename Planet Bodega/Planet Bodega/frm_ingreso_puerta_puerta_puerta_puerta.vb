Public Class frm_ingreso_puerta_puerta_puerta_puerta
    Dim clase As New class_library
    Dim Codigo As String
    Public i As Integer
    Public Patron, NombrePatron, ConsImport, IdProveedor, CodProveedor As String
    Dim indicador As Boolean
    Dim ind_guardado As Boolean
    Public codigo_guardado As Short
    Public indicador_presionador_de_boton As Boolean  'este varible me indica cuando el formulario patron de distribucion se descarga si lo hace por accion del boton aceptar
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Public Sub cargarconsecutivos()
        clase.consultar("SELECT MAX(op_codigo) AS CODIGO FROM ordenproduccion", "consecutivo")
        If IsDBNull(clase.dt.Tables("consecutivo").Rows(0)("CODIGO")) Then
            txtConsOrden.Text = "1"
        Else
            txtConsOrden.Text = clase.dt.Tables("consecutivo").Rows(0)("CODIGO") + 1
        End If
        clase.consultar("SELECT MAX(cabsal_cod) AS ConsCabSal FROM cabsalidas_mercancia", "conscabsal")
        If IsDBNull(clase.dt.Tables("conscabsal").Rows(0)("ConsCabSal")) Then
            txtConsMov.Text = "1"
        Else
            txtConsMov.Text = clase.dt.Tables("conscabsal").Rows(0)("ConsCabSal") + 1
        End If
    End Sub



    Public Sub PrepararColumnas()
        'PREPARAMOS LAS COLUMNAS PARA RECIBIR LAS INFORMACION DE LOS ARTICULOS
        With grdArticulo
            .Columns.Add("0", "Codigo")
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).ReadOnly = True
            .Columns(0).Width = 80
            .Columns.Add("1", "Referencia")
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(1).ReadOnly = True
            .Columns(1).Width = 150
            .Columns.Add("2", "Descripción")
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(2).Width = 150
            .Columns(2).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .Columns(2).ReadOnly = True
            .Columns.Add("3", "Cant")
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).ReadOnly = False
            .Columns(3).Width = 50
            .Columns.Add("4", "CodPatrón")
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).ReadOnly = True
            .Columns(4).Visible = False
            .Columns.Add("5", "Patrón")
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).ReadOnly = True
            .Columns(5).Width = 220
            .Columns.Add("6", "Registro Importación")
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            '.Columns(6).ReadOnly = True
            .Columns(6).Width = 150
            .Columns.Add("7", "Fecha Registro")
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(7).ReadOnly = True
            .Columns(7).Width = 100
            .Columns.Add("8", "IdProv")
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).ReadOnly = True
            .Columns(8).Visible = False
            .Columns.Add("9", "Prov")
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).ReadOnly = True
            .Columns(9).Width = 70
            .RowHeadersWidth = 4
        End With
    End Sub

    Sub guardar_datos()
        clase.consultar("SELECT ar_codigo AS CODIGO, ar_referencia AS REFERENCIA, ar_descripcion AS DESCRIPCION, ar_precio1 AS PRECIO_1, ar_precio2 AS PRECIO_2 FROM articulos WHERE ar_codigobarras='" & txtArticulo.Text & "' OR ar_codigo='" & txtArticulo.Text & "'", "articulo")
        If clase.dt.Tables("articulo").Rows.Count > 0 Then

            If ind_guardado = False Then
                clase.agregar_registro("INSERT INTO `cabmer_puertapuerta`(`cabpuerta_codigo`,`cabpuerta_importacion`,`cabpuerta_operario`,`cabpuerta_procesada`) VALUES ( NULL,'" & ConsImport & "','" & UCase(txtOperario.Text) & "',False)")
                codigo_guardado = calcular_consecutivo_cabpuerta()
                ind_guardado = True
            End If

            clase.consultar1("SELECT detpuerta_cantidad FROM detmer_puertapuerta WHERE (detpuerta_codigopuerta =" & codigo_guardado & " AND detpuerta_articulo =" & clase.dt.Tables("articulo").Rows(0)("CODIGO") & ")", "art")
            If clase.dt1.Tables("art").Rows.Count > 0 Then

                MessageBox.Show("Este producto ya se encuentra agregado en el listado, modifique el listado. ", "PRODUCTO YA AGREGADO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Dim a As Integer
                For a = 0 To grdArticulo.RowCount - 1
                    If txtArticulo.Text = grdArticulo(0, a).Value.ToString Or txtArticulo.Text = grdArticulo(1, a).Value.ToString Then
                        grdArticulo.Rows(a).Cells(3).Selected = True
                        grdArticulo.CurrentCell = grdArticulo.Rows(a).Cells(3)
                        txtArticulo.Text = ""
                        txtArticulo.Focus()
                        Exit For
                    End If
                Next

            Else
                ''consultar el ultimo registro de importacion del articulo
                'Dim Registro, FechaReg As String
                'Dim Fec As Date
                'clase.consultar1("SELECT detalle_importacion_detcajas.detcab_registrodian AS REG,detalle_importacion_detcajas.detcab_fecharegistrodian AS FECREG FROM asociaciones_codigos INNER JOIN articulos ON (asociaciones_codigos.asc_postcodart = articulos.ar_codigo) INNER JOIN detalle_importacion_detcajas ON (asociaciones_codigos.asc_precodref = detalle_importacion_detcajas.detcab_coditem) WHERE (articulos.ar_codigo =" & txtArticulo.Text & ");", "registro")

                'If clase.dt1.Tables("registro").Rows.Count = 0 Then
                '    'MsgBox("a")
                '    clase.agregar_registro("INSERT INTO `detmer_puertapuerta`(`detpuerta_codigo`,`detpuerta_codigopuerta`,`detpuerta_articulo`,`detpuerta_cantidad`,`detpuerta_patron`,`detpuerta_regimp`,`detpuerta_fechaimp`) VALUES ( NULL,'" & codigo_guardado & "','" & clase.dt.Tables("articulo").Rows(0)("CODIGO") & "',NULL,NULL,NULL,NULL)")
                '    Registro = ""
                '    FechaReg = ""
                'Else
                '    'MsgBox("b")
                '    If IsDBNull(clase.dt1.Tables("registro").Rows(0)("FECREG")) Then
                '        clase.agregar_registro("INSERT INTO `detmer_puertapuerta`(`detpuerta_codigo`,`detpuerta_codigopuerta`,`detpuerta_articulo`,`detpuerta_cantidad`,`detpuerta_patron`,`detpuerta_regimp`,`detpuerta_fechaimp`) VALUES ( NULL,'" & codigo_guardado & "','" & clase.dt.Tables("articulo").Rows(0)("CODIGO") & "',NULL,NULL,NULL,NULL)")
                '        Registro = ""
                '        FechaReg = ""
                '    Else
                '        Fec = clase.dt1.Tables("registro").Rows(0)("FECREG")
                '        Registro = clase.dt1.Tables("registro").Rows(0)("REG")
                '        FechaReg = Fec.ToString("yyyy-MM-dd")
                '        clase.agregar_registro("INSERT INTO `detmer_puertapuerta`(`detpuerta_codigo`,`detpuerta_codigopuerta`,`detpuerta_articulo`,`detpuerta_cantidad`,`detpuerta_patron`,`detpuerta_regimp`,`detpuerta_fechaimp`) VALUES ( NULL,'" & codigo_guardado & "','" & clase.dt.Tables("articulo").Rows(0)("CODIGO") & "',NULL,NULL,'" & Registro & "','" & FechaReg & "')")
                '    End If
                'End If

                clase.agregar_registro("INSERT INTO `detmer_puertapuerta`(`detpuerta_codigo`,`detpuerta_codigopuerta`,`detpuerta_articulo`,`detpuerta_cantidad`,`detpuerta_patron`,`detpuerta_regimp`,`detpuerta_fechaimp`) VALUES ( NULL,'" & codigo_guardado & "','" & clase.dt.Tables("articulo").Rows(0)("CODIGO") & "',NULL,NULL,NULL,NULL)")

                i = grdArticulo.Rows.Add()
                grdArticulo(0, i).Value = clase.dt.Tables("articulo").Rows(0)("CODIGO")
                grdArticulo(1, i).Value = clase.dt.Tables("articulo").Rows(0)("REFERENCIA")
                grdArticulo(2, i).Value = clase.dt.Tables("articulo").Rows(0)("DESCRIPCION")
                'grdArticulo(6, i).Value = Registro
                'grdArticulo(7, i).Value = FechaReg
                grdArticulo.CurrentCell = grdArticulo.Rows(i).Cells(3)
                txtArticulo.Text = ""
                txtArticulo.Enabled = False
                btnConsulta.Enabled = False
            End If
        Else
            MessageBox.Show("El articulo con el codigo digitado no existe.", "ARTICULO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtArticulo.Enabled = True
            txtArticulo.Text = ""
            txtArticulo.Focus()
        End If
    End Sub

    Function calcular_consecutivo_cabpuerta() As Short
        clase.consultar2("SELECT MAX(cabpuerta_codigo) AS consecutivo FROM cabmer_puertapuerta", "maximo")
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

    Private Sub grdArticulo_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdArticulo.CellClick
        If grdArticulo.RowCount > 0 Then
            Dim Y As Integer = grdArticulo.CurrentCell.RowIndex
            Codigo = grdArticulo(0, [Y]).Value.ToString
        End If
    End Sub

    Private Sub grdArticulo_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles grdArticulo.CellEndEdit
        Dim col As Short = grdArticulo.CurrentCell.ColumnIndex
        If col = 3 Then
            clase.actualizar("UPDATE detmer_puertapuerta SET detpuerta_cantidad = " & grdArticulo.Item(3, grdArticulo.CurrentRow.Index).Value & " WHERE detpuerta_codigopuerta = " & codigo_guardado & " AND detpuerta_articulo = " & grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value & "")
        End If
        If col = 6 Then
            clase.actualizar("UPDATE detmer_puertapuerta SET detpuerta_regimp = " & grdArticulo.Item(6, grdArticulo.CurrentRow.Index).Value & " WHERE detpuerta_codigopuerta = " & codigo_guardado & " AND detpuerta_articulo = " & grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value & "")
        End If
        If col = 7 Then
            If IsDate(grdArticulo(7, grdArticulo.CurrentRow.Index).Value) = False Then
                MessageBox.Show("Fecha no valida", "Planet Love", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim Fbase As Date = grdArticulo(7, grdArticulo.CurrentRow.Index).Value
                Dim Fecha As String = Fbase.ToString("yyyy-MM-dd")
                clase.actualizar("UPDATE detmer_puertapuerta SET detpuerta_fechaimp = '" & Fecha & "' WHERE detpuerta_codigopuerta = " & codigo_guardado & " AND detpuerta_articulo = " & grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value & "")
            End If
        End If
    End Sub

    Private Sub grdArticulo_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles grdArticulo.EditingControlShowing
        'referencia a la celda  
        Dim validar As TextBox = CType(e.Control, TextBox)
        'agregar el controlador de eventos para el KeyPress  
        AddHandler validar.KeyPress, AddressOf validar_Keypress
    End Sub

    Private Sub validar_Keypress( _
      ByVal sender As Object, _
      ByVal e As System.Windows.Forms.KeyPressEventArgs)
        'obtener indice de la columna  
        Dim columna As Integer = grdArticulo.CurrentCell.ColumnIndex
        'comprobar si la celda en edición corresponde a la columna 5
        If columna = 3 Or columna = 6 Or columna = 7 Then
            'Obtener caracter
            Dim caracter As Char = e.KeyChar
            'comprobar si el caracter es un número o el retroceso  
            If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
                If caracter = "-" Or caracter = "/" Then
                    Exit Sub
                End If
                Me.Text = e.KeyChar
                e.KeyChar = Chr(0)
            End If
        End If
    End Sub

    Private Sub grdArticulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdArticulo.KeyPress
        Dim p As Short = grdArticulo.CurrentRow.Index

        If grdArticulo(3, p).Value = Nothing Then
            MessageBox.Show("DEBES LLENAR LA CANTIDAD DE ESTE PRODUCTO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If Asc(e.KeyChar) = 13 Then
            'ELEGIR EL PROVEEDOR
            frm_proveedores_mercancia_no_importada_puerta_puerta.ShowDialog()
            frm_proveedores_mercancia_no_importada_puerta_puerta.Dispose()

            grdArticulo(8, p).Value = IdProveedor
            grdArticulo(9, p).Value = CodProveedor

            'ELEGIR EL PATRON DE DISTRIBUCION
            frm_patron_mercancia_no_importada_puerta_puerta.ShowDialog()
            frm_patron_mercancia_no_importada_puerta_puerta.Dispose()

            'VALIDAR SI LA CANTIDAD CUMPLE CON LAS REQUERIDAS POR EL PATRON
            If indicador_presionador_de_boton = True Then
                clase.consultar("SELECT cabpatrondist.pat_codigo AS PATRON, cabpatrondist.pat_nombre AS NOMBRE, SUM(detpatrondist.dp_cantidad) AS SUMPATRON FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo)WHERE (cabpatrondist.pat_codigo ='" & Patron & "')GROUP BY cabpatrondist.pat_codigo;", "patron")
                If grdArticulo(3, p).Value >= clase.dt.Tables("patron").Rows(0)("SUMPATRON") Then
                    grdArticulo(4, p).Value = Patron
                    grdArticulo(5, p).Value = clase.dt.Tables("patron").Rows(0)("NOMBRE")
                    clase.actualizar("UPDATE detmer_puertapuerta SET detpuerta_proveedor=" & grdArticulo.Item(8, p).Value & ",detpuerta_patron = " & grdArticulo.Item(4, p).Value & " WHERE detpuerta_codigopuerta = " & codigo_guardado & " AND detpuerta_articulo = " & grdArticulo.Item(0, p).Value & "")
                Else
                    clase.actualizar("UPDATE detmer_puertapuerta SET detpuerta_cantidad = NULL WHERE detpuerta_codigopuerta = " & codigo_guardado & " AND detpuerta_articulo = " & grdArticulo.Item(0, p).Value & "")
                    MessageBox.Show("LA CANTIDAD NO CUMPLE CON LAS REQUERIDAS PARA SURTIR LAS TIENDAS", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    grdArticulo.CurrentCell = grdArticulo.Rows(p).Cells(3)
                    grdArticulo(3, p).Value = ""
                    Exit Sub
                End If
                'FIN ELECCION DE PATRON
                btnConsulta.Enabled = True
                txtArticulo.Enabled = True
                txtArticulo.Focus()
                indicador_presionador_de_boton = False
            End If
        End If
    End Sub

    Private Sub btnImportacion_Click(sender As Object, e As EventArgs) Handles btnImportacion.Click
        frm_importacion_mercancia_no_importada_puerta_puerta.ShowDialog()
        frm_importacion_mercancia_no_importada_puerta_puerta.Dispose()
    End Sub

    Private Sub txtOperario_TextChanged(sender As Object, e As EventArgs)
        txtOperario.CharacterCasing = CharacterCasing.Upper
    End Sub
    Public Sub Limpiar()
        indicador = True
        ' GroupBox1.Enabled = True
        txtImportacion.Text = ""
        txtOperario.Text = ""
        grdArticulo.Rows.Clear()
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
        cargarconsecutivos()
        ind_guardado = False
    End Sub
    Private Sub btnDeshacer_Click(sender As Object, e As EventArgs) Handles btnDeshacer.Click
        Limpiar()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        'VALIDAR FECHA
        For b = 0 To grdArticulo.RowCount - 1
            If IsDate(grdArticulo(7, b).Value) = False Then
                MessageBox.Show("Formato de fecha no reconocido en la fila " & b & " revise y vuelva a intentarlo.", "Planet Love", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Next

        If grdArticulo.RowCount = 0 Then
            Exit Sub
        End If
        clase.consultar("SELECT* FROM detmer_puertapuerta WHERE (detpuerta_codigopuerta =" & codigo_guardado & ")", "detmer")
        If clase.dt.Tables("detmer").Rows.Count > 0 Then
            Dim x As Short
            Dim indvalidad As Boolean = False
            For x = 0 To clase.dt.Tables("detmer").Rows.Count - 1
                If IsDBNull(clase.dt.Tables("detmer").Rows(x)("detpuerta_cantidad")) Or IsDBNull(clase.dt.Tables("detmer").Rows(x)("detpuerta_patron")) Then
                    MessageBox.Show("Debe establecer los patrones y/o cantidades para cada item agregado.", "DATOS INCOMPLETOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    grdArticulo.CurrentCell = grdArticulo.Item(3, x)
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

        'TOMAR CONSECUTIVO DE LA TABLA CONSECUTIVO_CODIGOBARRA_IMPORTACION
        clase.consultar("SELECT * FROM consecutivo_codigobarra_importacion", "consecutivocaja")
        Dim ConsecutivoCaja As String = clase.dt.Tables("consecutivocaja").Rows(0)("codigo_ultimo") + 1

        Dim a As Integer
        '1. ENTRADAMERCANCIA
        For a = 0 To grdArticulo.RowCount - 1
            clase.consultar("SELECT entradamercancia.* FROM entradamercancia WHERE (com_codigoart =" & grdArticulo.Item(0, a).Value & " AND com_codigoimp =" & ConsImport & ")", "entradamercancia")
            If clase.dt.Tables("entradamercancia").Rows.Count > 0 Then
                Dim cant As Integer = clase.dt.Tables("entradamercancia").Rows(0)("com_unidades")
                clase.actualizar("update entradamercancia set com_unidades = " & cant + grdArticulo(3, a).Value & " WHERE (com_codigoart =" & grdArticulo.Item(0, a).Value & " AND com_codigoimp =" & ConsImport & ")")
            Else
                clase.agregar_registro("INSERT INTO entradamercancia(com_codigoimp,com_codigoart,com_unidades,com_danado,com_operario) VALUES('" & ConsImport & "','" & grdArticulo(0, a).Value.ToString & "','" & grdArticulo(3, a).Value.ToString & "','0','" & txtOperario.Text & "')")
            End If

        Next

        '2. DETALLE_IMPORTACION_CABCAJAS
        clase.agregar_registro("INSERT INTO detalle_importacion_cabcajas(det_codigoimportacion,det_caja,det_codigoproveedor,det_fecharecepcion) VALUES('" & ConsImport & "','" & ConsecutivoCaja & "',NULL,'" & clase.FormatoFecha(Date.Today) & "')")

        '3. CABSALIDAS_MERCANCIA
        clase.agregar_registro("INSERT INTO cabsalidas_mercancia(cabsal_fecha,calsal_estado,cabsal_liquidador,cabsal_proveedor) VALUES('" & clase.FormatoFecha(Date.Today) & "','1','" & UCase(txtOperario.Text) & "',NULL)")
        clase.consultar("SELECT MAX(cabsal_cod) AS ConsCabSal FROM cabsalidas_mercancia", "conscabsal")
        Dim ConsCabSal As String = clase.dt.Tables("conscabsal").Rows(0)("ConsCabSal")


        '4. DETALLE_IMPORTACION_DETCAJAS
        For a = 0 To grdArticulo.RowCount - 1
            If grdArticulo(7, a).Value = "" Then
                clase.agregar_registro("INSERT INTO detalle_importacion_detcajas(detcab_codigocaja,detcab_referencia,detcab_descripcion,detcab_marca,detcab_composicion,detcab_cantidad,detcab_unimedida,detcab_registrodian,detcab_fecharegistrodian) VALUES('" & ConsecutivoCaja & "','" & grdArticulo(2, a).Value.ToString & "','" & grdArticulo(1, a).Value.ToString & "','SIN MARCA','NO REGISTRA','0','Unidad(s)',NULL,NULL)")
            Else
                Dim Fbase As Date = grdArticulo(7, a).Value
                Dim FechaF As String = Fbase.ToString("yyyy-MM-dd")
                clase.agregar_registro("INSERT INTO detalle_importacion_detcajas(detcab_codigocaja,detcab_referencia,detcab_descripcion,detcab_marca,detcab_composicion,detcab_cantidad,detcab_unimedida,detcab_registrodian,detcab_fecharegistrodian) VALUES('" & ConsecutivoCaja & "','" & grdArticulo(2, a).Value.ToString & "','" & grdArticulo(1, a).Value.ToString & "','SIN MARCA','NO REGISTRA','0','Unidad(s)','" & grdArticulo(6, a).Value.ToString & "','" & FechaF & "')")
                clase.agregar_registro("insert into registros_de_importaciones (articulo, registro, fecha) values ('" & grdArticulo(0, a).Value.ToString & "', '" & grdArticulo(6, a).Value.ToString & "', '" & FechaF & "')")

            End If

            '5. DETSALIDAS_MERCANCIA
            clase.consultar("SELECT MAX(detcab_coditem) AS ConsDetCaja FROM detalle_importacion_detcajas", "consdetimport")
            Dim ConsDetImport As String = clase.dt.Tables("consdetimport").Rows(0)("ConsDetCaja")
            clase.agregar_registro("INSERT INTO detsalidas_mercancia(det_salidacodigo,det_codref,det_cant,det_procesado) VALUES('" & ConsCabSal & "','" & ConsDetImport & "','" & grdArticulo(3, a).Value.ToString & "','1')")
            '----
            clase.consultar2("select* from asociaciones_codigos where asc_postcodart = " & grdArticulo(0, a).Value.ToString & "", "asociaciones")
            If clase.dt2.Tables("asociaciones").Rows.Count > 0 Then
                clase.actualizar("UPDATE asociaciones_codigos SET asc_precodref = " & ConsDetImport & " WHERE asc_postcodart = " & grdArticulo(0, a).Value.ToString & "")
            Else
                clase.agregar_registro("INSERT INTO asociaciones_codigos(asc_precodref,asc_postcodart) VALUES('" & ConsDetImport & "','" & grdArticulo(0, a).Value.ToString & "')")
            End If
        Next
        cargarconsecutivos()

        '6. ORDENPRODUCCION
        clase.agregar_registro("INSERT INTO ordenproduccion(op_codigo,op_fecha,op_hora,op_realizadapor,op_recibida,op_codigoimportacion,op_procesado,op_estado) VALUES('" & txtConsOrden.Text & "','" & clase.FormatoFecha(Date.Today) & "', '" & clase.FormatoHora(TimeOfDay) & "', '" & UCase(txtOperario.Text) & "', '0','" & ConsImport & "','0','A')")

        '7. DEORDENPROD
        For a = 0 To grdArticulo.RowCount - 1
            clase.agregar_registro("INSERT INTO deordenprod(do_idcaborden,do_codigo,do_unidades,do_patrondist,do_estado) VALUES('" & txtConsOrden.Text & "','" & grdArticulo(0, a).Value.ToString & "','" & grdArticulo(3, a).Value.ToString & "','" & grdArticulo(4, a).Value.ToString & "','A')")
        Next

        '8. DETALLE_PROVEEDORES ARTICULO
        For d = 0 To grdArticulo.RowCount - 1
            clase.consultar("SELECT * FROM detalle_proveedores_articulos WHERE codigo_articulo='" & grdArticulo(0, d).Value & "'", "prov")
            If clase.dt.Tables("prov").Rows.Count = 0 Then
                clase.agregar_registro("INSERT INTO detalle_proveedores_articulos(codigo_articulo,codigo_proveedor) VALUES('" & grdArticulo(0, d).Value & "','" & grdArticulo(8, d).Value & "')")
            ElseIf grdArticulo(8, d).Value <> clase.dt.Tables("prov").Rows(0)("codigo_proveedor") Then
                clase.agregar_registro("INSERT INTO detalle_proveedores_articulos(codigo_articulo,codigo_proveedor) VALUES('" & grdArticulo(0, d).Value & "','" & grdArticulo(8, d).Value & "')")
            End If
        Next

        'ACTUALIZAR TABLA CONSECUTIVO_CODIGOBARRA_IMPORTACION
        ConsecutivoCaja = Val(ConsecutivoCaja) + 1
        clase.actualizar("UPDATE consecutivo_codigobarra_importacion SET codigo_ultimo='" & ConsecutivoCaja & "'")
        clase.actualizar("UPDATE cabmer_puertapuerta set cabpuerta_procesada = TRUE WHERE cabpuerta_codigo = " & codigo_guardado & "")
        ind_guardado = False
        codigo_guardado = vbEmpty
        MigrarArticulosLineasProveedores()
        '  CrearMovimientoCompraLovePOS(ConsImport, txtOperario.Text.ToUpper)
        MessageBox.Show("Datos guardados satisfactoriamente.", "Operación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Limpiar()

    End Sub

    Private Sub btnConsulta_Click(sender As Object, e As EventArgs) Handles btnConsulta.Click
        frm_busqueda_articulos_puerta_puerta.ShowDialog()
        frm_busqueda_articulos_puerta_puerta.Dispose()
        txtArticulo.Focus()
        If txtArticulo.Text <> "" Then
            SendKeys.Send("{ENTER}")
        End If
    End Sub

    Private Sub txtConsMov_GotFocus(sender As Object, e As EventArgs) Handles txtConsMov.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtImportacion_GotFocus(sender As Object, e As EventArgs) Handles txtImportacion.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtProveedor_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(txtImportacion, "Importación") = False Then Exit Sub
        If clase.validar_cajas_text(txtOperario, "Operario") = False Then Exit Sub
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
        If grdArticulo.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(grdArticulo.CurrentCell.RowIndex) Then
            clase.borradoautomatico("delete from detmer_puertapuerta WHERE detpuerta_codigopuerta = " & codigo_guardado & " AND detpuerta_articulo = " & grdArticulo.Item(0, grdArticulo.CurrentRow.Index).Value & " ")
            grdArticulo.Rows.Remove(grdArticulo.Rows(grdArticulo.CurrentCell.RowIndex))

        End If
    End Sub

    Private Sub grdArticulo_KeyDown(sender As Object, e As KeyEventArgs) Handles grdArticulo.KeyDown
        If e.KeyCode = Keys.Return Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frm_mercancia_no_guardad_suspedidas_puerta_puerta.ShowDialog()
        frm_mercancia_no_guardad_suspedidas_puerta_puerta.Dispose()
    End Sub

    Sub llenar_grilla_mercancia(codigo As Short) 'esta funcion la utilizo para cargar los movimientos suspendidos, 
        clase.consultar("SELECT detmer_puertapuerta.detpuerta_articulo, articulos.ar_referencia, articulos.ar_descripcion, detmer_puertapuerta.detpuerta_cantidad, cabpatrondist.pat_codigo, cabpatrondist.pat_nombre,detmer_puertapuerta.detpuerta_regimp,detmer_puertapuerta.detpuerta_fechaimp,proveedores.prv_codigo,proveedores.prv_codigoasignado FROM detmer_puertapuerta INNER JOIN cabmer_puertapuerta ON (detmer_puertapuerta.detpuerta_codigopuerta = cabmer_puertapuerta.cabpuerta_codigo) INNER JOIN articulos ON (detmer_puertapuerta.detpuerta_articulo = articulos.ar_codigo) INNER JOIN proveedores ON (detmer_puertapuerta.detpuerta_proveedor = proveedores.prv_codigo) LEFT JOIN cabpatrondist ON (detmer_puertapuerta.detpuerta_patron = cabpatrondist.pat_codigo) WHERE (cabmer_puertapuerta.cabpuerta_codigo =" & codigo & ")", "mov")
        If clase.dt.Tables("mov").Rows.Count > 0 Then
            grdArticulo.Rows.Clear()
            Dim c As Short
            For c = 0 To clase.dt.Tables("mov").Rows.Count - 1
                grdArticulo.RowCount = grdArticulo.RowCount + 1
                grdArticulo.Item(0, c).Value = clase.dt.Tables("mov").Rows(c)("detpuerta_articulo")
                grdArticulo.Item(1, c).Value = clase.dt.Tables("mov").Rows(c)("ar_referencia")
                grdArticulo.Item(2, c).Value = clase.dt.Tables("mov").Rows(c)("ar_descripcion")

                If IsDBNull(clase.dt.Tables("mov").Rows(c)("detpuerta_cantidad")) Then
                    grdArticulo.Item(3, c).Value = ""
                Else
                    grdArticulo.Item(3, c).Value = clase.dt.Tables("mov").Rows(c)("detpuerta_cantidad")
                End If
                If IsDBNull(clase.dt.Tables("mov").Rows(c)("pat_codigo")) Then
                    grdArticulo.Item(4, c).Value = ""
                    grdArticulo.Item(5, c).Value = ""
                Else
                    grdArticulo.Item(4, c).Value = clase.dt.Tables("mov").Rows(c)("pat_codigo")
                    grdArticulo.Item(5, c).Value = clase.dt.Tables("mov").Rows(c)("pat_nombre")
                End If
                If IsDBNull(clase.dt.Tables("mov").Rows(c)("detpuerta_regimp")) Then
                    grdArticulo(6, c).Value = ""
                Else
                    grdArticulo(6, c).Value = clase.dt.Tables("mov").Rows(c)("detpuerta_regimp")
                End If
                If IsDBNull(clase.dt.Tables("mov").Rows(c)("detpuerta_fechaimp")) Then
                    grdArticulo(7, c).Value = ""
                Else
                    grdArticulo(7, c).Value = clase.dt.Tables("mov").Rows(c)("detpuerta_fechaimp")
                End If
                If IsDBNull(clase.dt.Tables("mov").Rows(c)("prv_codigo")) Then
                    grdArticulo.Item(8, c).Value = ""
                    grdArticulo.Item(9, c).Value = ""
                Else
                    grdArticulo.Item(8, c).Value = clase.dt.Tables("mov").Rows(c)("prv_codigo")
                    grdArticulo.Item(9, c).Value = clase.dt.Tables("mov").Rows(c)("prv_codigoasignado")
                End If
            Next
        Else
            grdArticulo.Rows.Clear()
        End If
    End Sub

    Private Sub txtConsOrden_GotFocus1(sender As Object, e As EventArgs) Handles txtConsOrden.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub btnProveedor_Click(sender As Object, e As EventArgs)
        frm_proveedores_mercancia_no_importada_puerta_puerta.ShowDialog()
        frm_proveedores_mercancia_no_importada_puerta_puerta.Dispose()
    End Sub

    Private Sub frm_ingreso_puerta_puerta_puerta_puerta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PrepararColumnas()
        cargarconsecutivos()
        indicador = False
        ind_guardado = False
        indicador_presionador_de_boton = False
    End Sub

    Private Sub txtArticulo_TextChanged(sender As Object, e As EventArgs) Handles txtArticulo.TextChanged

    End Sub
End Class