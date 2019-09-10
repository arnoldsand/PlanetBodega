Public Class frmIngresoNoImportados
    Dim clase As New class_library
    Dim Codigo As String
    Public i As Integer
    Public Patron, NombrePatron, ConsImport As String
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        End
    End Sub
    Public Sub cargarconsecutivos()
        clase.consultar("SELECT MAX(op_codigo) AS CODIGO FROM ordenproduccion", "consecutivo")
        txtConsOrden.Text = clase.dt.Tables("consecutivo").Rows(0)("CODIGO") + 1
        clase.consultar("SELECT MAX(cabsal_cod) AS ConsCabSal FROM cabsalidas_mercancia", "conscabsal")
        txtConsMov.Text = clase.dt.Tables("conscabsal").Rows(0)("ConsCabSal") + 1
    End Sub
    Private Sub frmIngresoNoImportados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PrepararColumnas()
        cargarconsecutivos()
    End Sub

    Public Sub PrepararColumnas()
        'PREPARAMOS LAS COLUMNAS PARA RECIBIR LAS INFORMACION DE LOS ARTICULOS
        With grdArticulo
            .Columns.Add("0", "CODIGO")
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).ReadOnly = True
            .Columns(0).Width = 50
            .Columns.Add("1", "REFERENCIA")
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).ReadOnly = True
            .Columns(1).Width = 200
            .Columns.Add("2", "DESCRIPCION")
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).Width = 200
            .Columns(2).DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .Columns(2).ReadOnly = True
            .Columns.Add("3", "CANT")
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).ReadOnly = False
            .Columns(3).Width = 50
            .Columns.Add("4", "CODPATRON")
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).ReadOnly = True
            .Columns(4).Visible = False
            .Columns.Add("5", "PATRON")
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).ReadOnly = True

            .RowHeadersWidth = 4
        End With
    End Sub

    Public Sub Agregar()
        clase.consultar("SELECT ar_codigo AS CODIGO, ar_referencia AS REFERENCIA, ar_descripcion AS DESCRIPCION, ar_precio1 AS PRECIO_1, ar_precio2 AS PRECIO_2 FROM articulos WHERE ar_codigobarras='" & txtArticulo.Text & "' OR ar_codigo='" & txtArticulo.Text & "'", "articulo")
        'VERIFICAR SI EL PRODUCTO EXISTE
        If clase.dt.Tables("articulo").Rows.Count = 0 Then
            MessageBox.Show("ARTICULO NO EXISTE", "PLANET LOVE")
            txtArticulo.Enabled = True
            txtArticulo.Text = ""
            txtArticulo.Focus()
            Exit Sub
        Else
            'VERIFICAR SI EL PRODUCTO YA ESTA, EN CASO DE QUE SI LO ESTE LO PODRA MODIFICAR
            Dim a As Integer
            For a = 0 To grdArticulo.RowCount - 1
                If txtArticulo.Text = grdArticulo(0, a).Value.ToString Or txtArticulo.Text = grdArticulo(1, a).Value.ToString Then
                    MessageBox.Show("ESTE PRODUCTO YA SE ENCUENTRA EN LA TABLA, PUEDE MODIFICAR LA CANTIDAD", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    grdArticulo.Rows(a).Cells(3).Selected = True
                    grdArticulo.CurrentCell = grdArticulo.Rows(a).Cells(3)
                    Exit Sub
                End If
            Next
            'AGREGAMOS UNA NUEVA FILA CON LOS DATOS DEL PRODUCTO DIGITADO
            i = grdArticulo.Rows.Add()
            grdArticulo(0, i).Value = clase.dt.Tables("articulo").Rows(0)("CODIGO")
            grdArticulo(1, i).Value = clase.dt.Tables("articulo").Rows(0)("REFERENCIA")
            grdArticulo(2, i).Value = clase.dt.Tables("articulo").Rows(0)("DESCRIPCION")
            grdArticulo.CurrentCell = grdArticulo.Rows(i).Cells(3)
            txtArticulo.Text = ""
            ' txtArticulo.Enabled = False
        End If
    End Sub

    Private Sub txtArticulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtArticulo.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub txtArticulo_LostFocus(sender As Object, e As EventArgs) Handles txtArticulo.LostFocus
        'VALIDAMOS SI LOS DATOS DEL ENCABEZADO ESTAN BIEN DIGITADOS
        If clase.validar_cajas_text(txtOperario, "NOMBRE DE OPERARIO") And clase.validar_cajas_text(txtImportacion, "IMPORTACION") And clase.validar_cajas_text(txtProveedor, "PROVEEDOR") Then
            btnDeshacer.Enabled = True
            btnGuardar.Enabled = True
            If txtArticulo.Text <> "" Then
                Agregar()
            End If
        End If
    End Sub

    Private Sub grdArticulo_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdArticulo.CellClick
        Dim Y As Integer = grdArticulo.CurrentCell.RowIndex
        Codigo = grdArticulo(0, [Y]).Value.ToString
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
        If columna = 3 Then
            'Obtener caracter
            Dim caracter As Char = e.KeyChar
            'comprobar si el caracter es un número o el retroceso  
            If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
                Me.Text = e.KeyChar
                e.KeyChar = Chr(0)
            End If
        End If
    End Sub

    Private Sub grdArticulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdArticulo.KeyPress
        If grdArticulo(3, i).Value = Nothing Then
            MessageBox.Show("DEBES LLENAR LA CANTIDAD DE ESTE PRODUCTO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If Asc(e.KeyChar) = 13 Then
            'ELEGIR EL PATRON DE DISTRIBUCION
            frm_patron_distribucion2.ShowDialog()
            frm_patron_distribucion2.Dispose()
            'VALIDAR SI LA CANTIDAD CUMPLE CON LAS REQUERIDAS POR EL PATRON
            clase.consultar("SELECT cabpatrondist.pat_codigo AS PATRON, cabpatrondist.pat_nombre AS NOMBRE, SUM(detpatrondist.dp_cantidad) AS SUMPATRON FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo)WHERE (cabpatrondist.pat_codigo ='" & Patron & "')GROUP BY cabpatrondist.pat_codigo;", "patron")
            If grdArticulo(3, i).Value >= clase.dt.Tables("patron").Rows(0)("SUMPATRON") Then
                grdArticulo(4, i).Value = Patron
                grdArticulo(5, i).Value = clase.dt.Tables("patron").Rows(0)("NOMBRE")
            Else
                MessageBox.Show("LA CANTIDAD NO CUMPLE CON LAS REQUERIDAS PARA SURTIR LAS TIENDAS", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                grdArticulo.CurrentCell = grdArticulo.Rows(i).Cells(3)
                grdArticulo(3, i).Value = ""
                Exit Sub
            End If

            'FIN ELECCION DE PATRON
            txtArticulo.Enabled = True
            txtArticulo.Focus()
        End If
    End Sub

    Private Sub btnImportacion_Click(sender As Object, e As EventArgs) Handles btnImportacion.Click
        frm_listado_importacion8.ShowDialog()
    End Sub

    Private Sub btnProveedor_Click(sender As Object, e As EventArgs) Handles btnProveedor.Click
        frm_proveedores8.ShowDialog()
        frm_proveedores8.Dispose()
    End Sub

    Private Sub txtOperario_TextChanged(sender As Object, e As EventArgs)
        txtOperario.CharacterCasing = CharacterCasing.Upper
    End Sub
    Public Sub Limpiar()
        clase.Limpiar_Cajas(Me)
        grdArticulo.Rows.Clear()
        btnDeshacer.Enabled = False
        btnGuardar.Enabled = False
    End Sub
    Private Sub btnDeshacer_Click(sender As Object, e As EventArgs) Handles btnDeshacer.Click
        Limpiar()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        'TOMAR CONSECUTIVO DE LA TABLA CONSECUTIVO_CODIGOBARRA_IMPORTACION
        clase.consultar("SELECT * FROM consecutivo_codigobarra_importacion", "consecutivocaja")
        Dim ConsecutivoCaja As String = clase.dt.Tables("consecutivocaja").Rows(0)("codigo_ultimo")

        Dim a As Integer
        '1. ENTRADAMERCANCIA
        For a = 0 To grdArticulo.RowCount - 1
            clase.agregar_registro("INSERT INTO entradamercancia(com_codigoimp,com_codigoart,com_unidades,com_danado,com_operario) VALUES('" & ConsImport & "','" & grdArticulo(0, a).Value.ToString & "','" & grdArticulo(3, a).Value.ToString & "','0','" & txtOperario.Text & "')")
        Next

        '2. DETALLE_IMPORTACION_CABCAJAS
        clase.agregar_registro("INSERT INTO detalle_importacion_cabcajas(det_codigoimportacion,det_caja,det_codigoproveedor,det_fecharecepcion) VALUES('" & ConsImport & "','" & ConsecutivoCaja & "','" & txtProveedor.Text & "','" & clase.FormatoFecha(Date.Today) & "')")

        '3. CABSALIDAS_MERCANCIA
        clase.agregar_registro("INSERT INTO cabsalidas_mercancia(cabsal_fecha,calsal_estado,cabsal_liquidador,cabsal_proveedor) VALUES('" & clase.FormatoFecha(Date.Today) & "','1','" & txtOperario.Text & "','" & txtProveedor.Text & "')")
        clase.consultar("SELECT MAX(cabsal_cod) AS ConsCabSal FROM cabsalidas_mercancia", "conscabsal")
        Dim ConsCabSal As String = clase.dt.Tables("conscabsal").Rows(0)("ConsCabSal")

        '4. DETALLE_IMPORTACION_DETCAJAS
        For a = 0 To grdArticulo.RowCount - 1
            clase.agregar_registro("INSERT INTO detalle_importacion_detcajas(detcab_codigocaja,detcab_referencia,detcab_descripcion,detcab_cantidad) VALUES('" & ConsecutivoCaja & "','" & grdArticulo(2, a).Value.ToString & "','" & grdArticulo(1, a).Value.ToString & "','" & grdArticulo(3, a).Value.ToString & "')")

            '5. DETSALIDAS_MERCANCIA
            clase.consultar("SELECT MAX(detcab_coditem) AS ConsDetCaja FROM detalle_importacion_detcajas", "consdetimport")
            Dim ConsDetImport As String = clase.dt.Tables("consdetimport").Rows(0)("ConsDetCaja")
            clase.agregar_registro("INSERT INTO detsalidas_mercancia(det_salidacodigo,det_codref,det_cant,det_procesado) VALUES('" & ConsCabSal & "','" & ConsDetImport & "','" & grdArticulo(3, a).Value.ToString & "','1')")

            '6. ASOCIACIONES_CODIGO
            clase.agregar_registro("INSERT INTO asociaciones_codigos(asc_precodref,asc_postcodart) VALUES('" & ConsDetImport & "','" & grdArticulo(0, a).Value.ToString & "')")
        Next

        '7. ORDENPRODUCCION
        clase.agregar_registro("INSERT INTO ordenproduccion(op_codigo,op_fecha,op_hora,op_realizadapor,op_recibida,op_codigoimportacion,op_procesado,op_estado) VALUES('" & txtConsOrden.Text & "','" & clase.FormatoFecha(Date.Today) & "', '" & clase.FormatoHora(TimeOfDay) & "', '" & txtOperario.Text & "', '0','" & ConsImport & "','0','A')")

        '8. DEORDENPROD
        For a = 0 To grdArticulo.RowCount - 1
            clase.agregar_registro("INSERT INTO deordenprod(do_idcaborden,do_codigo,do_unidades,do_patrondist,do_estado) VALUES('" & txtConsOrden.Text & "','" & grdArticulo(0, a).Value.ToString & "','" & grdArticulo(3, a).Value.ToString & "','" & grdArticulo(4, a).Value.ToString & "','A')")
        Next

        'ACTUALIZAR TABLA CONSECUTIVO_CODIGOBARRA_IMPORTACION
        ConsecutivoCaja = Val(ConsecutivoCaja) + 1
        clase.actualizar("UPDATE consecutivo_codigobarra_importacion SET codigo_ultimo='" & ConsecutivoCaja & "'")

        MessageBox.Show("DATOS GUARDADOS CON EXITO, GRACIAS", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Limpiar()
        cargarconsecutivos()
    End Sub

    Private Sub btnConsulta_Click(sender As Object, e As EventArgs) Handles btnConsulta.Click
        frmBusqueda.ShowDialog()
        txtArticulo.Focus()
        If txtArticulo.Text <> "" Then
            SendKeys.Send("{ENTER}")
        End If
    End Sub

    Private Sub txtConsOrden_GotFocus(sender As Object, e As EventArgs) Handles txtConsOrden.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtConsMov_GotFocus(sender As Object, e As EventArgs) Handles txtConsMov.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtImportacion_GotFocus(sender As Object, e As EventArgs) Handles txtImportacion.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtProveedor_GotFocus(sender As Object, e As EventArgs) Handles txtProveedor.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    

   
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(txtImportacion, "Importación") = False Then Exit Sub
        If clase.validar_cajas_text(txtProveedor, "Proveedor") = False Then Exit Sub
        If clase.validar_cajas_text(txtOperario, "Operario") = False Then Exit Sub
        txtImportacion.Enabled = False
        txtProveedor.Enabled = False
        txtOperario.Enabled = False
        btnImportacion.Enabled = False
        btnProveedor.Enabled = False
        txtArticulo.Enabled = True
        btnConsulta.Enabled = True
        txtArticulo.Focus()
        Button1.Enabled = False
    End Sub

    Private Sub txtArticulo_TextChanged(sender As Object, e As EventArgs) Handles txtArticulo.TextChanged

    End Sub
End Class
