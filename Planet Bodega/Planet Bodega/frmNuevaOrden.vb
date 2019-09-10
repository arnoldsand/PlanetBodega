Public Class frmNuevaOrden
    Dim clase As New class_library
    Public i As Integer
    Public Orden, Patron, NombrePatron As String
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        txtArticulo.Enabled = True
        btnActualizar.Visible = False
        frmOrdenProduccion.Ver = False
        frmOrdenProduccion.Enabled = True
        frmOrdenProduccion.Show()
        Me.Close()
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
                    grdArticulo.Rows(a).Cells(5).Selected = True
                    grdArticulo.CurrentCell = grdArticulo.Rows(a).Cells(5)
                    Exit Sub
                End If
            Next
            'AGREGAMOS UNA NUEVA FILA CON LOS DATOS DEL PRODUCTO DIGITADO
            i = grdArticulo.Rows.Add()
            grdArticulo(0, i).Value = clase.dt.Tables("articulo").Rows(0)("CODIGO")
            grdArticulo(1, i).Value = clase.dt.Tables("articulo").Rows(0)("REFERENCIA")
            grdArticulo(2, i).Value = clase.dt.Tables("articulo").Rows(0)("DESCRIPCION")
            grdArticulo(3, i).Value = clase.dt.Tables("articulo").Rows(0)("PRECIO_1")
            grdArticulo(4, i).Value = clase.dt.Tables("articulo").Rows(0)("PRECIO_2")
            grdArticulo.CurrentCell = grdArticulo.Rows(i).Cells(5)

            txtArticulo.Text = ""
            txtArticulo.Enabled = False
        End If
    End Sub
    
    Public Sub PrepararColumnas()
        'PREPARAMOS LAS COLUMNAS PARA RECIBIR LAS INFORMACION DE LOS ARTICULOS
        grdArticulo.Columns.Add("0", "CODIGO")
        grdArticulo.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(0).ReadOnly = True
        grdArticulo.Columns.Add("1", "REFERENCIA")
        grdArticulo.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(1).ReadOnly = True
        grdArticulo.Columns(1).Width = 200
        grdArticulo.Columns.Add("2", "DESCRIPCION")
        grdArticulo.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(2).Width = 200
        grdArticulo.Columns(2).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        grdArticulo.Columns(2).ReadOnly = True
        grdArticulo.Columns.Add("3", "PRECIO 1")
        grdArticulo.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(3).ReadOnly = True
        grdArticulo.Columns.Add("4", "PRECIO 2")
        grdArticulo.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(4).ReadOnly = True
        grdArticulo.Columns.Add("5", "CANTIDAD")
        grdArticulo.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(5).ReadOnly = False
        grdArticulo.Columns.Add("6", "CODPATRON")
        grdArticulo.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(6).ReadOnly = True
        grdArticulo.Columns(6).Visible = False
        grdArticulo.Columns.Add("7", "PATRON")
        grdArticulo.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdArticulo.Columns(7).ReadOnly = True

        grdArticulo.RowHeadersWidth = 4
    End Sub


    Private Sub frmNuevaOrden_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If frmOrdenProduccion.Ver = False Then
            'SELECCIONAMOS EL CONSECUTIVO QUE SIGUE EN LA TABLA ORDENPRODUCCION
            Try
                clase.consultar("SELECT MAX(op_codigo) AS codigo FROM ordenproduccion", "codigo")
                txtOrden.Text = clase.dt.Tables("codigo").Rows(0)("codigo") + 1
            Catch ex As Exception
                txtOrden.Text = 1
            End Try
            PrepararColumnas()
        End If

    End Sub

    Private Sub txtArticulo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtArticulo.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub txtArticulo_LostFocus(sender As Object, e As EventArgs) Handles txtArticulo.LostFocus
        'VALIDAMOS SI LOS DATOS DEL ENCABEZADO ESTAN BIEN DIGITADOS
        If clase.validar_cajas_text(txtOperario, "NOMBRE DE OPERARIO") And clase.validar_cajas_text(txtImportacion, "IMPORTACION") Then
            btnDeshacer.Enabled = True
            btnEliminar.Enabled = True
            btnGuardar.Enabled = True
            If txtArticulo.Text <> "" Then
                Agregar()
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
        If columna = 5 Then

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
        If grdArticulo(5, i).Value = Nothing Then
            MessageBox.Show("DEBES LLENAR LA CANTIDAD DE ESTE PRODUCTO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If Asc(e.KeyChar) = 13 Then
            'ELEGIR EL PATRON DE DISTRIBUCION
            frm_patron_distribucion1.ShowDialog()
            frm_patron_distribucion1.Dispose()
            'VALIDAR SI LA CANTIDAD CUMPLE CON LAS REQUERIDAS POR EL PATRON
            clase.consultar("SELECT cabpatrondist.pat_codigo AS PATRON, cabpatrondist.pat_nombre AS NOMBRE, SUM(detpatrondist.dp_cantidad) AS SUMPATRON FROM detpatrondist INNER JOIN cabpatrondist ON (detpatrondist.dp_idpatron = cabpatrondist.pat_codigo)WHERE (cabpatrondist.pat_codigo ='" & Patron & "')GROUP BY cabpatrondist.pat_codigo;", "patron")
            If grdArticulo(5, i).Value >= clase.dt.Tables("patron").Rows(0)("SUMPATRON") Then
                grdArticulo(6, i).Value = Patron
                grdArticulo(7, i).Value = clase.dt.Tables("patron").Rows(0)("NOMBRE")
            Else
                MessageBox.Show("LA CANTIDAD NO CUMPLE CON LAS REQUERIDAS PARA SURTIR LAS TIENDAS", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                grdArticulo.CurrentCell = grdArticulo.Rows(i).Cells(5)
                grdArticulo(5, i).Value = ""
                Exit Sub
            End If

            'FIN ELECCION DE PATRON
            txtArticulo.Enabled = True
            txtArticulo.Focus()
        End If
    End Sub

    Private Sub btnImportacion_Click(sender As Object, e As EventArgs) Handles btnImportacion.Click
        frm_seleccionarimportacion1.ShowDialog()
    End Sub
    Public Sub Limpiar()
        clase.Limpiar_Cajas(Me)
        grdArticulo.Rows.Clear()
        btnEliminar.Enabled = False
        btnDeshacer.Enabled = False
        btnGuardar.Enabled = False
    End Sub
    Private Sub btnDeshacer_Click(sender As Object, e As EventArgs) Handles btnDeshacer.Click
        Limpiar()
    End Sub

    Private Sub grdArticulo_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdArticulo.CellContentClick
        Dim Y As Integer = grdArticulo.CurrentCell.RowIndex
        Orden = grdArticulo(0, [Y]).Value.ToString
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If Orden = "" Then
            MessageBox.Show("DEBES SELECCIONAR UN ITEM", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        grdArticulo.Rows.RemoveAt(grdArticulo.CurrentCell.RowIndex)
    End Sub
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Dim a, b As Integer
        For a = 0 To grdArticulo.RowCount - 1
            If grdArticulo(5, a).Value = Nothing Then
                MessageBox.Show("LA FILA " & a + 1 & " NO TIENE CANTIDAD", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                grdArticulo.CurrentCell = grdArticulo.Rows(a).Cells(5)
                Exit Sub
            End If
        Next
        If MessageBox.Show("DESEA GUARDAR LA ORDEN DE PRODUCCION", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
            'GUARDAR ENCABEZADO DE ORDEN PRODUCCION
            clase.agregar_registro("INSERT INTO ordenproduccion(op_codigo,op_fecha,op_hora,op_realizadapor,op_recibida,op_codigoimportacion,op_procesado,op_estado) VALUES('" & txtOrden.Text & "','" & clase.FormatoFecha(Date.Today) & "', '" & clase.FormatoHora(TimeOfDay) & "', '" & txtOperario.Text & "', '0','" & frm_seleccionarimportacion1.CodImport & "','0','A')")

            'GUARDAR DETALLE ORDEN DE PRODUCCION
            For b = 0 To grdArticulo.RowCount - 1
                clase.agregar_registro("INSERT INTO deordenprod(do_idcaborden,do_codigo,do_unidades,do_patrondist,do_estado) VALUES('" & txtOrden.Text & "','" & grdArticulo(0, b).Value.ToString & "','" & grdArticulo(5, b).Value.ToString & "','" & grdArticulo(6, b).Value.ToString & "','A')")
            Next

            'FIN GUARDADO
            MessageBox.Show("LA ORDEN SE GUARDO CON EXITO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            frmOrdenProduccion.LlenarGrid()
            frmOrdenProduccion.Enabled = True
            frmOrdenProduccion.Show()

            Me.Close()
        End If
    End Sub

    Private Sub btnActualizar_Click(sender As Object, e As EventArgs) Handles btnActualizar.Click
        If MessageBox.Show("DESEA ACTUALIZAR ESTA INFORMACION", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
            'VALIDAMOS QUE EL GRIDVIEW NO TENGA CELDAS EN BLANCO
            Dim a, b As Integer
            Dim Importacion
            For a = 0 To grdArticulo.RowCount - 1
                If grdArticulo(5, a).Value = Nothing Then
                    MessageBox.Show("LA FILA " & a + 1 & " NO TIENE CANTIDAD", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    grdArticulo.CurrentCell = grdArticulo.Rows(a).Cells(5)
                    Exit Sub
                End If
            Next
            If frm_seleccionarimportacion1.CodImport = "" Then
                Importacion = frmOrdenProduccion.CodImp
            Else
                Importacion = frm_seleccionarimportacion1.CodImport
            End If
            'ACTUALIZAMOS EL ENCABEZADO
            clase.actualizar("UPDATE ordenproduccion SET op_realizadapor='" & txtOperario.Text & "', op_codigoimportacion='" & Importacion & "' WHERE op_codigo='" & txtOrden.Text & "'")

            'ACTUALIZAMOS EL DETALLE
            For b = 0 To grdArticulo.RowCount - 1
                clase.actualizar("UPDATE deordenprod SET do_unidades='" & grdArticulo(5, b).Value.ToString & "', do_patrondist='" & grdArticulo(6, b).Value.ToString & "' WHERE do_idcaborden='" & txtOrden.Text & "' AND do_codigo='" & grdArticulo(0, b).Value.ToString & "'")
            Next

            MessageBox.Show("ORDEN ACTUALIZADA CON EXITO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            frmOrdenProduccion.LlenarGrid()
            frmOrdenProduccion.Show()
            Me.Close()
        End If
    End Sub

    Private Sub txtOperario_TextChanged(sender As Object, e As EventArgs) Handles txtOperario.TextChanged
        txtOperario.CharacterCasing = CharacterCasing.Upper
    End Sub

    
End Class