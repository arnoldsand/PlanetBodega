Public Class frm_patron_despacho
    Dim clase As New class_library
    Dim articulo_codigoart() As Integer
    Dim articulo_bodega() As Short
    Dim articulo_gondola() As String
    Dim articulo_cantidad() As Integer
    Dim cantidad_gondolas_donde_existe_articulo As Integer
    Dim codpedido As Integer
    Dim codtienda As Integer
    Private Sub frm_patron_despacho_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        codpedido = frm_mantenimiento_pedidos.DataGridView1.Item(0, frm_mantenimiento_pedidos.DataGridView1.CurrentRow.Index).Value
        clase.consultar("SELECT cabped_tienda FROM cabpedido WHERE (cabped_codigo =" & codpedido & ")", "codtienda")
        codtienda = clase.dt.Tables("codtienda").Rows(0)("cabped_tienda")
        TextBox18.Text = codpedido
        TextBox1.Text = nombretienda(codtienda)
        clase.consultar1("SELECT desp_pedido FROM patron_despacho WHERE (desp_pedido =" & codpedido & ")", "pedido1")
        If clase.dt1.Tables("pedido1").Rows.Count > 0 Then
            DataGridView1.ColumnCount = 13
            preparar_columnas()
            llenar_listado1()
            Button3.Enabled = True
            '   Button2.Enabled = True
            Button4.Enabled = False
        Else
            DataGridView1.ColumnCount = 13
            preparar_columnas()
            llenar_listado()
            Button3.Enabled = False
            '   Button2.Enabled = False
            Button4.Enabled = True
        End If
    End Sub

    Function buscar_todos_los_codigos_asociados(codigoarticulo As Integer) As Integer
        clase.consultar("SELECT ar_referencia, ar_descripcion, ar_linea, ar_sublinea1, ar_sublinea2, ar_sublinea3, ar_sublinea4 FROM articulos WHERE (ar_codigo =" & codigoarticulo & ")", "datoscodigo")
        With clase.dt.Tables("datoscodigo")
            clase.consultar1(buscar_codigos_a_apartir_de_la_referencia(.Rows(0)("ar_referencia"), .Rows(0)("ar_descripcion"), .Rows(0)("ar_linea"), .Rows(0)("ar_sublinea1"), .Rows(0)("ar_sublinea2"), .Rows(0)("ar_sublinea3"), .Rows(0)("ar_sublinea4")), "buscodigos")
        End With
        Dim d As Integer
        Dim sql2 As String = "SELECT SUM(desp_cant) AS cant FROM patron_despacho WHERE ((desp_pedido =" & codpedido & ") AND ("
        Dim indconsulta As Boolean = False
        For d = 0 To clase.dt1.Tables("buscodigos").Rows.Count - 1
            If indconsulta = False Then
                sql2 = sql2 & "(desp_articulo =" & clase.dt1.Tables("buscodigos").Rows(d)("ar_codigo") & ")"
                indconsulta = True
            Else
                sql2 = sql2 & " OR (desp_articulo =" & clase.dt1.Tables("buscodigos").Rows(d)("ar_codigo") & ")"
            End If
        Next
        sql2 = sql2 & "))"
        clase.consultar(sql2, "cantcodigos")
        If IsDBNull(clase.dt.Tables("cantcodigos").Rows(0)("cant")) Then
            Return 0
        Else
            Return clase.dt.Tables("cantcodigos").Rows(0)("cant")
        End If
    End Function

    Function nombretienda(codigo As Short) As String
        clase.consultar("select* from tiendas where id = " & codigo & "", "tabla2")
        Return clase.dt.Tables("tabla2").Rows(0)("tienda")
    End Function

    Sub llenar_listado()
        Dim cant As Integer
        clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, detped_articulo, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4, detpedido.detped_cant FROM articulos INNER JOIN detpedido ON (articulos.ar_codigo = detpedido.detped_articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (detpedido.detped_codpedido =" & codpedido & ")", "detpedido")
        If clase.dt.Tables("detpedido").Rows.Count > 0 Then
            Dim x As Integer
            For x = 0 To clase.dt.Tables("detpedido").Rows.Count - 1
                Application.DoEvents()
                With DataGridView1
                    .RowCount = .RowCount + 1
                    .Item(0, x).Value = clase.dt.Tables("detpedido").Rows(x)("ar_referencia")
                    .Item(0, x).ReadOnly = True
                    .Item(1, x).Value = clase.dt.Tables("detpedido").Rows(x)("ar_descripcion")
                    .Item(1, x).ReadOnly = True
                    .Item(2, x).Value = clase.dt.Tables("detpedido").Rows(x)("ln1_nombre")
                    .Item(2, x).ReadOnly = True
                    .Item(3, x).Value = clase.dt.Tables("detpedido").Rows(x)("sl1_nombre")
                    .Item(3, x).ReadOnly = True
                    .Item(4, x).Value = clase.dt.Tables("detpedido").Rows(x)("ar_linea")
                    .Item(4, x).ReadOnly = True
                    .Item(5, x).Value = clase.dt.Tables("detpedido").Rows(x)("ar_sublinea1")
                    .Item(5, x).ReadOnly = True
                    .Item(6, x).Value = clase.dt.Tables("detpedido").Rows(x)("ar_sublinea2")
                    .Item(6, x).ReadOnly = True
                    .Item(7, x).Value = clase.dt.Tables("detpedido").Rows(x)("ar_sublinea3")
                    .Item(7, x).ReadOnly = True
                    .Item(8, x).Value = clase.dt.Tables("detpedido").Rows(x)("ar_sublinea4")
                    .Item(8, x).ReadOnly = True
                    .Item(9, x).Value = clase.dt.Tables("detpedido").Rows(x)("detped_cant")
                    .Item(9, x).ReadOnly = True
                    cant = calcular_existencia(x)
                    .Item(10, x).Value = cant
                    .Item(10, x).ReadOnly = True
                    .Item(11, x).Value = cant - calcular_despacho_virtual(x)
                    .Item(11, x).ReadOnly = True
                    .Item(12, x).Value = clase.dt.Tables("detpedido").Rows(x)("detped_articulo")

                End With
            Next
        End If
    End Sub

    Sub llenar_listado1()
        Dim cant As Integer
        Dim cant1 As Integer
        clase.consultar_global("SELECT detpedido.detped_articulo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4, detpedido.detped_cant FROM articulos INNER JOIN detpedido ON (articulos.ar_codigo = detpedido.detped_articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (detpedido.detped_codpedido =" & codpedido & ") ", "detpedido")
        If clase.dt_global.Tables("detpedido").Rows.Count > 0 Then
            Dim x As Integer
            For x = 0 To clase.dt_global.Tables("detpedido").Rows.Count - 1
                Application.DoEvents()
                With DataGridView1
                    .RowCount = .RowCount + 1
                    .Item(0, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ar_referencia")
                    .Item(0, x).ReadOnly = True
                    .Item(1, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ar_descripcion")
                    .Item(1, x).ReadOnly = True
                    .Item(2, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ln1_nombre")
                    .Item(2, x).ReadOnly = True
                    .Item(3, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("sl1_nombre")
                    .Item(3, x).ReadOnly = True
                    .Item(4, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ar_linea")
                    .Item(4, x).ReadOnly = True
                    .Item(5, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ar_sublinea1")
                    .Item(5, x).ReadOnly = True
                    .Item(6, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ar_sublinea2")
                    .Item(6, x).ReadOnly = True
                    .Item(7, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ar_sublinea3")
                    .Item(7, x).ReadOnly = True
                    .Item(8, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("ar_sublinea4")
                    .Item(8, x).ReadOnly = True
                    .Item(9, x).Value = clase.dt_global.Tables("detpedido").Rows(x)("detped_cant")
                    .Item(9, x).ReadOnly = True
                    cant = calcular_existencia(x)
                    .Item(10, x).Value = cant
                    .Item(10, x).ReadOnly = True
                    'cant1 = buscar_todos_los_codigos_asociados(clase.dt_global.Tables("detpedido").Rows(x)("detped_articulo"))
                    .Item(11, x).Value = cant - calcular_despacho_virtual(x) + cant1
                    .Item(11, x).ReadOnly = True
                    '.Item(12, x).Value = cant1
                End With
            Next
        End If
    End Sub

    Function calcular_existencia(d As Integer) As Integer
        With DataGridView1
            clase.consultar1(buscar_codigos_a_apartir_de_la_referencia(.Item(0, d).Value, .Item(1, d).Value, .Item(4, d).Value, .Item(5, d).Value, .Item(6, d).Value, .Item(7, d).Value, .Item(8, d).Value), "tbl")
        End With
        Dim z As Integer
        Dim ind As Boolean = False
        Dim Sql As String = "SELECT SUM(inv_cantidad) AS cant FROM inventario_bodega WHERE (inv_cantidad >0) AND ("
        For z = 0 To clase.dt1.Tables("tbl").Rows.Count - 1
            If ind = False Then
                Sql = Sql & " (inv_codigoart = " & clase.dt1.Tables("tbl").Rows(z)("ar_codigo") & ")"
                ind = True
            Else
                Sql = Sql & "  OR (inv_codigoart = " & clase.dt1.Tables("tbl").Rows(z)("ar_codigo") & ")"
            End If
        Next
        Sql = Sql & ")"
        clase.consultar1(Sql, "total")
        If IsDBNull(clase.dt1.Tables("total").Rows(0)("cant")) Then
            Return 0
        Else
            Return clase.dt1.Tables("total").Rows(0)("cant")
        End If
    End Function

    Function calcular_despacho_virtual(d As Integer) As Integer
        With DataGridView1
            clase.consultar1(buscar_codigos_a_apartir_de_la_referencia(.Item(0, d).Value, .Item(1, d).Value, .Item(4, d).Value, .Item(5, d).Value, .Item(6, d).Value, .Item(7, d).Value, .Item(8, d).Value), "tbl")
        End With
        Dim z As Integer
        Dim ind As Boolean = False
        Dim Sql As String = "SELECT SUM(patron_despacho.desp_cant) AS cant FROM cabpedido INNER JOIN patron_despacho ON (cabpedido.cabped_codigo = patron_despacho.desp_pedido) WHERE (cabpedido.cabped_fecha_est_llegada IS NULL) AND ("
        For z = 0 To clase.dt1.Tables("tbl").Rows.Count - 1
            If ind = False Then
                Sql = Sql & " (patron_despacho.desp_articulo = " & clase.dt1.Tables("tbl").Rows(z)("ar_codigo") & ")"
                ind = True
            Else
                Sql = Sql & " OR (patron_despacho.desp_articulo = " & clase.dt1.Tables("tbl").Rows(z)("ar_codigo") & ")"
            End If
        Next
        Sql = Sql & ")"
        clase.consultar1(Sql, "total")
        If IsDBNull(clase.dt1.Tables("total").Rows(0)("cant")) Then
            Return 0
        Else
            Return clase.dt1.Tables("total").Rows(0)("cant")
        End If
    End Function

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 140
            .Columns(1).Width = 180
            .Columns(2).Width = 130
            .Columns(3).Width = 130
            .Columns(4).Visible = False
            .Columns(5).Visible = False
            .Columns(6).Visible = False
            .Columns(7).Visible = False
            .Columns(8).Visible = False
            .Columns(12).Visible = False
            .Columns(9).Width = 60
            .Columns(10).Width = 70
            .Columns(11).Width = 70
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Linea"
            .Columns(3).HeaderText = "Sublinea"
            .Columns(9).HeaderText = "Cant Sugerida"
            .Columns(10).HeaderText = "Existencia"
            .Columns(11).HeaderText = "Existencia Virtual"

            ' .Columns(12).HeaderText = "Cant"
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '  .Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            Dim d As Integer = .CurrentCell.RowIndex
            clase.consultar1(buscar_codigos_a_apartir_de_la_referencia(.Item(0, d).Value, .Item(1, d).Value, .Item(4, d).Value, .Item(5, d).Value, .Item(6, d).Value, .Item(7, d).Value, .Item(8, d).Value), "tbl")
        End With
        clase.consultar("select ar_foto from articulos where ar_codigo = " & clase.dt1.Tables("tbl").Rows(0)("ar_codigo") & "", "tablita")
        If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
            PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
        Else
            PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
        End If
        SetImage(PictureBox1)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        ' obtener indice de la columna   
        Dim columna As Integer = DataGridView1.CurrentCell.ColumnIndex
        ' comprobar si la celda en edición corresponde a la columna 12
        If (columna = 12) Then
            ' referencia a la celda   
            Dim validar As TextBox = CType(e.Control, TextBox)
            ' agregar el controlador de eventos para el KeyPress   
            AddHandler validar.KeyPress, AddressOf validar_Keypress
        End If
    End Sub

    Private Sub validar_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        ' Obtener caracter   
        Dim caracter As Char = e.KeyChar
        ' comprobar si el caracter es un número o el retroceso   
        If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
            'Me.Text = e.KeyChar   
            e.KeyChar = Chr(0)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '  procedimiento_para_grabar_patrones()
        ' procedimiento_para_grabar_patrones1()  'este procedimiento es improvisado por el descueadre en la bodega, cuando se haga el inventario se debe usar el procedimiento de arriba
        procedimiento_improvisado()  ' segundo proceso improvisado implementado
    End Sub


    Sub procedimiento_improvisado()
        Dim x As Integer
        Dim bod As Short
        For x = 0 To DataGridView1.RowCount - 1
            bod = buscar_bodega_articulo(DataGridView1.Item(12, x).Value)
            If bod <> 0 Then
                clase.agregar_registro("INSERT INTO `patron_despacho`(`desp_pedido`,`desp_bodega`,`desp_gondola`,`desp_articulo`,`desp_cant`) VALUES ( '" & codpedido & "','" & bod & "','" & buscar_gondola_articulo(DataGridView1.Item(12, x).Value) & "','" & DataGridView1.Item(12, x).Value & "','" & DataGridView1.Item(9, x).Value & "')")
            End If
        Next
        MessageBox.Show("Los patrones de despachos se han guardado satisfactoriamente.", "PATRONES ESTABLECIDOS", MessageBoxButtons.OK, MessageBoxIcon.Information)

        'inhabilito la pregunta e impresion de patrones, para imprimirlo con el boton imprimir del formulario mantenimiento de pedidos
        '______
        'Dim v As String
        'v = MessageBox.Show("¿Desea imprimir el patrón creado?", "IMPRIMIR PATRON", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        'If v = 6 Then
        '    imprimir_patron_pedido()
        'End If
        Me.Close()
    End Sub

    Function buscar_gondola_articulo(articulo As Integer) As String
        clase.consultar("select* from articulos_gondolas where articulo = " & articulo & "", "busqueda")
        If clase.dt.Tables("busqueda").Rows.Count > 0 Then
            Return clase.dt.Tables("busqueda").Rows(0)("gondola")
        Else
            Return ""
        End If
    End Function

    Function buscar_bodega_articulo(articulo As Integer) As Short
        clase.consultar("select* from articulos_gondolas where articulo = " & articulo & "", "busqueda")
        If clase.dt.Tables("busqueda").Rows.Count > 0 Then
            Return clase.dt.Tables("busqueda").Rows(0)("bodega")
        Else
            Return 0
        End If
    End Function

    Sub procedimiento_para_grabar_patrones1()  'quitar despues del inventario
        Dim a As Integer
        Dim existencia_referencia As Integer
        Dim existencia_referencia_virtual As Integer
        Dim cantidad_requerida As Integer
        Dim articulos_ya_despachados As Integer
        ' Dim saldo As Integer
        For x = 0 To DataGridView1.RowCount - 1
            existencia_referencia = DataGridView1.Item(10, x).Value
            existencia_referencia_virtual = DataGridView1.Item(11, x).Value
            cantidad_requerida = asignar_cantidad_requerida1(DataGridView1.Item(11, x).Value)
            If ((cantidad_requerida > 0)) Then 'verifico que la cantidad requerida sea superior a 0
                buscar_existencia_bodega_gondola_cantidad1(x)
                articulos_ya_despachados = 0
                For a = 0 To cantidad_gondolas_donde_existe_articulo - 1
                    Application.DoEvents()
                    ' MsgBox(DataGridView1.Item(0, x).Value & " -  " & articulo_bodega(a) & " - " & articulo_gondola(a))
                    '  saldo = saldo_gondola(articulo_bodega(a), articulo_gondola(a), articulo_codigoart(a))

                    'MsgBox(saldo)
                    ' MsgBox(DataGridView1.Item(0, x).Value & " / " & articulo_gondola(a) & " ----" & articulo_cantidad(a) - saldo & "---- " & articulo_cantidad(a) & "----" & saldo)
                    '  If (articulo_cantidad(a) - saldo) <= (cantidad_requerida - articulos_ya_despachados) Then 'verifico si la existencia del codigo actual es suficiente para cubrir lo requerido o debo completar con la existencia del codigo siguiente
                    'If (articulo_cantidad(a) - saldo) > 0 Then
                    clase.agregar_registro("INSERT INTO `patron_despacho`(`desp_pedido`,`desp_bodega`,`desp_gondola`,`desp_articulo`,`desp_cant`) VALUES ( '" & codpedido & "','" & articulo_bodega(a) & "','" & articulo_gondola(a) & "','" & articulo_codigoart(a) & "','" & cantidad_requerida & "')")
                    Exit For
                    'articulos_ya_despachados = articulos_ya_despachados + (articulo_cantidad(a) - saldo)
                    '  End If
                    ' Else
                    'clase.agregar_registro("INSERT INTO `patron_despacho`(`desp_pedido`,`desp_bodega`,`desp_gondola`,`desp_articulo`,`desp_cant`) VALUES ( '" & codpedido & "','" & articulo_bodega(a) & "','" & articulo_gondola(a) & "','" & articulo_codigoart(a) & "','" & cantidad_requerida - articulos_ya_despachados & "')")
                    '    termino_asignacion = True
                    ' Exit For
                    'End If
                Next
            End If
        Next
        MessageBox.Show("Los patrones de despachos se han guardado satisfactoriamente.", "PATRONES ESTABLECIDOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Dim v As String
        v = MessageBox.Show("¿Desea imprimir el patrón creado?", "IMPRIMIR PATRON", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            imprimir_patron_pedido()
        End If
        Me.Close()
    End Sub

    Sub procedimiento_para_grabar_patrones()
        Dim a As Integer
        Dim existencia_referencia As Integer
        Dim existencia_referencia_virtual As Integer
        Dim cantidad_requerida As Integer
        Dim articulos_ya_despachados As Integer
        Dim saldo As Integer
        For x = 0 To DataGridView1.RowCount - 1
            existencia_referencia = DataGridView1.Item(10, x).Value
            existencia_referencia_virtual = DataGridView1.Item(11, x).Value
            cantidad_requerida = asignar_cantidad_requerida(DataGridView1.Item(11, x).Value)
            If ((cantidad_requerida > 0)) Then 'verifico que la cantidad requerida sea superior a 0
                buscar_existencia_bodega_gondola_cantidad(x)
                articulos_ya_despachados = 0
                For a = 0 To cantidad_gondolas_donde_existe_articulo - 1
                    Application.DoEvents()
                    ' MsgBox(DataGridView1.Item(0, x).Value & " -  " & articulo_bodega(a) & " - " & articulo_gondola(a))
                    saldo = saldo_gondola(articulo_bodega(a), articulo_gondola(a), articulo_codigoart(a))

                    'MsgBox(saldo)
                    ' MsgBox(DataGridView1.Item(0, x).Value & " / " & articulo_gondola(a) & " ----" & articulo_cantidad(a) - saldo & "---- " & articulo_cantidad(a) & "----" & saldo)
                    If (articulo_cantidad(a) - saldo) <= (cantidad_requerida - articulos_ya_despachados) Then 'verifico si la existencia del codigo actual es suficiente para cubrir lo requerido o debo completar con la existencia del codigo siguiente
                        If (articulo_cantidad(a) - saldo) > 0 Then
                            clase.agregar_registro("INSERT INTO `patron_despacho`(`desp_pedido`,`desp_bodega`,`desp_gondola`,`desp_articulo`,`desp_cant`) VALUES ( '" & codpedido & "','" & articulo_bodega(a) & "','" & articulo_gondola(a) & "','" & articulo_codigoart(a) & "','" & (articulo_cantidad(a) - saldo) & "')")
                            articulos_ya_despachados = articulos_ya_despachados + (articulo_cantidad(a) - saldo)
                        End If
                    Else
                        clase.agregar_registro("INSERT INTO `patron_despacho`(`desp_pedido`,`desp_bodega`,`desp_gondola`,`desp_articulo`,`desp_cant`) VALUES ( '" & codpedido & "','" & articulo_bodega(a) & "','" & articulo_gondola(a) & "','" & articulo_codigoart(a) & "','" & cantidad_requerida - articulos_ya_despachados & "')")
                        '    termino_asignacion = True
                        Exit For
                    End If
                Next
            End If
        Next
        MessageBox.Show("Los patrones de despachos se han guardado satisfactoriamente.", "PATRONES ESTABLECIDOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Dim v As String
        v = MessageBox.Show("¿Desea imprimir el patrón creado?", "IMPRIMIR PATRON", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            imprimir_patron_pedido()
        End If
        Me.Close()
    End Sub

    Private Sub buscar_existencia_bodega_gondola_cantidad(fila As Integer)
        With DataGridView1
            clase.consultar(buscar_codigos_a_apartir_de_la_referencia(.Item(0, fila).Value, .Item(1, fila).Value, .Item(4, fila).Value, .Item(5, fila).Value, .Item(6, fila).Value, .Item(7, fila).Value, .Item(8, fila).Value), "tbl1")
        End With
        Dim sql1 As String = "SELECT inv_codigoart,inv_bodega, inv_gondola, inv_cantidad FROM inventario_bodega WHERE (inv_cantidad >0) AND ("
        Dim p As Integer
        Dim ind As Boolean = False
        For p = 0 To clase.dt.Tables("tbl1").Rows.Count - 1
            If ind = False Then
                sql1 = sql1 & "(inv_codigoart =" & clase.dt.Tables("tbl1").Rows(p)("ar_codigo") & ")"
                ind = True
            Else
                sql1 = sql1 & " OR (inv_codigoart =" & clase.dt.Tables("tbl1").Rows(p)("ar_codigo") & ")"
            End If
        Next
        sql1 = sql1 & ") ORDER BY inv_bodega, inv_gondola ASC"
        clase.consultar(sql1, "cantidad")
        If clase.dt.Tables("cantidad").Rows.Count > 0 Then
            Dim y As Integer
            ReDim articulo_bodega(clase.dt.Tables("cantidad").Rows.Count)
            ReDim articulo_gondola(clase.dt.Tables("cantidad").Rows.Count)
            ReDim articulo_cantidad(clase.dt.Tables("cantidad").Rows.Count)
            ReDim articulo_codigoart(clase.dt.Tables("cantidad").Rows.Count)
            cantidad_gondolas_donde_existe_articulo = clase.dt.Tables("cantidad").Rows.Count
            For y = 0 To clase.dt.Tables("cantidad").Rows.Count - 1
                articulo_bodega(y) = clase.dt.Tables("cantidad").Rows(y)("inv_bodega")
                articulo_gondola(y) = clase.dt.Tables("cantidad").Rows(y)("inv_gondola")
                articulo_cantidad(y) = clase.dt.Tables("cantidad").Rows(y)("inv_cantidad")
                articulo_codigoart(y) = clase.dt.Tables("cantidad").Rows(y)("inv_codigoart")
            Next
        Else
            cantidad_gondolas_donde_existe_articulo = 0
            ReDim articulo_codigoart(0)
            ReDim articulo_bodega(0)
            ReDim articulo_gondola(0)
            ReDim articulo_cantidad(0)
        End If
    End Sub

    Private Sub buscar_existencia_bodega_gondola_cantidad1(fila As Integer)   'quitar despues del inventario
        With DataGridView1
            clase.consultar(buscar_codigos_a_apartir_de_la_referencia(.Item(0, fila).Value, .Item(1, fila).Value, .Item(4, fila).Value, .Item(5, fila).Value, .Item(6, fila).Value, .Item(7, fila).Value, .Item(8, fila).Value), "tbl1")
        End With
        Dim sql1 As String = "SELECT inv_codigoart,inv_bodega, inv_gondola, inv_cantidad FROM inventario_bodega WHERE ("
        Dim p As Integer
        Dim ind As Boolean = False
        For p = 0 To clase.dt.Tables("tbl1").Rows.Count - 1
            If ind = False Then
                sql1 = sql1 & "(inv_codigoart =" & clase.dt.Tables("tbl1").Rows(p)("ar_codigo") & ")"
                ind = True
            Else
                sql1 = sql1 & " OR (inv_codigoart =" & clase.dt.Tables("tbl1").Rows(p)("ar_codigo") & ")"
            End If
        Next
        sql1 = sql1 & ") ORDER BY inv_bodega, inv_gondola ASC"
        clase.consultar(sql1, "cantidad")
        If clase.dt.Tables("cantidad").Rows.Count > 0 Then
            Dim y As Integer
            ReDim articulo_bodega(clase.dt.Tables("cantidad").Rows.Count)
            ReDim articulo_gondola(clase.dt.Tables("cantidad").Rows.Count)
            ReDim articulo_cantidad(clase.dt.Tables("cantidad").Rows.Count)
            ReDim articulo_codigoart(clase.dt.Tables("cantidad").Rows.Count)
            cantidad_gondolas_donde_existe_articulo = clase.dt.Tables("cantidad").Rows.Count
            For y = 0 To clase.dt.Tables("cantidad").Rows.Count - 1
                articulo_bodega(y) = clase.dt.Tables("cantidad").Rows(y)("inv_bodega")
                articulo_gondola(y) = clase.dt.Tables("cantidad").Rows(y)("inv_gondola")
                articulo_cantidad(y) = clase.dt.Tables("cantidad").Rows(y)("inv_cantidad")
                articulo_codigoart(y) = clase.dt.Tables("cantidad").Rows(y)("inv_codigoart")
            Next
        Else
            cantidad_gondolas_donde_existe_articulo = 0
            ReDim articulo_codigoart(0)
            ReDim articulo_bodega(0)
            ReDim articulo_gondola(0)
            ReDim articulo_cantidad(0)
        End If
    End Sub

    Function referencia_desde_codigo(codigo As Integer) As String
        clase.consultar("select* from articulos where ar_codigo = " & codigo & "", "ab")
        Return clase.dt.Tables("ab").Rows(0)("ar_referencia")
    End Function

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox18_GotFocus(sender As Object, e As EventArgs) Handles TextBox18.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Function saldo_gondola(bodega As Integer, gondola As String, articulo As Integer) As Integer
        Dim sql3 As String = "SELECT SUM(patron_despacho.desp_cant) AS cant FROM cabpedido INNER JOIN patron_despacho ON (cabpedido.cabped_codigo = patron_despacho.desp_pedido) WHERE (cabpedido.cabped_fecha_est_llegada IS NULL AND patron_despacho.desp_bodega = " & bodega & " AND patron_despacho.desp_gondola = '" & gondola & "') AND (patron_despacho.desp_articulo = " & articulo & ")"
        clase.consultar(sql3, "tabla")
        If IsDBNull(clase.dt.Tables("tabla").Rows(0)("cant")) Then
            Return 0
        Else
            Return clase.dt.Tables("tabla").Rows(0)("cant")
        End If
    End Function

    Sub imprimir_patron_pedido()
        Dim gondola As String = ""
        Dim m_excel As Microsoft.Office.Interop.Excel.Application
        m_excel = CreateObject("Excel.Application")
        m_excel.Workbooks.Open("C:\Data\patrondespacho.xls")
        m_excel.Visible = False
        m_excel.Worksheets("Hoja1").cells(3, 1).value = "PEDIDO No: " & TextBox18.Text
        clase.consultar("SELECT tiendas.tienda FROM tiendas INNER JOIN cabpedido ON (tiendas.id = cabpedido.cabped_tienda) WHERE (cabpedido.cabped_codigo =" & TextBox18.Text & ")", "tabla2")
        m_excel.Worksheets("Hoja1").cells(4, 1).value = "DESTINATARIO: " & clase.dt.Tables("tabla2").Rows(0)("tienda")
        clase.consultar("SELECT patron_despacho.desp_articulo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, bodegas.bod_abreviatura, patron_despacho.desp_gondola, patron_despacho.desp_cant, patron_despacho.desp_bodega FROM bodegas INNER JOIN patron_despacho ON (bodegas.bod_codigo = patron_despacho.desp_bodega) INNER JOIN articulos ON (articulos.ar_codigo = patron_despacho.desp_articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (patron_despacho.desp_pedido =" & TextBox18.Text & ") ORDER BY patron_despacho.desp_bodega, patron_despacho.desp_gondola ASC", "listapatron")
        If clase.dt.Tables("listapatron").Rows.Count > 0 Then
            Dim t As Integer
            For t = 0 To clase.dt.Tables("listapatron").Rows.Count - 1
                With clase.dt.Tables("listapatron")
                    gondola = .Rows(t)("bod_abreviatura") & " - " & .Rows(t)("desp_gondola")
                    m_excel.Worksheets("Hoja1").cells(7 + t, 2).value = gondola
                    m_excel.Worksheets("Hoja1").cells(7 + t, 3).value = .Rows(t)("desp_articulo")
                    m_excel.Worksheets("Hoja1").cells(7 + t, 4).value = .Rows(t)("ar_referencia")
                    m_excel.Worksheets("Hoja1").cells(7 + t, 5).value = .Rows(t)("ar_descripcion")
                    m_excel.Worksheets("Hoja1").cells(7 + t, 6).value = .Rows(t)("ln1_nombre")
                    m_excel.Worksheets("Hoja1").cells(7 + t, 7).value = .Rows(t)("sl1_nombre")
                    m_excel.Worksheets("Hoja1").cells(7 + t, 8).value = .Rows(t)("desp_cant")
                End With
            Next
            m_excel.Application.ActiveWorkbook.PrintOutEx()
            If Not m_excel Is Nothing Then
                m_excel.Application.ActiveWorkbook.Saved = True
                m_excel.Quit()
                m_excel.Application.Quit()
                m_excel = Nothing
            End If
        End If
    End Sub

    ' Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    clase.consultar("SELECT desp_pedido FROM patron_despacho WHERE (desp_pedido =" & TextBox18.Text & ")", "impresion")
    '    If clase.dt.Tables("impresion").Rows.Count > 0 Then
    ' Dim v As String
    '         v = MessageBox.Show("¿Desea Imprimir el patron de despacho?", "IMPRIMIR PATRON DE DESPACHO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
    '         If v = 6 Then
    '            imprimir_patron_pedido()
    '        End If
    '    Else
    '       MessageBox.Show("El patrón de despacho aun no ha sido creado.", "IMPRIMIR PATRON DE DESPACHO", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '   End If
    '  End Sub

    Function cantidad_tipo_tienda(tienda As Integer) As Integer
        clase.consultar("SELECT cantidades_x_tipo.cantidad FROM cantidades_x_tipo INNER JOIN tiendas ON (cantidades_x_tipo.tipo = tiendas.tipotienda) WHERE (tiendas.id =" & tienda & ")", "tipo")
        Return clase.dt.Tables("tipo").Rows(0)("cantidad")
    End Function

    Function asignar_cantidad_requerida(existencia_virtual As Integer) As Integer
        Dim cantiptienda As Integer = cantidad_tipo_tienda(codtienda)
        If cantiptienda > existencia_virtual Then
            Return cantiptienda - existencia_virtual
        Else
            Return cantiptienda
        End If
    End Function

    Function asignar_cantidad_requerida1(existencia_virtual As Integer) As Integer  'quitar despues de inventario
        Return cantidad_tipo_tienda(codtienda)
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frm_distribucion_pedido.ShowDialog()
        frm_distribucion_pedido.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_ajustar_gongola_pedido.ShowDialog()
        frm_ajustar_gongola_pedido.Dispose()
    End Sub
End Class