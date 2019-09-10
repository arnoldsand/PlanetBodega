Public Class frm_salida_de_mercancia_desde_hand_held
    Dim clase As New class_library
    Dim conflicto As Integer
    Dim pedido As Integer
    Dim unidades As Integer
    Dim cantref As Integer
    Private tiendaMigrada As Boolean

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        conflicto = 0
        conexion.Open()
        Dim cm1 As New OleDbCommand("select* from detalle_entrada where codigoarticulo <> 'N'", conexion)
        Dim c1 As New OleDbDataAdapter
        c1.SelectCommand = cm1
        Dim dt4 As New DataSet
        c1.Fill(dt4, "transferencia")
        conexion.Close()
        If dt4.Tables("transferencia").Rows.Count > 0 Then
            restablecer()
            pedido = dt4.Tables("transferencia").Rows(0)("pedido")
            'clase.consultar("SELECT* FROM cabpedido WHERE (cabped_codigo =" & pedido & " AND cabped_fecha_est_llegada IS NOT NULL)", "pedido")
            'If clase.dt.Tables("pedido").Rows.Count > 0 Then
            '    Dim v As String
            '    MessageBox.Show("El pedido No " & pedido & " ya fue despachado, no se puede volver a procesar", "PEDIDO YA DESPACHADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Exit Sub
            'End If
            unidades = calcular_unidades(dt4.Tables("transferencia"))
            TextBox1.Text = pedido
            TextBox4.Text = destino_x_pedido(pedido)
            TextBox2.Text = unidades
            cantref = dt4.Tables("transferencia").Rows.Count
            TextBox3.Text = cantref
            clase.consultar1("SELECT pedido FROM transferencia_hand_held WHERE (pedido =" & pedido & ")", "tabla3")
            If clase.dt1.Tables("tabla3").Rows.Count > 0 Then clase.borradoautomatico("Delete from transferencia_hand_held WHERE (pedido =" & pedido & ")")
            Dim v2 As String = MessageBox.Show("Se importará el pedido No " & pedido & " con " & dt4.Tables("transferencia").Rows.Count & " articulo(s) y " & unidades & " unidad(es) . ¿Desea continuar?", "IMPORTAR  PEDIDO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v2 = 6 Then
                Dim x As Integer
                For x = 0 To dt4.Tables("transferencia").Rows.Count - 1
                    If verificar_existencia_articulo(dt4.Tables("transferencia").Rows(x)("codigoarticulo")) = True Then
                        clase.agregar_registro("INSERT INTO `transferencia_hand_held`(`pedido`,`cod_articulo`,`cant`) VALUES ( '" & dt4.Tables("transferencia").Rows(x)("pedido") & "','" & convertir_codigobarra_a_codigo_normal(dt4.Tables("transferencia").Rows(x)("codigoarticulo")) & "','" & dt4.Tables("transferencia").Rows(x)("cantidad") & "')")
                    End If
                Next
                Dim cm3 As New OleDbCommand("delete from detalle_entrada")
                conexion.Open()
                cm3.Connection = conexion
                cm3.ExecuteNonQuery()
                conexion.Close()
                clase.consultar("SELECT transferencia_hand_held.cod_articulo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, transferencia_hand_held.cant FROM articulos INNER JOIN transferencia_hand_held ON (articulos.ar_codigo = transferencia_hand_held.cod_articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (transferencia_hand_held.pedido =" & pedido & ")", "transferencia")
                If clase.dt.Tables("transferencia").Rows.Count > 0 Then
                    Dim p As Integer
                    Dim t As Integer = 0
                    '  Dim diferencia As Integer
                    For p = 0 To clase.dt.Tables("transferencia").Rows.Count - 1
                        With DataGridView1
                            '  If verificar_existencia_articulo_en_pedido(clase.dt.Tables("transferencia").Rows(p)("cod_articulo"), pedido) = True Then
                            .RowCount = .RowCount + 1
                            ' este punto se obviará ya que la revisión de transferencia se hará en otro punto por tal razon ya no es necesario tener en cuenta el patron de despacho, es mas creo que este último se eliminará
                            'diferencia = cantidad_articulo_pedido(clase.dt.Tables("transferencia").Rows(p)("cod_articulo"), pedido) - clase.dt.Tables("transferencia").Rows(p)("cant")
                            'If diferencia <> 0 Then
                            '    conflicto = conflicto + 1
                            '    .Rows(t).DefaultCellStyle.ForeColor = Color.Red
                            'End If
                            .Item(0, t).Value = clase.dt.Tables("transferencia").Rows(p)("cod_articulo")
                            .Item(1, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_referencia")
                            .Item(2, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_descripcion")
                            .Item(3, t).Value = clase.dt.Tables("transferencia").Rows(p)("ln1_nombre")
                            .Item(4, t).Value = clase.dt.Tables("transferencia").Rows(p)("sl1_nombre")
                            .Item(5, t).Value = clase.dt.Tables("transferencia").Rows(p)("cant")
                            ' .Item(6, t).Value = diferencia
                            '.Item(7, p).Value = "No pertenece al Pedido"
                            t = t + 1
                            'End If
                        End With
                    Next
                    MessageBox.Show("Los articulos se importaron satisfactoriamente.", "ARTICULOS IMPORTADOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Button4.Enabled = False

                    Button1.Enabled = True
                End If
            End If

        Else
            MessageBox.Show("No se encontraron datos para importar.", "IMPOSIBLE IMPORTAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Sub llenar_listado()
        conflicto = 0
        DataGridView1.RowCount = 0
        clase.consultar("SELECT transferencia_hand_held.cod_articulo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, transferencia_hand_held.cant FROM articulos INNER JOIN transferencia_hand_held ON (articulos.ar_codigo = transferencia_hand_held.cod_articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (transferencia_hand_held.pedido =" & pedido & ")", "transferencia")
        If clase.dt.Tables("transferencia").Rows.Count > 0 Then
            '  restablecer()
            Dim p As Integer
            Dim t As Integer = 0
            Dim diferencia As Integer
            For p = 0 To clase.dt.Tables("transferencia").Rows.Count - 1
                With DataGridView1
                    If verificar_existencia_articulo_en_pedido(clase.dt.Tables("transferencia").Rows(p)("cod_articulo"), pedido) = True Then
                        .RowCount = .RowCount + 1
                        diferencia = cantidad_articulo_pedido(clase.dt.Tables("transferencia").Rows(p)("cod_articulo"), pedido) - clase.dt.Tables("transferencia").Rows(p)("cant")
                        If diferencia <> 0 Then
                            conflicto = conflicto + 1
                            .Rows(t).DefaultCellStyle.ForeColor = Color.Red
                        End If
                        .Item(0, t).Value = clase.dt.Tables("transferencia").Rows(p)("cod_articulo")
                        .Item(1, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_referencia")
                        .Item(2, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_descripcion")
                        .Item(3, t).Value = clase.dt.Tables("transferencia").Rows(p)("ln1_nombre")
                        .Item(4, t).Value = clase.dt.Tables("transferencia").Rows(p)("sl1_nombre")
                        .Item(5, t).Value = clase.dt.Tables("transferencia").Rows(p)("cant")
                        .Item(6, t).Value = diferencia
                        '.Item(7, p).Value = "No pertenece al Pedido"
                        t = t + 1
                    End If
                End With
            Next

        End If
    End Sub

    Sub restablecer()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        DataGridView1.RowCount = 0
    End Sub

    Function verificar_existencia_articulo_en_pedido(articulo As Integer, ped As Integer) As Boolean
        clase.consultar1("SELECT * FROM patron_despacho WHERE (desp_articulo =" & articulo & " AND desp_pedido =" & ped & ")", "articulo")
        If clase.dt1.Tables("articulo").Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub frm_salida_de_mercancia_desde_hand_held_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inicializar_cadena_access()
        With DataGridView1
            .ColumnCount = 8
            .Columns(0).Width = 70
            .Columns(1).Width = 150
            .Columns(2).Width = 150
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 60
            .Columns(6).Width = 60
            .Columns(7).Width = 300
            .Columns(0).HeaderText = "Código"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Línea"
            .Columns(4).HeaderText = "Sublínea"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "Faltante"
            .Columns(7).HeaderText = "Observaciones"
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        crear_encabezado_transferencia()
    End Sub

    Function calcular_unidades(dt8 As DataTable) As Integer
        Dim z As Integer
        Dim cant As Integer = 0
        For z = 0 To dt8.Rows.Count - 1
            cant = cant + dt8.Rows(z)("cantidad")
        Next
        Return cant
    End Function

    Function destino_x_pedido(ped As Integer) As String
        clase.consultar1("SELECT tiendas.tienda FROM tiendas INNER JOIN cabpedido ON (tiendas.id = cabpedido.cabped_tienda) WHERE (cabpedido.cabped_codigo =" & ped & ")", "nombretienda")
        If clase.dt1.Tables("nombretienda").Rows.Count > 0 Then
            Return clase.dt1.Tables("nombretienda").Rows(0)("tienda")
        Else
            MessageBox.Show("El número de pedido relacionado en la captura no fue encontrado. Verifique si existe y vuelva a intentarlo.", "NUMERO DE PEDIDO NO ENCOTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End If

    End Function

    Function codigo_destino_x_pedido(ped As Integer) As String
        clase.consultar1("SELECT tiendas.id FROM tiendas INNER JOIN cabpedido ON (tiendas.id = cabpedido.cabped_tienda) WHERE (cabpedido.cabped_codigo =" & ped & ")", "nombretienda")
        Return clase.dt1.Tables("nombretienda").Rows(0)("id")
    End Function

    Function cantidad_articulo_pedido(articulo As Integer, ped As Integer) As Integer
        clase.consultar1("SELECT SUM(desp_cant) AS cant FROM patron_despacho WHERE (desp_pedido =" & ped & " AND desp_articulo =" & articulo & ")", "tabla1")
        If IsDBNull(clase.dt1.Tables("tabla1").Rows(0)("cant")) Then
            Return 0
        Else
            Return clase.dt1.Tables("tabla1").Rows(0)("cant")
        End If
    End Function

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            estado_gondola_pedido.ShowDialog()
            estado_gondola_pedido.Dispose()
            llenar_listado()
        End If
    End Sub

    Function VerificarExistenciaDeCodigosSQLServer() As Boolean
        Dim b As Integer
        Dim existe As Boolean = True
        For b = 0 To DataGridView1.RowCount - 1
            existe = VerificarExistenciaArticuloLovePOS(DataGridView1.Item(0, b).Value)
            If existe = False Then
                Exit For
            End If
        Next
        Return existe
    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            If clase.validar_cajas_text(TextBox5, "Operario") = False Then Exit Sub
            Dim fecha As Date = Now
            'verificar si la tienda fue migrada
            clase.consultar("SELECT* FROM tiendas WHERE id = " & codigo_destino_x_pedido(TextBox1.Text) & "", "migrado")
            If clase.dt.Tables("migrado").Rows.Count > 0 Then
                If clase.dt.Tables("migrado").Rows(0)("MigradoLovePos") = True Then
                    tiendaMigrada = True
                Else
                    tiendaMigrada = False
                End If
            Else
                tiendaMigrada = False
            End If
            If (tiendaMigrada = True) And (VerificarExistenciaDeCodigosSQLServer() = False) Then
                MessageBox.Show("Algunos de los codigos capturados no han sigo migrados a LovePOS, por favor realize la migración y vuelva a intentarlo.", "CODIGOS NO ENCONTRADOS", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Exit Sub
            End If
            clase.consultar2("SELECT cod_articulo, SUM(cant) AS cantidad FROM transferencia_hand_held WHERE (pedido =" & TextBox1.Text & ") GROUP BY cod_articulo", "transferencia")
            Dim x As Short
            Dim totinv, tottra, transferido As Integer
            For x = 0 To clase.dt2.Tables("transferencia").Rows.Count - 1
                clase.consultar1("SELECT inv_cantidad, inv_transferido FROM inventario_bodega WHERE (inv_codigoart =" & clase.dt2.Tables("transferencia").Rows(x)("cod_articulo") & ")", "tablita")
                If clase.dt1.Tables("tablita").Rows.Count > 0 Then
                    totinv = clase.dt1.Tables("tablita").Rows(0)("inv_cantidad")
                    tottra = clase.dt2.Tables("transferencia").Rows(x)("cantidad")
                    transferido = comprobar_nulidad_de_integer(clase.dt1.Tables("tablita").Rows(0)("inv_transferido"))
                    clase.actualizar("UPDATE inventario_bodega SET inv_cantidad = " & totinv - tottra & ", inv_transferido = " & transferido + tottra & " WHERE (inv_codigoart =" & clase.dt2.Tables("transferencia").Rows(x)("cod_articulo") & ")")
                Else
                    clase.agregar_registro("INSERT INTO `inventario_bodega`(`inv_codigo`,`inv_codigoart`,`inv_transferido`,`inv_cantidad`) VALUES ( NULL,'" & clase.dt2.Tables("transferencia").Rows(x)("cod_articulo") & "','" & clase.dt2.Tables("transferencia").Rows(x)("cantidad") & "','" & -1 * clase.dt2.Tables("transferencia").Rows(x)("cantidad") & "')")
                End If
                clase.agregar_registro("insert into `dettransferencia`(`dt_numero`,`dt_trnumero`,`dt_gondola`,`dt_codarticulo`,`dt_cantidad`,`dt_costo`,`dt_venta1`,`dt_venta2`,`dt_costo2`) values ( NULL,'" & TextBox6.Text & "',NULL,'" & clase.dt2.Tables("transferencia").Rows(x)("cod_articulo") & "','" & clase.dt2.Tables("transferencia").Rows(x)("cantidad") & "','" & Str(RecuperarCostoFiscal(clase.dt2.Tables("transferencia").Rows(x)("cod_articulo"))) & "','" & Str(precio_venta1(clase.dt2.Tables("transferencia").Rows(x)("cod_articulo"))) & "','" & Str(precio_venta2(clase.dt2.Tables("transferencia").Rows(x)("cod_articulo"))) & "','" & Str(RecuperarCostoReal(clase.dt2.Tables("transferencia").Rows(x)("cod_articulo"))) & "')")
            Next
            clase.actualizar("UPDATE cabtransferencia SET tr_estado = TRUE, tr_operador = '" & UCase(TextBox5.Text) & "', tr_destino = " & codigo_destino_x_pedido(TextBox1.Text) & "   WHERE tr_numero = " & TextBox6.Text & "")
            clase.consultar1("select cabped_transferencia from cabpedido where cabped_codigo = " & TextBox1.Text & "", "cons")
            Dim cadena_transferencia As String
            If IsDBNull(clase.dt1.Tables("cons").Rows(0)("cabped_transferencia")) Then
                cadena_transferencia = TextBox6.Text
            Else
                cadena_transferencia = clase.dt1.Tables("cons").Rows(0)("cabped_transferencia") & "-" & TextBox6.Text
            End If
            clase.actualizar("UPDATE cabpedido SET cabped_fecha_est_llegada = '" & fecha.ToString("yyyy-MM-dd") & "', cabped_transferencia = '" & cadena_transferencia & "' WHERE cabped_codigo = " & TextBox1.Text & "")
            clase.borradoautomatico("DELETE FROM transferencia_hand_held WHERE pedido = " & TextBox1.Text & "")
            'VOY POR AQUI !!!!!
            DataGridView1.Rows.Clear()
            TextBox1.Text = ""
            TextBox4.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Enabled = True
            PictureBox1.Image = Nothing
            crear_encabezado_transferencia()
        End If
    End Sub


    Private Sub crear_encabezado_transferencia()
        TextBox6.Text = consecutivo_transferencia()
        Dim fecha As Date = Now
        clase.agregar_registro("INSERT INTO `cabtransferencia`(`tr_numero`,`tr_destino`,`tr_origen`,`tr_bodega`,`tr_fecha`,`tr_hora`,`tr_estado`,`tr_operador`,`tr_revisada`,`tr_finalizada`) VALUES ( '" & TextBox6.Text & "',NULL,NULL,'4','" & fecha.ToString("yyyy-MM-dd") & "','" & Now.ToString("HH:mm") & "',FALSE,NULL,FALSE,FALSE)")
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

    Function precio_costo(articulo As Double) As Double
        clase.consultar("select ar_costo from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("ar_costo")
    End Function

    Function precio_venta1(articulo As Double) As Double
        clase.consultar("select ar_precio1 from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("ar_precio1")
    End Function

    Function precio_venta2(articulo As Double) As Double
        clase.consultar("select ar_precio2 from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("ar_precio2")
    End Function

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        With DataGridView1
            clase.consultar("select ar_foto from articulos where ar_codigo = " & .Item(0, .CurrentCell.RowIndex).Value & "", "tablita")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim v1 As String = MessageBox.Show("¿Desea deshacer los datos importados actuales?", "DESHACER DATOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v1 = 6 Then
            DataGridView1.RowCount = 0
            pedido = vbEmpty
            conflicto = 0
            cantref = 0
            unidades = 0
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            Button1.Enabled = False
            Button4.Enabled = True
            PictureBox1.Image = Nothing
            clase.borradoautomatico("DELETE FROM transferencia_hand_held WHERE pedido = " & TextBox1.Text & "")
        End If
    End Sub

    Private Sub TextBox6_GotFocus1(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub DateTimePicker1_Click(sender As Object, e As EventArgs) Handles DateTimePicker1.Click

    End Sub
End Class