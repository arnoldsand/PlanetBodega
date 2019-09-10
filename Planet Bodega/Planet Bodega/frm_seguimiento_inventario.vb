Public Class frm_seguimiento_inventario
    Dim clase As New class_library
    Dim pedido As Integer
    Dim unidades As Integer
    Dim cantref As Integer

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        conexion.Open()
        Dim cm1 As New OleDbCommand("select* from detalle_entrada where codigoarticulo <> 'N'", conexion)
        Dim c1 As New OleDbDataAdapter
        c1.SelectCommand = cm1
        Dim dt4 As New DataSet
        c1.Fill(dt4, "transferencia")
        conexion.Close()
        If dt4.Tables("transferencia").Rows.Count > 0 Then
            Dim pedido As Integer = dt4.Tables("transferencia").Rows(0)("pedido")
            unidades = frm_salida_de_mercancia_desde_hand_held.calcular_unidades(dt4.Tables("transferencia"))
            TextBox1.Text = pedido
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
                clase.consultar("SELECT transferencia_hand_held.cod_articulo, articulos.ar_referencia, articulos.ar_descripcion, transferencia_hand_held.cant, colores.colornombre, tallas.nombretalla FROM articulos INNER JOIN transferencia_hand_held ON (articulos.ar_codigo = transferencia_hand_held.cod_articulo) INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) WHERE (transferencia_hand_held.pedido =" & pedido & ")", "transferencia")
                If clase.dt.Tables("transferencia").Rows.Count > 0 Then
                    Dim p As Integer
                    Dim t As Integer = 0
                    Dim cantid As Integer = 0
                    Dim stok As Integer = 0
                    For p = 0 To clase.dt.Tables("transferencia").Rows.Count - 1
                        With DataGridView1
                            .RowCount = .RowCount + 1
                            .Item(0, t).Value = clase.dt.Tables("transferencia").Rows(p)("cod_articulo")
                            .Item(1, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_referencia")
                            .Item(2, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_descripcion")
                            .Item(3, t).Value = clase.dt.Tables("transferencia").Rows(p)("colornombre")
                            .Item(4, t).Value = clase.dt.Tables("transferencia").Rows(p)("nombretalla")
                            cantid = clase.dt.Tables("transferencia").Rows(p)("cant")
                            .Item(5, t).Value = cantid
                            stok = cant_en_inventario(clase.dt.Tables("transferencia").Rows(p)("cod_articulo"))
                            .Item(6, t).Value = stok
                            .Item(7, t).Value = cantid - stok
                            .Item(8, t).Value = frm_mantenimiento_pedidos.opciones_gondolas(clase.dt.Tables("transferencia").Rows(p)("cod_articulo"))
                            t = t + 1

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

    Private Sub frm_seguimiento_inventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inicializar_cadena_access()
    End Sub

    Function cant_en_inventario(articulo As Long) As Integer
        clase.consultar1("SELECT inv_cantidad FROM inventario_bodega WHERE inv_codigoart = " & articulo, "inventario")
        If clase.dt1.Tables("inventario").Rows.Count > 0 Then
            Return clase.dt1.Tables("inventario").Rows(0)("inv_cantidad")
        Else
            Return 0
        End If
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
            cantref = 0
            unidades = 0
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            Button1.Enabled = False
            Button4.Enabled = True
            PictureBox1.Image = Nothing
            clase.borradoautomatico("DELETE FROM transferencia_hand_held WHERE pedido = " & TextBox1.Text & "")
        End If
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim v As String = MessageBox.Show("¿Desea realizar los ajustes sugeridos al inventario?", "CONFIRMAR AJUSTES", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            If clase.validar_cajas_text(TextBox5, "Operario") = False Then Exit Sub
            Dim crear_ajuste_positivo As Boolean = False
            Dim crear_ajuste_negativos As Boolean = False
            Dim x As Short
            For x = 0 To DataGridView1.RowCount - 1 ' con este bucle verificamos si es necesario crear un ajuste positivo
                If DataGridView1(7, x).Value > 0 Then
                    crear_ajuste_positivo = True
                    Exit For
                End If
            Next
            For x = 0 To DataGridView1.RowCount - 1 ' con este bucle verificamos si es necesario crear un ajuste negativo
                If DataGridView1(7, x).Value < 0 Then
                    crear_ajuste_negativos = True
                    Exit For
                End If
            Next
            Dim cons_pos As Integer = vbEmpty
            Dim cons_neg As Integer = vbEmpty
            If crear_ajuste_positivo = True Then
                clase.agregar_registro("INSERT INTO `cabajuste`(`cabaj_codigo`,`cabaj_fecha`,`cabaj_hora`,`cabaj_tipo_ajuste`,`cabaj_operario`,`cabaj_observaciones`,`cabaj_procesado`) VALUES ( NULL,'" & Now.ToString("yyyy-MM-dd") & "','" & Now.ToString("HH:mm:ss") & "','6','" & UCase(TextBox5.Text) & "','AJUSTE POR SEGUIMIENTO',FALSE)")
                cons_pos = rescartar_consecutivo_cabajuste()
            End If
            If crear_ajuste_negativos = True Then
                clase.agregar_registro("INSERT INTO `cabajuste`(`cabaj_codigo`,`cabaj_fecha`,`cabaj_hora`,`cabaj_tipo_ajuste`,`cabaj_operario`,`cabaj_observaciones`,`cabaj_procesado`) VALUES ( NULL,'" & Now.ToString("yyyy-MM-dd") & "','" & Now.ToString("HH:mm:ss") & "','7','" & UCase(TextBox5.Text) & "','AJUSTE POR SEGUIMIENTO',FALSE)")
                cons_neg = rescartar_consecutivo_cabajuste()
            End If
            Dim existencia, ajust_pos, ajust_neg As Integer
            For x = 0 To DataGridView1.RowCount - 1
                clase.consultar2("select inv_cantidad, inv_ajustado_pos, inv_ajustado_neg from inventario_bodega where inv_codigoart = " & DataGridView1.Item(0, x).Value, "stock") 'busco las existencias, para modificarlas mas adelante
                If clase.dt2.Tables("stock").Rows.Count > 0 Then ' si esta en la tabla inventario_bodega
                    existencia = clase.dt2.Tables("stock").Rows(0)("inv_cantidad")
                    ajust_pos = comprobar_nulidad_de_integer(clase.dt2.Tables("stock").Rows(0)("inv_ajustado_pos"))
                    ajust_neg = comprobar_nulidad_de_integer(clase.dt2.Tables("stock").Rows(0)("inv_ajustado_neg"))
                    Select Case DataGridView1.Item(7, x).Value  'lleno los detale de la(s) tablas de ajustes los positivos apartes de los negativos
                        Case Is > 0
                            clase.agregar_registro("INSERT INTO `detajuste`(`detaj_codigo`,`detaj_codigo_ajuste`,`detaj_articulo`,`detaj_cantidad`,`detaj_cantidad_anterior`,`detaj_precio_costo`,`detaj_precio_venta1`,`detaj_precio_venta2`) VALUES ( NULL,'" & cons_pos & "','" & DataGridView1.Item(0, x).Value & "','" & DataGridView1.Item(7, x).Value & "','" & frm_ajustes.cant_anterior_antes_de_ajuste(DataGridView1.Item(0, x).Value) & "','" & Str(frm_detalle_ajuste.precio_costo(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta1(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta2(DataGridView1.Item(0, x).Value)) & "')")
                            clase.actualizar("UPDATE inventario_bodega SET inv_cantidad = " & existencia + DataGridView1.Item(7, x).Value & ", inv_ajustado_pos = " & ajust_pos + DataGridView1.Item(7, x).Value & " WHERE inv_codigoart = " & DataGridView1.Item(0, x).Value)
                        Case Is < 0
                            clase.agregar_registro("INSERT INTO `detajuste`(`detaj_codigo`,`detaj_codigo_ajuste`,`detaj_articulo`,`detaj_cantidad`,`detaj_cantidad_anterior`,`detaj_precio_costo`,`detaj_precio_venta1`,`detaj_precio_venta2`) VALUES ( NULL,'" & cons_neg & "','" & DataGridView1.Item(0, x).Value & "','" & DataGridView1.Item(7, x).Value & "','" & frm_ajustes.cant_anterior_antes_de_ajuste(DataGridView1.Item(0, x).Value) & "','" & Str(frm_detalle_ajuste.precio_costo(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta1(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta2(DataGridView1.Item(0, x).Value)) & "')")
                            clase.actualizar("UPDATE inventario_bodega SET inv_cantidad = " & existencia + DataGridView1.Item(7, x).Value & ", inv_ajustado_neg = " & ajust_neg + System.Math.Abs(DataGridView1.Item(7, x).Value) & " WHERE inv_codigoart = " & DataGridView1.Item(0, x).Value)
                    End Select
                Else 'si no esta en la tabla inventario_bodega
                    Select Case DataGridView1.Item(7, x).Value  'lleno los detale de la(s) tablas de ajustes los positivos apartes de los negativos
                        Case Is > 0
                            clase.agregar_registro("INSERT INTO `detajuste`(`detaj_codigo`,`detaj_codigo_ajuste`,`detaj_articulo`,`detaj_cantidad`,`detaj_cantidad_anterior`,`detaj_precio_costo`,`detaj_precio_venta1`,`detaj_precio_venta2`) VALUES ( NULL,'" & cons_pos & "','" & DataGridView1.Item(0, x).Value & "','" & DataGridView1.Item(7, x).Value & "','" & frm_ajustes.cant_anterior_antes_de_ajuste(DataGridView1.Item(0, x).Value) & "','" & Str(frm_detalle_ajuste.precio_costo(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta1(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta2(DataGridView1.Item(0, x).Value)) & "')")
                            clase.agregar_registro("INSERT INTO inventario_bodega (inv_codigoart, inv_ajustado_pos, inv_cantidad) VALUES ('" & DataGridView1.Item(0, x).Value & "', '" & ajust_pos + DataGridView1.Item(7, x).Value & "', '" & existencia + DataGridView1.Item(7, x).Value & "')")
                        Case Is < 0
                            clase.agregar_registro("INSERT INTO `detajuste`(`detaj_codigo`,`detaj_codigo_ajuste`,`detaj_articulo`,`detaj_cantidad`,`detaj_cantidad_anterior`,`detaj_precio_costo`,`detaj_precio_venta1`,`detaj_precio_venta2`) VALUES ( NULL,'" & cons_neg & "','" & DataGridView1.Item(0, x).Value & "','" & DataGridView1.Item(7, x).Value & "','" & frm_ajustes.cant_anterior_antes_de_ajuste(DataGridView1.Item(0, x).Value) & "','" & Str(frm_detalle_ajuste.precio_costo(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta1(DataGridView1.Item(0, x).Value)) & "','" & Str(frm_detalle_ajuste.precio_venta2(DataGridView1.Item(0, x).Value)) & "')")
                            clase.agregar_registro("INSERT INTO inventario_bodega (inv_codigoart, inv_ajustado_pos, inv_cantidad) VALUES ('" & DataGridView1.Item(0, x).Value & "', '" & ajust_neg + System.Math.Abs(DataGridView1.Item(7, x).Value) & "', '" & existencia + DataGridView1.Item(7, x).Value & "')")
                    End Select
                End If
            Next
            clase.actualizar("UPDATE cabajuste SET cabaj_procesado = TRUE WHERE cabaj_codigo = " & cons_pos & " OR cabaj_codigo = " & cons_neg)
            MessageBox.Show("Los ajuste por seguimiento fueron cargados exitosamente.", "AJUSTES REALIZADOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Enabled = True
            pedido = vbEmpty
            unidades = 0
            cantref = 0
            DataGridView1.Rows.Clear()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            PictureBox1.Image = Nothing
        End If
    End Sub


    Function rescartar_consecutivo_cabajuste() As Integer
        clase.consultar1("select MAX(cabaj_codigo) as maximo from cabajuste", "max")
        If clase.dt1.Tables("max").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("max").Rows(0)("maximo")) Then
                Return 1
            Else
                Return clase.dt1.Tables("max").Rows(0)("maximo")
            End If
        Else
            Return 1
        End If
    End Function
End Class