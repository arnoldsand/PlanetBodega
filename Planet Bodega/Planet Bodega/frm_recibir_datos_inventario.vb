Public Class frm_recibir_datos_inventario
    Dim clase As New class_library
    Dim pedido As Integer
    Dim unidades As Integer
    Dim cantref As Integer
    Dim bodega As Short

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        conexion.Open()
        Dim cm1 As New OleDbCommand("select* from detalle_entrada where codigoarticulo <> 'N'", conexion)
        Dim c1 As New OleDbDataAdapter
        c1.SelectCommand = cm1
        Dim dt4 As New DataSet
        c1.Fill(dt4, "transferencia")
        conexion.Close()
        If dt4.Tables("transferencia").Rows.Count > 0 Then
            pedido = dt4.Tables("transferencia").Rows(0)("pedido")

            clase.consultar1("SELECT secciones_inventario.*, bodegas.* FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_codigo =" & pedido & ")", "inv")
            If clase.dt1.Tables("inv").Rows.Count > 0 Then
                TextBox4.Text = clase.dt1.Tables("inv").Rows(0)("secc_operario")
                TextBox6.Text = clase.dt1.Tables("inv").Rows(0)("secc_conteo")
                TextBox7.Text = clase.dt1.Tables("inv").Rows(0)("bod_nombre")
                TextBox5.Text = clase.dt1.Tables("inv").Rows(0)("secc_gondola")
                bodega = clase.dt1.Tables("inv").Rows(0)("bod_codigo")  'reiniciar esta varible despues de guardar
            Else
                MessageBox.Show("El número de captura asociado no esta relacionado con ninguna sección del inventario. Debe cambiar el número antes de guardar la captura.", "SECCIÓN NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End If
            unidades = calcular_unidades(dt4.Tables("transferencia"))
            TextBox1.Text = pedido
            TextBox2.Text = unidades
            cantref = dt4.Tables("transferencia").Rows.Count
            TextBox3.Text = cantref
            clase.consultar1("SELECT captura FROM tabla_capturas_inventarios WHERE (captura =" & pedido & ")", "tabla3")
            If clase.dt1.Tables("tabla3").Rows.Count > 0 Then clase.borradoautomatico("Delete from tabla_capturas_inventarios WHERE (captura =" & pedido & ")")
            Dim v2 As String = MessageBox.Show("Se importará el pedido No " & pedido & " con " & dt4.Tables("transferencia").Rows.Count & " articulo(s) y " & unidades & " unidad(es) . ¿Desea continuar?", "IMPORTAR  PEDIDO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v2 = 6 Then
                Dim x As Integer
                For x = 0 To dt4.Tables("transferencia").Rows.Count - 1
                    If verificar_existencia_articulo(dt4.Tables("transferencia").Rows(x)("codigoarticulo")) = True Then
                        clase.agregar_registro("INSERT INTO `tabla_capturas_inventarios`(`captura`,`articulo`,`cant`) VALUES ( '" & dt4.Tables("transferencia").Rows(x)("pedido") & "','" & convertir_codigobarra_a_codigo_normal(dt4.Tables("transferencia").Rows(x)("codigoarticulo")) & "','" & dt4.Tables("transferencia").Rows(x)("cantidad") & "')")
                    End If
                Next
                Dim cm3 As New OleDbCommand("delete from detalle_entrada")
                conexion.Open()
                cm3.Connection = conexion
                '  cm3.ExecuteNonQuery()
                conexion.Close()
                clase.consultar("SELECT tabla_capturas_inventarios.articulo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, tabla_capturas_inventarios.cant FROM articulos INNER JOIN tabla_capturas_inventarios ON (articulos.ar_codigo = tabla_capturas_inventarios.articulo) INNER JOIN linea1 ON (linea1.ln1_codigo = articulos.ar_linea) INNER JOIN sublinea1 ON (sublinea1.sl1_codigo = articulos.ar_sublinea1) WHERE (tabla_capturas_inventarios.captura =" & pedido & ")", "transferencia")
                If clase.dt.Tables("transferencia").Rows.Count > 0 Then
                    Dim p As Integer
                    Dim t As Integer = 0
                    For p = 0 To clase.dt.Tables("transferencia").Rows.Count - 1
                        With dataGridView1
                            .RowCount = .RowCount + 1
                            .Item(0, t).Value = clase.dt.Tables("transferencia").Rows(p)("articulo")
                            .Item(1, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_referencia")
                            .Item(2, t).Value = clase.dt.Tables("transferencia").Rows(p)("ar_descripcion")
                            .Item(3, t).Value = clase.dt.Tables("transferencia").Rows(p)("ln1_nombre")
                            .Item(4, t).Value = clase.dt.Tables("transferencia").Rows(p)("sl1_nombre")
                            .Item(5, t).Value = clase.dt.Tables("transferencia").Rows(p)("cant")
                            t = t + 1
                        End With
                    Next
                    MessageBox.Show("Los articulos se importaron satisfactoriamente.", "ARTICULOS IMPORTADOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    TextBox4.Focus()
                    Button1.Enabled = False
                    Button3.Enabled = True
                    Button2.Enabled = True
                    Button8.Enabled = True
                End If
            End If

        Else
            MessageBox.Show("No se encontraron datos para importar.", "IMPOSIBLE IMPORTAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Function calcular_unidades(dt8 As DataTable) As Integer
        Dim z As Integer
        Dim cant As Integer = 0
        For z = 0 To dt8.Rows.Count - 1
            cant = cant + dt8.Rows(z)("cantidad")
        Next
        Return cant
    End Function

    Private Sub frm_recibir_datos_inventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inicializar_cadena_access()
        With dataGridView1
            .ColumnCount = 6
            .Columns(0).Width = 70
            .Columns(1).Width = 150
            .Columns(2).Width = 150
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 60
            .Columns(0).HeaderText = "Código"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Línea"
            .Columns(4).HeaderText = "Sublínea"
            .Columns(5).HeaderText = "Cant"
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As EventArgs) Handles TextBox7.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        frm_cambiar_numero_seccion.ShowDialog()
        frm_cambiar_numero_seccion.Dispose()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If dataGridView1.RowCount > 0 Then
            clase.consultar1("SELECT secciones_inventario.*, bodegas.* FROM secciones_inventario INNER JOIN bodegas ON (secciones_inventario.secc_bodega = bodegas.bod_codigo) WHERE (secciones_inventario.secc_codigo =" & TextBox1.Text & ")", "inv")
            If clase.dt1.Tables("inv").Rows.Count = 0 Then
                MessageBox.Show("El número de captura asociado no esta relacionado con ninguna sección del inventario. Debe cambiar el número antes de guardar la captura.", "SECCIÓN NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If clase.validar_cajas_text(TextBox4, "Operario") = False Then Exit Sub
            clase.consultar("SELECT detinv_codigo FROM detinventario WHERE (detinv_cod_inv =" & frm_modulo_inventario.TextBox1.Text & " AND detinv_bodega =" & bodega & " AND detinv_gondola ='" & UCase(TextBox5.Text) & "')", "gondola")
            If clase.dt.Tables("gondola").Rows.Count > 0 Then
                Dim x As Short
                Dim sql1 As String = ""
                For x = 0 To dataGridView1.RowCount - 1
                    Select Case TextBox6.Text
                        Case 1
                            sql1 = "UPDATE detinventario SET detinv_cant_cont1 = " & dataGridView1.Item(5, x).Value & " WHERE (detinv_cod_inv =" & frm_modulo_inventario.TextBox1.Text & " AND detinv_bodega =" & bodega & " AND detinv_gondola ='" & UCase(TextBox5.Text) & "' AND detinv_articulo = " & dataGridView1.Item(5, x).Value & ")"
                        Case 2
                            sql1 = "UPDATE detinventario SET detinv_cant_cont2 = " & dataGridView1.Item(5, x).Value & " WHERE (detinv_cod_inv =" & frm_modulo_inventario.TextBox1.Text & " AND detinv_bodega =" & bodega & " AND detinv_gondola ='" & UCase(TextBox5.Text) & "' AND detinv_articulo = " & dataGridView1.Item(5, x).Value & ")"
                        Case 3
                            sql1 = "UPDATE detinventario SET detinv_cant_cont3 = " & dataGridView1.Item(5, x).Value & " WHERE (detinv_cod_inv =" & frm_modulo_inventario.TextBox1.Text & " AND detinv_bodega =" & bodega & " AND detinv_gondola ='" & UCase(TextBox5.Text) & "' AND detinv_articulo = " & dataGridView1.Item(5, x).Value & ")"
                    End Select
                    clase.actualizar(sql1)
                Next
            Else
                Dim sql1 As String = ""
                For x = 0 To dataGridView1.RowCount - 1
                    Select Case TextBox6.Text
                        Case 1
                            sql1 = "INSERT INTO `detinventario`(`detinv_codigo`,`detinv_cod_inv`,`detinv_bodega`,`detinv_gondola`,`detinv_articulo`,`detinv_cant_cont1`,`detinv_cant_cont2`,`detinv_cant_cont3`,`detinv_cant_definitiva`) VALUES ( NULL,'" & frm_modulo_inventario.TextBox1.Text & "','" & bodega & "','" & TextBox5.Text & "','" & dataGridView1.Item(0, x).Value & "','" & dataGridView1.Item(5, x).Value & "',NULL,NULL,NULL)"
                        Case 2
                            sql1 = "INSERT INTO `detinventario`(`detinv_codigo`,`detinv_cod_inv`,`detinv_bodega`,`detinv_gondola`,`detinv_articulo`,`detinv_cant_cont1`,`detinv_cant_cont2`,`detinv_cant_cont3`,`detinv_cant_definitiva`) VALUES ( NULL,'" & frm_modulo_inventario.TextBox1.Text & "','" & bodega & "','" & TextBox5.Text & "','" & dataGridView1.Item(0, x).Value & "',NULL,'" & dataGridView1.Item(5, x).Value & "',NULL,NULL)"
                        Case 3
                            sql1 = "INSERT INTO `detinventario`(`detinv_codigo`,`detinv_cod_inv`,`detinv_bodega`,`detinv_gondola`,`detinv_articulo`,`detinv_cant_cont1`,`detinv_cant_cont2`,`detinv_cant_cont3`,`detinv_cant_definitiva`) VALUES ( NULL,'" & frm_modulo_inventario.TextBox1.Text & "','" & bodega & "','" & TextBox5.Text & "','" & dataGridView1.Item(0, x).Value & "',NULL,NULL,'" & dataGridView1.Item(5, x).Value & "',NULL)"
                    End Select
                    clase.agregar_registro(sql1)
                Next
            End If
        End If
    End Sub
End Class