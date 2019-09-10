Public Class frm_entrada_mercancia
    Dim clase As New class_library
    '  Dim coditem As Integer
    Dim codigoconsecutivo As Integer
    Dim ind_codigo_cargado As Boolean
    Dim importacion As Integer
    Dim salida As String
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_entrada_mercancia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Focus()
        DataGridView2.ColumnCount = 5
        preparar_datagridview2()
        TextBox5.Text = determinar_consecutivo_caja()
        codigoconsecutivo = determinar_consecutivo_caja()
        ind_codigo_cargado = False
        salida = ""
    End Sub

    Function determinar_consecutivo_caja() As Integer
        Dim valor1 As Integer = 0
        clase.consultar("select* from consecutivo_codigobarra_importacion", "tblresult")
        If clase.dt.Tables("tblresult").Rows.Count > 0 Then
            valor1 = clase.dt.Tables("tblresult").Rows(0)("codigo_ultimo") + 1
            Return valor1
        Else
            Return valor1
        End If
    End Function

    Function hallar_codigo_importacion_apartir_salida(sal As Integer) As Integer
        clase.consultar("SELECT detalle_importacion_cabcajas.det_codigoimportacion FROM  detsalidas_mercancia INNER JOIN cabsalidas_mercancia ON (detsalidas_mercancia.det_salidacodigo = cabsalidas_mercancia.cabsal_cod) INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (cabsalidas_mercancia.cabsal_cod =" & sal & ")", "resultadostabla")
        If clase.dt.Tables("resultadostabla").Rows.Count > 0 Then
            Return clase.dt.Tables("resultadostabla").Rows(0)("det_codigoimportacion")
        End If
    End Function

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        salida = TextBox1.Text
        If clase.validar_cajas_text(TextBox1, "Número Movimiento") = False Then Exit Sub
        clase.consultar("select* from cabsalidas_mercancia where cabsal_cod = " & TextBox1.Text & "", "table123")
        If clase.dt.Tables("table123").Rows.Count > 0 Then
            If ind_codigo_cargado = True Then ' esta verificacion solo se va a hacer cuando el codigo de la importacion se halla cargado en "importacion"
                If importacion <> hallar_codigo_importacion_apartir_salida(TextBox1.Text) Then
                    MessageBox.Show("El movimiento que intenta visualizar pertenece a una caja relacionada en otra importación. No se pueden agregar items de este número de movimiento.", "CAJA PERTENECE A OTRA IMPORTACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox1.Text = ""
                    TextBox1.Focus()
                    restablecer_datagrid2()
                    Exit Sub
                End If
            End If
            clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_descripcion AS Descripcion, detsalidas_mercancia.det_cant AS Cant, detsalidas_mercancia.det_procesado AS Procesado, detsalidas_mercancia.det_codref, detalle_importacion_detcajas.detcab_cantidad, detsalidas_mercancia.det_codigo FROM  detalle_importacion_detcajas INNER JOIN detsalidas_mercancia ON (detalle_importacion_detcajas.detcab_coditem = detsalidas_mercancia.det_codref) WHERE (detsalidas_mercancia.det_salidacodigo = " & TextBox1.Text & ")", "resultados")
            If clase.dt.Tables("resultados").Rows.Count > 0 Then
                DataGridView2.Columns.Clear()
                DataGridView2.DataSource = clase.dt.Tables("resultados")
                preparar_datagridview2()
                Button2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
                Button6.Enabled = True
            End If

        Else
            MessageBox.Show("No se encontró ningún movimiento que coincida con el criterio escrito.", "MOVIMIENTO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Text = ""
            TextBox1.Focus()
            restablecer_datagrid2()
        End If
        ' determinar cuando se va a generara una caja apartir de varias, que sean de importaciones iguales
    End Sub

    Private Sub preparar_datagridview2()
        With DataGridView2
            .Columns(0).Width = 235
            .Columns(1).Width = 235
            .Columns(2).Width = 80
            .Columns(3).Width = 80
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripcion"
            .Columns(2).HeaderText = "Cant"
            .Columns(3).HeaderText = "Procesado"
        End With
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        TextBox2.Text = DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value
        TextBox3.Text = DataGridView2.Item(1, DataGridView2.CurrentRow.Index).Value
        TextBox4.Text = ""
        ' coditem = DataGridView2.Item(4, DataGridView2.CurrentRow.Index).Value
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox4, "Cantidad") = False Then Exit Sub
        If DataGridView2.Item(3, DataGridView2.CurrentRow.Index).Value = True Then
            MessageBox.Show("Ya esta referencia fue procesado y no puede ser devuelta a bodega.", "REFERENCIA PROCESADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            MessageBox.Show("Debe seleccionar una referencia para agregar.", "AGREGAR REFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If hallar_referencia_existente(DataGridView2.Item(4, DataGridView2.CurrentRow.Index).Value) Then 'le paso como parametro el codigo de la referencia en la bodega muerta
            MessageBox.Show("La referenca que intenta agregar ya fue agregada.", "REFERENCIA AGREGADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If ind_codigo_cargado = False Then  ' el codigo de importacion se asignará la primera vez que se agregue un item
            importacion = hallar_codigo_importacion_apartir_salida(salida)
            ind_codigo_cargado = True
        End If
        DataGridView1.RowCount = DataGridView1.RowCount + 1
        DataGridView1.Item(0, DataGridView1.RowCount - 1).Value = TextBox2.Text ' nombre referencia
        DataGridView1.Item(1, DataGridView1.RowCount - 1).Value = TextBox3.Text 'descripcion
        DataGridView1.Item(2, DataGridView1.RowCount - 1).Value = TextBox4.Text ' cantidad devuelta
        DataGridView1.Item(3, DataGridView1.RowCount - 1).Value = DataGridView2.Item(4, DataGridView2.CurrentRow.Index).Value ' codigo de la referencia en la bodega muerta
        DataGridView1.Item(4, DataGridView1.RowCount - 1).Value = DataGridView2.Item(5, DataGridView2.CurrentRow.Index).Value ' cantidad de la referencia en las  cajas 
        DataGridView1.Item(5, DataGridView1.RowCount - 1).Value = DataGridView2.Item(6, DataGridView2.CurrentRow.Index).Value ' codigo del item en la salida
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)
    End Sub

    Function hallar_referencia_existente(ref As String) As Boolean
        Dim d As Short
        For d = 0 To DataGridView1.RowCount - 1
            If ref = DataGridView1.Item(3, d).Value Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.RowCount = 0 Then
            MessageBox.Show("Debe agregar por lo menos un item para registrar la entrada a bodega.", "AGREGAR ITEMS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If clase.validar_cajas_text(TextBox5, "Número Caja") = False Then Exit Sub
        clase.consultar("select* from detalle_importacion_cabcajas where det_caja = " & TextBox5.Text & "", "busqueda")
        If clase.dt.Tables("busqueda").Rows.Count > 0 Then
            MessageBox.Show("El número de caja escrito no se puede utilizar porque fue asignado antes a otra caja, por favor especifique otro y vuelva a intentarlo.", "NÚMERO DE CAJA YA EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox5.Text = ""
            TextBox5.Focus()
            Exit Sub
        End If
        Dim v As String = MessageBox.Show("¿Desea registrar la entrada a bodega generada en este momento?", "GENERAR ENTRADA A BODEGA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            Dim fecha As String = Now.ToString("yyyy-MM-dd")
            clase.agregar_registro("insert into detalle_importacion_cabcajas (`det_codigoimportacion`,`det_caja`,`det_codigoproveedor`,`det_peso`,`det_hoja`,`det_fecharecepcion`) values ('" & importacion & "', '" & TextBox5.Text & "', '" & hallar_codigo_proveedor(salida) & "', NULL, NULL, '" & fecha & "')")
            Dim a As Short
            For a = 0 To DataGridView1.RowCount - 1
                clase.agregar_registro("INSERT INTO `detalle_importacion_detcajas`(`detcab_codigocaja`,`detcab_referencia`,`detcab_descripcion`,`detcab_marca`,`detcab_composicion`,`detcab_cantidad`,`detcab_unimedida`) VALUES ('" & TextBox5.Text & "','" & DataGridView1.Item(0, a).Value & "','" & DataGridView1.Item(1, a).Value & "','" & hallar_valor_del_parametro(DataGridView1.Item(3, a).Value, "marca") & "','" & hallar_valor_del_parametro(DataGridView1.Item(3, a).Value, "composicion") & "','" & hallar_valor_del_parametro(DataGridView1.Item(3, a).Value, "cantidad") & "','" & hallar_valor_del_parametro(DataGridView1.Item(3, a).Value, "unimedida") & "')")
                clase.actualizar("UPDATE `detsalidas_mercancia` SET `det_procesado`=True WHERE `det_codigo`=" & DataGridView1.Item(5, a).Value & " ")
            Next
            If TextBox5.Text = codigoconsecutivo Then
                clase.actualizar("UPDATE consecutivo_codigobarra_importacion set codigo_ultimo = " & TextBox5.Text & "")
            End If
            restablecer()
            TextBox5.Text = determinar_consecutivo_caja()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        restablecer()
    End Sub

    Private Sub restablecer()
        TextBox1.Text = ""
        TextBox1.Focus()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DataGridView2.DataSource = Nothing
        DataGridView1.RowCount = 0
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = True
        salida = ""
        importacion = vbEmpty
        ind_codigo_cargado = False
    End Sub

    Private Sub restablecer_datagrid2()
        DataGridView2.DataSource = Nothing
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.validar_numeros(e)
    End Sub

    Function hallar_codigo_proveedor(sal As Integer) As Integer
        clase.consultar("SELECT detalle_importacion_cabcajas.det_codigoproveedor FROM  detsalidas_mercancia INNER JOIN cabsalidas_mercancia ON (detsalidas_mercancia.det_salidacodigo = cabsalidas_mercancia.cabsal_cod) INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) WHERE (cabsalidas_mercancia.cabsal_cod =" & sal & " )", "rest")
        If clase.dt.Tables("rest").Rows.Count > 0 Then
            Return clase.dt.Tables("rest").Rows(0)("det_codigoproveedor")
        End If
    End Function

    Function hallar_valor_del_parametro(coditem As Integer, parametrobuscar As String) As String 'marca, composicion, unimedida, cantidad
        clase.consultar("SELECT detcab_" & parametrobuscar & " FROM detalle_importacion_detcajas WHERE (detcab_coditem =" & coditem & ")", "busq")
        If clase.dt.Tables("busq").Rows.Count > 0 Then
            Return clase.dt.Tables("busq").Rows(0)("detcab_" & parametrobuscar)
        Else
            Return ""
        End If
    End Function
   
End Class