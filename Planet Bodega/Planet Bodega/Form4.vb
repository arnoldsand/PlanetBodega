Public Class Form4
    Dim clase As New class_library
    Dim ean13 As New clase_ean13_codigo_de_barar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'consulta sql para hallar los saldos iniciales =>"SELECT detajuste.detaj_articulo, SUM(detajuste.detaj_cantidad) AS cant FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) WHERE (cabajuste.cabaj_codigo =152) OR (cabajuste.cabaj_codigo =153) OR (cabajuste.cabaj_codigo =154) GROUP BY detajuste.detaj_articulo"
        clase.consultar("SELECT detajuste.detaj_articulo, SUM(detajuste.detaj_cantidad) AS cant FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) WHERE (cabajuste.cabaj_fecha BETWEEN '2014-05-27' AND '2014-07-31' AND cabajuste.cabaj_tipo_ajuste =7 AND cabajuste.cabaj_procesado =TRUE) OR (cabajuste.cabaj_tipo_ajuste =2) OR (cabajuste.cabaj_tipo_ajuste =3) GROUP BY detajuste.detaj_articulo", "saldos")

        Dim i As Integer
        Dim cont As Integer = 0
        For i = 0 To clase.dt.Tables("saldos").Rows.Count - 1
            Application.DoEvents()
            TextBox1.Text = i + 1
            ' MsgBox("update inventario_bodega set inv_saldo_inicial = " & clase.dt.Tables("saldos").Rows(i)("cant") & " where inv_codigoart = " & clase.dt.Tables("saldos").Rows(i)("detaj_articulo") & "")
            clase.consultar1("select inv_cantidad from inventario_bodega where inv_codigoart = " & clase.dt.Tables("saldos").Rows(i)("detaj_articulo") & "", "act")
            If clase.dt1.Tables("act").Rows.Count > 0 Then
                clase.actualizar("update inventario_bodega set inv_ajustado_neg = " & System.Math.Abs(clase.dt.Tables("saldos").Rows(i)("cant")) & " where inv_codigoart = " & clase.dt.Tables("saldos").Rows(i)("detaj_articulo") & "")
            Else
                cont += 1
            End If

        Next
        MsgBox("Operación Exitosa " & cont)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clase.consultar1("SELECT cabajuste.cabaj_bodega, detajuste.detaj_gondola, detajuste.detaj_articulo, detajuste.detaj_cantidad FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) WHERE (cabajuste.cabaj_codigo =152) OR (cabajuste.cabaj_codigo =153) OR (cabajuste.cabaj_codigo =154)", "inventario")
        Dim x As Integer
        Dim cant As Short
        For x = 0 To clase.dt1.Tables("inventario").Rows.Count - 1
            Application.DoEvents()
            TextBox1.Text = x + 1
            With clase.dt1.Tables("inventario")
                cant = .Rows(x)("detaj_cantidad")
                clase.agregar_registro("INSERT INTO `detinventario`(`detinv_codigo`,`detinv_cod_inv`,`detinv_bodega`,`detinv_gondola`,`detinv_articulo`,`detinv_cant_cont1`,`detinv_cant_cont2`,`detinv_cant_cont3`,`detinv_cant_ajust`) VALUES ( NULL,'1','" & .Rows(x)("cabaj_bodega") & "','" & .Rows(x)("detaj_gondola") & "','" & .Rows(x)("detaj_articulo") & "','" & cant & "','" & cant & "','" & cant & "','" & cant & "')")
            End With
        Next
        MsgBox("Operacion terminada exitosamaente")
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("select ar_codigo, ar_referencia, ar_descripcion, ar_precio1, ar_costo from articulos", "articulos")
        If clase.dt.Tables("articulos").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("articulos")
        Else
            DataGridView1.DataSource = Nothing

        End If
        '  clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2, articulos.ar_precioanterior, articulos.ar_precioanterior2, articulos.ar_precioanttixy1, articulos.ar_precioanttixy2 FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (sublinea2.sl2_sl1codigo = sublinea1.sl1_codigo) AND (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN sublinea3 ON (sublinea3.sl3_sl2codigo = sublinea2.sl2_codigo) AND (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) AND (sublinea1.sl1_ln1codigo = linea1.ln1_codigo) LEFT JOIN sublinea4 ON (sublinea4.sl4_sl3codigo = sublinea3.sl3_codigo) AND (articulos.ar_sublinea4 = sublinea4.sl4_codigo) INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) ORDER BY articulos.ar_fechaingreso DESC ", "art")
        '  DataGridView1.DataSource = clase.dt.Tables("art")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        clase.consultar("select* from inventario_bodega", "inv")
        Dim z As Integer
        For z = 0 To clase.dt.Tables("inv").Rows.Count - 1
            Application.DoEvents()
            TextBox1.Text = z + 1
            With clase.dt.Tables("inv")
                If ((comprobar_nulidad_de_integer(.Rows(z)("inv_saldo_inicial")) + comprobar_nulidad_de_integer(.Rows(z)("inv_ajustado_pos")) + comprobar_nulidad_de_integer(.Rows(z)("inv_ingresado"))) - (comprobar_nulidad_de_integer(.Rows(z)("inv_transferido")) + comprobar_nulidad_de_integer(.Rows(z)("inv_ajustado_neg")))) <> .Rows(z)("inv_cantidad") Then

                    clase.agregar_registro("insert into codigos_descuadrados (codigo, inventario_inicial, inv_ajustado_pos, inv_ingresado, inv_tranferido, inv_ajustado_neg) values ('" & .Rows(z)("inv_codigoart") & "', '" & .Rows(z)("inv_saldo_inicial") & "', '" & .Rows(z)("inv_ajustado_pos") & "', '" & .Rows(z)("inv_ingresado") & "', '" & .Rows(z)("inv_transferido") & "', '" & .Rows(z)("inv_ajustado_neg") & "')")

                End If


            End With
        Next

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ean13.imprimir_codigos_de_barra_19_x_15_203dpi(119381, 3, "Generic / Text Only")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim cadena As String = "\\maquinaubuntu\Fotos Articulos\119445.jpg"
        Dim strArray() As String
        strArray = Split(cadena, "\")
        Dim intCount As Short
        For intCount = LBound(strArray) To UBound(strArray)
            MsgBox(strArray(intCount))
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        clase.borradoautomatico("delete from reporteenvio")
        clase.actualizar("INSERT INTO `reporteenvio`(`codigo`,`empresa`,`direccionremitente`,`telefono`,`ciudadremitente`,`guia`,`fechaenvio`,`transferencia`,`piezas`,`destino`,`direcciondestino`,`ciudaddestino`) SELECT envios.id_envio, empresas.Nombreempresa, empresas.direccion, empresas.telefono, empresas.ciudad, envios.nro_guia, envios.Fecha_envio, detalleenvios.nro_transfer, envios.unidades, tiendas.tienda, tiendas.direccion, CONCAT(tiendas.ciudad, ', ', tiendas.departamento) AS ciudad FROM envios INNER JOIN detalleenvios ON (envios.id_envio = detalleenvios.id_envio) INNER JOIN tiendas ON (tiendas.id = envios.id_tiendaRec) INNER JOIN empresas ON (empresas.idempresa = tiendas.Idempresa) WHERE (envios.id_envio =" & TextBox3.Text & ")")

        frm_reporte_envio.ShowDialog()
        frm_reporte_envio.Dispose()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        MsgBox(frm_movimiento_de_mercancia.hallar_ano_bisiesto(TextBox4.Text))
    End Sub
End Class