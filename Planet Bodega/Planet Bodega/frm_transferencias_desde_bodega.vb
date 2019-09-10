Imports System.IO
Public Class frm_transferencias_desde_bodega
    Dim clase As New class_library
    Dim tiendaMigrada As Boolean

    Private Sub frmTransferencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = consecutivo_transferencia()
        clase.llenar_combo(ComboBox9, "select* from tiendas where estado = TRUE order by tienda ASC", "tienda", "id")
        TextBox2.Text = FormatCurrency(0, 0)
        TextBox3.Text = FormatCurrency(0, 0)
        TextBox4.Text = FormatCurrency(0, 0)
        Dim fecha As Date = Now
        clase.agregar_registro("INSERT INTO `cabtransferencia`(`tr_numero`,`tr_destino`,`tr_origen`,`tr_bodega`,`tr_fecha`,`tr_estado`,`tr_operador`,`tr_revisada`,`tr_finalizada` ) VALUES ( '" & TextBox1.Text & "',NULL,NULL,'4','" & fecha.ToString("yyyy-MM-dd") & "',FALSE,NULL,FALSE,FALSE)")
        DataGridView1.ColumnCount = 7
        preparar_columnas()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
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

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_combobox(ComboBox9, "Destino") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox18, "Operario") = False Then Exit Sub
        Button2.Enabled = False
        Button6.Enabled = True
        Button4.Enabled = True
        Button3.Enabled = True
        ComboBox9.Enabled = False
        TextBox18.Enabled = False
        TextBox5.Enabled = True
        '  Button1.Enabled = False
        clase.actualizar("UPDATE cabtransferencia SET tr_destino = " & ComboBox9.SelectedValue & ", tr_operador = '" & UCase(TextBox18.Text) & "', tr_revisada = FALSE, tr_finalizada = FALSE WHERE tr_numero = " & TextBox1.Text & "")
        clase.consultar("SELECT* FROM tiendas WHERE id = " & ComboBox9.SelectedValue & "", "migrado")
        If clase.dt.Tables("migrado").Rows.Count > 0 Then
            If clase.dt.Tables("migrado").Rows(0)("MigradoLovePos") = True Then
                tiendaMigrada = True
            Else
                tiendaMigrada = False
            End If
        Else
            tiendaMigrada = False
        End If
        TextBox5.Focus()
        Me.AcceptButton = Button3
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_cajas_text(TextBox5, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox5.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos." & campo_codigobarra_largo(ComboBox9.SelectedValue) & " ='" & TextBox5.Text & "')", "tabla11")
            '  codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox5.Text)            esta linea seguramente debe ser quitada
        End If
        If Len(TextBox5.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, articulos.ar_costo, articulos.ar_precio1, articulos.ar_precio2 FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos." & campo_codigobarra_corto(ComboBox9.SelectedValue) & " =" & TextBox5.Text & ")", "tabla11")
            ' codigo_articulo = TextBox5.Text     esta linea seguramente debe ser quitada
        End If
        If Len(TextBox5.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox5.Text = ""
            TextBox5.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            codigo_articulo = clase.dt.Tables("tabla11").Rows(0)("ar_codigo")
            If ((tiendaMigrada = True) And (VerificarExistenciaArticuloLovePOS(codigo_articulo) = False)) Then
                MessageBox.Show("El codigo capturado no ha sigo migrado a LovePOS, por favor realize la migración y vuelva a intentarlo.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox5.Text = ""
                TextBox5.Focus()
                Exit Sub
            End If
            With clase.dt.Tables("tabla11")
                clase.consultar1("select* from dettransferencia where dt_trnumero = " & TextBox1.Text & " AND dt_codarticulo = " & codigo_articulo & "", "result")
                If clase.dt1.Tables("result").Rows.Count > 0 Then
                    Dim cant As Integer = clase.dt1.Tables("result").Rows(0)("dt_cantidad")
                    clase.actualizar("UPDATE dettransferencia SET dt_cantidad = " & cant + 1 & " WHERE dt_trnumero = " & TextBox1.Text & " AND dt_codarticulo = " & codigo_articulo & "")
                Else
                    clase.agregar_registro("INSERT INTO `dettransferencia`(`dt_trnumero`,`dt_gondola`,`dt_codarticulo`,`dt_cantidad`,`dt_danado`,`dt_costo`,`dt_venta1`,`dt_venta2`,`dt_costo2`) VALUES ('" & TextBox1.Text & "',NULL,'" & codigo_articulo & "','1','0','" & Str(RecuperarCostoFiscal(codigo_articulo)) & "','" & Str(.Rows(0)("ar_precio1")) & "','" & Str(.Rows(0)("ar_precio2")) & "','" & Str(RecuperarCostoReal(codigo_articulo)) & "')")
                End If
            End With
            EstablecerFoto(codigo_articulo)
            llenar_grilla(TextBox1.Text)
            Dim x As Short
            For x = 0 To DataGridView1.RowCount - 1
                If codigo_articulo = DataGridView1.Item(0, x).Value Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, x)
                    Exit For
                End If
            Next
            TextBox5.Text = ""
            TextBox5.Focus()
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox5.Text = ""
            TextBox5.Focus()
        End If
    End Sub

    Sub llenar_grilla(transferencias As Integer)
        Dim destinotrans As Integer = vbEmpty 'hallar destino transferencia apartir del numero
        clase.consultar("SELECT tr_destino FROM cabtransferencia WHERE (tr_numero =" & transferencias & ")", "numerotransferencia")
        If clase.dt.Tables("numerotransferencia").Rows.Count > 0 Then
            destinotrans = clase.dt.Tables("numerotransferencia").Rows(0)("tr_destino")
        End If
        clase.consultar("SELECT articulos." & campo_codigobarra_corto(destinotrans) & ", articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, dettransferencia.dt_cantidad, FORMAT(dt_venta1 * dt_cantidad, 'Currency') FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (dettransferencia.dt_trnumero =" & transferencias & ")", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("tabla")
            preparar_columnas()
            TextBox2.Text = FormatCurrency(precion_costo(TextBox1.Text), 0)
            TextBox3.Text = FormatCurrency(precion_venta1(TextBox1.Text), 0)
            TextBox4.Text = FormatCurrency(precion_venta2(TextBox1.Text), 0)
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 7
            preparar_columnas()
            TextBox2.Text = FormatCurrency(0, 0)
            TextBox3.Text = FormatCurrency(0, 0)
            TextBox4.Text = FormatCurrency(0, 0)
        End If
    End Sub

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"
            .Columns(5).HeaderText = "Cant"
            .Columns(6).HeaderText = "Precio Venta"
            .Columns(0).Width = 80
            .Columns(1).Width = 100
            .Columns(2).Width = 180
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 60
            .Columns(6).Width = 90
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        frm_eliminar3.ShowDialog()
        frm_eliminar3.Dispose()
    End Sub

    Sub cantidad_unidades()
        Dim i As Integer
        Dim cant_filas As Integer = 0
        For i = 0 To DataGridView1.RowCount - 1
            cant_filas = cant_filas + DataGridView1.Item(8, i).Value
        Next
        '  TextBox1.Text = cant_filas
        '   TextBox3.Text = DataGridView1.RowCount
    End Sub

    Function precion_costo(transferencia As Integer) As Double
        clase.consultar("SELECT SUM(dt_costo * dt_cantidad) AS costo FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ") GROUP BY dt_trnumero", "table")
        If IsDBNull(clase.dt.Tables("table").Rows(0)("costo")) Then
            Return 0
        End If
        Return clase.dt.Tables("table").Rows(0)("costo")
    End Function

    Function precion_venta1(transferencia As Integer) As Double
        clase.consultar("SELECT SUM(dt_venta1 * dt_cantidad) AS venta1 FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ") GROUP BY dt_trnumero", "table")
        If IsDBNull(clase.dt.Tables("table").Rows(0)("venta1")) Then
            Return 0
        End If
        Return clase.dt.Tables("table").Rows(0)("venta1")
    End Function

    Function precion_venta2(transferencia As Integer) As Double
        clase.consultar("SELECT SUM(dt_venta2 * dt_cantidad) AS Venta2 FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ") GROUP BY dt_trnumero", "table")
        If IsDBNull(clase.dt.Tables("table").Rows(0)("venta2")) Then
            Return 0
        End If
        Return clase.dt.Tables("table").Rows(0)("venta2")
    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frm_suspendidas.ShowDialog()
        frm_suspendidas.Dispose()
        If TextBox5.Enabled = True Then
            TextBox5.Focus()
        End If
    End Sub

    Sub establecer_encabezado(transferencia As Integer)
        clase.consultar("select* from cabtransferencia where tr_numero = " & transferencia & "", "tbl1")
        If clase.dt.Tables("tbl1").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tbl1").Rows(0)("tr_destino")) Then
                ComboBox9.SelectedValue = -1
            Else
                ComboBox9.SelectedValue = clase.dt.Tables("tbl1").Rows(0)("tr_destino")
            End If
            TextBox18.Text = clase.dt.Tables("tbl1").Rows(0)("tr_operador") & ""
            ComboBox9.Enabled = False
            TextBox18.Enabled = False
            TextBox5.Enabled = True
            TextBox5.Focus()
        End If
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            Dim v As String = MessageBox.Show("Desea Cerrar la transferencia actual, no podrá modificarla posteriormente. ¿Desea Continuar? ", "CERRAR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                clase.actualizar("UPDATE cabtransferencia SET tr_estado = TRUE, tr_finalizada = TRUE, tr_hora = '" & Now.ToString("HH:mm:ss") & "'  WHERE tr_numero = " & TextBox1.Text & "")

                Dim a As Short
                Dim totinv, tottra, transferido As Integer
                For a = 0 To DataGridView1.RowCount - 1
                    clase.consultar("SELECT inv_cantidad, inv_transferido FROM inventario_bodega WHERE (inv_codigoart =" & DataGridView1.Item(0, a).Value & ")", "tablita")
                    If clase.dt.Tables("tablita").Rows.Count > 0 Then
                        totinv = clase.dt.Tables("tablita").Rows(0)("inv_cantidad")
                        tottra = DataGridView1.Item(5, a).Value
                        transferido = comprobar_nulidad_de_integer(clase.dt.Tables("tablita").Rows(0)("inv_transferido"))
                        clase.actualizar("UPDATE inventario_bodega SET inv_cantidad = " & totinv - tottra & ", inv_transferido = " & transferido + tottra & " WHERE (inv_codigoart =" & DataGridView1.Item(0, a).Value & ")")
                    Else
                        clase.agregar_registro("INSERT INTO `inventario_bodega`(`inv_codigo`,`inv_codigoart`,`inv_transferido`,`inv_cantidad`) VALUES ( NULL,'" & DataGridView1.Item(0, a).Value & "','" & DataGridView1.Item(5, a).Value & "','" & -1 * DataGridView1.Item(5, a).Value & "')")
                    End If
                Next

                'Dim k As String = MessageBox.Show("¿Desea enviar el archivo de transferencia por e-mail al destino?", "ENVIAR ARCHIVO DE TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                ' If k = 6 Then
                'construir_archivo(TextBox1.Text)
                ' End If
                'imprimir_hoja_transferencia(TextBox1.Text)
                MessageBox.Show("La Transferencia se cerró Exitosamente.", "TRANSFERENCIA REALIZADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox1.Text = consecutivo_transferencia()
                ' numerotransferencia = TextBox1.Text
                guardar_encabezado_transferencia()
                TextBox2.Text = FormatCurrency(0, 0)
                TextBox3.Text = FormatCurrency(0, 0)
                TextBox4.Text = FormatCurrency(0, 0)
                preparar_columnas()
                TextBox18.Text = ""
                TextBox18.Enabled = True
                DataGridView1.DataSource = Nothing
                ComboBox9.SelectedIndex = -1
                ComboBox9.Enabled = True
                Button2.Enabled = True
                Button3.Enabled = False
                Button6.Enabled = False
                Button4.Enabled = False
                TextBox5.Enabled = False
                '    Button1.Enabled = True
                TextBox5.Text = ""
            End If
        End If
    End Sub

    Sub guardar_encabezado_transferencia()
        Dim fecha As Date = Now
        clase.agregar_registro("INSERT INTO `cabtransferencia`(`tr_numero`,`tr_destino`,`tr_origen`,`tr_bodega`,`tr_fecha`,`tr_estado`,`tr_operador`) VALUES ( '" & TextBox1.Text & "',NULL,NULL,'4','" & fecha.ToString("yyyy-MM-dd") & "',FALSE,NULL)")
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs)
        clase.validar_numeros(e)
    End Sub



    Function calcular_cant_articulos_x_ordenproduccion(articulo As Long, ordenproduccion As Short, tienda As Short) As Short
        clase.consultar("SELECT SUM(cantidad) AS cantidad FROM distribucion_orden_de_produccion WHERE (ordendeproduccion =" & ordenproduccion & " AND articulo =" & articulo & " AND tienda = " & tienda & ") GROUP BY articulo", "articulo")
        If clase.dt.Tables("articulo").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("articulo").Rows(0)("cantidad")) Then
                Return 0
            Else
                Return clase.dt.Tables("articulo").Rows(0)("cantidad")
            End If
        Else
            Return 0
        End If
    End Function

    Public Sub EstablecerFoto(IdArticulo As Integer)
        clase.consultar("select ar_foto from articulos where ar_codigo = " & IdArticulo & "", "tablita")
        If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
            PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
        Else
            PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
        End If
        SetImage(PictureBox1)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Rows.Count = 0 Then Exit Sub
        EstablecerFoto(DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value)
    End Sub
End Class