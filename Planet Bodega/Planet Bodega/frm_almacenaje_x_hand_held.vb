Public Class frm_almacenaje_x_hand_held
    Dim clase As New class_library
    Dim pedido As Integer
    Dim unidades As Integer
    Dim cantref As Integer
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

    Private Sub frm_almacenaje_x_hand_held_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox1, "select* from bodegas order by bod_codigo asc", "bod_nombre", "bod_codigo")
        DataGridView1.ColumnCount = 6
        preparar_columnas()
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

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_entradas_mercancias_suspendidas.ShowDialog()
        frm_entradas_mercancias_suspendidas.Dispose()
        TextBox5.Focus()
    End Sub

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"
            .Columns(5).HeaderText = "Cant"
            .Columns(0).Width = 80
            .Columns(1).Width = 150
            .Columns(2).Width = 180
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 80
        End With
    End Sub

    Sub EstablecerFoto(IdArticulo As Integer)
        Try
            clase.consultar("select ar_foto from articulos where ar_codigo = " & IdArticulo & "", "tablita")
            If System.IO.File.Exists(clase.dt.Tables("tablita").Rows(0)("ar_foto")) Then
                PictureBox1.Image = Image.FromFile(clase.dt.Tables("tablita").Rows(0)("ar_foto"))
            Else
                PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            End If
            SetImage(PictureBox1)
        Catch
            PictureBox1.Image = Global.WindowsApplication1.My.Resources.Resources.sinfoto
            SetImage(PictureBox1)
        End Try
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        inicializar_cadena_access()
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
            unidades = calcular_unidades(dt4.Tables("transferencia"))
            TextBox1.Text = pedido
            TextBox2.Text = unidades
            cantref = dt4.Tables("transferencia").Rows.Count
            TextBox3.Text = cantref
            clase.consultar1("SELECT pedido FROM entrada_hand_held WHERE (pedido =" & pedido & ")", "tabla3")
            If clase.dt1.Tables("tabla3").Rows.Count > 0 Then clase.borradoautomatico("Delete from entrada_hand_held WHERE (pedido =" & pedido & ")")
            Dim v2 As String = MessageBox.Show("Se importará el pedido No " & pedido & " con " & dt4.Tables("transferencia").Rows.Count & " articulo(s) y " & unidades & " unidad(es) . ¿Desea continuar?", "IMPORTAR  PEDIDO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v2 = 6 Then
                Dim x As Integer
                For x = 0 To dt4.Tables("transferencia").Rows.Count - 1
                    If verificar_existencia_articulo(dt4.Tables("transferencia").Rows(x)("codigoarticulo")) = True Then
                        clase.agregar_registro("INSERT INTO `entrada_hand_held`(`pedido`,`cod_articulo`,`cant`) VALUES ( '" & dt4.Tables("transferencia").Rows(x)("pedido") & "','" & convertir_codigobarra_a_codigo_normal(dt4.Tables("transferencia").Rows(x)("codigoarticulo")) & "','" & dt4.Tables("transferencia").Rows(x)("cantidad") & "')")
                    End If
                Next
                Dim cm3 As New OleDbCommand("delete from detalle_entrada")
                conexion.Open()
                cm3.Connection = conexion
                cm3.ExecuteNonQuery()
                conexion.Close()
                clase.consultar("SELECT entrada_hand_held.cod_articulo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla, entrada_hand_held.cant FROM articulos INNER JOIN entrada_hand_held ON (articulos.ar_codigo = entrada_hand_held.cod_articulo) INNER JOIN colores ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (entrada_hand_held.pedido =" & pedido & ")", "transferencia")
                If clase.dt.Tables("transferencia").Rows.Count > 0 Then
                    DataGridView1.Columns.Clear()
                    DataGridView1.DataSource = clase.dt.Tables("transferencia")
                    acomodar_columnas()
                    MessageBox.Show("Los articulos se importaron satisfactoriamente.", "ARTICULOS IMPORTADOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Else
            MessageBox.Show("No se encontraron datos para importar.", "IMPOSIBLE IMPORTAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Sub acomodar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Color"
            .Columns(4).HeaderText = "Talla"
            .Columns(5).HeaderText = "Cant"
            .Columns(0).Width = 80
            .Columns(1).Width = 120
            .Columns(2).Width = 140
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 80
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Sub restablecer()
        Button4.Enabled = False
        Button1.Enabled = True
        Button2.Enabled = True
        DataGridView1.DataSource = Nothing
        TextBox5.Text = ""
        TextBox5.Focus()
        TextBox2.Text = ""
        unidades = 0
        TextBox3.Text = ""
        cantref = 0
    End Sub

    Function verificar_existencia_articulo(articulo As String) As Boolean
        clase.consultar("SELECT ar_codigo FROM articulos WHERE (ar_codigobarras ='" & articulo & "')", "articulo")
        If clase.dt.Tables("articulo").Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function calcular_unidades(dt8 As DataTable) As Integer
        Dim z As Integer
        Dim cant As Integer = 0
        For z = 0 To dt8.Rows.Count - 1
            cant = cant + dt8.Rows(z)("cantidad")
        Next
        Return cant
    End Function

    Function cant_articulos_x_gondola(gondola As String, bodega As Short, articulo As Integer) As Integer
        clase.consultar("SELECT inv_cantidad FROM inventario_bodega WHERE (inv_bodega =" & bodega & " AND inv_gondola ='" & gondola & "' AND inv_codigoart =" & articulo & ")", "rest")
        If clase.dt.Tables("rest").Rows.Count > 0 Then
            Return clase.dt.Tables("rest").Rows(0)("inv_cantidad")
        Else
            Return 0
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(TextBox6, "Entrada") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox5, "Góndola") = False Then Exit Sub
        If verificar_existencia_gondola(UCase(TextBox5.Text), ComboBox1.SelectedValue) = True Then
            Dim v As String = MessageBox.Show("¿Desea guardar la captura hecha en este momento?", "GUARDAR ENTRADA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then

                clase.borradoautomatico("delete from detalle_entrada where detent_cod_entrada = " & TextBox6.Text & " and detent_gondola = '" & UCase(TextBox5.Text) & "'")
                Dim i As Short
                For i = 0 To DataGridView1.RowCount - 1
                    With DataGridView1        '"INSERT INTO `detajuste`(`detaj_codigo_ajuste`,`detaj_gondola`,`detaj_articulo`,`detaj_cantidad`,`detaj_cantidad_anterior`,`detaj_precio_costo`,`detaj_precio_venta1`,`detaj_precio_venta2`) VALUES ('" & TextBox6.Text & "','" & UCase(TextBox5.Text) & "','" & .Item(0, i).Value & "','" & .Item(5, i).Value & "','" & calcular_existencia_anterior(.Item(0, i).Value, ComboBox1.SelectedValue, UCase(TextBox5.Text)) & "','" & Str(precio_costo(.Item(0, i).Value)) & "','" & Str(precio_venta1(.Item(0, i).Value)) & "','" & Str(precio_venta2(.Item(0, i).Value)) & "')"   
                        clase.agregar_registro("INSERT INTO detalle_entrada (detent_cod_entrada, detent_gondola, detent_articulo, detent_cantidad) VALUES ('" & TextBox6.Text & "', '" & UCase(TextBox5.Text) & "', '" & DataGridView1.Item(0, i).Value & "', '" & DataGridView1.Item(5, i).Value & "')")
                    End With
                Next
                'restablecer()
                DataGridView1.DataSource = Nothing
                TextBox3.Text = ""
                TextBox2.Text = ""
                TextBox1.Text = ""
                TextBox5.Text = ""
                Button1.Enabled = False
                Button2.Enabled = False
                Button4.Enabled = True
                unidades = 0
                cantref = 0
                TextBox6.Text = ""
                ComboBox1.SelectedIndex = -1
            End If
        Else
            MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox5.Text = ""
            TextBox5.Focus()
        End If
    End Sub

    Function verificar_existencia_gondola(gondo As String, bodega As Short) As Boolean
        Dim gondola(cant_gondola_bodega(bodega)) As String
        clase.consultar("SELECT detbodega.dtbod_bloque, detbodega.dtbod_cant_gondola FROM detbodega INNER JOIN bodegas ON (detbodega.dtbod_bodega = bodegas.bod_codigo) WHERE (bodegas.bod_codigo =" & bodega & ")", "tabla1")
        If clase.dt.Tables("tabla1").Rows.Count > 0 Then
            Dim a, b As Short
            Dim cont As Integer = -1
            For a = 0 To clase.dt.Tables("tabla1").Rows.Count - 1
                For b = 1 To clase.dt.Tables("tabla1").Rows(a)("dtbod_cant_gondola")
                    cont += 1
                    gondola(cont) = bloque(clase.dt.Tables("tabla1").Rows(a)("dtbod_bloque") - 1) & b
                Next
            Next
            Dim i As Integer
            For i = 0 To cont
                If gondo = gondola(i) Then
                    Return True
                    Exit Function
                End If
            Next
            Return False
        End If
    End Function

    Function cant_gondola_bodega(bodega As Short) As Integer
        clase.consultar("SELECT SUM(detbodega.dtbod_cant_gondola) AS totalgondola FROM detbodega INNER JOIN bodegas ON (detbodega.dtbod_bodega = bodegas.bod_codigo) WHERE (bodegas.bod_codigo =" & bodega & ")", "tablatotal")
        If clase.dt.Tables("tablatotal").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("tablatotal").Rows(0)("totalgondola")) Then
                Return 0
                Exit Function
            End If
            Return clase.dt.Tables("tablatotal").Rows(0)("totalgondola")
        End If
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.ColumnCount = 6
        preparar_columnas()
        ComboBox1.SelectedIndex = -1
        TextBox6.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = True
        pedido = 0
        cantref = 0
        unidades = 0
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Rows.Count = 0 Then Exit Sub
        EstablecerFoto(DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value)
    End Sub
End Class