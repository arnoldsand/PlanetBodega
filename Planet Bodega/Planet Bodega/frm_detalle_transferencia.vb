Public Class frm_detalle_transferencia
    Dim clase As New class_library
    Dim columna As New DataGridViewCheckBoxColumn

    Private Sub frm_detalle_transferencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar()
    End Sub

    Private Sub llenar()
        clase.consultar("SELECT tiendas.tienda, cabtransferencia.tr_operador, SUM(dt_costo * dt_cantidad) AS Costo, SUM(dt_venta1 * dt_cantidad) AS Precio1, SUM(dt_venta2 * dt_cantidad) AS Precio2 FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas  ON (cabtransferencia.tr_destino = tiendas.id) WHERE (cabtransferencia.tr_numero =" & frm_mantenimiento_de_transfencia.DataGridView2.Item(0, frm_mantenimiento_de_transfencia.DataGridView2.CurrentRow.Index).Value & ") GROUP BY dettransferencia.dt_trnumero", "tabla5")
        TextBox1.Text = frm_mantenimiento_de_transfencia.DataGridView2.Item(0, frm_mantenimiento_de_transfencia.DataGridView2.CurrentRow.Index).Value
        TextBox5.Text = clase.dt.Tables("tabla5").Rows(0)("tienda")
        TextBox18.Text = clase.dt.Tables("tabla5").Rows(0)("tr_operador")
        TextBox2.Text = FormatCurrency(clase.dt.Tables("tabla5").Rows(0)("Costo"), 0)
        TextBox3.Text = FormatCurrency(clase.dt.Tables("tabla5").Rows(0)("Precio1"), 0)
        TextBox4.Text = FormatCurrency(clase.dt.Tables("tabla5").Rows(0)("Precio2"), 0)
        clase.consultar("SELECT dettransferencia.dt_codarticulo AS Codigo, articulos.ar_referencia AS Referencia, articulos.ar_descripcion AS Descripcion, colores.colornombre AS Color, tallas.nombretalla AS Talla, dettransferencia.dt_cantidad AS Cant, FORMAT(SUM(dt_costo * dt_cantidad), 'Currency') AS PrecioCosto, FORMAT(SUM(dt_venta1 * dt_cantidad), 'Currency') AS PrecioVenta1 FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) WHERE (dettransferencia.dt_trnumero =" & frm_mantenimiento_de_transfencia.DataGridView2.Item(0, frm_mantenimiento_de_transfencia.DataGridView2.CurrentRow.Index).Value & ") GROUP BY dettransferencia.dt_codarticulo", "tabla6")
        If clase.dt.Tables("tabla6").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("tabla6")
            preparar_columnas()
            DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewCheckBoxColumn() {Me.columna})
            DataGridView1.Columns(8).Width = 20
        Else
            DataGridView1.DataSource = Nothing
        End If
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

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 70
            .Columns(1).Width = 110
            .Columns(2).Width = 150
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 50
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Dim a As Short
        For a = 0 To 7
            DataGridView1.Columns(a).ReadOnly = True
        Next
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox18_GotFocus(sender As Object, e As EventArgs) Handles TextBox18.GotFocus
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

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        clase.consultar("SELECT tr_finalizada FROM cabtransferencia WHERE (tr_numero =" & TextBox1.Text & ")", "transferencia")
        If clase.dt.Tables("transferencia").Rows(0)("tr_finalizada") = False Then
            MessageBox.Show("No se puede enviar una transferencia  que no se ha revisado.", "NO SE PUEDE ENVIAR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        construir_archivo(TextBox1.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            Dim v As String = MessageBox.Show("¿Desea Imprimir la transferencia seleccionada?", "IMPRIMIR TRANSFERENCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                clase.consultar("SELECT tr_finalizada FROM cabtransferencia WHERE (tr_numero =" & TextBox1.Text & ")", "transferencia")
                If clase.dt.Tables("transferencia").Rows(0)("tr_finalizada") = False Then
                    MessageBox.Show("No se puede imprimir una transferencia  que no se ha revisado.", "NO SE PUEDE IMPRIMIR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                imprimir_hoja_transferencia(TextBox1.Text)
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim a As Integer
        Dim cont As Short = 0
        For a = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(9, a).Value = True Then
                cont += 1
            End If
        Next
        If cont = 0 Then
            MessageBox.Show("Debe seleccionar por lo menos un item para eliminar de la transferencia", "SELECCIONAR ITEM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim v As String = MessageBox.Show("¿Esta seguro que desea eliminar de la transferencia los articulos seleccionados?", "CONFIRME ELIMINACIÓN", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If v = 6 Then
            If (DataGridView1.RowCount - cont) < 1 Then
                MessageBox.Show("No se pueden eliminar todos los items de la transferencia, si lo que desea es anularla utilize la opción anular transferencia del modulo mantenimiento de transferencia.", "OPERACIÓN NO PERMITIDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                deseleccionar()
                Exit Sub
            End If
            clase.consultar("SELECT* FROM cabtransferencia WHERE (tr_numero =" & TextBox1.Text & ")", "esttra")
            If clase.dt.Tables("esttra").Rows.Count > 0 Then
                If clase.dt.Tables("esttra").Rows(0)("tr_revisada") = True Then
                    Dim tiptra As Boolean
                    If IsDBNull(clase.dt.Tables("esttra").Rows(0)("tr_bodega")) Then
                        tiptra = True
                    Else
                        tiptra = False
                    End If
                    Dim x As Short
                    For x = 0 To DataGridView1.RowCount - 1
                        If DataGridView1.Item(9, x).Value = True Then
                            clase.agregar_registro("INSERT INTO faltante_transferencia (rev_transferencia, rev_articulo, rev_faltante) VALUES ('" & TextBox1.Text & "', '" & DataGridView1.Item(0, x).Value & "', '" & -1 * DataGridView1.Item(5, x).Value & "')")
                            clase.agregar_registro("INSERT INTO eliminados_transferencia (trelim_idtransferencia, trelim_articulo, trelim_cantidad) VALUES ('" & TextBox1.Text & "', '" & DataGridView1.Item(0, x).Value & "', '" & DataGridView1.Item(5, x).Value & "')")
                            clase.borradoautomatico("Delete from dettransferencia where dt_trnumero = " & TextBox1.Text & " AND dt_codarticulo = " & DataGridView1.Item(0, x).Value & "")
                        End If
                    Next
                    If tiptra = True Then
                        clase.consultar1("SELECT * FROM faltante_transferencia WHERE (rev_transferencia =" & TextBox1.Text & ")", "revision")
                        If clase.dt1.Tables("revision").Rows.Count > 0 Then
                            Dim z As Short
                            For z = 0 To clase.dt1.Tables("revision").Rows.Count - 1
                                sumar_al_inventario(clase.dt1.Tables("revision").Rows(z)("rev_articulo"), clase.dt1.Tables("revision").Rows(z)("rev_faltante"))
                            Next
                        End If
                    End If
                    MessageBox.Show("Los articulos fueron eliminados satisfactoriamente.", "ARTICULOS ELIMINADOS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    DataGridView1.Columns.Clear()
                    DataGridView1.DataSource = Nothing
                    llenar()
                Else
                    MessageBox.Show("Para realizar esta acción la transferencia debe de haber pasado primero por el proceso de revisión.", "REVISAR TRANSFERENCIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub

    Private Sub deseleccionar()
        Dim p As Short
        For p = 0 To DataGridView1.RowCount - 1
            DataGridView1.Item(9, p).Value = False
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            Dim FormularioEdicion As New frm_edicion_cantidad(TextBox1.Text, DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value, DataGridView1.Item(5, DataGridView1.CurrentCell.RowIndex).Value)
            FormularioEdicion.ShowDialog()
            FormularioEdicion.Dispose()
            llenar()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Rows.Count = 0 Then Exit Sub
        EstablecerFoto(DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value)
    End Sub
End Class