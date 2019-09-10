Public Class frm_detalle_recibo
    Dim clase As New class_library
    Private Sub frm_detalle_recibo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '  preparar_columnas()
        clase.consultar("select* from cabrecibos_bodegas where recbod_codigo = " & frm_mantenimiento_recibo_bodega.DataGridView2.Item(0, frm_mantenimiento_recibo_bodega.DataGridView2.CurrentCell.RowIndex).Value & "", "listo")
        If clase.dt.Tables("listo").Rows.Count > 0 Then
            txtproveedor.Text = clase.dt.Tables("listo").Rows(0)("recbod_proveedor")
            txtnit.Text = clase.dt.Tables("listo").Rows(0)("recbod_nit")
            txtciudad.Text = clase.dt.Tables("listo").Rows(0)("recbod_ciudad") & ""
            txttelefono.Text = clase.dt.Tables("listo").Rows(0)("recbod_telefono") & ""
            txtentrega.Text = clase.dt.Tables("listo").Rows(0)("recbod_entrega") & ""
            txtrecibe.Text = clase.dt.Tables("listo").Rows(0)("recbod_recibe") & ""
            txtreferencia.Text = clase.dt.Tables("listo").Rows(0)("recbod_referencia")
            txttotal.Text = FormatCurrency(calcular_total_recibo(frm_mantenimiento_recibo_bodega.DataGridView2.Item(0, frm_mantenimiento_recibo_bodega.DataGridView2.CurrentCell.RowIndex).Value))
            clase.consultar1("SELECT detrec_referencia, detrec_descripcion, detrec_cantidad, detrec_dcto, detrec_precio, FORMAT(detrec_cantidad * (detrec_precio -   detrec_precio*(detrec_dcto/100)), 'Currency') As total  FROM detrecibo_bodegas WHERE (detrec_codrecibo =" & frm_mantenimiento_recibo_bodega.DataGridView2.Item(0, frm_mantenimiento_recibo_bodega.DataGridView2.CurrentCell.RowIndex).Value & ")", "con")
            '  MsgBox(clase.dt1.Tables("con").Rows.Count)
            Dim a As Short
            For a = 0 To clase.dt1.Tables("con").Rows.Count - 1
                DataGridView2.RowCount = DataGridView2.RowCount + 1
                DataGridView2.Item(0, a).Value = clase.dt1.Tables("con").Rows(a)("detrec_referencia")
                DataGridView2.Item(1, a).Value = clase.dt1.Tables("con").Rows(a)("detrec_descripcion")
                DataGridView2.Item(2, a).Value = clase.dt1.Tables("con").Rows(a)("detrec_cantidad")
                DataGridView2.Item(3, a).Value = clase.dt1.Tables("con").Rows(a)("detrec_dcto")
                DataGridView2.Item(4, a).Value = clase.dt1.Tables("con").Rows(a)("detrec_precio")
                DataGridView2.Item(5, a).Value = clase.dt1.Tables("con").Rows(a)("total")
            Next
            With DataGridView2
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            'If clase.dt1.Tables("con").Rows.Count > 0 Then
            '    '     MsgBox("hola")
            '    DataGridView2.DataSource = clase.dt1.Tables("con")
            '    preparar_columnas()
            'Else
            '    DataGridView2.DataSource = Nothing
            '    DataGridView2.ColumnCount = 6
            '    preparar_columnas()
            'End If

        End If
    End Sub

    Private Sub preparar_columnas()
        With DataGridView2
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Cant"
            .Columns(3).HeaderText = "Dcto"
            .Columns(4).HeaderText = "Precio"
            .Columns(5).HeaderText = "Total"
            .Columns(0).Width = 100
            .Columns(1).Width = 200
            .Columns(2).Width = 50
            .Columns(3).Width = 50
            .Columns(4).Width = 80
            .Columns(5).Width = 80
           
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Function calcular_total_recibo(codrecibo As Short) As Double
        clase.consultar2("SELECT SUM(detrec_cantidad * (detrec_precio -   detrec_precio*(detrec_dcto/100))) AS total FROM detrecibo_bodegas WHERE (detrec_codrecibo =" & codrecibo & ") GROUP BY detrec_codrecibo", "total")
        If clase.dt2.Tables("total").Rows.Count > 0 Then
            Return clase.dt2.Tables("total").Rows(0)("total")
        Else
            Return 0
        End If
    End Function

    
    Private Sub txttotal_GotFocus(sender As Object, e As EventArgs) Handles txttotal.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub btneliminar_Click(sender As Object, e As EventArgs) Handles btneliminar.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            If DataGridView2.RowCount > 2 Then
                Dim v As String = MessageBox.Show("¿Desea borrar la fila seleccionada?", "BORRAR FILA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If v = 6 Then
                    DataGridView2.Rows.Remove(DataGridView2.Rows(DataGridView2.CurrentCell.RowIndex))
                End If
            Else
                MessageBox.Show("No puede eliminar todos los registros de la orden.", "NO SE PUEDE ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub btnguardar_Click(sender As Object, e As EventArgs) Handles btnguardar.Click
        If clase.validar_cajas_text(txtproveedor, "Proveedor") = False Then Exit Sub
        If clase.validar_cajas_text(txtentrega, "Entrega") = False Then Exit Sub
        If clase.validar_cajas_text(txtrecibe, "Recibe") = False Then Exit Sub
        If clase.validar_cajas_text(txtreferencia, "Observaciones") = False Then Exit Sub
        Dim v As String = MessageBox.Show("¿Desea guardar los cambios efectuados en la orden actual?", "GUARDAR CAMBIO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            clase.actualizar("UPDATE cabrecibos_bodegas SET recbod_proveedor = " & comprobar_nulidad(txtproveedor.Text.ToUpper()) & ", recbod_nit = " & comprobar_nulidad(txtnit.Text.ToUpper()) & ", recbod_ciudad = " & comprobar_nulidad(txtciudad.Text.ToUpper()) & ", recbod_telefono = " & comprobar_nulidad(txttelefono.Text.ToUpper()) & ", recbod_entrega = '" & txtentrega.Text.ToUpper() & "', recbod_recibe = '" & txtrecibe.Text.ToUpper() & "' WHERE recbod_codigo = " & frm_mantenimiento_recibo_bodega.DataGridView2.Item(0, frm_mantenimiento_recibo_bodega.DataGridView2.CurrentCell.RowIndex).Value & "")
            clase.borradoautomatico("DELETE from detrecibo_bodegas WHERE detrec_codrecibo = " & frm_mantenimiento_recibo_bodega.DataGridView2.Item(0, frm_mantenimiento_recibo_bodega.DataGridView2.CurrentCell.RowIndex).Value & "")
            Dim a As Short
            For a = 0 To DataGridView2.RowCount - 2
                clase.agregar_registro("INSERT INTO `detrecibo_bodegas`(`detrec_codigo`,`detrec_codrecibo`,`detrec_referencia`,`detrec_descripcion`,`detrec_cantidad`,`detrec_dcto`,`detrec_precio`) VALUES ( NULL,'" & frm_mantenimiento_recibo_bodega.DataGridView2.Item(0, frm_mantenimiento_recibo_bodega.DataGridView2.CurrentCell.RowIndex).Value & "','" & DataGridView2.Item(0, a).Value & "','" & DataGridView2.Item(1, a).Value & "','" & DataGridView2.Item(2, a).Value & "','" & DataGridView2.Item(3, a).Value & "','" & DataGridView2.Item(4, a).Value & "')")
            Next
            Me.Close()
        End If
    End Sub

    Private Sub validar_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim columna As Integer = DataGridView2.CurrentCell.ColumnIndex
        If (columna = 2) Or (columna = 3) Or (columna = 4) Then
            ' Obtener caracter   
            Dim caracter As Char = e.KeyChar
            ' comprobar si el caracter es un número o el retroceso   
            If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
                'Me.Text = e.KeyChar   
                e.KeyChar = Chr(0)
            End If
        End If
    End Sub

    Private Sub DataGridView2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit
        Dim columna As Integer = DataGridView2.CurrentCell.ColumnIndex
        If columna = 0 Or columna = 1 Then
            DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value = UCase(DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value)
            DataGridView2.Item(1, DataGridView2.CurrentCell.RowIndex).Value = UCase(DataGridView2.Item(1, DataGridView2.CurrentCell.RowIndex).Value)
        End If
        If columna = 2 Or columna = 3 Or columna = 4 Then
            If DataGridView2.Item(2, DataGridView2.CurrentCell.RowIndex).Value.ToString() <> "" And DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value.ToString() <> "" And DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value.ToString() <> "" Then
                DataGridView2.Item(5, DataGridView2.CurrentCell.RowIndex).Value = DataGridView2.Item(2, DataGridView2.CurrentCell.RowIndex).Value * (DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value - (DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value * (DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value / 100)))
                If columna = 4 Or columna = 2 Or columna = 3 Then
                    'DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value = FormatCurrency(DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value)
                    DataGridView2.Item(5, DataGridView2.CurrentCell.RowIndex).Value = FormatCurrency(DataGridView2.Item(5, DataGridView2.CurrentCell.RowIndex).Value)
                End If
            End If
            If columna = 2 Or columna = 3 Then
                If DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value.ToString() = "" Then
                    DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value = "0"
                End If
            End If
            calcular_total()
        End If
        'If columna = 2 Then
        '    If DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value = "" Then
        '        DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value = "0"
        '    End If
        'End If
    End Sub

    Private Sub calcular_total()
        Dim a As Short
        Dim sum As Double = 0
        For a = 0 To DataGridView2.RowCount - 2
            sum = sum + DataGridView2.Item(5, a).Value
        Next
        txttotal.Text = FormatCurrency(sum)
    End Sub
End Class