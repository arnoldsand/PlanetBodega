Public Class frm_nuevo_recibo
    Dim clase As New class_library

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
            If DataGridView2.Item(2, DataGridView2.CurrentCell.RowIndex).Value <> "" And DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value <> "" And DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value <> "" Then
                DataGridView2.Item(5, DataGridView2.CurrentCell.RowIndex).Value = DataGridView2.Item(2, DataGridView2.CurrentCell.RowIndex).Value * (DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value - (DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value * (DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value / 100)))
                If columna = 4 Or columna = 2 Or columna = 3 Then
                    'DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value = FormatCurrency(DataGridView2.Item(4, DataGridView2.CurrentCell.RowIndex).Value)
                    DataGridView2.Item(5, DataGridView2.CurrentCell.RowIndex).Value = FormatCurrency(DataGridView2.Item(5, DataGridView2.CurrentCell.RowIndex).Value)
                End If
            End If
            If columna = 2 Or columna = 3 Then
                If DataGridView2.Item(3, DataGridView2.CurrentCell.RowIndex).Value = "" Then
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

    
    'Private Sub DataGridView2_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEnter
    '    System.Windows.Forms.SendKeys.Send("{TAB}")
    'End Sub

    Private Sub DataGridView2_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView2.EditingControlShowing

        ' obtener indice de la columna   
        Dim columna As Short = DataGridView2.CurrentCell.ColumnIndex

        ' comprobar si la celda en edición corresponde a la columna 6 
        '    If (columna = 2) Or (columna = 4) Or (columna = 3) Then
        'MsgBox("assa")
        ' referencia a la celda   
        Dim validar As TextBox = CType(e.Control, TextBox)
        ' agregar el controlador de eventos para el KeyPress   
        AddHandler validar.KeyPress, AddressOf validar_Keypress
        ' End If

    End Sub

    Private Sub frm_nuevo_recibo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView2.Columns(5).ReadOnly = True
        DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        txttotal.Text = FormatCurrency(0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btnaceptar.Click
        If clase.validar_cajas_text(txtproveedor, "Proveedor") = False Then Exit Sub
        If clase.validar_cajas_text(txtentrega, "Entrega") = False Then Exit Sub
        If clase.validar_cajas_text(txtrecibe, "Recibe") = False Then Exit Sub
        txtproveedor.Enabled = False
        txtproveedor.Text = UCase(txtproveedor.Text)
        txtnit.Enabled = False
        txtnit.Text = UCase(txtnit.Text)
        txtciudad.Enabled = False
        txtciudad.Text = UCase(txtciudad.Text)
        txttelefono.Enabled = False
        txttelefono.Text = UCase(txttelefono.Text)
        txtentrega.Enabled = False
        txtentrega.Text = UCase(txtentrega.Text)
        txtrecibe.Enabled = False
        DataGridView2.Enabled = True
        btnaceptar.Enabled = False
        btneliminar.Enabled = True
        btndeshacer.Enabled = True
        btnguardar.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btneliminar.Click
        If DataGridView2.RowCount = 1 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            Dim v As Short = MessageBox.Show("¿Desea eliminar la fila seleccionada?", "ELIMINAR FILA", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If v = 6 Then
                DataGridView2.Rows.Remove(DataGridView2.Rows(DataGridView2.CurrentRow.Index))
                calcular_total()
            End If
        End If
    End Sub

    Private Sub txttotal_GotFocus(sender As Object, e As EventArgs) Handles txttotal.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles btndeshacer.Click
        Dim v As String = MessageBox.Show("¿Desea deshacer la orden actual?", "DESHACER ORDEN ACTUAL", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If v = 6 Then
            DataGridView2.Enabled = False
            DataGridView2.Rows.Clear()
            txttotal.Text = FormatCurrency(0)
            txtproveedor.Text = ""
            txtnit.Text = ""
            txtciudad.Text = ""
            txttelefono.Text = ""
            txtentrega.Text = ""
            txtrecibe.Text = ""
            txtproveedor.Enabled = True
            txtnit.Enabled = True
            txtciudad.Enabled = True
            txttelefono.Enabled = True
            txtrecibe.Enabled = True
            txtentrega.Enabled = True
            txtproveedor.Focus()
            btnaceptar.Enabled = True
            btneliminar.Enabled = False
            btndeshacer.Enabled = False
            btnguardar.Enabled = False
        End If
    End Sub

    Private Sub btnguardar_Click(sender As Object, e As EventArgs) Handles btnguardar.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If clase.validar_cajas_text(txtreferencia, "Observaciones") = False Then Exit Sub
        Dim a As Short
        Dim ind As Boolean = False
        For a = 0 To DataGridView2.RowCount - 2
            If DataGridView2.Item(0, a).Value = "" Or DataGridView2.Item(1, a).Value = "" Or DataGridView2.Item(2, a).Value = "" Or DataGridView2.Item(3, a).Value = "" Or DataGridView2.Item(4, a).Value = "" Then
                MessageBox.Show("Debe completar los campos vacios.", "COMPLETAR CAMPOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DataGridView2.CurrentCell = DataGridView2.Item(0, a)
                ind = True
            End If
        Next
        If ind = False Then
            Dim v As String = MessageBox.Show("¿Desea guardar la orden creada?", "GUARDAR ORDEN", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If v = 6 Then
                Dim fecha As Date = Now
                clase.agregar_registro("INSERT INTO `cabrecibos_bodegas`(`recbod_codigo`,`recbod_fecha`,`recbod_hora`,`recbod_proveedor`,`recbod_nit`,`recbod_entrega`,`recbod_recibe`,`recbod_referencia`,`recbod_ciudad`,`recbod_telefono`) VALUES ( NULL,'" & fecha.ToString("yyyy-MM-dd") & "','" & fecha.ToString("HH:mm:ss") & "','" & txtproveedor.Text & "'," & comprobar_nulidad(txtnit.Text) & "," & comprobar_nulidad(txtentrega.Text) & "," & comprobar_nulidad(txtrecibe.Text) & "," & comprobar_nulidad(txtciudad.Text) & "," & comprobar_nulidad(txtreferencia.Text) & "," & comprobar_nulidad(txttelefono.Text) & ")")
                clase.consultar("select max(recbod_codigo) as maximo from cabrecibos_bodegas", "maximo")
                Dim consecutivo As Short = clase.dt.Tables("maximo").Rows(0)("maximo")
                Dim i As Short
                For i = 0 To DataGridView2.RowCount - 2
                    clase.agregar_registro("INSERT INTO `detrecibo_bodegas`(`detrec_codigo`,`detrec_codrecibo`,`detrec_referencia`,`detrec_descripcion`,`detrec_cantidad`,`detrec_dcto`,`detrec_precio`) VALUES ( NULL,'" & consecutivo & "','" & DataGridView2.Item(0, i).Value & "','" & DataGridView2.Item(1, i).Value & "','" & DataGridView2.Item(2, i).Value & "','" & Val(DataGridView2.Item(3, i).Value) & "','" & DataGridView2.Item(4, i).Value & "')")
                Next
                Me.Close()
            End If
        End If
    End Sub

    Sub imprimir_hoja_recibo(codigorecibo As Short)
        clase.consultar("select* from cabrecibos_bodega where recbod_codigo = " & codigorecibo & "", "rec")
        If clase.dt.Tables("rec").Rows.Count > 0 Then
            Dim m_excel As Microsoft.Office.Interop.Excel.Application
            m_excel = CreateObject("Excel.Application")
            m_excel.Workbooks.Open("C:\Data\actaentrega.xls")
            m_excel.Visible = True
            'm_excel.Worksheets("Hoja1").cells(1, 1).value = 
        End If
    End Sub
End Class