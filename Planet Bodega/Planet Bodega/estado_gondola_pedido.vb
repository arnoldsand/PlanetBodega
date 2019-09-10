Public Class estado_gondola_pedido
    Dim clase As New class_library
    Dim cantidad As Integer
    Dim pedido As Integer
    Private Sub estado_gondola_pedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tot As Integer = 0
        pedido = frm_salida_de_mercancia_desde_hand_held.TextBox1.Text
        cantidad = frm_salida_de_mercancia_desde_hand_held.DataGridView1.Item(5, frm_salida_de_mercancia_desde_hand_held.DataGridView1.CurrentCell.RowIndex).Value
        TextBox1.Text = frm_salida_de_mercancia_desde_hand_held.DataGridView1.Item(0, frm_salida_de_mercancia_desde_hand_held.DataGridView1.CurrentCell.RowIndex).Value
        TextBox2.Text = frm_salida_de_mercancia_desde_hand_held.DataGridView1.Item(1, frm_salida_de_mercancia_desde_hand_held.DataGridView1.CurrentCell.RowIndex).Value
        With DataGridView1
            .ColumnCount = 4
            .Columns(0).Width = 140
            .Columns(1).Width = 60
            .Columns(2).Width = 60
            .Columns(0).HeaderText = "Bodega"
            .Columns(1).HeaderText = "Góndola"
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).HeaderText = "Cant"
            .Columns(3).Visible = False
        End With
        clase.consultar("SELECT bodegas.bod_nombre, patron_despacho.desp_bodega, patron_despacho.desp_gondola, patron_despacho.desp_cant FROM bodegas INNER JOIN patron_despacho ON (bodegas.bod_codigo = patron_despacho.desp_bodega) WHERE (patron_despacho.desp_pedido =" & pedido & " AND patron_despacho.desp_articulo =" & TextBox1.Text & ")", "estado")
        If clase.dt.Tables("estado").Rows.Count > 0 Then
            Dim x As Integer
            For x = 0 To clase.dt.Tables("estado").Rows.Count - 1
                DataGridView1.RowCount = DataGridView1.RowCount + 1
                DataGridView1.Item(0, x).Value = clase.dt.Tables("estado").Rows(x)("bod_nombre")
                DataGridView1.Item(0, x).ReadOnly = True
                DataGridView1.Item(1, x).Value = clase.dt.Tables("estado").Rows(x)("desp_gondola")
                DataGridView1.Item(1, x).ReadOnly = True
                DataGridView1.Item(2, x).Value = clase.dt.Tables("estado").Rows(x)("desp_cant")
                DataGridView1.Item(3, x).Value = clase.dt.Tables("estado").Rows(x)("desp_bodega")
            Next
        End If
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        ' obtener indice de la columna   
        Dim columna As Integer = DataGridView1.CurrentCell.ColumnIndex
        ' comprobar si la celda en edición corresponde a la columna 12
        If (columna = 2) Then
            ' referencia a la celda   
            Dim validar As TextBox = CType(e.Control, TextBox)
            ' agregar el controlador de eventos para el KeyPress   
            AddHandler validar.KeyPress, AddressOf validar_Keypress
        End If
    End Sub

    Private Sub validar_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        ' Obtener caracter   
        Dim caracter As Char = e.KeyChar
        ' comprobar si el caracter es un número o el retroceso   
        If Not Char.IsNumber(caracter) And (caracter = ChrW(Keys.Back)) = False Then
            'Me.Text = e.KeyChar   
            e.KeyChar = Chr(0)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim x As Integer
        Dim cd As Integer = 0
        For x = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(2, x).Value = "" Then
                cd = cd + 1
            End If
        Next
        If cd = DataGridView1.RowCount Then
            MessageBox.Show("Debe escribir valores validos para realizar ajustes al pedido.", "AJUSTES", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim can As Integer = 0
        For x = 0 To DataGridView1.RowCount - 1
            can = can + DataGridView1(2, x).Value
        Next
        If can <> cantidad Then
            MessageBox.Show("La suma de las existencias en góndolas del pedido debe ser igual a la cantidad global capturada en el hand held.", "CANTIDADES EN GÓNDOLAS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        For x = 0 To DataGridView1.RowCount - 1
            clase.actualizar("update patron_despacho set desp_cant = " & Val(DataGridView1.Item(2, x).Value) & " where desp_pedido = " & pedido & " and desp_bodega = " & DataGridView1(3, x).Value & " and desp_gondola = '" & DataGridView1(1, x).Value & "' and desp_articulo = " & TextBox1.Text & "")
        Next
        Me.Close()
    End Sub
End Class