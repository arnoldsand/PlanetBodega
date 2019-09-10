Public Class frm_colores_tallas
    Dim clase As class_library = New class_library
    Private Sub frm_colores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With DataGridView1
            .RowHeadersWidth = 20
            .Columns(0).Width = 180
            .Columns(1).Width = 100
        End With
        If ind_creacion_colores = True Then
            DataGridView1.RowCount = colores.Length
            Dim c As Integer
            For c = 0 To colores.Length - 1
                DataGridView1.Item(0, c).Value = colores(c)
                DataGridView1.Item(1, c).Value = arraytallas(c)
            Next
        End If
        clase.llenar_combo_datagrid(cmbcolores, "select* from colores order by colornombre asc", "colornombre", "cod_color")
        clase.llenar_combo_datagrid(tallas, "select* from tallas order by nombretalla asc", "nombretalla", "codigo_talla")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Function hallar_codigo_color(ByVal color As String) As Integer
        clase.consultar("select* from colores where colornombre = '" & color & "'", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            Return clase.dt.Tables("tabla").Rows(0)("cod_color")
        End If
    End Function

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        ' obtener indice de la columna   
        Dim columna As Integer = DataGridView1.CurrentCell.ColumnIndex
        ' comprobar si la celda en edición corresponde a la columna 6 
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

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim cant_filas_a_restar As Short
        If comprobar_existencia_fila_nuevoregistro(DataGridView1) Then
            cant_filas_a_restar = 2

        Else
            cant_filas_a_restar = 1

        End If

        Dim lista(2, DataGridView1.RowCount - cant_filas_a_restar) As Integer
        Dim lista1(2, DataGridView1.RowCount - cant_filas_a_restar) As Integer
        Dim a, b, c As Integer
        For a = 0 To DataGridView1.RowCount - cant_filas_a_restar
            For c = 0 To 1
                lista(c, a) = DataGridView1.Item(c, a).Value
                lista1(c, a) = DataGridView1.Item(c, a).Value
            Next
        Next
        For a = 0 To DataGridView1.RowCount - cant_filas_a_restar
            For b = 0 To DataGridView1.RowCount - cant_filas_a_restar
                If (lista(0, a) = lista1(0, b)) And (lista(1, a) = lista1(1, b)) And (a <> b) Then
                    MessageBox.Show("Hay uno o más combinaciones de colores y tallas repetidos, verifique y vuelvalo a intentar.", "COLORES DUPLICADOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next
        Next
        ReDim colores(DataGridView1.RowCount - cant_filas_a_restar)
        ReDim arraytallas(DataGridView1.RowCount - cant_filas_a_restar)
        For a = 0 To DataGridView1.RowCount - cant_filas_a_restar
            colores(a) = DataGridView1.Item(0, a).Value
            arraytallas(a) = DataGridView1.Item(1, a).Value
        Next
        Me.Close()
        ind_creacion_colores = True
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        crear_colores.ShowDialog()
        crear_colores.Dispose()
        clase.llenar_combo_datagrid(cmbcolores, "select* from colores order by colornombre asc", "colornombre", "cod_color")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        crear_tallas.ShowDialog()
        crear_colores.Dispose()
        clase.llenar_combo_datagrid(tallas, "select* from tallas order by nombretalla asc", "nombretalla", "codigo_talla")
    End Sub
End Class