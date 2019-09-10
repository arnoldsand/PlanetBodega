Public Class frm_reporte_de_danado
    Dim clase As New class_library
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.Close()
    End Sub

    Private Sub frm_reporte_de_danado_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnCount = 7
        preparar_columnas()
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 80
            .Columns(1).Width = 150
            .Columns(2).Width = 150
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 100
            .Columns(6).Width = 50
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Costo"
            .Columns(4).HeaderText = "Precio 1"
            .Columns(5).HeaderText = "Precio 2"
            .Columns(6).HeaderText = "Cant"
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

  

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If clase.validar_cajas_text(TextBox3, "Codigo Articulo") = False Then Exit Sub
        Dim codigo_articulo As Long
        If Len(TextBox3.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox3.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox3.Text)
        End If
        If Len(TextBox3.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox3.Text & ")", "tabla11")
            codigo_articulo = TextBox3.Text
        End If
        If Len(TextBox3.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox3.Text = ""
            TextBox3.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            Dim a As Short
            Dim indupdate As Boolean = False ' esta variable indica si el registro fue actualizado para no tener que insertar un registro nuevo
            For a = 0 To DataGridView1.RowCount - 1
                If codigo_articulo = DataGridView1.Item(0, a).Value Then
                    Dim cant As Short = DataGridView1.Item(6, a).Value
                    DataGridView1.Item(6, a).Value = cant + 1
                    DataGridView1.CurrentCell = DataGridView1.Item(0, a)
                    indupdate = True
                    Exit For
                End If
            Next
            If indupdate = False Then
                DataGridView1.RowCount = DataGridView1.RowCount + 1
                DataGridView1.Item(0, DataGridView1.RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_codigo")
                DataGridView1.Item(1, DataGridView1.RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_referencia")
                DataGridView1.Item(2, DataGridView1.RowCount - 1).Value = clase.dt.Tables("tabla11").Rows(0)("ar_descripcion")
                DataGridView1.Item(3, DataGridView1.RowCount - 1).Value = FormatCurrency(precio_costo(clase.dt.Tables("tabla11").Rows(0)("ar_codigo")))
                DataGridView1.Item(4, DataGridView1.RowCount - 1).Value = FormatCurrency(precio_venta1(clase.dt.Tables("tabla11").Rows(0)("ar_codigo")))
                DataGridView1.Item(5, DataGridView1.RowCount - 1).Value = FormatCurrency(precio_venta2(clase.dt.Tables("tabla11").Rows(0)("ar_codigo")))
                DataGridView1.Item(6, DataGridView1.RowCount - 1).Value = 1
                DataGridView1.CurrentCell = DataGridView1.Item(0, DataGridView1.RowCount - 1)
            End If
            TextBox3.Text = ""
            TextBox3.Focus()

        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox3.Text = ""
            TextBox3.Focus()
        End If
    End Sub


    Function precio_costo(articulo As Double) As Double
        clase.consultar1("select ar_costo from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt1.Tables("tabla").Rows(0)("ar_costo")
    End Function

    Function precio_venta1(articulo As Double) As Double
        clase.consultar1("select ar_precio1 from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt1.Tables("tabla").Rows(0)("ar_precio1")
    End Function

    Function precio_venta2(articulo As Double) As Double
        clase.consultar1("select ar_precio2 from articulos where ar_codigo = " & articulo & "", "tabla")
        Return clase.dt1.Tables("tabla").Rows(0)("ar_precio2")
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim v As String = MessageBox.Show("¿Desea deshacer los articulos capturados e iniciar una nueva captura?", "DESHACER", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            DataGridView1.Rows.Clear()
            TextBox3.Text = ""
            TextBox3.Focus()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frm_eliminar_danado.ShowDialog()
        frm_eliminar_danado.Dispose()
    End Sub
End Class