Public Class frm_ajustar_gondola_pedido2
    Dim clase As New class_library
    Dim bloque As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ajustar_gondola_pedido2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(ComboBox1, "select* from bodegas order by bod_codigo asc", "bod_nombre", "bod_codigo")

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        clase.validar_numeros(e)
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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If clase.validar_cajas_text(TextBox3, "Pedido") = False Then Exit Sub
        If clase.validar_combobox(ComboBox1, "Bodega") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox18, "Góndola") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox1, "Articulo") = False Then Exit Sub
        If clase.validar_cajas_text(TextBox2, "Cantidad") = False Then Exit Sub
        If verificar_existencia_gondola(UCase(TextBox18.Text), ComboBox1.SelectedValue) = False Then
            MessageBox.Show("La góndola especificada no existe.", "GÓNDOLA NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox18.Text = ""
            TextBox18.Focus()
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            MessageBox.Show("Debe escribir un valor válido para realizar ajuste.", "ESPECIFICAR CANTIDAD", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Text = ""
            TextBox2.Focus()
            Exit Sub
        End If
        Dim codigo_articulo As Long
        If Len(TextBox1.Text) = 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigobarras ='" & TextBox1.Text & "')", "tabla11")
            codigo_articulo = convertir_codigobarra_a_codigo_normal(TextBox1.Text)
        End If
        If Len(TextBox1.Text) < 13 Then
            clase.consultar("SELECT articulos.ar_referencia, articulos.ar_descripcion, colores.colornombre, tallas.nombretalla FROM colores INNER JOIN articulos ON (colores.cod_color = articulos.ar_color) INNER JOIN tallas ON (tallas.codigo_talla = articulos.ar_talla) WHERE (articulos.ar_codigo =" & TextBox1.Text & ")", "tabla11")
            codigo_articulo = TextBox1.Text
        End If
        If Len(TextBox1.Text) > 13 Then
            MessageBox.Show("El codigo digitado es invalido", "CODIGO INVALIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Text = ""
            TextBox1.Focus()
            Exit Sub
        End If
        If clase.dt.Tables("tabla11").Rows.Count > 0 Then
            clase.consultar1("SELECT desp_cant FROM patron_despacho WHERE (desp_pedido =" & TextBox3.Text & " AND desp_bodega =" & ComboBox1.SelectedValue & " AND desp_gondola ='" & UCase(TextBox18.Text) & "' AND desp_articulo =" & codigo_articulo & ")", "patronpedido")
            If clase.dt1.Tables("patronpedido").Rows.Count > 0 Then
                Dim v1 As String = MessageBox.Show("¿Desea ajustar el pedido con los parametros especificados?", "AJUSTAR ARTICULO GÓNDOLA", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If v1 = 6 Then
                    clase.actualizar("UPDATE patron_despacho SET desp_cant = " & Val(TextBox2.Text) & " WHERE (desp_pedido =" & TextBox3.Text & " AND desp_bodega =" & ComboBox1.SelectedValue & " AND desp_gondola ='" & UCase(TextBox18.Text) & "' AND desp_articulo =" & codigo_articulo & ")")
                    restablecer()
                End If
            Else
                MessageBox.Show("El articulo con los parametros especificado no pertenece el pedido.", "ARTICULO NO PERTENECE A PEDIDO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                restablecer()
            End If
        Else
            MessageBox.Show("El código de articulo digitado no fue encontrado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Text = ""
            TextBox1.Focus()
        End If
    End Sub

    Sub recalcular_patron()
        '  clase.borradoautomatico("delete from patron_despacho where desp_pedido = " & frm_patron_despacho.TextBox18.Text & "")
        MessageBox.Show("interrupcion")
        frm_patron_despacho.procedimiento_para_grabar_patrones()
    End Sub

    Sub restablecer()
        TextBox3.Text = ""
        ComboBox1.SelectedValue = -1
        TextBox18.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Focus()
    End Sub

    Function calcular_consecutivo() As Integer
        clase.consultar1("SELECT MAX(cabaj_codigo) AS cantidad FROM cabajuste", "tabla2")
        If IsDBNull(clase.dt1.Tables("tabla2").Rows(0)("cantidad")) Then
            Return 1
        Else
            Return clase.dt1.Tables("tabla2").Rows(0)("cantidad") + 1
        End If

    End Function

    Function precio_costo(articulo As Integer) As Double
        clase.consultar("select ar_costo from articulos where ar_codigo = " & articulo & "", "costo")
        If clase.dt.Tables("costo").Rows.Count > 0 Then
            Return clase.dt.Tables("costo").Rows(0)("ar_costo")
        Else
            Return 0
        End If
    End Function

    Function precio_venta1(articulo As Integer) As Double
        clase.consultar("select ar_precio1 from articulos where ar_codigo = " & articulo & "", "venta")
        If clase.dt.Tables("venta").Rows.Count > 0 Then
            Return clase.dt.Tables("venta").Rows(0)("ar_precio1")
        Else
            Return 0
        End If
    End Function

    Function precio_venta2(articulo As Integer) As Double
        clase.consultar("select ar_precio2 from articulos where ar_codigo = " & articulo & "", "venta")
        If clase.dt.Tables("venta").Rows.Count > 0 Then
            Return clase.dt.Tables("venta").Rows(0)("ar_precio2")
        Else
            Return 0
        End If
    End Function


    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        clase.validar_numeros(e)
    End Sub
End Class