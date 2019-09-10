Public Class frm_cambio_precios_puerta_puerta
    Dim clase As New class_library
    Dim costo, costo2, precio1, precio2 As Double
    Dim filaactual As Short
    Dim preciosactuales() As Double

    Private Sub frm_cambio_precios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        costo = Val(txtCosto.Text)
        costo2 = Val(txtCosto2.Text)
        precio1 = Val(txtPrecio1.Text)
        precio2 = Val(txtPrecio2.Text)
        filaactual = frm_busqueda_articulos_puerta_puerta.grdConsulta.CurrentCell.RowIndex
        preciosactuales = {txtPrecio1.Text, txtPrecio2.Text}
    End Sub

    Private Sub txtCosto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCosto.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub txtPrecio1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrecio1.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub txtPrecio2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrecio2.KeyPress
        clase.validar_numeros(e)
        clase.enter(e)
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtCosto, "Costo 1") = False Then Exit Sub
        If clase.validar_cajas_text(txtCosto2, "Costo 2") = False Then Exit Sub
        If clase.validar_cajas_text(txtPrecio1, "Precio 1") = False Then Exit Sub
        If clase.validar_cajas_text(txtPrecio2, "Precio 2") = False Then Exit Sub
        Dim fechaultimamodificacion As String = Now.ToString("yyyy-MM-dd")
        clase.consultar2("select* from articulos where ar_codigo = " & txtCodigo.Text & "", "articulo-ficha")
        Dim preciosafijar() As Double = {txtPrecio1.Text, txtPrecio2.Text}
        'VOY POR AQUI!!!!!!!!!!!!!!!!!!!!!!
        If MessageBox.Show("DESEA MODIFICAR LOS DATOS DEL PRODUCTO?", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            clase.actualizar("UPDATE articulos SET ar_migrado=FALSE, ar_costo=" & Str(txtCosto2.Text) & ", ar_costo2 = " & Str(txtCosto.Text) & ", ar_precio1=" & Val(txtPrecio1.Text) & ", ar_precio2=" & Val(txtPrecio2.Text) & ", ar_ultimamodificacion = '" & fechaultimamodificacion & "'" & precios_anteriores(preciosactuales, clase.dt2.Tables("articulo-ficha"), preciosafijar) & " WHERE ar_codigo=" & txtCodigo.Text & "")
            frm_busqueda_articulos_puerta_puerta.llenar_rejilla()
            frm_busqueda_articulos_puerta_puerta.grdConsulta.CurrentCell = frm_busqueda_articulos_puerta_puerta.grdConsulta.Item(0, filaactual)
            Me.Close()
        Else
            txtCosto.Text = costo
            txtCosto2.Text = costo2
            txtPrecio1.Text = precio1
            txtPrecio2.Text = precio2
            txtCosto.Focus()
        End If
    End Sub

    Private Sub txtCodigo_GotFocus(sender As Object, e As EventArgs) Handles txtCodigo.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub


    Private Sub txtReferencia_GotFocus(sender As Object, e As EventArgs) Handles txtReferencia.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtDescripcion_GotFocus(sender As Object, e As EventArgs) Handles txtDescripcion.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub
End Class