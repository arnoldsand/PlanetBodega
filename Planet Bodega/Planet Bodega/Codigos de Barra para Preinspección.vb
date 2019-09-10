Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class Codigos_de_Barra_para_Preinspección
    Dim clase As New class_library
    Dim Imprimir As New clase_ean13_codigo_de_barar
    Dim Printer As New Printer
    Dim Indicador As Boolean

    Private Sub Codigos_de_Barra_para_Preinspección_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim impresora As Printer
        cbImpresora.Items.Clear()
        For Each impresora In Printers
            cbImpresora.Items.Add(impresora.DeviceName)
        Next
        cbImpresora.Text = ultima_impresora_usada()
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        If Indicador = False Then
            clase.borradoautomatico("DELETE FROM consecutivo_codigobarra_importacion WHERE codigodebarras IS NOT NULL")
        End If
        Me.Close()
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs)
        If clase.validar_cajas_text(txtCompañia, "Codigo Compañia") = False Then Exit Sub
        If clase.validar_cajas_text(txtCantidad, "Cantidad de Codigos") = False Then Exit Sub
        clase.consultar("SELECT prv_codigoasignado FROM proveedores WHERE (prv_codigoasignado =" & txtCompañia.Text.Trim & ")", "tabla")
        If clase.dt.Tables("tabla").Rows.Count = 0 Then
            MessageBox.Show("El codigo de compañia especificado no está asociado a ninguna compañia.", "CODIGO NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCompañia.Text = ""
            txtCantidad.Text = ""
            txtCompañia.Focus()
            Exit Sub
        End If
        clase.consultar("select MAX(codigo_ultimo) as codigo_ultimo from consecutivo_codigobarra_importacion", "tblmx")
        If clase.dt.Tables("tblmx").Rows.Count > 0 Then
            Dim maximo As Long = clase.dt.Tables("tblmx").Rows(0)("codigo_ultimo")
            Dim c As Long
            For c = 1 To Val(txtCantidad.Text)
                clase.agregar_registro("insert into consecutivo_codigobarra_importacion (codigo_ultimo, codigodebarras) values('" & maximo + c & "','" & "" & txtCompañia.Text.Trim & "-" & maximo + c & "" & "')")
            Next
        End If

        clase.consultar1("select codigo_ultimo as Consecutivo, codigodebarras as Codigo from consecutivo_codigobarra_importacion WHERE codigodebarras IS NOT NULL", "print")
        With grdCodigos
            .DataSource = clase.dt1.Tables("print")
            .RowHeadersVisible = False
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With

        txtCompañia.Text = ""
        txtCantidad.Text = ""
        txtCompañia.Focus()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Indicador = True
        Dim Page, Prov As Integer
        Dim Control As Integer = 1
        Dim Cont As Integer

        For i = 0 To grdCodigos.RowCount - 1
            Prov = Mid(grdCodigos(1, i).Value, 1, 4)
            Page = VerificarPaginas(Prov)

            If Control = Prov Then
                Control = Prov
                Cont = Cont + 1
            Else
                Control = Prov
                Cont = 1
            End If

            Imprimir.ImprimirCodigoPreispeccion(grdCodigos(1, i).Value, cbImpresora.Text, Cont & " DE " & Page)
        Next
        clase.consultar("select max(codigo_ultimo) as maximo from consecutivo_codigobarra_importacion", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            Dim codultimo As Integer = clase.dt.Tables("tbl").Rows(0)("maximo")
            clase.borradoautomatico("delete from consecutivo_codigobarra_importacion where codigo_ultimo <> " & codultimo & "")
            clase.actualizar("update consecutivo_codigobarra_importacion set codigodebarras=NULL where codigo_ultimo = " & codultimo & "")
        End If
    End Sub

    Function ultima_impresora_usada() As String
        clase.consultar("select* from informacion", "impresora")
        If IsDBNull(clase.dt.Tables("impresora").Rows(0)("ultima_impresora_usada")) Then
            Return ""
        Else
            Return clase.dt.Tables("impresora").Rows(0)("ultima_impresora_usada")
        End If
    End Function

    Private Sub txtCantidad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCantidad.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub txtCompañia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCompañia.KeyPress
        clase.enter(e)
        clase.validar_numeros(e)
    End Sub

    Private Sub btnAceptar_Click_1(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtCompañia, "Codigo Compañia") = False Then Exit Sub
        If clase.validar_cajas_text(txtCantidad, "Cantidad de Codigos") = False Then Exit Sub
        clase.consultar("SELECT prv_codigoasignado FROM proveedores WHERE (prv_codigoasignado =" & txtCompañia.Text.Trim & ")", "tabla")
        If clase.dt.Tables("tabla").Rows.Count = 0 Then
            MessageBox.Show("El codigo de compañia especificado no está asociado a ninguna compañia.", "CODIGO NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCompañia.Text = ""
            txtCantidad.Text = ""
            txtCompañia.Focus()
            Exit Sub
        End If
        clase.consultar("select MAX(codigo_ultimo) as codigo_ultimo from consecutivo_codigobarra_importacion", "tblmx")
        If clase.dt.Tables("tblmx").Rows.Count > 0 Then
            Dim maximo As Long = clase.dt.Tables("tblmx").Rows(0)("codigo_ultimo")
            Dim c As Long
            For c = 1 To Val(txtCantidad.Text)
                clase.agregar_registro("insert into consecutivo_codigobarra_importacion (codigo_ultimo, codigodebarras) values('" & maximo + c & "','" & "" & txtCompañia.Text.Trim & maximo + c & "" & "')")
            Next
        End If
        clase.consultar1("select codigo_ultimo as Consecutivo, codigodebarras as Codigo from consecutivo_codigobarra_importacion WHERE codigodebarras IS NOT NULL", "print")
        With grdCodigos
            .DataSource = clase.dt1.Tables("print")
            .RowHeadersVisible = False
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        txtCompañia.Text = ""
        txtCantidad.Text = ""
        txtCompañia.Focus()
    End Sub

    Function VerificarPaginas(ByVal Codigo As Integer) As Integer
        Dim Page, ProvVerif As Integer

        For i = 0 To grdCodigos.RowCount - 1
            ProvVerif = Mid(grdCodigos(1, i).Value, 1, 4)

            If ProvVerif = Codigo Then
                Page = Page + 1
            End If
        Next
        Return Page
    End Function

    
End Class