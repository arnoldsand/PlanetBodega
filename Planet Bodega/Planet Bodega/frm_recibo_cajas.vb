Public Class frm_recibo_cajas
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_recibo_cajas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Todas las Cajas"
        dtgridcajas.ColumnCount = 5
        preparar_grilla_caja()
    End Sub

    Private Sub preparar_grilla_caja()
        With dtgridcajas
            .Columns(0).HeaderText = "Caja"
            .Columns(1).HeaderText = "Proveedor"
            .Columns(2).HeaderText = "Peso"
            .Columns(3).HeaderText = "Hoja"
            .Columns(4).HeaderText = "Recepción"
            .Columns(0).Width = 100
            .Columns(1).Width = 250
            .Columns(2).Width = 80
            .Columns(3).Width = 80
            .Columns(4).Width = 150
        End With
    End Sub


    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If clase.validar_cajas_text(textBox2, "Codigo Caja") = False Then Exit Sub
        clase.consultar("select* from detalle_importacion_cabcajas where det_caja = '" & textBox2.Text & "'", "table")
        If clase.dt.Tables("table").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("table").Rows(0)("det_fecharecepcion")) = False Then
                MessageBox.Show("La caja con el codigo escrito ya fue recibida.", "CAJA YA RECIBIDA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                textBox2.Text = ""
                textBox2.Focus()
                Exit Sub
            End If
            importacion = clase.dt.Tables("table").Rows(0)("det_codigoimportacion")
            TextBox1.Text = buscar_nombre_importacion(importacion)
            Dim fecha As String = Now.ToString("yyyy-MM-dd")
            clase.actualizar("update detalle_importacion_cabcajas SET det_fecharecepcion = '" & fecha & "' WHERE det_caja = '" & textBox2.Text & "'")
            llenar_listado_cajas(importacion, ComboBox1.Text)
            textBox2.Text = ""
            textBox2.Focus()
        Else
            MessageBox.Show("No se encontró ninguna caja con el codigo especificado.", "CODIGO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            textBox2.Text = ""
            textBox2.Focus()
        End If
    End Sub

    Sub llenar_listado_cajas(imp As Integer, tipo As String)
        Dim consultasql As String = ""
        Select Case tipo
            Case "Todas las Cajas"
                consultasql = "SELECT det_caja AS Caja, prv_codigoasignado AS Proveedor, det_peso AS Peso, det_hoja AS Hoja, det_fecharecepcion AS Recepcion FROM proveedores INNER JOIN detalle_importacion_cabcajas ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & imp & ") ORDER BY det_caja ASC"
            Case "Cajas Recibidas"
                consultasql = "SELECT det_caja AS Caja, prv_codigoasignado AS Proveedor, det_peso AS Peso, det_hoja AS Hoja, det_fecharecepcion AS Recepcion FROM proveedores INNER JOIN detalle_importacion_cabcajas ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & imp & " AND detalle_importacion_cabcajas.det_fecharecepcion IS NOT NULL) ORDER BY det_caja ASC"
            Case "Cajas en Transito"
                consultasql = "SELECT det_caja AS Caja, prv_codigoasignado AS Proveedor, det_peso AS Peso, det_hoja AS Hoja, det_fecharecepcion AS Recepcion FROM proveedores INNER JOIN detalle_importacion_cabcajas ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & imp & " AND detalle_importacion_cabcajas.det_fecharecepcion IS NULL) ORDER BY det_caja ASC"
        End Select
        clase.consultar(consultasql, "tableresult")
        If clase.dt.Tables("tableresult").Rows.Count > 0 Then
            dtgridcajas.Columns.Clear()
            dtgridcajas.DataSource = clase.dt.Tables("tableresult")
            preparar_grilla_caja()
        Else
            dtgridcajas.Columns.Clear()
            dtgridcajas.DataSource = Nothing
            dtgridcajas.ColumnCount = 5
            preparar_grilla_caja()
        End If
    End Sub

    

    Function buscar_nombre_importacion(cod As Integer) As String
        clase.consultar("select* from cabimportacion where imp_codigo = " & cod & "", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            Return clase.dt.Tables("tabla").Rows(0)("imp_nombrefecha")
        Else
            Return ""
        End If
    End Function

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_listado_importacion2.ShowDialog()
        frm_listado_importacion2.Dispose()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        llenar_listado_cajas(importacion, ComboBox1.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        inicializar_cadena_access()
        conexion.Open()
        Dim cm1 As New OleDbCommand("select* from detalle_entrada where codigoarticulo <> 'N'", conexion)
        Dim c1 As New OleDbDataAdapter
        c1.SelectCommand = cm1
        Dim dt4 As New DataSet
        c1.Fill(dt4, "transferencia")
        conexion.Close()
        Dim ind As Boolean = False
        If dt4.Tables("transferencia").Rows.Count > 0 Then
            Dim v As String = MessageBox.Show("Se encontraron " & dt4.Tables("transferencia").Rows.Count & " registros ¿Desea importarlos en este momento?", "IMPORTAR DATOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                Dim a As Short
                For a = 0 To dt4.Tables("transferencia").Rows.Count - 1
                    clase.consultar("SELECT det_codigoimportacion, det_caja FROM detalle_importacion_cabcajas WHERE (det_caja ='" & dt4.Tables("transferencia").Rows(a)("codigoarticulo") & "')", "cajas")
                    If clase.dt.Tables("cajas").Rows.Count > 0 Then
                        If ind = False Then
                            importacion = clase.dt.Tables("cajas").Rows(a)("det_codigoimportacion")
                            ind = True
                        End If
                        clase.actualizar("UPDATE detalle_importacion_cabcajas SET det_fecharecepcion = '" & Now.ToString("yyyy-MM-dd") & "' WHERE (det_caja ='" & dt4.Tables("transferencia").Rows(a)("codigoarticulo") & "')")
                    End If
                Next
                MessageBox.Show("Los datos se impotaron exitosamente.", "DATOS IMPORTADOS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TextBox1.Text = buscar_nombre_importacion(importacion)
                llenar_listado_cajas(importacion, ComboBox1.Text)
                Dim cm3 As New OleDbCommand("delete from detalle_entrada")
                conexion.Open()
                cm3.Connection = conexion
                cm3.ExecuteNonQuery()
                conexion.Close()
            End If
        Else
            MessageBox.Show("No se encontraron registros para importar.", "SIN DATOS PARA IMPORTAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class