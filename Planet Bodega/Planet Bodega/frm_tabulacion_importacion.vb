Public Class frm_tabulacion_importacion
    Dim clase As New class_library
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frm_listado_importacion.ShowDialog()
        frm_listado_importacion.Dispose()
    End Sub

    Sub procedimiento_para_llenar_grilla_cajas()
        dtgridcajas.Columns.Clear()
        llenar_listado()
    End Sub

    Private Sub frm_tabulacion_importacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgridcajas.ColumnCount = 5
        preparar_grilla_caja()

    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Public Sub llenar_listado_referencias(ref As String, marca As String, caja As String, proveedor As String, subpartida As String)
        'MsgBox("SELECT detalle_importacion_detcajas.detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_descripcion AS Descripcion, detalle_importacion_detcajas.detcab_marca AS Marca, detalle_importacion_detcajas.detcab_composicion AS Composicion, detalle_importacion_detcajas.detcab_cantidad AS Cant, detalle_importacion_detcajas.detcab_unimedida AS Medida, detalle_importacion_detcajas.detcab_codigocaja AS Caja, proveedores.prv_nombre AS Proveedor, detalle_importacion_cabcajas.det_fecharecepcion AS Recepcion, detalle_importacion_cabcajas.det_hoja AS Hoja FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detalle_importacion_detcajas.detcab_referencia LIKE '" & ref & "%' AND detalle_importacion_detcajas.detcab_marca LIKE '" & marca & "%' AND detalle_importacion_detcajas.detcab_codigocaja LIKE '" & caja & "%' AND proveedores.prv_codigo LIKE '" & proveedor & "%' AND detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & ")")
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia AS Referencia, detalle_importacion_detcajas.detcab_descripcion AS Descripcion, detalle_importacion_detcajas.detcab_marca AS Marca, detalle_importacion_detcajas.detcab_composicion AS Composicion, detsalidas_mercancia.det_cant AS Cant, detalle_importacion_detcajas.detcab_unimedida AS Medida, detalle_importacion_detcajas.detcab_codigocaja AS Caja, proveedores.prv_codigoasignado AS Proveedor, detalle_importacion_cabcajas.det_fecharecepcion AS Recepcion, detalle_importacion_detcajas.detcab_costodolares AS Costo_USD, detalle_importacion_detcajas.detcab_costopesos_x_pieza AS Costo_COP, detalle_importacion_detcajas.detcab_subpartida AS Subpartida, detalle_importacion_detcajas.detcab_registrodian AS Registro_DIAN, detalle_importacion_detcajas.detcab_fecharegistrodian AS Fecha_Registro, detalle_importacion_cabcajas.det_hoja AS Hoja, detalle_importacion_detcajas.detcab_origen AS Origen FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) INNER JOIN detsalidas_mercancia ON (detalle_importacion_detcajas.detcab_coditem = detcab_coditem) WHERE (detalle_importacion_detcajas.detcab_referencia LIKE '" & ref & "%' AND detalle_importacion_detcajas.detcab_marca LIKE '" & marca & "%' AND detalle_importacion_detcajas.detcab_codigocaja LIKE '" & caja & "%' AND proveedores.prv_codigo LIKE '" & proveedor & "%' AND detalle_importacion_detcajas.detcab_subpartida like '" & subpartida & "%' AND detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & ")", "resultados_tab")
        If clase.dt.Tables("resultados_tab").Rows.Count > 0 Then
            datagridreferencia.DataSource = clase.dt.Tables("resultados_tab")
            With datagridreferencia
                .Columns(0).Width = 100
                .Columns(1).Width = 150
                .Columns(2).Width = 100
                .Columns(3).Width = 180
                .Columns(4).Width = 50
                .Columns(5).Width = 80
                .Columns(6).Width = 70
                .Columns(7).Width = 80
                .Columns(8).Width = 100
                .Columns(9).Width = 80
                .Columns(10).Width = 80
                .Columns(12).Width = 150
                .Columns(13).Width = 100
            End With
        Else
            datagridreferencia.DataSource = Nothing
        End If
    End Sub

    Private Sub llenar_listado()
        clase.consultar("SELECT detalle_importacion_cabcajas.det_caja, proveedores.prv_codigoasignado, detalle_importacion_cabcajas.det_peso, detalle_importacion_cabcajas.det_hoja, detalle_importacion_cabcajas.det_fecharecepcion FROM detalle_importacion_cabcajas INNER JOIN proveedores ON (detalle_importacion_cabcajas.det_codigoproveedor = proveedores.prv_codigo) WHERE (detalle_importacion_cabcajas.det_codigoimportacion = " & cod_importacion & ") order by detalle_importacion_cabcajas.det_caja asc", "tblresult")
        If clase.dt.Tables("tblresult").Rows.Count > 0 Then
            dtgridcajas.DataSource = clase.dt.Tables("tblresult")
            preparar_grilla_caja()
        End If
    End Sub

    Private Sub preparar_grilla_caja()
        With dtgridcajas
            .Columns(0).HeaderText = "Cod Caja"
            .Columns(1).HeaderText = "Proveedor"
            .Columns(2).HeaderText = "Peso"
            .Columns(3).HeaderText = "Hoja"
            .Columns(4).HeaderText = "Fecha Recepción"
            .Columns(0).Width = 100
            .Columns(1).Width = 150
            .Columns(2).Width = 110
            .Columns(3).Width = 50
            .Columns(4).Width = 150
        End With
    End Sub

   

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If clase.validar_combobox(ComboBox1, "Criterio de Filtrado") = False Then Exit Sub
        frm_filtrado.ShowDialog()
        frm_filtrado.Dispose()
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        TextBox1.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        clase.consultar("SELECT detalle_importacion_detcajas.* FROM detalle_importacion_detcajas INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN detsalidas_mercancia ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detalle_importacion_cabcajas.det_codigoimportacion =" & cod_importacion & ")", "borrar")
        If clase.dt.Tables("borrar").Rows.Count > 0 Then
            MessageBox.Show("No se puede eliminar los datos de la importación ya que hay movimientos generados apartir de ella.", "NO SE PUEDE ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim v As String = MessageBox.Show("¿Desea eliminar los datos asociados a esta importación?, Recuerde que no podrá deshacer esta acción después.", "ELIMINAR DATOS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                clase.borradoautomatico("DELETE FROM detalle_importacion_cabcajas WHERE det_codigoimportacion = " & cod_importacion & "")
                dtgridcajas.DataSource = Nothing
                dtgridcajas.ColumnCount = 5
                preparar_grilla_caja()
                MessageBox.Show("Datos eliminados exitosamente.", "OPERACIÓN EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub
End Class