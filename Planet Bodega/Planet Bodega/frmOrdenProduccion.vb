Public Class frmOrdenProduccion
    Dim clase As New class_library
    Dim Consulta As String
    Public Ver, Procesado As Boolean
    Public CodImp As String
    Private Sub frmOrdenProduccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LlenarGrid()
        If grdConsulta.RowCount > 0 Then
            grdConsulta.Item(0, 0).Selected = True
            Consulta = grdConsulta(0, 0).Value.ToString
        End If
        'VARIABLE VER SE UTILIZA PARA SABER SI SE CARGA EL FORMULARIO NUEVAORDEN PARA ACTUALIZAR O PARA UNA NUEVA ORDEN
        Ver = False
    End Sub

    Public Sub LlenarGrid()
        clase.consultar("SELECT ordenproduccion.op_codigo AS ORDEN, ordenproduccion.op_fecha AS FECHA, ordenproduccion.op_hora AS HORA, ordenproduccion.op_realizadapor AS HECHO_POR, cabimportacion.imp_nombrefecha AS IMPORTACION, ordenproduccion.op_procesado AS PROCESADO, SUM(deordenprod.do_unidades) AS TOTAL FROM ordenproduccion INNER JOIN cabimportacion  ON (ordenproduccion.op_codigoimportacion = cabimportacion.imp_codigo) INNER JOIN deordenprod ON (deordenprod.do_idcaborden = ordenproduccion.op_codigo)WHERE (ordenproduccion.op_estado ='A')GROUP BY ORDEN order by op_codigo desc", "orden")
        grdConsulta.DataSource = clase.dt.Tables("orden")

        grdConsulta.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdConsulta.Columns(0).Width = 50
        grdConsulta.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdConsulta.Columns(1).Width = 80
        grdConsulta.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdConsulta.Columns(2).Width = 80
        grdConsulta.Columns(3).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        grdConsulta.Columns(3).Width = 150
        grdConsulta.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdConsulta.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdConsulta.Columns(4).Width = 150
        grdConsulta.Columns(4).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        grdConsulta.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdConsulta.Columns(5).Width = 150
        grdConsulta.Columns(5).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        grdConsulta.Columns(6).Width = 50
        grdConsulta.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdConsulta.RowHeadersWidth = 4
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Me.Enabled = False
        frmNuevaOrden.Show()
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        '  frmMenu.Show()
        Me.Close()
    End Sub

    Private Sub btnVer_Click(sender As Object, e As EventArgs) Handles btnVer.Click
        If Consulta = "" Then
            MessageBox.Show("DEBES SELECCIONAR UNA ORDEN", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Ver = True
        clase.consultar("SELECT cabimportacion.imp_codigo AS CODIMP, ordenproduccion.op_codigo AS ORDEN, ordenproduccion.op_realizadapor AS HECHO_POR, cabimportacion.imp_nombrefecha AS IMPORTACION FROM ordenproduccion INNER JOIN cabimportacion ON (ordenproduccion.op_codigoimportacion = cabimportacion.imp_codigo) WHERE (ordenproduccion.op_estado ='A' AND ordenproduccion.op_codigo='" & Consulta & "');", "encabezado")
        CodImp = clase.dt.Tables("encabezado").Rows(0)("CODIMP")

        'CARGAMOS ENCABEZADO EN FRMNUEVAORDEN
        With frmNuevaOrden
            .txtArticulo.Enabled = False
            .btnGuardar.Visible = False
            .btnActualizar.Visible = True
            .txtOrden.Text = clase.dt.Tables("encabezado").Rows(0)("ORDEN")
            .txtOperario.Text = clase.dt.Tables("encabezado").Rows(0)("HECHO_POR")
            .txtImportacion.Text = clase.dt.Tables("encabezado").Rows(0)("IMPORTACION")
        End With

        'CARGAMOS DETALLE EN FRMNUEVAORDEN
        frmNuevaOrden.PrepararColumnas()
        clase.consultar("SELECT deordenprod.do_codigo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, articulos.ar_precio1 AS PRECIO1, articulos.ar_precio2 AS PRECIO2, deordenprod.do_unidades AS CANTIDAD, deordenprod.do_patrondist AS IDPATRON FROM articulos INNER JOIN deordenprod ON (articulos.ar_codigo = deordenprod.do_codigo) INNER JOIN ordenproduccion ON (deordenprod.do_idcaborden = ordenproduccion.op_codigo)WHERE (ordenproduccion.op_codigo = '" & Consulta & "');", "detalle")
        Dim i As Integer
        Dim cod, ref, des, p1, p2, cant, idpatron, patron As String
        For i = 0 To clase.dt.Tables("detalle").Rows.Count - 1
            cod = clase.dt.Tables("detalle").Rows(i)("CODIGO")
            ref = clase.dt.Tables("detalle").Rows(i)("REFERENCIA")
            des = clase.dt.Tables("detalle").Rows(i)("DESCRIPCION")
            p1 = clase.dt.Tables("detalle").Rows(i)("PRECIO1")
            p2 = clase.dt.Tables("detalle").Rows(i)("PRECIO2")
            cant = clase.dt.Tables("detalle").Rows(i)("CANTIDAD")
            idpatron = clase.dt.Tables("detalle").Rows(i)("IDPATRON")
            clase.consultar1("SELECT pat_nombre FROM cabpatrondist WHERE pat_codigo='" & clase.dt.Tables("detalle").Rows(i)("IDPATRON") & "'", "patron")
            patron = clase.dt1.Tables("patron").Rows(0)("pat_nombre")

            frmNuevaOrden.grdArticulo.Rows.Add(cod, ref, des, p1, p2, cant, idpatron, patron)
        Next

        'SI YA ESTA PROCESADA NO PERMITE MODIFICAR CANTIDADES
        If Procesado = True Then
            frmNuevaOrden.grdArticulo.Enabled = False
        End If
        frmNuevaOrden.Show()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If Consulta = "" Then
            MessageBox.Show("DEBES SELECCIONAR UNA ORDEN", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If MessageBox.Show("ESTA SEGURO DE BORRAR ESTA ORDEN", "PLANET LOVE", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
            'BORRAMOS EL ENCABEZADO
            clase.actualizar("UPDATE ordenproduccion SET op_estado='B' WHERE op_codigo='" & Consulta & "'")

            'BORRAMOS EL DETALLE
            clase.actualizar("UPDATE deordenprod SET do_estado='B' WHERE do_idcaborden='" & Consulta & "'")

            MessageBox.Show("LA ORDEN FUE BORRADA CON EXITO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LlenarGrid()
        End If
    End Sub

    Private Sub grdConsulta_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdConsulta.CellClick
        Dim Y As Integer = grdConsulta.CurrentCell.RowIndex
        Consulta = grdConsulta(0, [Y]).Value.ToString
        Procesado = grdConsulta(5, [Y]).Value.ToString
    End Sub
End Class
