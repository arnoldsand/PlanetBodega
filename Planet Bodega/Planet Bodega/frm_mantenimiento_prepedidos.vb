Public Class frm_mantenimiento_prepedidos
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_mantenimiento_pedidos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Pendientes"
        clase.consultar("select* from tiendas order by tienda asc", "tiendas")
        If clase.dt.Tables("tiendas").Rows.Count > 0 Then
            cmbtiendas.Items.Clear()
            cmbtiendas.Items.Add("TODAS")
            Dim i As Integer
            For i = 0 To clase.dt.Tables("tiendas").Rows.Count - 1
                cmbtiendas.Items.Add(clase.dt.Tables("tiendas").Rows(i)("tienda"))
            Next
        End If
        DataGridView1.ColumnCount = 6
        preparar_columnas()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If clase.validar_combobox(cmbtiendas, "Tienda") = False Then Exit Sub
        llenar_listado()

    End Sub

    Function codigo_tienda(tienda As String) As Short
        clase.consultar("select* from tiendas where tienda = '" & tienda & "'", "result")
        Return clase.dt.Tables("result").Rows(0)("id")
    End Function

    Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 50
            .Columns(1).Width = 80
            .Columns(2).Width = 60
            .Columns(3).Width = 150
            .Columns(4).Width = 200
            .Columns(5).Width = 60
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Fecha"
            .Columns(2).HeaderText = "Hora"
            .Columns(3).HeaderText = "Tienda"
            .Columns(4).HeaderText = "Operario"
            .Columns(5).HeaderText = "Estado"
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then

            frm_solicitudes_prepedidos.ShowDialog()
            frm_solicitudes_prepedidos.Dispose()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            clase.consultar("SELECT solpro_revisado FROM solicitudes_productos WHERE (solpro_codigo =" & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & ")", "pedido")
            Select Case clase.dt.Tables("pedido").Rows(0)("solpro_revisado")
                Case False
                    clase.actualizar("UPDATE solicitudes_productos SET solpro_revisado = TRUE WHERE solpro_codigo = " & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")
                Case True
                    clase.actualizar("UPDATE solicitudes_productos SET solpro_revisado = FALSE WHERE solpro_codigo = " & DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value & "")
            End Select
            llenar_listado()
        End If
    End Sub

    Sub llenar_listado()
        Dim sql As String = ""
        If cmbtiendas.Text = "TODAS" Then
            Select Case ComboBox1.Text
                Case "Todos"
                    sql = "SELECT solicitudes_productos.solpro_codigo, solicitudes_productos.solpro_fecha, solicitudes_productos.solpro_hora, tiendas.tienda, solicitudes_productos.solpro_operario1, solicitudes_productos.solpro_revisado FROM tiendas INNER JOIN solicitudes_productos ON (tiendas.id = solicitudes_productos.solpro_tienda) ORDER BY solicitudes_productos.solpro_codigo DESC"
                Case "Revisados"
                    sql = "SELECT solicitudes_productos.solpro_codigo, solicitudes_productos.solpro_fecha, solicitudes_productos.solpro_hora, tiendas.tienda, solicitudes_productos.solpro_operario1, solicitudes_productos.solpro_revisado FROM tiendas INNER JOIN solicitudes_productos ON (tiendas.id = solicitudes_productos.solpro_tienda) WHERE (solicitudes_productos.solpro_revisado =TRUE) ORDER BY solicitudes_productos.solpro_codigo DESC"
                Case "Pendientes"
                    sql = "SELECT solicitudes_productos.solpro_codigo, solicitudes_productos.solpro_fecha, solicitudes_productos.solpro_hora, tiendas.tienda, solicitudes_productos.solpro_operario1, solicitudes_productos.solpro_revisado FROM tiendas INNER JOIN solicitudes_productos ON (tiendas.id = solicitudes_productos.solpro_tienda) WHERE (solicitudes_productos.solpro_revisado =FALSE) ORDER BY solicitudes_productos.solpro_codigo DESC"
            End Select
        Else
            Select Case ComboBox1.Text
                Case "Todos"
                    sql = "SELECT solicitudes_productos.solpro_codigo, solicitudes_productos.solpro_fecha, solicitudes_productos.solpro_hora, tiendas.tienda, solicitudes_productos.solpro_operario1, solicitudes_productos.solpro_revisado FROM tiendas INNER JOIN solicitudes_productos ON (tiendas.id = solicitudes_productos.solpro_tienda) WHERE (solicitudes_productos.solpro_tienda = " & codigo_tienda(cmbtiendas.Text) & ")ORDER BY solicitudes_productos.solpro_codigo DESC"
                Case "Revisados"
                    sql = "SELECT solicitudes_productos.solpro_codigo, solicitudes_productos.solpro_fecha, solicitudes_productos.solpro_hora, tiendas.tienda, solicitudes_productos.solpro_operario1, solicitudes_productos.solpro_revisado FROM tiendas INNER JOIN solicitudes_productos ON (tiendas.id = solicitudes_productos.solpro_tienda) WHERE (solicitudes_productos.solpro_tienda = " & codigo_tienda(cmbtiendas.Text) & " AND solicitudes_productos.solpro_revisado =TRUE) ORDER BY solicitudes_productos.solpro_codigo DESC"
                Case "Pendientes"
                    sql = "SELECT solicitudes_productos.solpro_codigo, solicitudes_productos.solpro_fecha, solicitudes_productos.solpro_hora, tiendas.tienda, solicitudes_productos.solpro_operario1, solicitudes_productos.solpro_revisado FROM tiendas INNER JOIN solicitudes_productos ON (tiendas.id = solicitudes_productos.solpro_tienda) WHERE (solicitudes_productos.solpro_tienda = " & codigo_tienda(cmbtiendas.Text) & " AND solicitudes_productos.solpro_revisado =FALSE) ORDER BY solicitudes_productos.solpro_codigo DESC"
            End Select
        End If
        clase.consultar(sql, "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("tabla")
            preparar_columnas()
        Else
            MessageBox.Show("No se encontraron prepedidos con los parámetros especificados.", "NO SE ENCONTRARON PREPEDIDOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            Dim v As String = MessageBox.Show("¿Desea imprimir el pedido seleccionado?", "IMPRIMIR PEDIDO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                imprimir_pedido(DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value)
            End If
        End If
    End Sub

    Sub imprimir_pedido(pedido As Integer)
        clase.dt_global.Tables.Clear()
        clase.consultar_global("SELECT cabpedido.cabped_operario, cabpedido.cabped_fecha, cabpedido.cabped_tienda, detpedido.detped_cant, detpedido.detped_articulo FROM cabpedido INNER JOIN detpedido ON (cabpedido.cabped_codigo = detpedido.detped_codpedido) WHERE (cabpedido.cabped_codigo =" & pedido & ")", "pedido")
        If clase.dt_global.Tables("pedido").Rows.Count > 0 Then
            Dim a As Integer
            For a = 0 To clase.dt_global.Tables("pedido").Rows.Count - 1
                With clase.dt_global.Tables("pedido")
                    clase.agregar_registro("INSERT INTO `reporte_pedidos`(`pedido`,`articulo`,`cantidad`,`opciones`) VALUES ( '" & pedido & "','" & .Rows(a)("detped_articulo") & "','" & .Rows(a)("detped_cant") & "','" & opciones_gondolas(.Rows(a)("detped_articulo")) & "')")
                End With
            Next
            Dim m_excel As Microsoft.Office.Interop.Excel.Application
            m_excel = CreateObject("Excel.Application")
            m_excel.Workbooks.Open("C:\Data\pedido.xls")
            m_excel.Visible = False
            m_excel.Worksheets("Hoja1").cells(3, 1).value = "PEDIDO A BODEGA No: " & pedido
            clase.consultar("SELECT * FROM cabpedido WHERE (cabped_codigo =" & pedido & ")", "conspedido")
            Dim operador As String = clase.dt.Tables("conspedido").Rows(0)("cabped_operario")
            Dim fecha As Date = clase.dt.Tables("conspedido").Rows(0)("cabped_fecha")
            Dim codtienda As Integer = clase.dt.Tables("conspedido").Rows(0)("cabped_tienda")
            m_excel.Worksheets("Hoja1").cells(4, 1).value = "REMITENTE: " & nombretienda(codtienda)
            m_excel.Worksheets("Hoja1").cells(5, 1).value = "GENERADO POR: " & operador
            m_excel.Worksheets("Hoja1").cells(5, 5).value = "FECHA: " & fecha.ToString("dd/MM/yyyy")
            clase.consultar1("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, reporte_pedidos.opciones, reporte_pedidos.cantidad FROM reporte_pedidos INNER JOIN articulos  ON (reporte_pedidos.articulo = articulos.ar_codigo) INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1  ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) WHERE (reporte_pedidos.pedido =" & pedido & ") ORDER BY reporte_pedidos.opciones ASC", "excel")
            If clase.dt1.Tables("excel").Rows.Count > 0 Then
                Dim p As Integer
                For p = 0 To clase.dt1.Tables("excel").Rows.Count - 1
                    With clase.dt1.Tables("excel")
                        m_excel.Worksheets("Hoja1").cells(8 + p, 1).value = .Rows(p)("ar_codigo")
                        m_excel.Worksheets("Hoja1").cells(8 + p, 2).value = .Rows(p)("ar_referencia")
                        m_excel.Worksheets("Hoja1").cells(8 + p, 3).value = .Rows(p)("ar_descripcion")
                        m_excel.Worksheets("Hoja1").cells(8 + p, 4).value = .Rows(p)("cantidad")
                        m_excel.Worksheets("Hoja1").cells(8 + p, 5).value = .Rows(p)("opciones")
                    End With
                Next
            End If
            m_excel.Application.ActiveWorkbook.PrintOutEx()
            If Not m_excel Is Nothing Then
                m_excel.Application.ActiveWorkbook.Saved = True
                m_excel.Quit()
                m_excel.Application.Quit()
                m_excel = Nothing
            End If
            clase.borradoautomatico("delete from reporte_pedidos where pedido = " & pedido & "")
        End If
    End Sub

    Function nombretienda(codigo As Short) As String
        clase.consultar("select* from tiendas where id = " & codigo & "", "tabla2")
        Return clase.dt.Tables("tabla2").Rows(0)("tienda")
    End Function

    Function bodega(codigo As Short) As String
        clase.consultar("SELECT bod_abreviatura FROM bodegas WHERE (bod_codigo =" & codigo & ")", "tabla2")
        Return clase.dt.Tables("tabla2").Rows(0)("bod_abreviatura")
    End Function



    Function opciones_gondolas(articulo As Integer) As String '(queda como solución defintiva)  esta funcion me arroja la ubicacion(es) del articulo teniendo en cuenta la solucion temporal de la ubicaciones registradas en la tabla gondolas_articulos, en el futuro estas ubicaciones de deben deducir de la tabla inventario bodega
        clase.consultar("SELECT bodegas.bod_abreviatura, articulos_gondolas.gondola FROM bodegas INNER JOIN articulos_gondolas ON (bodegas.bod_codigo = articulos_gondolas.bodega) WHERE (articulos_gondolas.articulo =" & articulo & ")", "tabla5")
        Dim i As Integer
        Dim cadena As String = ""
        Dim ind As Boolean = False
        If clase.dt.Tables("tabla5").Rows.Count > 0 Then
            For i = 0 To clase.dt.Tables("tabla5").Rows.Count - 1   ' la cadena "cadena" debe ser construida en orden descendente es decir la ultima ubicación del articulo sera la primera
                If ind = False Then
                    cadena = clase.dt.Tables("tabla5").Rows((clase.dt.Tables("tabla5").Rows.Count - 1) - i)("bod_abreviatura") & "-" & clase.dt.Tables("tabla5").Rows((clase.dt.Tables("tabla5").Rows.Count - 1) - i)("gondola")
                    ind = True
                Else
                    cadena = cadena & ", " & clase.dt.Tables("tabla5").Rows((clase.dt.Tables("tabla5").Rows.Count - 1) - i)("bod_abreviatura") & "-" & clase.dt.Tables("tabla5").Rows((clase.dt.Tables("tabla5").Rows.Count - 1) - i)("gondola")
                End If
            Next
            Return cadena
        Else
            Return ""
        End If
    End Function
End Class