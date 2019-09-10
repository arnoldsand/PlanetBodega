Public Class frmInforme
    Public ConsImport
    Dim clase As New class_library
    Dim Import, Ref, Prov As Boolean
    Dim i As Integer = 0
    Dim page As Integer = 0
    Dim Total As Integer = 0

    Private Sub btnImportacion_Click(sender As Object, e As EventArgs) Handles btnImportacion.Click
        frm_listado_importacion_consulta.ShowDialog()
        frm_listado_importacion_consulta.Dispose()
        btnAceptar.PerformClick()
    End Sub

    Private Sub btnProveedor_Click(sender As Object, e As EventArgs) Handles btnProveedor.Click
        frm_proveedores_consulta.ShowDialog()
        frm_proveedores_consulta.Dispose()
        btnAceptar.PerformClick()
    End Sub

    Private Sub txtReferencia_TextChanged(sender As Object, e As EventArgs) Handles txtReferencia.TextChanged
        If txtImportacion.Text <> "" Then
            Ref = True
            Import = False
            Prov = False
            LlenarGrid()
        End If
    End Sub

    Public Sub LlenarGrid()
        grdImportacion.DataSource = Nothing
        If Import = True Then
            clase.consultar("SELECT articulos.ar_codigo AS Codigo, articulos.ar_referencia AS Referencia, articulos.ar_descripcion AS Descripcion, colores.colornombre AS Color, tallas.nombretalla AS Talla, linea1.ln1_nombre AS Linea, articulos.ar_costo AS Costo, articulos.ar_precio1 AS Precio1, articulos.ar_precio2 AS Precio2, entradamercancia.com_unidades AS Cant FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) WHERE (entradamercancia.com_codigoimp='" & ConsImport & "');", "consulta")
        End If

        If Ref = True Then
            clase.consultar("SELECT articulos.ar_codigo AS Codigo, articulos.ar_referencia AS Referencia, articulos.ar_descripcion AS Descripcion, colores.colornombre AS Color, tallas.nombretalla AS Talla, linea1.ln1_nombre AS Linea, articulos.ar_costo AS Costo, articulos.ar_precio1 AS Precio1, articulos.ar_precio2 AS Precio2, entradamercancia.com_unidades AS Cant FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo)  INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) WHERE (entradamercancia.com_codigoimp='" & ConsImport & "' AND articulos.ar_referencia LIKE '%" & txtReferencia.Text & "%');", "consulta")
        End If
        If Prov = True Then
            clase.consultar("SELECT articulos.ar_codigo AS Codigo ,articulos.ar_referencia AS Referencia, articulos.ar_descripcion AS Descripcion, colores.colornombre AS Color, tallas.nombretalla AS Talla, linea1.ln1_nombre AS Linea, articulos.ar_costo AS Costo, articulos.ar_precio1 AS Precio1, articulos.ar_precio2 AS Precio2, entradamercancia.com_unidades AS Cant FROM entradamercancia INNER JOIN articulos ON (entradamercancia.com_codigoart = articulos.ar_codigo) INNER JOIN detalle_proveedores_articulos ON (detalle_proveedores_articulos.codigo_articulo = articulos.ar_codigo) INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo)  INNER JOIN colores ON (articulos.ar_color = colores.cod_color) INNER JOIN tallas ON (articulos.ar_talla = tallas.codigo_talla) WHERE (entradamercancia.com_codigoimp ='" & ConsImport & "' AND detalle_proveedores_articulos.codigo_proveedor ='" & txtProveedor.Text & "');", "consulta")
        End If

        If clase.dt.Tables("consulta").Rows.Count = 0 Then
            grdImportacion.DataSource = Nothing
            Exit Sub
        End If

        With grdImportacion
            .DataSource = clase.dt.Tables("consulta")
            .RowHeadersVisible = False

            '.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Width = 70
            .Columns(1).Width = 250
            .Columns(2).Width = 250
            .Columns(3).Width = 250
            .Columns(4).Width = 100
            .Columns(4).DefaultCellStyle.Format = "C"
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(5).Width = 150
            .Columns(5).DefaultCellStyle.Format = "C"
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(6).Width = 150
            .Columns(6).DefaultCellStyle.Format = "C"
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(7).Width = 150
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
        Import = False
        Ref = False
        Prov = False

    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtImportacion, "Importacion") = False Then Exit Sub
        If txtProveedor.Text <> "" Then
            Prov = True
        Else
            Import = True
        End If
        LlenarGrid()
    End Sub

    Private Sub frmInforme_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Prov = False
        Ref = False
        Import = False
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        SelectPrint.Document = DocToPrint
        Dim Result As DialogResult = SelectPrint.ShowDialog()

        If Result = Windows.Forms.DialogResult.OK Then
            DocToPrint.Print()
        End If
    End Sub

    Private Sub DocToPrint_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles DocToPrint.PrintPage
        Dim linesPerPage As Double = 0
        Dim count As Integer = 0
        Dim FuenteTitulo As System.Drawing.Font = New Font("Arial", 12, FontStyle.Bold)
        Dim FuenteContenido As System.Drawing.Font = New Font("Arial", 9, FontStyle.Bold)
        Dim FuenteGrid As System.Drawing.Font = New Font("Arial", 8)


        'Imprimir Encabezado
        e.Graphics.DrawString("LISTADO ARTICULOS POR IMPORTACION", FuenteTitulo, Brushes.Black, 250, 30)
        e.Graphics.DrawString("Nombre de la Importacion: " & txtImportacion.Text, FuenteContenido, Brushes.Black, 40, 60)
        If txtProveedor.Text <> "" Then
            e.Graphics.DrawString("Proveedor: " & txtProveedor.Text, FuenteContenido, Brushes.Black, 40, 80)
        End If
        Dim py As Integer = 150

        'Imprimir Detalle
        e.Graphics.DrawString("Codigo", FuenteContenido, Brushes.Black, 10, 130)
        e.Graphics.DrawString("Referencia", FuenteContenido, Brushes.Black, 90, 130)
        e.Graphics.DrawString("Descripción", FuenteContenido, Brushes.Black, 180, 130)

        e.Graphics.DrawString("Color", FuenteContenido, Brushes.Black, 280, 130)
        e.Graphics.DrawString("Talla", FuenteContenido, Brushes.Black, 350, 130)

        e.Graphics.DrawString("Linea", FuenteContenido, Brushes.Black, 440, 130)
        e.Graphics.DrawString("Costo", FuenteContenido, Brushes.Black, 520, 130)
        e.Graphics.DrawString("Precio1", FuenteContenido, Brushes.Black, 590, 130)
        e.Graphics.DrawString("Precio2", FuenteContenido, Brushes.Black, 670, 130)
        e.Graphics.DrawString("Cant", FuenteContenido, Brushes.Black, 740, 130)

        linesPerPage = (e.MarginBounds.Height / FuenteGrid.GetHeight(e.Graphics)) - 33

        While count < linesPerPage AndAlso i < grdImportacion.Rows.Count
            e.Graphics.DrawString(grdImportacion(0, i).Value.ToString, FuenteGrid, Brushes.Black, 10, py)
            e.Graphics.DrawString(grdImportacion(1, i).Value.ToString, FuenteGrid, Brushes.Black, 90, py)
            e.Graphics.DrawString(grdImportacion(2, i).Value.ToString, FuenteGrid, Brushes.Black, 180, py)

            e.Graphics.DrawString(grdImportacion(3, i).Value.ToString, FuenteGrid, Brushes.Black, 280, py)
            e.Graphics.DrawString(grdImportacion(4, i).Value.ToString, FuenteGrid, Brushes.Black, 350, py)

            e.Graphics.DrawString(grdImportacion(5, i).Value.ToString, FuenteGrid, Brushes.Black, 440, py)
            e.Graphics.DrawString(grdImportacion(6, i).Value.ToString, FuenteGrid, Brushes.Black, 520, py)
            e.Graphics.DrawString(grdImportacion(7, i).Value.ToString, FuenteGrid, Brushes.Black, 590, py)
            e.Graphics.DrawString(grdImportacion(8, i).Value.ToString, FuenteGrid, Brushes.Black, 670, py)
            e.Graphics.DrawString(grdImportacion(9, i).Value.ToString, FuenteGrid, Brushes.Black, 740, py)
            py = py + 20
            Total = Total + grdImportacion(9, i).Value.ToString

            count += 1
            i += 1
        End While

        If i < grdImportacion.Rows.Count Then
            e.HasMorePages = True
            page = page + 1
            e.Graphics.DrawString("Pagina " & page, FuenteContenido, Brushes.Black, 670, 20)
        Else
            e.Graphics.DrawString("Total:   " & Total, FuenteContenido, Brushes.Black, 700, py + 15)
            page = page + 1
            e.Graphics.DrawString("Pagina " & page, FuenteContenido, Brushes.Black, 670, 20)
            e.HasMorePages = False
            i = 0
            Total = 0
            page = 0
        End If
    End Sub

    Private Sub txtProveedor_GotFocus(sender As Object, e As EventArgs) Handles txtProveedor.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtProveedor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtProveedor.KeyPress
        clase.enter(e)
    End Sub

    Private Sub txtImportacion_GotFocus(sender As Object, e As EventArgs) Handles txtImportacion.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        frm_editar_cantidades.ShowDialog()
        frm_editar_cantidades.Dispose()
        btnAceptar_Click(Nothing, Nothing)
    End Sub
End Class
