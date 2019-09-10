Public Class frmVerificarExistenciaFotosArticulos
    Dim clase As New class_library

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Integer
        Dim ArticulosSinFotoGrande As Integer = 0
        Dim ArticulosSinFotoPequeña As Integer = 0
        clase.consultar1("SELECT articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, articulos.ar_fechaingreso, articulos.ar_codigo, articulos.ar_foto FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) INNER JOIN detalle_proveedores_articulos ON (detalle_proveedores_articulos.codigo_articulo = articulos.ar_codigo) WHERE (sublinea2.sl2_nombre ='GOLDFILLED' AND articulos.ar_fechaingreso BETWEEN '2016-08-01' AND '2016-11-20' AND detalle_proveedores_articulos.codigo_proveedor =25)", "lista")
        If clase.dt1.Tables("lista").Rows.Count > 0 Then
            ProgressBar1.Maximum = clase.dt1.Tables("lista").Rows.Count
            For i = 0 To clase.dt1.Tables("lista").Rows.Count - 1
                Application.DoEvents()
                dtgListaArticulos.RowCount = dtgListaArticulos.RowCount + 1
                With clase.dt1.Tables("lista")
                    dtgListaArticulos.Item(0, i).Value = .Rows(i)("ar_codigo")
                    dtgListaArticulos.Item(1, i).Value = .Rows(i)("ar_referencia")
                    dtgListaArticulos.Item(2, i).Value = .Rows(i)("ar_descripcion")
                    dtgListaArticulos.Item(3, i).Value = .Rows(i)("ln1_nombre")
                    dtgListaArticulos.Item(4, i).Value = .Rows(i)("sl1_nombre")
                    dtgListaArticulos.Item(5, i).Value = .Rows(i)("sl2_nombre")
                    dtgListaArticulos.Item(6, i).Value = False
                    dtgListaArticulos.Item(7, i).Value = False
                End With
                dtgListaArticulos.CurrentCell = dtgListaArticulos.Item(0, i)
                If System.IO.File.Exists(clase.dt1.Tables("lista").Rows(i)("ar_foto")) = True Then
                    dtgListaArticulos.Item(6, i).Value = True
                    Dim cadena As String() = Split(clase.dt1.Tables("lista").Rows(i)("ar_foto"), ".")
                    If System.IO.File.Exists(cadena(0) & "mini.jpg") = True Then
                        dtgListaArticulos.Item(7, i).Value = True
                    Else
                        dtgListaArticulos.Item(7, i).Value = False
                        ArticulosSinFotoPequeña += 1
                        txtFotosPequeñasFaltantes.Text = ArticulosSinFotoPequeña
                    End If
                Else
                    dtgListaArticulos.Item(6, i).Value = False
                    ArticulosSinFotoGrande += 1
                    txtFotosGrandesFaltantes.Text = ArticulosSinFotoGrande

                    Dim cadena As String() = Split(clase.dt1.Tables("lista").Rows(i)("ar_foto"), ".")
                    If System.IO.File.Exists(cadena(0) & "mini.jpg") = True Then
                        dtgListaArticulos.Item(7, i).Value = True
                    Else
                        dtgListaArticulos.Item(7, i).Value = False
                        ArticulosSinFotoPequeña += 1
                        txtFotosPequeñasFaltantes.Text = ArticulosSinFotoPequeña
                    End If
                End If
                ProgressBar1.Increment(1)
            Next
        Else
            dtgListaArticulos = Nothing
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        clase.consultar("SELECT  articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, sublinea1.sl1_nombre, sublinea2.sl2_nombre, articulos.ar_fechaingreso, articulos.ar_foto FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) INNER JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (articulos.ar_sublinea2 = sublinea2.sl2_codigo) WHERE (sublinea2.sl2_nombre ='GOLDFILLED' AND articulos.ar_fechaingreso BETWEEN '2016-10-01' AND '2016-11-20')", "lista")
        If clase.dt.Tables("lista").Rows.Count > 0 Then
            ProgressBar1.Maximum = clase.dt.Tables("lista").Rows.Count
            Dim x As Integer
            Dim cadena As String()
            Dim indicador As Boolean
            For x = 0 To clase.dt.Tables("lista").Rows.Count - 1
                Application.DoEvents()
                indicador = False
                cadena = Split(clase.dt.Tables("lista").Rows(x)("ar_foto"), ".")
                If (System.IO.File.Exists(clase.dt.Tables("lista").Rows(x)("ar_foto")) = True) And (System.IO.File.Exists(cadena(0) & "mini.jpg") = False) Then
                    crear_foto_pequeño(clase.dt.Tables("lista").Rows(x)("ar_foto"), cadena(0) & "mini.jpg")
                    indicador = True
                End If
                If indicador = False Then
                    If (System.IO.File.Exists(clase.dt.Tables("lista").Rows(x)("ar_foto")) = False) And (System.IO.File.Exists(cadena(0) & "mini.jpg") = True) Then
                        crear_foto_grande(cadena(0) & "mini.jpg", clase.dt.Tables("lista").Rows(x)("ar_foto"))
                    End If
                End If
                ProgressBar1.Increment(1)
            Next
        End If
    End Sub

    Function GridAExcel(ByVal DGV As DataGridView) As Boolean
        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet
        Try
            exLibro = exApp.Workbooks.Add
            exHoja = exLibro.Worksheets.Add()
            ' ¿Cuantas columnas y cuantas filas?
            Dim NCol As Integer = DGV.ColumnCount
            Dim NRow As Integer = DGV.RowCount
            'recorremos todas las filas, y por cada fila todas las columnas
            'y vamos escribiendo.

            For i As Integer = 1 To NCol
                exHoja.Cells.Item(1, i) = DGV.Columns(i - 1).Name.ToString
            Next
            For Fila As Integer = 0 To NRow - 1
                For Col As Integer = 0 To NCol - 1
                    exHoja.Cells.Item(Fila + 2, Col + 1) =
                    DGV.Rows(Fila).Cells(Col).Value()
                Next
            Next
            'Titulo en negrita, Alineado
            exHoja.Rows.Item(1).Font.Bold = 1
            exHoja.Rows.Item(1).HorizontalAlignment = 3
            exHoja.Columns.AutoFit()
            'para visualizar el libro
            exApp.Application.Visible = True
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")
            Return False
        End Try
        Return True
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GridAExcel(dtgListaArticulos)
    End Sub
End Class