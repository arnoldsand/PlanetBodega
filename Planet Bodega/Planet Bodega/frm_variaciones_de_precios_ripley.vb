Public Class frm_variaciones_de_precios_ripley
    Dim clase As New class_library
    Private Sub frm_variaciones_de_precios_ripley_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim costoripley As Double = frm_planillas_ripley_codigos_nuevos.porcentaje_costo_ripley()
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, Format(articulos.ar_precio1 * " & Str(costoripley) & ", 'Fixed'), articulos.ar_precio1, articulos.ar_precio2 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) LEFT JOIN codigos_ripley ON (articulos.ar_codigo = codigos_ripley.codigoart) WHERE (articulos.ar_ultimamodificacion BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "') AND articulos.ar_precioanterior IS NOT NULL AND codigos_ripley.id IS NOT NULL", "muznovar")
        If clase.dt.Tables("muznovar").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("muznovar")
            prepara_columnas()
            textBox2.Text = clase.dt.Tables("muznovar").Rows.Count
            Button4.Enabled = True
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 7
            prepara_columnas()
            textBox2.Text = 0
            Button4.Enabled = False
            MessageBox.Show("No se encontraron articulos creados en el intervalo de fechas especificado.", "NO SE ENCONTRARON ARTICULOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub prepara_columnas()
        With dataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Referencia"
            .Columns(2).HeaderText = "Descripción"
            .Columns(3).HeaderText = "Línea"
            .Columns(4).HeaderText = "Costo"
            .Columns(5).HeaderText = "Precio 1"
            .Columns(6).HeaderText = "Precio 2"
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Width = 80
            .Columns(1).Width = 150
            .Columns(2).Width = 180
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 100
            .Columns(6).Width = 100
        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim m_excel As Microsoft.Office.Interop.Excel.Application
        m_excel = CreateObject("Excel.Application")
        m_excel.Workbooks.Open("C:\Data\CAMBIO_PVP_Y_COSTO.xlsx")
        m_excel.Visible = True
        Dim costoripley As Double = frm_planillas_ripley_codigos_nuevos.porcentaje_costo_ripley()
        Dim iva As Double = frm_planillas_ripley_codigos_nuevos.porcentaje_de_iva()
        clase.consultar("SELECT CONCAT(ar_descripcion, ' ', ar_referencia) AS descripcion, ar_referencia, ar_codigobarras, ar_precio1 * " & Str(costoripley) & "  AS Costo, ar_precio1, 1 - ((ar_precio1 * " & Str(costoripley) & " * " & Str(iva + 1) & ")/ar_precio1) AS margen, ar_precioanterior * " & Str(costoripley) & "  AS costoant, ar_precioanterior AS pvpant, 1 - ((ar_precioanterior * " & Str(costoripley) & " * " & Str(iva + 1) & ")/ar_precioanterior) AS margenant FROM articulos LEFT JOIN codigos_ripley ON (articulos.ar_codigo = codigos_ripley.codigoart) WHERE (ar_ultimamodificacion BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "') AND articulos.ar_precioanterior IS NOT NULL AND codigos_ripley.id IS NOT NULL", "ripley")
        Dim x As Integer
        For x = 0 To clase.dt.Tables("ripley").Rows.Count - 1
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 2).value = x + 1
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 3).value = clase.dt.Tables("ripley").Rows(x)("descripcion")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 13).value = clase.dt.Tables("ripley").Rows(x)("ar_referencia")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 26).value = clase.dt.Tables("ripley").Rows(x)("ar_codigobarras")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 28).value = clase.dt.Tables("ripley").Rows(x)("Costo")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 29).value = clase.dt.Tables("ripley").Rows(x)("ar_precio1")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 30).value = clase.dt.Tables("ripley").Rows(x)("margen")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 34).value = "PLANET LOVE"
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 124).value = clase.dt.Tables("ripley").Rows(x)("costoant")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 125).value = clase.dt.Tables("ripley").Rows(x)("pvpant")
            m_excel.Worksheets("Planilla de Ingreso").cells(23 + x, 126).value = clase.dt.Tables("ripley").Rows(x)("margenant")
        Next
        If Not m_excel Is Nothing Then
            m_excel.Application.ActiveWorkbook.Close(SaveChanges:=True)
            m_excel.Quit()
            m_excel = Nothing
        End If
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            FileCopy("C:\Data\CAMBIO_PVP_Y_COSTO.xlsx", SaveFileDialog1.FileName)
            FileCopy("C:\Data\CAMBIO_PVP_Y_COSTO1.xlsx", "C:\Data\CAMBIO_PVP_Y_COSTO.xlsx")
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 7
            prepara_columnas()
            textBox2.Text = "0"
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class