Public Class frm_planillas_ripley_codigos_nuevos
    Dim clase As New class_library
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, articulos.ar_precio1, articulos.ar_precio2 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) LEFT JOIN codigos_ripley ON (articulos.ar_codigo = codigos_ripley.codigoart) WHERE (articulos.ar_ultimamodificacion BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "' AND codigos_ripley.id IS NULL)", "muznovar")
        If clase.dt.Tables("muznovar").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("muznovar")
            prepara_columnas()
            textBox2.Text = clase.dt.Tables("muznovar").Rows.Count
            Button4.Enabled = True
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 6
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
            .Columns(4).HeaderText = "Precio 1"
            .Columns(5).HeaderText = "Precio 2"
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).Width = 80
            .Columns(1).Width = 150
            .Columns(2).Width = 180
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 100
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim m_excel As Microsoft.Office.Interop.Excel.Application
        m_excel = CreateObject("Excel.Application")
        m_excel.Workbooks.Open("C:\Data\planilla_ripley.xlsx")
        m_excel.Visible = False
        Dim costoripley As Double = porcentaje_costo_ripley()
        Dim iva As Double = porcentaje_de_iva()

        clase.consultar("SELECT CONCAT(ar_descripcion, ' ', ar_referencia) AS descripcion, ar_referencia, ar_codigobarras, ar_precio1 * " & Str(costoripley) & "  AS Costo, ar_precio1, 1 - ((ar_precio1 * " & Str(costoripley) & " * " & Str(iva + 1) & ")/ar_precio1) AS margen, ar_codigo FROM articulos LEFT JOIN codigos_ripley ON (articulos.ar_codigo = codigos_ripley.codigoart) WHERE (ar_ultimamodificacion BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "') AND codigos_ripley.id IS NULL", "ripley")
        Dim x As Integer
        For x = 0 To clase.dt.Tables("ripley").Rows.Count - 1
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 2).value = x + 1
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 3).value = clase.dt.Tables("ripley").Rows(x)("descripcion")
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 13).value = clase.dt.Tables("ripley").Rows(x)("ar_referencia")
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 26).value = clase.dt.Tables("ripley").Rows(x)("ar_codigobarras")
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 28).value = clase.dt.Tables("ripley").Rows(x)("Costo")
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 29).value = clase.dt.Tables("ripley").Rows(x)("ar_precio1")
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 30).value = clase.dt.Tables("ripley").Rows(x)("margen")
            m_excel.Worksheets("PLANILLA DE CODIFICACION 2").cells(23 + x, 34).value = "PLANET LOVE"

            clase.agregar_registro("INSERT INTO codigos_ripley(codigoart) VALUES('" & clase.dt.Tables("ripley").Rows(x)("ar_codigo") & "')")
        Next
        If Not m_excel Is Nothing Then
            m_excel.Application.ActiveWorkbook.Close(SaveChanges:=True)
            m_excel.Quit()
            m_excel = Nothing
        End If
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            FileCopy("C:\Data\planilla_ripley.xlsx", SaveFileDialog1.FileName)
            FileCopy("C:\Data\planilla_ripley1.xlsx", "C:\Data\planilla_ripley.xlsx")
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 6
            prepara_columnas()
            textBox2.Text = "0"
        End If
    End Sub

    Function porcentaje_costo_ripley() As Double
        clase.consultar1("select porcentaje_costo_ripley from informacion", "ripley")
        Return clase.dt1.Tables("ripley").Rows(0)("porcentaje_costo_ripley")
    End Function

    Function porcentaje_de_iva() As Double
        clase.consultar1("select porcentaje_de_iva from informacion", "ripley")
        Return clase.dt1.Tables("ripley").Rows(0)("porcentaje_de_iva")
    End Function

    Private Sub frm_planillas_ripley_codigos_nuevos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class