Public Class frm_mantenimiento_recibo_bodega
    Dim clase As New class_library
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frm_nuevo_recibo.ShowDialog()
        frm_nuevo_recibo.Dispose()
        llenar_lista()
    End Sub

    Private Sub llenar_lista()
        clase.consultar("SELECT cabrecibos_bodegas.recbod_codigo, cabrecibos_bodegas.recbod_fecha, cabrecibos_bodegas.recbod_hora, cabrecibos_bodegas.recbod_proveedor, cabrecibos_bodegas.recbod_entrega, cabrecibos_bodegas.recbod_recibe FROM cabrecibos_bodegas ORDER BY cabrecibos_bodegas.recbod_codigo DESC", "lista")
        If clase.dt.Tables("lista").Rows.Count > 0 Then
            DataGridView2.DataSource = clase.dt.Tables("lista")
            preparar_columnas()
        Else
            DataGridView2.DataSource = Nothing
        End If
    End Sub

    Private Sub frm_mantenimiento_recibo_bodega_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_lista()
    End Sub

    Private Sub preparar_columnas()
        With DataGridView2
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Fecha"
            .Columns(2).HeaderText = "Hora"
            .Columns(3).HeaderText = "Proveedor"
            .Columns(4).HeaderText = "Entrega"
            .Columns(5).HeaderText = "Recibe"
            .Columns(0).Width = 70
            .Columns(1).Width = 80
            .Columns(2).Width = 60
            .Columns(3).Width = 150
            .Columns(4).Width = 150
            .Columns(5).Width = 150
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            frm_detalle_recibo.ShowDialog()
            frm_detalle_recibo.Dispose()
            llenar_lista()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView2.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView2.CurrentCell.RowIndex) Then
            Dim codrecibo As Short = DataGridView2.Item(0, DataGridView2.CurrentCell.RowIndex).Value
            Dim m_excel As Microsoft.Office.Interop.Excel.Application
            m_excel = CreateObject("Excel.Application")
            m_excel.Workbooks.Open("C:\Data\formatoentrada.xls")
            m_excel.Visible = False
            clase.consultar("SELECT cabrecibos_bodegas.recbod_fecha, cabrecibos_bodegas.recbod_hora, cabrecibos_bodegas.recbod_proveedor, cabrecibos_bodegas.recbod_nit, cabrecibos_bodegas.recbod_entrega, cabrecibos_bodegas.recbod_recibe, cabrecibos_bodegas.recbod_ciudad, detrecibo_bodegas.detrec_referencia, detrecibo_bodegas.detrec_descripcion, detrecibo_bodegas.detrec_cantidad, detrecibo_bodegas.detrec_dcto, detrecibo_bodegas.detrec_precio FROM detrecibo_bodegas INNER JOIN cabrecibos_bodegas ON (detrecibo_bodegas.detrec_codrecibo = cabrecibos_bodegas.recbod_codigo) WHERE (cabrecibos_bodegas.recbod_codigo =" & codrecibo & ")", "codrecibo")
            m_excel.Worksheets("Hoja1").cells(2, 1).value = clase.dt.Tables("codrecibo").Rows(0)("recbod_proveedor") & ""
            m_excel.Worksheets("Hoja1").cells(3, 1).value = "Nit: " & clase.dt.Tables("codrecibo").Rows(0)("recbod_nit") & ""
            m_excel.Worksheets("Hoja1").cells(4, 1).value = clase.dt.Tables("codrecibo").Rows(0)("recbod_ciudad") & ""
            m_excel.Worksheets("Hoja1").cells(6, 2).value = clase.dt.Tables("codrecibo").Rows(0)("recbod_entrega") & ""
            m_excel.Worksheets("Hoja1").cells(7, 2).value = clase.dt.Tables("codrecibo").Rows(0)("recbod_recibe") & ""
            m_excel.Worksheets("Hoja1").cells(4, 6).value = "Fecha: " & clase.dt.Tables("codrecibo").Rows(0)("recbod_fecha")
            m_excel.Worksheets("Hoja1").cells(5, 6).value = "Hora: " & clase.dt.Tables("codrecibo").Rows(0)("recbod_hora").ToString()
            Dim a As Short
            For a = 0 To clase.dt.Tables("codrecibo").Rows.Count - 1
                m_excel.Worksheets("Hoja1").cells(9 + a, 1).value = clase.dt.Tables("codrecibo").Rows(a)("detrec_referencia")
                m_excel.Worksheets("Hoja1").cells(9 + a, 2).value = clase.dt.Tables("codrecibo").Rows(a)("detrec_descripcion")
                m_excel.Worksheets("Hoja1").cells(9 + a, 3).value = clase.dt.Tables("codrecibo").Rows(a)("detrec_cantidad")
                m_excel.Worksheets("Hoja1").cells(9 + a, 4).value = clase.dt.Tables("codrecibo").Rows(a)("detrec_dcto")
                m_excel.Worksheets("Hoja1").cells(9 + a, 5).value = clase.dt.Tables("codrecibo").Rows(a)("detrec_precio")
                m_excel.Worksheets("Hoja1").cells(9 + a, 6).value = clase.dt.Tables("codrecibo").Rows(a)("detrec_precio") * clase.dt.Tables("codrecibo").Rows(a)("detrec_cantidad") - (clase.dt.Tables("codrecibo").Rows(a)("detrec_precio") * clase.dt.Tables("codrecibo").Rows(a)("detrec_cantidad") * (clase.dt.Tables("codrecibo").Rows(a)("detrec_dcto") / 100))
            Next
            m_excel.Application.ActiveWorkbook.PrintOutEx()
            If Not m_excel Is Nothing Then
                m_excel.Application.ActiveWorkbook.Saved = True
                m_excel.Quit()
                m_excel = Nothing
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class