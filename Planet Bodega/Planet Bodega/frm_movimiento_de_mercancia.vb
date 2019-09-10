Public Class frm_movimiento_de_mercancia
    Dim clase As New class_library

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_movimiento_de_mercancia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.llenar_combo(cmbtienda, "select* from tiendas order by tienda asc", "tienda", "id")
        clase.llenar_combo(cmbano, "SELECT DISTINCT(YEAR(tr_fecha)) AS ano FROM cabtransferencia ORDER BY ano desc", "ano", "ano")
        'clase.consultar()
    End Sub


    Function hallar_codigo(mes As String) As Short
        Select Case mes
            Case "Enero"
                Return 1
            Case "Febrero"
                Return 2
            Case "Marzo"
                Return 3
            Case "Abril"
                Return 4
            Case "Mayo"
                Return 5
            Case "Junio"
                Return 6
            Case "Julio"
                Return 7
            Case "Agosto"
                Return 8
            Case "Septiembre"
                Return 9
            Case "Octubre"
                Return 10
            Case "Noviembre"
                Return 11
            Case "Diciembre"
                Return 12
        End Select
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If clase.validar_combobox(cmbtienda, "Tienda") = False Then Exit Sub
        If clase.validar_combobox(cmbano, "Año") = False Then Exit Sub
        If clase.validar_combobox(cmbmes, "Mes") = False Then Exit Sub
        Dim fechainicio As Date = "01/" & hallar_codigo(cmbmes.Text) & "/" & cmbano.SelectedValue
        Dim fechafinal As Date = final_mes(hallar_codigo(cmbmes.Text), cmbano.SelectedValue) & "/" & hallar_codigo(cmbmes.Text) & "/" & cmbano.SelectedValue
        Dim sql As String = "SELECT cabtransferencia.tr_numero AS Numero, cabtransferencia.tr_fecha AS Fecha, cabtransferencia.tr_operador AS Operario, cabtransferencia.tr_revisor AS Revisor, SUM(dettransferencia.dt_cantidad) AS Cant, COUNT(dettransferencia.dt_numero) AS Referencias FROM dettransferencia INNER JOIN cabtransferencia ON (dettransferencia.dt_trnumero = cabtransferencia.tr_numero) INNER JOIN tiendas  ON (cabtransferencia.tr_destino = tiendas.id) WHERE (cabtransferencia.tr_destino =" & cmbtienda.SelectedValue & " AND cabtransferencia.tr_fecha BETWEEN '" & fechainicio.ToString("yyyy-MM-dd") & "' AND '" & fechafinal.ToString("yyyy-MM-dd") & "' AND cabtransferencia.tr_revisada =TRUE AND cabtransferencia.tr_finalizada =TRUE) GROUP BY cabtransferencia.tr_numero, cabtransferencia.tr_fecha, cabtransferencia.tr_operador, cabtransferencia.tr_revisor, tiendas.email"
        ' MsgBox(sql)
        clase.consultar(sql, "movimiento")
        If clase.dt.Tables("movimiento").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("movimiento")
            GridAExcel(DataGridView1)
            cmbtienda.SelectedIndex = -1
            cmbmes.SelectedIndex = -1
            cmbano.SelectedIndex = -1
            MessageBox.Show("El archivo fue generado exitosamente.", "ARCHIVO GENERADO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            DataGridView1.DataSource = Nothing
            MessageBox.Show("No se encontraron transferencias de esta tienda realizadas en la fecha indicada.", "MOVIMIENTO DE MERCANCÍA NO EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub



    Function GridAExcel(ByVal DGV As DataGridView) As Boolean
        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet
        ' Try
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
        exApp.Application.Visible = False
        Dim fecha As Date = Now
        Dim nombre As String = fecha.ToString("dd-MM-yyyy-HHmmss")
        exApp.ActiveWorkbook.SaveAs("C:\Data\Movimientos\movimiento_mercancia_" & cmbtienda.Text & "_" & nombre & ".xlsx")
        exApp.ActiveWorkbook.Close(True)
        exApp.Quit()
        enviar_actualizacion_x_correo(hallar_email(cmbtienda.SelectedValue), "MOVIMIENTO DE MERCANCIA " & cmbmes.Text & " " & cmbano.Text, "auxiliarbodega@planetloveonline.com", "", "C:\Data\Movimientos\movimiento_mercancia_" & cmbtienda.Text & "_" & nombre & ".xlsx")
        exHoja = Nothing
        exLibro = Nothing
        exApp = Nothing
       


        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")
        '    Return False
        'End Try
        Return True
    End Function

    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Private Sub cerrar_excel()
        Dim p As Process
        For Each p In Process.GetProcesses()
            If Not p Is Nothing Then
                If p.ProcessName = "EXCEL" Then
                    p.Kill()
                End If
            End If
        Next
    End Sub


    Function final_mes(mes As Short, ano As Short) As Short
        Select Case mes
            Case 1
                Return 31
            Case 2
                Select Case hallar_ano_bisiesto(ano)
                    Case True
                        Return 29
                    Case False
                        Return 28
                End Select
            Case 3
                Return 31
            Case 4
                Return 30
            Case 5
                Return 31
            Case 6
                Return 30
            Case 7
                Return 31
            Case 8
                Return 31
            Case 9
                Return 30
            Case 10
                Return 31
            Case 11
                Return 30
            Case 12
                Return 31
        End Select
    End Function


    Function hallar_ano_bisiesto(ano As Short) As Boolean
        Dim ind1 As Boolean = False
        Dim inicio As Short = 1992
        Do While ind1 = False
            If inicio = ano Then
                ind1 = True
            Else
                inicio = inicio + 4
                If inicio > ano Then
                    Exit Do
                End If
            End If
        Loop
        If ind1 = False Then
            Return False
        Else
            Return True
        End If
    End Function
End Class