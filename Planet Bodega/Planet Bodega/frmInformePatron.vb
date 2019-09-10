Public Class frmInformePatron
    Dim clase As New class_library
    Dim CodigoPatron As String
    Dim i As Integer = 0
    Dim page As Integer = 0
    Dim Total As Integer = 0

    Public Sub LlenarGrid()
        grdTiendas.DataSource = Nothing

        clase.consultar("SELECT tiendas.tienda AS TIENDA, detpatrondist.dp_cantidad AS CANTIDAD FROM detpatrondist INNER JOIN tiendas ON (detpatrondist.dp_tienda = tiendas.id) WHERE (detpatrondist.dp_idpatron ='" & CodigoPatron & "');", "consulta")

        If clase.dt.Tables("consulta").Rows.Count = 0 Then
            grdTiendas.DataSource = Nothing
            Exit Sub
        End If

        With grdTiendas
            .DataSource = clase.dt.Tables("consulta")
            .RowHeadersVisible = False
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Focus()
        End With

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        If CodigoPatron = "" Then
            MessageBox.Show("No hay información para imprimir", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
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
        e.Graphics.DrawString("REPORTE PATRON DE DISTRIBUCION", FuenteTitulo, Brushes.Black, 250, 30)
        e.Graphics.DrawString("Patron: " & grdPatron(1, grdPatron.CurrentCell.RowIndex).Value.ToString, FuenteContenido, Brushes.Black, 100, 60)
        
        Dim py As Integer = 120

        'Imprimir Detalle
        e.Graphics.DrawString("TIENDA", FuenteContenido, Brushes.Black, 100, 100)
        e.Graphics.DrawString("CANTIDAD", FuenteContenido, Brushes.Black, 300, 100)

        linesPerPage = (e.MarginBounds.Height / FuenteGrid.GetHeight(e.Graphics)) - 33

        While count < linesPerPage AndAlso i < grdTiendas.Rows.Count
            e.Graphics.DrawString(grdTiendas(0, i).Value.ToString, FuenteGrid, Brushes.Black, 100, py)
            e.Graphics.DrawString(grdTiendas(1, i).Value.ToString, FuenteGrid, Brushes.Black, 330, py)

            py = py + 20
            Total = Total + grdTiendas(1, i).Value.ToString

            count += 1
            i += 1
        End While

        If i < grdTiendas.Rows.Count Then
            e.HasMorePages = True
            page = page + 1
            e.Graphics.DrawString("Pagina " & page, FuenteContenido, Brushes.Black, 670, 20)
        Else
            e.Graphics.DrawString("Total:   " & Total, FuenteContenido, Brushes.Black, 290, py + 15)
            page = page + 1
            e.Graphics.DrawString("Pagina " & page, FuenteContenido, Brushes.Black, 670, 20)
            e.HasMorePages = False
            i = 0
            Total = 0
            page = 0
        End If
    End Sub

    Private Sub frmInformePatron_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenar_listado("")
    End Sub

    Private Sub llenar_listado(criterio As String)
        clase.consultar("SELECT pat_codigo AS CODIGO,pat_nombre AS PATRON FROM cabpatrondist WHERE pat_nombre LIKE '" & criterio & "%'", "consulta")
        If clase.dt.Tables("consulta").Rows.Count = 0 Then
            Exit Sub
        End If
        With grdPatron
            .DataSource = clase.dt.Tables("consulta")
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .RowHeadersVisible = False
            .Columns(0).Visible = False
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        End With
    End Sub

    Private Sub grdPatron_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdPatron.CellClick
        CodigoPatron = grdPatron(0, grdPatron.CurrentCell.RowIndex).Value.ToString
        LlenarGrid()
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub txtcriterio_TextChanged(sender As Object, e As EventArgs) Handles txtcriterio.TextChanged
        llenar_listado(txtcriterio.Text)
    End Sub

    Private Sub grdPatron_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdPatron.CellContentClick

    End Sub
End Class
