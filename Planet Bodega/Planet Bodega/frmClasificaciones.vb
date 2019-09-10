
Public Class frmClasificaciones
    Dim clase As New class_library

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Descripción"
            .Columns(0).Width = 80
            .Columns(1).Width = 250
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        End With
    End Sub

    Private Sub frmClasificaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnCount = 2
        llenar_grilla()

    End Sub

    Sub llenar_grilla()
        clase.consultar("SELECT IdClasificacion, Clasificacion FROM CabClasificacion ORDER BY Clasificacion ASC", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("tbl")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 2
            preparar_columnas()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmAgregarClasificaciones.ShowDialog()
        frmAgregarClasificaciones.Dispose()
        llenar_grilla()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmEditarClasificacion.ShowDialog()
        frmEditarClasificacion.Dispose()
        llenar_grilla()
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub
End Class