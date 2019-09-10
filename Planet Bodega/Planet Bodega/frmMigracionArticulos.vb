Imports Microsoft.VisualBasic
Public Class frmMigracionArticulos
    Dim clase As New class_library
    Dim maximo As Integer

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, articulos.ar_precio1, articulos.ar_precio2  FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) WHERE articulos.ar_migrado=FALSE", "articulos")
        If clase.dt.Tables("articulos").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("articulos")
            prepara_columnas()
            textBox2.Text = clase.dt.Tables("articulos").Rows.Count
            Button4.Enabled = True
            Button2.Enabled = False
            ProgressBar1.Maximum = clase.dt.Tables("articulos").Rows.Count()
            maximo = clase.dt.Tables("articulos").Rows.Count()
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 6
            prepara_columnas()
            textBox2.Text = 0
            Button4.Enabled = False
            Button2.Enabled = True
            maximo = 0
            MessageBox.Show("No se encontraron articulos para migrar.", "MIGRACIÓN DE ARTICULOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MigrarArticulosLineasProveedores()
        Button4.Enabled = False
        dataGridView1.DataSource = Nothing
        Dim a As Integer
        For a = 1 To maximo
            ProgressBar1.Increment(1)
        Next
        textBox2.Text = ""
        MessageBox.Show("La migración de articulos se realizó satisfactoriamente.", "OPERACIÓN COMPLETADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

   
End Class