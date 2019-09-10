Public Class frm_seleccionar_importacion1
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Public CodigoImportacionSeleccionado As Integer
    Public NombreImportacion As String
    Public HuboSeleccion As Boolean

    Public Sub New()

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        CodigoImportacionSeleccionado = 0
        NombreImportacion = ""
        HuboSeleccion = False
    End Sub


    Private Sub frm_seleccionar_importacion1_Load(sender As Object, e As EventArgs) Handles Me.Load
        llenar_combo()
        With DataGridView1
            .Columns(0).Width = 350

        End With
    End Sub

    Sub llenar_combo()
        Dim c1 As New MySqlDataAdapter("select imp_nombrefecha as NOMBRE_IMPORTACION, imp_codigo from cabimportacion order by imp_fecha desc LIMIT 0, 100", clase.conex)
        Dim dt As New DataTable
        clase.Conectar()
        c1.SelectCommand.Connection = clase.conex
        c1.Fill(dt)
        clase.desconectar()
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentCell.RowIndex) Then
            codiimportation = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
            CodigoImportacionSeleccionado = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
            NombreImportacion = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            HuboSeleccion = True
            '  frm_almacenaje_articulos.TextBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            Me.Close()
        End If
    End Sub
End Class