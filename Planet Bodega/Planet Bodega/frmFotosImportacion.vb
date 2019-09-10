Public Class frmFotosImportacion
    Dim clase As New class_library
    Dim CodigoImportacion As Integer
    Dim dt As New DataTable
    Dim dt1 As New DataSet
    Dim RutaFacturas As String

    Public Sub New()

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

        ' columna.DisplayIndex = 3
        CrearDatatable()

        clase.consultar("select* from informacion", "ruta")
        RutaFacturas = clase.dt.Tables("ruta").Rows(0)("foto_factura")
    End Sub

    Private Sub CrearDatatable()


        dt1.Tables(0).Columns.Add("Codigo")
        dt1.Tables(0).Columns.Add("Proveedor")
        dt1.Tables(0).Columns.Add("Fotografia")
    End Sub

    Private Sub btnImportacion_Click(sender As Object, e As EventArgs) Handles btnImportacion.Click
        Dim formulario As New frm_seleccionar_importacion1
        formulario.ShowDialog()
        If formulario.HuboSeleccion Then
            CodigoImportacion = formulario.CodigoImportacionSeleccionado
            txtNombreImportacion.Text = formulario.NombreImportacion
            llenarGrilla()
        End If
        formulario.Dispose()
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub frmFotosImportacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub llenarGrilla()
        clase.consultar("SELECT proveedores.prv_codigoasignado AS Codigo, proveedores.prv_nombre AS Proveedor,'' AS Foto FROM entradamercancia INNER JOIN detalle_proveedores_articulos ON (entradamercancia.com_codigoart = detalle_proveedores_articulos.codigo_articulo) INNER JOIN proveedores ON (detalle_proveedores_articulos.codigo_proveedor = proveedores.prv_codigo) WHERE (entradamercancia.com_codigoimp =" & CodigoImportacion & ") GROUP BY proveedores.prv_codigoasignado, proveedores.prv_nombre ORDER BY Proveedores.prv_nombre ASC", "proveedores")
        dt = clase.dt.Tables("proveedores")
        dtgGrillaImportacion.DataSource = Nothing
        dtgGrillaImportacion.Columns.Clear()
        dtgGrillaImportacion.DataSource = dt
        prepararColumnas()
    End Sub

    Private Sub prepararColumnas()
        dtgGrillaImportacion.Columns(0).Width = 100
        dtgGrillaImportacion.Columns(1).Width = 200
        dtgGrillaImportacion.Columns(2).Width = 340
        dtgGrillaImportacion.Columns(0).ReadOnly = True
        dtgGrillaImportacion.Columns(1).ReadOnly = True
        dtgGrillaImportacion.Columns(2).ReadOnly = True
        dtgGrillaImportacion.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgGrillaImportacion.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dtgGrillaImportacion.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    End Sub

    Private Sub FiltrarGrilla(proveedor As String)
        dtgGrillaImportacion.DataSource = Nothing

        Dim filas() As DataRow = dt.Select("Proveedor LIKE '%" & proveedor & "%'")

        dt1.Tables(0).Rows.Clear()

        Dim row As DataRow
        For Each fila As DataRow In filas
            row = dt1.Tables(0).NewRow
            row("Codigo") = fila(0)
            row("Proveedor") = fila(1)
            row("Fotografia") = fila(2)
            dt1.Tables(0).Rows.Add(row)
        Next
        dtgGrillaImportacion.DataSource = dt1
        prepararColumnas()
    End Sub


    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnFactura.Click
        If dtgGrillaImportacion.RowCount > 0 And dtgGrillaImportacion.CurrentCell IsNot Nothing Then
            Dim openFileDialog1 As System.Windows.Forms.OpenFileDialog
            openFileDialog1 = New System.Windows.Forms.OpenFileDialog
            Dim RutaPdf As String
            With openFileDialog1
                .Title = "Cargar Foto"
                .FileName = ""
                .DefaultExt = ".pdf"
                .AddExtension = True
                .Filter = "Formato de Intercambio de Archivos PDF|*.pdf"
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    RutaPdf = (CType(.FileName, String))
                    If (RutaPdf.Length) <> 0 Then
                        dtgGrillaImportacion.Item(2, dtgGrillaImportacion.CurrentCell.RowIndex).Value = RutaPdf
                        VisorPDF.src = RutaPdf
                        For i As Integer = 0 To dt.Rows.Count - 1
                            If (dt.Rows(i)("Codigo") = dtgGrillaImportacion.Item(0, dtgGrillaImportacion.CurrentCell.RowIndex).Value) Then
                                dt.Rows(i)("Foto") = RutaPdf
                            End If
                        Next
                        'PictureBox1.Image = Image.FromFile(RutaPdf)
                        'SetImage(PictureBox1)
                    Else
                        MessageBox.Show("No se ha seleccionado ningun archivo compatible. Pulse aceptar para volverlo a intentar.", "SELECCIONAR ARCHIVO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End With
        End If
    End Sub

    Private Sub dtgGrillaImportacion_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtgGrillaImportacion.CellClick
        If dtgGrillaImportacion.Item(2, dtgGrillaImportacion.CurrentCell.RowIndex).Value = "" Then
            VisorPDF.src = RutaFacturas & "\SinPDF.pdf"
        Else
            VisorPDF.src = dtgGrillaImportacion.Item(2, dtgGrillaImportacion.CurrentCell.RowIndex).Value
        End If
    End Sub


    Private Sub txtProveedor_TextChanged(sender As Object, e As EventArgs) Handles txtProveedor.TextChanged
        FiltrarGrilla(txtProveedor.Text.Trim())
    End Sub
End Class