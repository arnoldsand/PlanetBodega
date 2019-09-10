Public Class frm_crear_articulos_lotes
    Dim clase As New class_library
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        frm_crear_importacion.ShowDialog()
        frm_crear_importacion.Dispose()
    End Sub

    Private Sub frm_crear_articulos_lotes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With dataGridView1
            .RowCount = 0
            .RowHeadersWidth = 4
           
        End With
        llenar_combo()
        ind_cargador_lote = False
    End Sub

    Sub llenar_combo()
        Dim c1 As New MySqlDataAdapter("select* from cabimportacion order by imp_fecha desc LIMIT 0, 10", clase.conex)
        Dim dt As New DataTable
        clase.Conectar()
        c1.SelectCommand.Connection = clase.conex
        c1.Fill(dt)
        clase.desconectar()
        If dt.Rows.Count > 0 Then
            comboBox1.DataSource = dt
            comboBox1.ValueMember = "imp_codigo"
            ComboBox1.DisplayMember = "imp_nombrefecha"
        Else
            ComboBox1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim openFileDialog1 As System.Windows.Forms.OpenFileDialog
        openFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Dim excelPathName As String
        With openFileDialog1
            .Title = "Cargar x Lote de Archivos"
            .FileName = ""
            .DefaultExt = ".xls"
            .AddExtension = True
            .Filter = "Excel Worksheets|*.xls; *.xlsx"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                excelPathName = (CType(.FileName, String))
                If (excelPathName.Length) <> 0 Then

                    dataGridView1.Columns.Clear()
                    clase.consultar("select* from detalleimportacion where codigo_importacion = " & cod_importacion & "", "tbl")
                    If clase.dt.Tables("tbl").Rows.Count > 0 Then
                        Dim v As String
                        v = MessageBox.Show("Ya hay articulos cargados en esta importación. ¿Desea borrar y reemplazarlos por los items nuevos", "REEMPLAZAR ITEMS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                        If v = 6 Then
                            clase.borrar("delete from detalleimportacion where codigo_importacion = " & cod_importacion & "", "Se borrarán los articulos que ya estan cargados en la actual importación.")
                            Cargar(dataGridView1, excelPathName, "Hoja1")
                        End If
                    End If

                Else
                    MessageBox.Show("No se ha seleccionado ningun archivo compatible. Pulse aceptar para volverlo a intentar.", "SELECCIONAR ARCHIVO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End With
    End Sub

    Private Sub button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    

    Sub Cargar(ByVal dgView As DataGridView, ByVal SLibro As String, ByVal sHoja As String)
        'HDR=YES : Con encabezado  
        Dim cs As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & SLibro & ";" & _
                           "Extended Properties=""Excel 8.0;HDR=YES"""
        '  Try
        ' cadena de conexión  
        Dim cn As New OleDbConnection(cs)
        If Not System.IO.File.Exists(SLibro) Then
            MsgBox("No se encontró el Libro: " & _
                    SLibro, MsgBoxStyle.Critical, _
                    "Ruta inválida")
            Exit Sub
        End If
        ' se conecta con la hoja sheet 1  
        Dim dAdapter As New OleDbDataAdapter("Select * From [" & sHoja & "$]", cs)
        Dim datos As New DataSet
        ' agrega los datos  
        Try
            dAdapter.Fill(datos)
        Catch ex As Exception
            MessageBox.Show("Se produjo un error al intentar leer el archivo, recuerde que el archivo debe estar abierto y correctamente configurado.", "ERROR AL LEER ARCHIVO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        With dataGridView1
            ' llena el DataGridView  
            .DataSource = datos.Tables(0)
            ind_cargador_lote = True
            'guardar_encabezados(datos)
            Button11.Enabled = True
            insertar_datos(datos)
            ' DefaultCellStyle: formato currency   
            'para los encabezados 1,2 y 3 del DataGrid  
            '  .Columns(1).DefaultCellStyle.Format = "c"
            '  .Columns(2).DefaultCellStyle.Format = "c"
            '  .Columns(3).DefaultCellStyle.Format = "c"
        End With
        ' Catch oMsg As Exception
        ' MsgBox(oMsg.Message, MsgBoxStyle.Critical)
        ' End Try
    End Sub

    Sub insertar_datos(ByVal data As DataSet)
        Dim campos As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
        Dim c1 As New MySqlDataAdapter("Select* from detalleimportacion", clase.conex)
        Dim dt1 As New DataSet
        Dim cms As New MySqlCommandBuilder(c1)
        Dim c As Integer
        Dim b As Short
        Dim dtr As DataRow
        clase.Conectar()
        c1.SelectCommand.Connection = clase.conex
        c1.InsertCommand = cms.GetInsertCommand
        c1.Fill(dt1, "detalleimportacion")
        clase.desconectar()
        For c = 0 To data.Tables(0).Rows.Count - 1
            Application.DoEvents()
            dtr = dt1.Tables(0).NewRow
            dtr("codigo_importacion") = cod_importacion
            b = 0
            Do While b <= data.Tables(0).Columns.Count - 1
                dtr(campos(b)) = data.Tables(0).Rows(c)(b)
                b += 1
            Loop
            dtr("procesado") = False
            dt1.Tables(0).Rows.Add(dtr)
        Next c
        clase.Conectar()
        c1.Update(dt1, "detalleimportacion")
        clase.desconectar()
    End Sub

    ' Sub guardar_encabezados(ByVal dt5 As DataSet)
    'Dim a As Short
    '   For a = 0 To dt5.Tables(0).Columns.Count - 1
    '      listado(a) = dt5.Tables(0).Columns(a).Caption
    ' Next
    ' End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If ComboBox1.Text = "" Then
            Exit Sub
        End If
        cod_importacion = ComboBox1.SelectedValue
        Button12.Enabled = True
        llenar_datagrid()
        clase.consultar("select* from cabimportacion where imp_codigo = " & ComboBox1.SelectedValue & "", "tabla1")
        If clase.dt.Tables("tabla1").Rows.Count > 0 Then
            Label3.Text = clase.dt.Tables("tabla1")(0)("imp_codigo") & " - " & clase.dt.Tables("tabla1")(0)("imp_nombrefecha")
        End If
        'Label3.Text = 
        Button1.Enabled = True
    End Sub

    Public Sub llenar_datagrid()
        clase.consultar("select A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, item from detalleimportacion where codigo_importacion = " & ComboBox1.SelectedValue & " and procesado = 0", "detalleimportacion")
        If clase.dt.Tables("detalleimportacion").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("detalleimportacion")
            Button11.Enabled = True
        Else
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = Nothing
            Button11.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frm_parametros_importacion.ShowDialog()
        frm_parametros_importacion.Dispose()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If ComboBox1.Text = "" Then
            Exit Sub
        End If
        clase.borrar("delete from cabimportacion where imp_codigo = " & ComboBox1.SelectedValue & "", "¿Desea eliminar los datos de la actual importación?")
        llenar_combo()
        dataGridView1.DataSource = Nothing
        'ComboBox1.DataSource = Nothing
        Button11.Enabled = False
        Button12.Enabled = False
        Button1.Enabled = False
    End Sub

    Private Sub button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button7.Click
        frm_filtrar_x_caja.ShowDialog()
        frm_filtrar_x_caja.Dispose()
    End Sub

    
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If IsNumeric(dataGridView1.CurrentCell.ColumnIndex) And IsNumeric(dataGridView1.CurrentCell.RowIndex) Then
            frm_creacion_articulos_lotes.ShowDialog()
            frm_creacion_articulos_lotes.Dispose()
        End If
    End Sub

   
End Class