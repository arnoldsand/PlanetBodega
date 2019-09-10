Public Class frm_crear_articulos_lotes
    Dim clase As New class_library
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frm_crear_importacion.ShowDialog()
        frm_crear_importacion.Dispose()
    End Sub

    Private Sub frm_crear_articulos_lotes_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub frm_crear_articulos_lotes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        salidamovimiento = vbEmpty
    End Sub

    'este codigo no lo quiero borrar porque creo que me sirve mas adelante
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frm_parametros_importacion.ShowDialog()
        frm_parametros_importacion.Dispose()
    End Sub

    Private Sub button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frm_filtrar_x_caja.ShowDialog()
        frm_filtrar_x_caja.Dispose()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dataGridView1.CurrentCell.ColumnIndex) And IsNumeric(dataGridView1.CurrentCell.RowIndex) Then
            '   If dataGridView1.Item(4, dataGridView1.CurrentCell.RowIndex).Value = True Then
            'MessageBox.Show("Ya este item fue procesado, no se mostrará.", "ITEM PROCESADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Exit Sub
            ' End If

            clase.consultar("SELECT detalle_importacion_detcajas.detcab_registrodian FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) WHERE (detsalidas_mercancia.det_codigo =" & dataGridView1.Item(5, dataGridView1.CurrentRow.Index).Value & ")", "regimp")
            If clase.dt.Tables("regimp").Rows.Count > 0 Then
                If IsDBNull(clase.dt.Tables("regimp").Rows(0)("detcab_registrodian")) Then
                    MessageBox.Show("No se encontró información aduanera del articulo, esta debe existir para continuar con el proceso de liquidación.", "INFORMACIÓN ADUANERA NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    frm_creacion_articulos_lotes.ShowDialog()
                    frm_creacion_articulos_lotes.Dispose()
                End If
            End If
            
            
        End If
    End Sub

    Private Sub preparar_columnas()
        With dataGridView1
            .Columns(0).Width = 180
            .Columns(1).Width = 200
            .Columns(2).Width = 180
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 2
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripcion"
            .Columns(2).HeaderText = "Proveedor"
            .Columns(3).HeaderText = "Cantidad"
            .Columns(4).HeaderText = "Procesado"
            .Columns(6).Visible = False
        End With
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        salidamovimiento = Val(textBox2.Text)
        If clase.validar_cajas_text(textBox2, "Número Movimiento") = False Then Exit Sub
        llenar_grilla(TextBox4.Text)
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Sub llenar_grilla(referen As String)
        clase.consultar("SELECT cabsalidas_mercancia.cabsal_liquidador FROM cabsalidas_mercancia WHERE (cabsalidas_mercancia.cabsal_cod =" & salidamovimiento & ")", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            TextBox1.Text = clase.dt.Tables("tabla").Rows(0)("cabsal_liquidador")
            Button11.Enabled = True
        Else
            MessageBox.Show("El número de movimiento digitado no existe.", "MOVIMIENTO NO EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Text = ""
            textBox2.Text = ""
            textBox2.Focus()
            Button11.Enabled = False
        End If
        clase.consultar("SELECT detalle_importacion_detcajas.detcab_referencia, detalle_importacion_detcajas.detcab_descripcion, proveedores.prv_codigoasignado, detsalidas_mercancia.det_cant, detsalidas_mercancia.det_procesado, detsalidas_mercancia.det_codigo, detalle_importacion_detcajas.detcab_coditem FROM detsalidas_mercancia INNER JOIN detalle_importacion_detcajas ON (detsalidas_mercancia.det_codref = detalle_importacion_detcajas.detcab_coditem) INNER JOIN detalle_importacion_cabcajas ON (detalle_importacion_detcajas.detcab_codigocaja = detalle_importacion_cabcajas.det_caja) INNER JOIN proveedores ON (proveedores.prv_codigo = detalle_importacion_cabcajas.det_codigoproveedor) WHERE (detsalidas_mercancia.det_salidacodigo =" & salidamovimiento & " and detalle_importacion_detcajas.detcab_referencia like '%" & UCase(referen) & "%')", "table")
        If clase.dt.Tables("table").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("table")
            preparar_columnas()
        Else
            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 7
            preparar_columnas()
        End If
    End Sub

  
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        llenar_grilla(TextBox4.Text)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim openFileDialog1 As System.Windows.Forms.OpenFileDialog
        openFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Dim excelPathName As String
        With openFileDialog1
            .Title = "Cargar Foto"
            .FileName = ""
            .DefaultExt = ".jpg"
            .AddExtension = True
            .Filter = "Formato de Intercambio de Archivos JPEG|*.jpg; *.jpeg;*.jpe;*.jfif"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                excelPathName = (CType(.FileName, String))
                If (excelPathName.Length) <> 0 Then
                    txtFactura.Text = excelPathName
                    ' PictureBox1.Image = Image.FromFile(excelPathName)
                    'SetImage(PictureBox1)


                Else
                    MessageBox.Show("No se ha seleccionado ningun archivo compatible. Pulse aceptar para volverlo a intentar.", "SELECCIONAR ARCHIVO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End With
    End Sub

    Private Sub txtFactura_Enter(sender As Object, e As EventArgs) Handles txtFactura.Enter
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub


End Class