
Public Class frm_actualizaciones_de_articulo
    Dim clase As New class_library


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_combobox(ComboBox6, "Marca") = False Then Exit Sub
        clase.consultar("SELECT articulos.ar_codigo, articulos.ar_referencia, articulos.ar_descripcion, linea1.ln1_nombre, articulos.ar_precio1, articulos.ar_precio2 FROM articulos INNER JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) WHERE (articulos.ar_ultimamodificacion BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "')", "muznovar")
        If clase.dt.Tables("muznovar").Rows.Count > 0 Then
            dataGridView1.Columns.Clear()
            dataGridView1.DataSource = clase.dt.Tables("muznovar")
            prepara_columnas()
            textBox2.Text = clase.dt.Tables("muznovar").Rows.Count
            Button4.Enabled = True
            Button2.Enabled = False
            ComboBox6.Enabled = False

        Else

            dataGridView1.DataSource = Nothing
            dataGridView1.ColumnCount = 6
            prepara_columnas()
            textBox2.Text = 0
            Button4.Enabled = False
            Button2.Enabled = True
            ComboBox6.Enabled = True
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

    Private Sub frm_actualizaciones_de_articulo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        prepara_columnas()
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim conexaccess As OleDb.OleDbConnection = New OleDb.OleDbConnection
        conexaccess.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\muznovar\muznovar.accdb;Jet OLEDB:Database Password=PL456324;"
        Dim sql As String = ""
        Dim consecutivo As Integer = vbEmpty
        Select Case ComboBox6.SelectedIndex
            Case 0
                sql = "SELECT *, ar_precio1 AS precio FROM articulos WHERE (ar_ultimamodificacion BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "')"
                consecutivo = consecutivo_actulizacion_marca(1)
            Case 1
                sql = "SELECT *, ar_precio2 AS precio FROM articulos WHERE (ar_ultimamodificacion BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "')"
                consecutivo = consecutivo_actulizacion_marca(2)
        End Select
        clase.consultar(sql, "muznovar1")
        If clase.dt.Tables("muznovar1").Rows.Count > 0 Then
            Dim a As Integer
            conexaccess.Open()
            Dim cmaccess1 As New OleDbCommand("delete from MUZNOVAR", conexaccess)
            cmaccess1.ExecuteNonQuery()
            Dim cont As Integer = 0
            For a = 0 To clase.dt.Tables("muznovar1").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("muznovar1")
                    Dim cmaccess As New OleDbCommand("INSERT INTO MUZNOVAR (AR_TIPO, AR_CODART, AR_CODIGO, AR_CONCESI, AR_DESCRI1, AR_DESCRI2, AR_GRUPO, AR_SUBGRU1, AR_SUBGRU2, AR_SUBGRU3, AR_SUBGRU4, AR_MARCA, AR_REFEREN, AR_MEDIDA, AR_TALLA, AR_COLOR, AR_TIPOVG, AR_UNIEMP, AR_PAGLIST, AR_LINEPAG, AR_GONDOLA, AR_FACTVEN, AR_TASACOM, AR_PREVEN1, AR_PREVEN2, AR_PREVEN3, AR_PREVAN1, AR_PREVAN2, AR_FECVEN1, AR_FECVAN1, AR_FECVAN2, AR_PREVREQ, AR_PESOREQ, AR_CANTREQ, AR_COSPROM, AR_FEULCOM,  AR_PRULCOM, AR_CAULCOM, AR_NPULCOM, AR_FEULVEN, AR_CANAFAM, AR_DESCODI, AR_FEDESCO, AR_CADESCO, AR_GENERIC, AR_CONTMIN, AR_HOSPITA, AR_IMPORTA, AR_SUBSIDI, AR_UNIMEDI, AR_ESTADIS, AR_IMPUES1, AR_IMPUES2, AR_IMPUES3, AR_TASADES, AR_FECHING, AR_TIPRET, AR_LINEA, AR_CLABCU, AR_CLABCM) VALUES ('1', '" & .Rows(a)("ar_codigobarras") & "', '" & colocar_espacios_vacios_al_inicio(.Rows(a)("ar_codigo"), 7) & "', ' 0', '" & colocar_espacios_vacios_al_final(Microsoft.VisualBasic.Left(.Rows(a)("ar_descripcion"), 40), 40) & "', '" & colocar_espacios_vacios_al_final(Microsoft.VisualBasic.Left(.Rows(a)("ar_descripcion"), 18), 18) & "', '00', '00','00', '00', '00', '   0', '" & colocar_espacios_vacios_al_final(Microsoft.VisualBasic.Left(.Rows(a)("ar_referencia"), 9), 15) & "', '            ', '      ', '  ', 'F', '    0', '   0', ' 0', '      1','1  ', ' ','" & colocar_espacios_vacios_al_inicio(.Rows(a)("precio"), 7) & "', '       0', '       0', '       0', '       0', '       0', '  /  /    ', '  /  /    ', 'F', 'F', 'F','" & colocar_espacios_vacios_al_inicio(Str(.Rows(a)("ar_costo")), 9) & "', '" & ultima_fecha_de_compra(.Rows(a)("AR_CODIGO")) & "', '" & colocar_espacios_vacios_al_inicio(ultimo_proveedor(.Rows(a)("AR_CODIGO")), 15) & "', '" & colocar_espacios_vacios_al_inicio(ultima_cant_comprada(.Rows(a)("AR_CODIGO")), 9) & "','      0','  /  /    ','F', 'F', '  /  /    ','0','F','F','F','F','F', '  ', 'F','19', '    0.00', ' 0', ' 0', '" & .Rows(a)("ar_fechaingreso") & "', '  ',' 0', ' ', ' ')", conexaccess)
                    cmaccess.ExecuteNonQuery()
                End With
                cont = cont + 1
            Next
            Dim cmaccess3 As New OleDbCommand("INSERT INTO MUZNOVAR (AR_CODIGO, AR_MARCA, AR_UNIEMP) VALUES ('9999999','" & colocar_espacios_vacios_al_inicio(consecutivo, 4) & "','" & colocar_espacios_vacios_al_inicio(cont, 5) & "')", conexaccess)
            cmaccess3.ExecuteNonQuery()
            If System.IO.File.Exists("C:\Data\muznovar\MUZNOVAR.TXT") = True Then 'verifico si existe
                System.IO.File.Delete("C:\Data\muznovar\MUZNOVAR.TXT") 'si exite lo elimino
            End If
            Dim cmaccess2 As New OleDbCommand("SELECT* INTO [TEXT;DATABASE=C:\Data\muznovar].MUZNOVAR.TXT FROM MUZNOVAR", conexaccess)
            Dim adpaccess As New OleDbDataAdapter
            adpaccess.SelectCommand = cmaccess2
            Dim dtaccess As New DataSet
            adpaccess.Fill(dtaccess, "muznovar")
            conexaccess.Close()
            frm_enviar_actualizacion.ShowDialog()
            frm_enviar_actualizacion.Dispose()
        End If
    End Sub

    Function ultimo_proveedor(articulo As Long) As Short
        clase.consultar1("SELECT codigo_proveedor FROM detalle_proveedores_articulos WHERE (codigo_articulo =" & articulo & ")", "proveedor")
        If clase.dt1.Tables("proveedor").Rows.Count > 0 Then
            Return clase.dt1.Tables("proveedor").Rows(clase.dt1.Tables("proveedor").Rows.Count - 1)("codigo_proveedor")
        Else
            Return vbEmpty
        End If
    End Function

    Function ultima_cant_comprada(articulo As Long) As Short
        clase.consultar1("SELECT com_unidades FROM entradamercancia WHERE (com_codigoart =" & articulo & ")", "ultcomprada")
        If clase.dt1.Tables("ultcomprada").Rows.Count > 0 Then
            Return clase.dt1.Tables("ultcomprada").Rows(clase.dt1.Tables("ultcomprada").Rows.Count - 1)("com_unidades")
        Else
            Return vbEmpty
        End If
    End Function

    Function ultima_fecha_de_compra(articulo As Long) As String
        clase.consultar1("SELECT cabimportacion.imp_fecha FROM entradamercancia INNER JOIN cabimportacion ON (entradamercancia.com_codigoimp = cabimportacion.imp_codigo) WHERE (entradamercancia.com_codigoart =" & articulo & ")", "fechacomprada")
        If clase.dt1.Tables("fechacomprada").Rows.Count > 0 Then
            Return clase.dt1.Tables("fechacomprada").Rows(clase.dt1.Tables("fechacomprada").Rows.Count - 1)("imp_fecha")
        Else
            Return vbEmpty
        End If
    End Function

    Function colocar_espacios_vacios_al_final(expresion As String, max As Short) As String
        Dim cant As Short = Len(expresion)
        If cant > max Then
            cant = max
        End If
        Dim i As Short = max - cant
        If i > 0 Then
            Dim a As Short
            For a = 1 To i
                expresion = expresion & " "
            Next
            Return expresion
        Else
            Return expresion
        End If
    End Function

    Function colocar_espacios_vacios_al_inicio(expresion As String, max As Short) As String
        Dim cant As Short = Len(expresion)
        If cant > max Then
            cant = max
        End If
        Dim i As Short = max - cant
        If i > 0 Then
            Dim a As Short
            For a = 1 To i
                expresion = " " & expresion
            Next
            Return expresion
        Else
            Return expresion
        End If
    End Function

    Function consecutivo_actulizacion_marca(marca As Short) As Integer  ' 1 => planet, 2=> tixy
        Select Case marca
            Case 1
                clase.consultar1("SELECT consecutivo_muznovar_planetlove as marca FROM informacion", "marc")
            Case 2
                clase.consultar1("SELECT consecutivo_muznovar_tixy as marca FROM informacion", "marc")
        End Select
        Return clase.dt1.Tables("marc").Rows(0)("marca")
    End Function

    Sub reinicar()
        dataGridView1.DataSource = Nothing
        dataGridView1.ColumnCount = 6
        prepara_columnas()
        ComboBox6.Enabled = True
        Button2.Enabled = True
        Button4.Enabled = False
        textBox2.Text = "0"
    End Sub
End Class