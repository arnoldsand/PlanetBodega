Public Class class_library
    Public conex As MySqlConnection

    Function validar_cajas_text(ByVal text As TextBox, ByVal campo_requerido As String) As Boolean
        Dim indicador As Boolean
        indicador = False
       
        If text.Text = "" Then
            MessageBox.Show("Debe escribir algo en el campo " & campo_requerido & ". Pulse Aceptar para volverlo a intentar.", "COMPLETAR CAMPOS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            text.Focus()
            indicador = False
        Else
            indicador = True
        End If

        If indicador = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Function validar_combobox(ByVal combo As ComboBox, ByVal campo_requerido As String) As Boolean
        Dim indicador As Boolean
        indicador = False

        If combo.Text = "" Then
            MessageBox.Show("Debe elegir algo en el campo " & campo_requerido & ". Pulse Aceptar para volverlo a intentar.", "COMPLETAR CAMPOS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            combo.Focus()
            indicador = False
        Else
            indicador = True
        End If

        If indicador = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public dt As DataSet = New DataSet
    Public dt_global As DataSet = New DataSet


    Public Sub Conectar()
        conex = New MySqlConnection()
        conex.ConnectionString = "Server=192.168.1.142;Database=erp-db-planetlove;Uid=root;Pwd=planetlove44;"
        conex.Open()
    End Sub

    Public Sub desconectar()
        conex.Close()
    End Sub

    Public Sub consultar(ByVal sql As String, ByVal tablaconsulta As String)
        Conectar()
        dt.Tables.Clear()
        Dim cm1 As New MySqlCommand(sql, conex)
        Dim c1 As New MySqlDataAdapter
        c1.SelectCommand = cm1
        c1.Fill(dt, tablaconsulta)
        desconectar()
    End Sub

    Public Sub consultar_global(ByVal sql As String, ByVal tablaconsulta As String)
        Conectar()
        ' dt.Tables.Clear()
        Dim cm1 As New MySqlCommand(sql, conex)
        Dim c1 As New MySqlDataAdapter
        c1.SelectCommand = cm1
        c1.Fill(dt_global, tablaconsulta)
        desconectar()
    End Sub

    Public Sub agregar_registro(ByVal Sql As String)
        Conectar()
        Dim cm2 As New MySqlCommand(Sql, conex)
        cm2.ExecuteNonQuery()
        desconectar()
    End Sub

    Public Sub borrar(ByVal sql As String, ByVal msj As String)
        Dim v As String = MessageBox.Show(msj, "ELIMINACIÓN DE REGISTROS", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If v = 6 Then
            Dim cm3 As New MySqlCommand(sql, conex)
            Try
                Conectar()
                cm3.Connection = conex
                cm3.ExecuteNonQuery()
                desconectar()
                MessageBox.Show("El registro se ha eliminado satisfactoriamente.", "REGISTRO ELIMINADO", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "ERROR AL INTENTAR ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
        
    End Sub

    Public Sub actualizar(ByVal sql As String)
        Dim cm3 As New MySqlCommand(sql, conex)
        Conectar()
        cm3.Connection = conex
        cm3.ExecuteNonQuery()
        desconectar()
    End Sub

    Public Sub llenar_combo(ByVal combo As ComboBox, ByVal sql As String, ByVal mostrar As String, ByVal valor As String)
        consultar(sql, "tabla4")
        If dt.Tables("tabla4").Rows.Count > 0 Then
            combo.DataSource = dt.Tables("tabla4")
            combo.DisplayMember = mostrar
            combo.ValueMember = valor
            combo.SelectedIndex = -1
        Else
            combo.DataSource = Nothing
        End If
    End Sub

    Public Sub llenar_combo_datagrid(ByVal combo As System.Windows.Forms.DataGridViewComboBoxColumn, ByVal sql As String, ByVal mostrar As String, ByVal valor As String)
        consultar(sql, "tabla4")
        If dt.Tables("tabla4").Rows.Count > 0 Then
            combo.DataSource = dt.Tables("tabla4")
            combo.DisplayMember = mostrar
            combo.ValueMember = valor
            ' combo.SelectedIndex = -1
        Else
            combo.DataSource = Nothing
        End If
    End Sub

    Public Sub enter(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
    End Sub

    Public Sub validar_numeros(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 8 Then
            e.Handled = False
        End If
    End Sub

    Function generar_codigo_tabla_articulos() As Long ' generar consecutivos de la tabla articulos
        consultar("SELECT MAX(ar_codigo) AS codigo from articulos", "tbl")

        If dt.Tables("tbl").Rows.Count > 0 Then
            If IsDBNull(dt.Tables("tbl")(0)("codigo")) Then
                generar_codigo_tabla_articulos = 1
                Exit Function
            End If
            generar_codigo_tabla_articulos = Val(dt.Tables("tbl")(0)("codigo")) + 1
        Else
            generar_codigo_tabla_articulos = 1
        End If
    End Function
End Class
