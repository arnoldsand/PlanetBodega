Public Class class_library
    Public conex As MySqlConnection
    Public conexaccess As OleDb.OleDbConnection = New OleDbConnection

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
        If combo.Text.Trim = "" Then
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

    Public Ds As DataSet = New DataSet
    Public dt As DataSet = New DataSet
    Public dt1 As DataSet = New DataSet
    Public dt2 As DataSet = New DataSet
    Public dt_global As DataSet = New DataSet
    Public dtSql As DataSet = New DataSet
    Public dtSql1 As DataSet = New DataSet
    Public dtaccess As DataSet = New DataSet
    Public c1global As New MySqlDataAdapter
    Public conexsqlserver As SqlClient.SqlConnection

    Public Sub Conectar()
        'Dim host, user, pass, data As String
        'Dim conexion1 As New OleDb.OleDbConnection
        'conexion1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\setting.accdb;Jet OLEDB:Database Password=PL456324;"
        'Dim cm111 As New OleDb.OleDbCommand("select* from setting", conexion1)
        'Dim c111 As New OleDb.OleDbDataAdapter
        'c111.SelectCommand = cm111
        'Dim dt111 As New DataTable
        'c111.Fill(dt111)
        'If dt111.Rows.Count > 0 Then
        '    host = dt111.Rows(0)("host")
        '    user = dt111.Rows(0)("user")
        '    pass = dt111.Rows(0)("pass")
        '    data = dt111.Rows(0)("data")

        'End If

        conex = New MySqlConnection()
        conex.ConnectionString = "Server=localhost;Database=erp-db-planetlove;Uid=root;Pwd=181288;"
        'conex.ConnectionString = "Server=localhost;Database=erp-db-planetlove;Uid=root;Pwd=181288;"
        'conex.ConnectionString = "Server=192.168.1.142;Database=pl_cajas;Uid=root;Pwd=planetlove44;"
        '    conex.ConnectionString = "Server=192.168.1.142;Database=erp-db-planetlove;Uid=root;Pwd=planetlove44;"
        ' conex.ConnectionString = "server=190.85.66.234;Uid=root;Pwd=planetlove44;Database=caja"
        conex.Open()
    End Sub

    'metodos para SQL Server
    Private Sub ConectarSQLServer()
        conexsqlserver = New SqlClient.SqlConnection
        ' conexsqlserver.ConnectionString = "Data Source=192.168.1.165\ICG;Initial Catalog=LovePOSCentral;Integrated Security=SSPI;"
        'conexsqlserver.ConnectionString = "data source = 192.168.1.165\ICG; initial catalog = LovePOSCentral; user id = sistemas; password = 9VYE648X4K"
        'conexsqlserver.ConnectionString = "data source = 190.165.165.45\ICG; initial catalog = LovePOSCentral; user id = sistemas; password = 9VYE648X4K"
        'cadena de conexion de prueba
        conexsqlserver.ConnectionString = "Data Source=.;Initial Catalog=LovePOS;Integrated Security=SSPI;"
        conexsqlserver.Open()
    End Sub

    Private Sub DesconectarSQLServer()
        conexsqlserver.Close()
    End Sub

    Public Sub ConsultarSQLServer(ByVal sql As String, ByVal tablaconsulta As String)
        ConectarSQLServer()
        dtSql.Tables.Clear()
        Dim cm1 As New SqlClient.SqlCommand(sql, conexsqlserver)
        Dim c1 As New SqlClient.SqlDataAdapter()
        c1.SelectCommand = cm1
        '  MsgBox(sql)
        c1.Fill(dtSql, tablaconsulta)
        DesconectarSQLServer()
    End Sub

    Public Sub ConsultarSQLServer1(ByVal sql As String, ByVal tablaconsulta As String)
        ConectarSQLServer()
        dtSql1.Tables.Clear()
        Dim cm1 As New SqlClient.SqlCommand(sql, conexsqlserver)
        Dim c1 As New SqlClient.SqlDataAdapter()
        c1.SelectCommand = cm1
        '  MsgBox(sql)
        c1.Fill(dtSql1, tablaconsulta)
        DesconectarSQLServer()
    End Sub

    Public Sub AgregarSQLServer(ByVal Sql As String)
        ConectarSQLServer()
        'MsgBox(Sql)
        Dim cm2 As New SqlClient.SqlCommand(Sql, conexsqlserver)
        cm2.ExecuteNonQuery()
        '     MsgBox(Sql)
        DesconectarSQLServer()
    End Sub

    Public Sub ActualizarSQLServer(ByVal Sql As String)
        ConectarSQLServer()
        'MsgBox(Sql)
        Dim cm2 As New SqlClient.SqlCommand(Sql, conexsqlserver)
        cm2.ExecuteNonQuery()
        '     MsgBox(Sql)
        DesconectarSQLServer()
    End Sub

    Public Sub conectar_access()
        conexaccess = New OleDb.OleDbConnection
        conexaccess.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\foxpos.accdb;Jet OLEDB:Database Password=PL456324;"
        conexaccess.Open()
    End Sub

    Public Sub desconectar_access()
        conexaccess.Close()
    End Sub




    Function consultarExcel(Sql As String, tabla As String, Ruta As String) As Boolean
        Try
            conectarexcel(Ruta)
            Ds.Tables.Clear()
            Dim cm As New OleDbCommand(Sql, conexaccess)
            Dim Da As New OleDbDataAdapter()
            Da.SelectCommand = cm
            Da.Fill(Ds, tabla)
            conexaccess.Close()
            If Ds.Tables(tabla).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            conexaccess.Close()
            MessageBox.Show("EL NOMBRE DE LA HOJA NO CORRESPONDE AL ARCHIVO", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub conectarexcel(File As String)
        conexaccess.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (File & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
        'cn.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.12.0;" & ("Data Source=" & (Archivo & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
        conexaccess.Open()
    End Sub

    Public Sub consultaraccess(sql As String, tablaconsulta As String)
        conectar_access()
        dtaccess.Tables.Clear()
        Dim cm1 As New OleDb.OleDbCommand(sql, conexaccess)
        Dim c1 As New OleDb.OleDbDataAdapter
        c1.SelectCommand = cm1
        c1.Fill(dtaccess, tablaconsulta)
        desconectar()
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
        '  MsgBox(sql)
        c1.Fill(dt, tablaconsulta)
        desconectar()
    End Sub

    Public Sub ConsultarConParametros(sql As String, tabla As String, NombreParametros As String)
        Conectar()
        dt.Tables.Clear()
        Dim cm1 As New MySqlCommand()
        cm1.Connection = conex
        cm1.CommandText = sql
        Dim param As New MySqlClient.MySqlParameter
        param.ParameterName = NombreParametros
        param.MySqlDbType = MySqlDbType.Int32
        param.Value = 0
        cm1.Parameters.Add(param)
        Dim c1 As New MySqlDataAdapter
        c1.SelectCommand = cm1
        '  MsgBox(sql)
        c1.Fill(dt, tabla)
        desconectar()
    End Sub

    Public Sub consultar1(ByVal sql As String, ByVal tablaconsulta As String)
        Conectar()
        dt1.Tables.Clear()
        Dim cm1 As New MySqlCommand(sql, conex)
        Dim c1 As New MySqlDataAdapter
        c1.SelectCommand = cm1
        c1.Fill(dt1, tablaconsulta)
        desconectar()
    End Sub

    Public Sub consultar2(ByVal sql As String, ByVal tablaconsulta As String)
        Conectar()
        dt2.Tables.Clear()
        Dim cm2 As New MySqlCommand(sql, conex)
        Dim c2 As New MySqlDataAdapter
        c2.SelectCommand = cm2
        c2.Fill(dt2, tablaconsulta)
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

    Public Sub consultar_global1(ByVal sql As String, ByVal tablaconsulta As String)
        Conectar()
        ' dt.Tables.Clear()
        Dim cm1 As New MySqlCommand(sql, conex)
        c1global.SelectCommand = cm1
        c1global.Fill(dt_global, tablaconsulta)
        desconectar()
    End Sub

    Public Sub agregar_registro(ByVal Sql As String)
        Conectar()
        'MsgBox(Sql)
        Dim cm2 As New MySqlCommand(Sql, conex)
        cm2.ExecuteNonQuery()
        '     MsgBox(Sql)
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

    Public Sub borradoautomatico(ByVal sql As String)
        Dim cm3 As New MySqlCommand(sql, conex)
        Try
            Conectar()
            cm3.Connection = conex
            cm3.ExecuteNonQuery()
            desconectar()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ERROR AL INTENTAR ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub actualizar(ByVal sql As String)
        Dim cm3 As New MySqlCommand(sql, conex)
        Conectar()
        '    MsgBox(sql)
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


    'funciones agregadas por cesar, revisarlas después
    Function FormatoFecha(Fbase As Date) As String
        FormatoFecha = Fbase.ToString("yyyy-MM-dd")
    End Function

    Function FormatoHora(Hbase As Date) As String
        FormatoHora = Hbase.ToString("HH:mm:ss")
    End Function

    Public Sub agregar_registro2(ByVal Sql As String)
        Conectar()
        Dim cm3 As New MySqlCommand(Sql, conex)
        cm3.ExecuteNonQuery()
        desconectar()
    End Sub

    Public Sub actualizar2(ByVal sql As String)
        Dim cm4 As New MySqlCommand(sql, conex)
        Conectar()
        cm4.Connection = conex
        cm4.ExecuteNonQuery()
        desconectar()
    End Sub

    Public Sub Limpiar_Cajas(ByVal f As Form)
        ' recorrer todos los controles del formulario indicado  
        For Each c As Control In f.Controls
            If TypeOf c Is TextBox Then
                c.Text = "" ' eliminar el texto  
            End If
        Next
    End Sub



End Class
