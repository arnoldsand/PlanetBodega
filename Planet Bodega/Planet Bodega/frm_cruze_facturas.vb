Public Class frm_cruze_facturas
    Dim clase As New class_library
    Private Sub frm_cruze_facturas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT cabfacturas.cabfact_codigo, cabfacturas.cabfact_fecha, proveedores.prv_codigoasignado, cabimportacion.imp_nombrefecha, cabfacturas.cabfect_funcionario FROM cabfacturas  INNER JOIN cabimportacion  ON (cabfacturas.cabfact_importacion = cabimportacion.imp_codigo) INNER JOIN proveedores ON (cabfacturas.cabfact_proveedor = proveedores.prv_codigo) WHERE (cabfacturas.cabfact_importacion =" & frm_creacion_articulos_lotes.TextBox7.Text & ")", "facturas")
        If clase.dt.Tables("facturas").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("facturas")
            preparar_columnas()
        Else
            MessageBox.Show("No hay facturas asociadas a este proveedor, no hay datos para mostrar.", "NO HAY FACTURAS REGISTRADAS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 5
            preparar_columnas()

        End If
        dtgridcajas.ColumnCount = 5
        preparar_columnas1()
    End Sub

    Private Sub preparar_columnas()
        With DataGridView1
            .Columns(0).Width = 50
            .Columns(1).Width = 80
            .Columns(2).Width = 80
            .Columns(3).Width = 140
            .Columns(4).Width = 140
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Fecha"
            .Columns(2).HeaderText = "Proveedor"
            .Columns(3).HeaderText = "Importación"
            .Columns(4).HeaderText = "Funcionario"
        End With
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub dtgridcajas_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub preparar_columnas1()
        With dtgridcajas
            .Columns(0).HeaderText = "Referencia"
            .Columns(1).HeaderText = "Descripción"
            .Columns(2).HeaderText = "Cant"
            .Columns(3).HeaderText = "Medida"
            .Columns(4).HeaderText = "Faltante Unds"
            .Columns(0).Width = 200
            .Columns(1).Width = 226
            .Columns(2).Width = 50
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        llenar_dtgrid_referencias("", DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value)
    End Sub

    Private Sub textBox2_TextChanged(sender As Object, e As EventArgs) Handles textBox2.TextChanged
        dtgridcajas.DataSource = Nothing
        dtgridcajas.ColumnCount = 5
        preparar_columnas1()
        clase.consultar("SELECT cabfacturas.cabfact_codigo, cabfacturas.cabfact_fecha, proveedores.prv_codigoasignado, cabimportacion.imp_nombrefecha, cabfacturas.cabfect_funcionario FROM cabfacturas INNER JOIN proveedores ON (cabfacturas.cabfact_proveedor = proveedores.prv_codigo) INNER JOIN cabimportacion ON (cabfacturas.cabfact_importacion = cabimportacion.imp_codigo) WHERE (proveedores.prv_codigoasignado LIKE '" & textBox2.Text & "%' AND cabimportacion.imp_codigo = " & frm_creacion_articulos_lotes.TextBox7.Text & ")", "facturas")
        If clase.dt.Tables("facturas").Rows.Count > 0 Then
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = clase.dt.Tables("facturas")
            preparar_columnas()
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.ColumnCount = 5
            preparar_columnas()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        llenar_dtgrid_referencias(TextBox1.Text, DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value)
    End Sub

    Sub llenar_dtgrid_referencias(referencia As String, fac_cod As Integer)
        clase.consultar("SELECT detfacturas.detfact_referencia, detfacturas.detfact_descripcion, detfacturas.detfact_cant, detfacturas.detfact_unimedida, detfact_faltante, detfact_codigo FROM detfacturas INNER JOIN cabfacturas ON (detfacturas.detfact_factura = cabfacturas.cabfact_codigo) WHERE (detfacturas.detfact_referencia LIKE '" & referencia & "%' AND cabfacturas.cabfact_codigo =" & fac_cod & ")", "detfacturas")
        If clase.dt.Tables("detfacturas").Rows.Count > 0 Then
            dtgridcajas.Columns.Clear()
            dtgridcajas.DataSource = clase.dt.Tables("detfacturas")
            preparar_columnas1()
        Else
            dtgridcajas.DataSource = Nothing
            dtgridcajas.ColumnCount = 5
            preparar_columnas1()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If dtgridcajas.RowCount = 0 Then
            Exit Sub
        End If
        clase.consultar("select* from detfacturas where detfact_codigo = " & dtgridcajas.Item(5, dtgridcajas.CurrentCell.RowIndex).Value & "", "dtfact")
        If IsDBNull(clase.dt.Tables("dtfact").Rows(0)("detfact_faltante")) Then
            frm_faltante_liquidacion.ShowDialog()
            frm_faltante_liquidacion.Dispose()
        Else
            MessageBox.Show("Ya hay un faltante reportado para esta referencia, no se puede volver a reportar.", "FALTANTE YA REPORTADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class
