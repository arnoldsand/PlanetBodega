Public Class frm_subir_codigos_ripley
    Dim clase As New class_library
    Dim Ruta As String
    Dim Archivo As Boolean
    Dim DsExcel As New DataSet

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        'SOLICITAMOS EL ARCHIVO DE EXCEL DE LA PRE INSPECCION
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Title = "SELECIONE ARCHIVO EXCEL"
        openFileDialog1.Filter = "Archivo Pre Inspeccion|*.xls;*.xlsx"

        'SE OBTIENE LA RUTA DEL ARCHIVO SELECCIONADO
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Ruta = openFileDialog1.FileName
            txtRuta.Text = openFileDialog1.FileName
            LlenarGrid()
            Archivo = True
        End If
        openFileDialog1.Dispose()
    End Sub

    Public Sub LlenarGrid()
        grdCodigosRipley.DataSource = Nothing
        Dim Hoja As String = ""
        'CONSULTAR EL ARCHIVO DE EXCEL Y LLENAR EL DATASET CON LA INFORMACION DEL MISMO
        Hoja = InputBox("POR FAVOR DIGITE EL NOMBRE DE LA HOJA DEL DOCUMENTO DE EXCEL", "PLANET LOVE")

        If clase.consultarExcel("select * from [" & Hoja & "$]", "excel", Ruta) = True Then
            grdCodigosRipley.DataSource = clase.Ds.Tables("excel")
            DsExcel = clase.Ds
            grdCodigosRipley.RowHeadersVisible = False
        Else
            Limpiar()
        End If
    End Sub

    Public Sub Limpiar()
        txtRuta.Text = ""
        grdCodigosRipley.DataSource = Nothing
        OpenFileDialog1.Dispose()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        ProgressBar1.Maximum = grdCodigosRipley.RowCount
        For i = 0 To grdCodigosRipley.Rows.Count - 1
            Application.DoEvents()

            If IsDBNull(grdCodigosRipley(3, i).Value) = False And IsDBNull(grdCodigosRipley(4, i).Value) = False Then
                clase.consultar("SELECT * FROM articulos WHERE ar_codigobarras='" & grdCodigosRipley(4, i).Value & "'", "consulta")
                If clase.dt.Tables("consulta").Rows.Count > 0 Then
                    clase.agregar_registro("INSERT INTO codigos_ripley(codigoart,codigoripley) VALUES('" & clase.dt.Tables("consulta").Rows(0)("ar_codigo") & "','" & grdCodigosRipley(3, i).Value & "')")
                End If
            End If
            ProgressBar1.Increment(1)
        Next
        MessageBox.Show("Datos guardados con éxito.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class