Public Class frm_suspendidas
    Dim clase As New class_library

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frm_suspendidas1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT tr_numero AS NumeroTransferencia FROM cabtransferencia WHERE (tr_bodega IS NOT NULL AND tr_estado = FALSE AND tr_destino IS NOT NULL AND tr_operador IS NOT NULL) ORDER BY NumeroTransferencia DESC", "tbl")
        If clase.dt.Tables("tbl").Rows.Count > 0 Then
            DataGridView1.DataSource = clase.dt.Tables("tbl")
            DataGridView1.Columns(0).Width = 300
        Else
            DataGridView1.DataSource = Nothing
            DataGridView1.RowCount = 1
            DataGridView1.Columns(0).Width = 300
            DataGridView1.Columns(0).HeaderText = "NumeroTransferencia"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(DataGridView1.CurrentRow.Index) Then
            frm_transferencias_desde_bodega.TextBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            frm_transferencias_desde_bodega.llenar_grilla(frm_transferencias_desde_bodega.TextBox1.Text)
            frm_transferencias_desde_bodega.Button3.Enabled = True
            frm_transferencias_desde_bodega.Button6.Enabled = True
            frm_transferencias_desde_bodega.Button4.Enabled = True
            frm_transferencias_desde_bodega.Button7.Enabled = True
            frm_transferencias_desde_bodega.Button2.Enabled = False
            frm_transferencias_desde_bodega.PictureBox1.Image = Nothing
            frm_transferencias_desde_bodega.establecer_encabezado(DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value)
            frm_transferencias_desde_bodega.AcceptButton = frm_transferencias_desde_bodega.Button3
            frm_transferencias_desde_bodega.TextBox5.Text = ""
            '     frmTransferencia.Button1.Enabled = False
            Me.Close()
        End If
    End Sub
End Class