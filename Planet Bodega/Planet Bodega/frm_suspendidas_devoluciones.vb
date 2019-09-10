﻿Public Class frm_suspendidas_devoluciones
    Dim clase As New class_library

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frm_suspendidas1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar("SELECT tr_numero AS NumeroTransferencia FROM cabtransferencia WHERE (tr_bodega IS NULL AND tr_estado = FALSE AND tr_destino IS NOT NULL AND tr_operador IS NOT NULL AND tr_opproduccion IS NULL AND tr_importacion IS NULL) ORDER BY NumeroTransferencia DESC", "tbl")
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
            frmTransferenciaDevoluciones.TextBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            frmTransferenciaDevoluciones.llenar_grilla(frmTransferenciaDevoluciones.TextBox1.Text)
            frmTransferenciaDevoluciones.Button3.Enabled = True
            frmTransferenciaDevoluciones.Button6.Enabled = True
            frmTransferenciaDevoluciones.Button4.Enabled = True
            frmTransferenciaDevoluciones.Button7.Enabled = True
            frmTransferenciaDevoluciones.Button2.Enabled = False
            frmTransferenciaDevoluciones.establecer_encabezado(DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value)
            frmTransferenciaDevoluciones.AcceptButton = frmTransferenciaDevoluciones.Button3
            frmTransferenciaDevoluciones.TextBox5.Text = ""
            '     frmTransferencia.Button1.Enabled = False
            Me.Close()
        End If
    End Sub
End Class