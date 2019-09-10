Public Class frm_finalizar_sin_revisar
    Dim clase As New class_library
    Dim ind As Boolean

    Dim Bodegaorigen As Integer
    Dim CodigoDestino As Integer
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_finalizar_sin_revisar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox5.Focus()
        ind = False
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub textBox2_GotFocus(sender As Object, e As EventArgs) Handles textBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
   
        If clase.validar_cajas_text(TextBox5, "Número de Transferencia") = False Then Exit Sub
        If ind = False Then
            clase.consultar("SELECT cabtransferencia.tr_operador, tiendas.tienda, cabtransferencia.tr_fecha, cabtransferencia.tr_fecha, cabtransferencia.tr_revisada, cabtransferencia.tr_finalizada, cabtransferencia.tr_revisor, cabtransferencia.tr_bodega, cabtransferencia.tr_destino  FROM cabtransferencia INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) WHERE (cabtransferencia.tr_estado =TRUE AND cabtransferencia.tr_operador IS NOT NULL AND cabtransferencia.tr_numero = " & TextBox5.Text & ")", "trans")
            If clase.dt.Tables("trans").Rows.Count > 0 Then
                If IsDBNull(clase.dt.Tables("trans").Rows(0)("tr_finalizada")) = False Then
                    If (clase.dt.Tables("trans").Rows(0)("tr_finalizada")) Then
                        MessageBox.Show("La transferencia digitada ya fue finalizada.", "TRANSFERENCIA FINALIZADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                End If

                If IsDBNull(clase.dt.Tables("trans").Rows(0)("tr_bodega")) Then
                    Bodegaorigen = 0
                Else
                    Bodegaorigen = clase.dt.Tables("trans").Rows(0)("tr_bodega")
                End If
                CodigoDestino = clase.dt.Tables("trans").Rows(0)("tr_destino")
                If clase.dt.Tables("trans").Rows(0)("tr_revisada") = False Then
                    textBox2.Text = clase.dt.Tables("trans").Rows(0)("tienda")
                    TextBox1.Text = clase.dt.Tables("trans").Rows(0)("tr_operador")
                    Dim fecha As Date = clase.dt.Tables("trans").Rows(0)("tr_fecha")
                    TextBox3.Text = fecha.ToString("dd/MM/yyyy")
                    Button1.Enabled = True
                    TextBox5.Enabled = False
                    TextBox4.Enabled = True
                    ind = True
                Else
                    MessageBox.Show("La transferencia digitada ya fue revisada.", "TRANSFERENCIA NO REVISADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox5.Text = ""
                End If
            Else
                MessageBox.Show("La transferencia digitada no fue encontrada.", "TRANSFERENCIA NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox5.Text = ""
            End If
        Else
            If TextBox4.Focus = True Then
                If clase.validar_cajas_text(TextBox4, "Contraseña") = False Then Exit Sub
                If TextBox4.Text = hallar_contrasena() Then
                    clase.actualizar("UPDATE cabtransferencia SET tr_finalizada = TRUE, tr_revisada = TRUE, tr_revisor = 'OPERADOR POR DEFECTO' WHERE tr_numero = " & TextBox5.Text & "")
                    MessageBox.Show("La transferencia fue finalizada exitosamente.", "TRANSFERENCIA FINALIZADA", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    SubirTransferenciaLovePos("OPERADOR POR DEFECTO", TextBox5.Text, CodigoDestino)
                    ' imprimir_hoja_transferencia(TextBox5.Text)
                    textBox2.Text = ""
                    TextBox1.Text = ""
                    TextBox3.Text = ""
                    TextBox5.Text = ""
                    TextBox5.Enabled = True
                    TextBox4.Text = ""
                    TextBox4.Enabled = False
                    TextBox5.Focus()

                Else
                    MessageBox.Show("La contraseña escrita es incorrecta.", "CONTRASEÑA INCORRECTA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox4.Text = ""
                    TextBox4.Focus()
                End If
            End If
        End If
    End Sub

    Function hallar_contrasena() As String
        clase.consultar("select password_impedir_revision from informacion ", "contra")
        If IsDBNull(clase.dt.Tables("contra").Rows(0)("password_impedir_revision")) Then
            Return ""
        Else
            Return clase.dt.Tables("contra").Rows(0)("password_impedir_revision")
        End If
    End Function

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        clase.validar_numeros(e)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        textBox2.Text = ""
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox5.Enabled = True
        TextBox4.Text = ""
        TextBox4.Enabled = False
        TextBox5.Focus()
        ind = False
    End Sub
End Class