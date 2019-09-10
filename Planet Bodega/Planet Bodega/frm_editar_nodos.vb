Public Class frm_editar_nodos
    Dim clase As New class_library
    Dim linea, sublinea1, sublinea2, sublinea3, sublinea4 As Short
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub


    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub frm_editar_nodos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim level As Integer = frm_lineas_sublineas.TreeView1.SelectedNode.Level
        Dim nodoseleccionado As TreeNode = frm_lineas_sublineas.TreeView1.SelectedNode
        Select Case level
            Case 0
                linea = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Text

            Case 1
                linea = nodoseleccionado.Parent.Name
                sublinea1 = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Parent.Text
                TextBox2.Text = nodoseleccionado.Text
                'TextBox3.Visible = False
                'TextBox4.Visible = False
            Case 2
                linea = nodoseleccionado.Parent.Parent.Name
                sublinea1 = nodoseleccionado.Parent.Name
                sublinea2 = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Parent.Parent.Text
                TextBox2.Text = nodoseleccionado.Parent.Text
                TextBox3.Text = nodoseleccionado.Text


                'TextBox4.Visible = False
            Case 3
                linea = nodoseleccionado.Parent.Parent.Parent.Name
                sublinea1 = nodoseleccionado.Parent.Parent.Name
                sublinea2 = nodoseleccionado.Parent.Name
                sublinea3 = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Parent.Parent.Parent.Text
                TextBox2.Text = nodoseleccionado.Parent.Parent.Text
                TextBox3.Text = nodoseleccionado.Parent.Text
                TextBox4.Text = nodoseleccionado.Text
            Case 4
                linea = nodoseleccionado.Parent.Parent.Parent.Parent.Name
                sublinea2 = nodoseleccionado.Parent.Parent.Parent.Name
                sublinea2 = nodoseleccionado.Parent.Parent.Name
                sublinea3 = nodoseleccionado.Parent.Name
                sublinea4 = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Parent.Parent.Parent.Parent.Text
                TextBox2.Text = nodoseleccionado.Parent.Parent.Parent.Text
                TextBox3.Text = nodoseleccionado.Parent.Parent.Text
                TextBox4.Text = nodoseleccionado.Parent.Text
                TextBox5.Text = nodoseleccionado.Text
        End Select
        TextBox6.Text = nodoseleccionado.Text

        TextBox6.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox6, "Nodo") = False Then Exit Sub
        Dim nodoseleccionado As TreeNode = frm_lineas_sublineas.TreeView1.SelectedNode
        Dim level As Integer = frm_lineas_sublineas.TreeView1.SelectedNode.Level
        Select level
            Case 0
                clase.actualizar("UPDATE linea1 SET ln1_nombre = '" & UCase(TextBox6.Text) & "', migrado = '0' WHERE ln1_codigo = " & nodoseleccionado.Name & "")
            Case 1
                clase.actualizar("UPDATE sublinea1 SET sl1_nombre = '" & UCase(TextBox6.Text) & "', migrado = '0' WHERE sl1_codigo = " & nodoseleccionado.Name & "")
            Case 2
                clase.actualizar("UPDATE sublinea2 SET sl2_nombre = '" & UCase(TextBox6.Text) & "', migrado = '0' WHERE sl2_codigo = " & nodoseleccionado.Name & "")
            Case 3
                clase.actualizar("UPDATE sublinea3 SET sl3_nombre = '" & UCase(TextBox6.Text) & "' WHERE sl3_codigo = " & nodoseleccionado.Name & "")
            Case 4
                clase.actualizar("UPDATE sublinea4 SET sl4_nombre = '" & UCase(TextBox6.Text) & "' WHERE sl4_codigo = " & nodoseleccionado.Name & "")
        End Select
        frm_lineas_sublineas.limpiar_dataset()
        frm_lineas_sublineas.inicializar_lineas()
        frm_lineas_sublineas.llenar_treeview()
        Me.Close()
    End Sub
End Class