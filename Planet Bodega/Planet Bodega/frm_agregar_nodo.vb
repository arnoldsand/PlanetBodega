Public Class frm_agregar_nodo
    Dim clase As New class_library
    Dim linea, sublinea1, sublinea2, sublinea3 As Short
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


    Private Sub frm_agregar_nodo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim level As Integer
        If frm_lineas_sublineas.TreeView1.SelectedNode Is Nothing Then
            Label5.Text = "Linea:"
          
            Exit Sub
        Else
            level = frm_lineas_sublineas.TreeView1.SelectedNode.Level
        End If
        Dim nodoseleccionado As TreeNode = frm_lineas_sublineas.TreeView1.SelectedNode
        Select Case level
            Case 0
                linea = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Text
                Label5.Text = "Sublinea 1:"
                '    Label19.Visible = False
            Case 1
                linea = nodoseleccionado.Parent.Name
                sublinea1 = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Parent.Text
                TextBox2.Text = nodoseleccionado.Text
                Label5.Text = "Sublinea 2:"
                'Label19.Visible = False
                'Label1.Visible = False
            Case 2
                linea = nodoseleccionado.Parent.Parent.Name
                sublinea1 = nodoseleccionado.Parent.Name
                sublinea2 = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Parent.Parent.Text
                TextBox2.Text = nodoseleccionado.Parent.Text
                TextBox3.Text = nodoseleccionado.Text
                Label5.Text = "Sublinea 3:"
                'Label19.Visible = False
                'Label1.Visible = False
                'Label2.Visible = False
            Case 3
                linea = nodoseleccionado.Parent.Parent.Parent.Name
                sublinea1 = nodoseleccionado.Parent.Parent.Name
                sublinea2 = nodoseleccionado.Parent.Name
                sublinea3 = nodoseleccionado.Name
                TextBox1.Text = nodoseleccionado.Parent.Parent.Parent.Text
                TextBox2.Text = nodoseleccionado.Parent.Parent.Text
                TextBox3.Text = nodoseleccionado.Parent.Text
                TextBox4.Text = nodoseleccionado.Text
                Label5.Text = "Sublinea 4:"
                'Label19.Visible = False
                'Label1.Visible = False
                'Label2.Visible = False
                'Label3.Visible = False
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs)
        System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If clase.validar_cajas_text(TextBox6, "Nodo") = False Then Exit Sub
        If frm_lineas_sublineas.TreeView1.SelectedNode Is Nothing Then
            clase.agregar_registro("INSERT INTO `linea1`(`ln1_nombre`) VALUES ('" & UCase(TextBox6.Text) & "')")
            frm_lineas_sublineas.limpiar_dataset()
            frm_lineas_sublineas.inicializar_lineas()
            frm_lineas_sublineas.llenar_treeview()
        Else
            Dim level As Integer = frm_lineas_sublineas.TreeView1.SelectedNode.Level
            Select Case level
                Case 0
                    clase.agregar_registro("INSERT INTO `sublinea1`(`sl1_ln1codigo`,`sl1_nombre`) VALUES ('" & frm_lineas_sublineas.TreeView1.SelectedNode.Name & "','" & UCase(TextBox6.Text) & "')")
                Case 1
                    clase.agregar_registro("INSERT INTO `sublinea2`(`sl2_sl1codigo`,`sl2_nombre`) VALUES ('" & frm_lineas_sublineas.TreeView1.SelectedNode.Name & "','" & UCase(TextBox6.Text) & "')")
                Case 2
                    clase.agregar_registro("INSERT INTO `sublinea3`(`sl3_sl2codigo`,`sl3_nombre`) VALUES ('" & frm_lineas_sublineas.TreeView1.SelectedNode.Name & "','" & UCase(TextBox6.Text) & "')")
                Case 3
                    clase.agregar_registro("INSERT INTO `sublinea4`(`sl4_sl3codigo`,`sl4_nombre`) VALUES ('" & frm_lineas_sublineas.TreeView1.SelectedNode.Name & "','" & UCase(TextBox6.Text) & "')")
            End Select
            frm_lineas_sublineas.limpiar_dataset()
            frm_lineas_sublineas.inicializar_lineas()
            frm_lineas_sublineas.llenar_treeview()
        End If
        TextBox6.Text = ""
        TextBox6.Focus()
    End Sub
End Class