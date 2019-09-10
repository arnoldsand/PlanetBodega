Public Class frm_lineas_sublineas
    Dim clase As New class_library
    Private Sub frm_lineas_sublineas_Click(sender As Object, e As EventArgs) Handles Me.Click
        TreeView1.SelectedNode = Nothing
    End Sub
    Private Sub frm_lineas_sublineas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        inicializar_lineas()

        llenar_treeview()

    End Sub

    Sub inicializar_lineas()

        consultarlineas("select* from linea1 order by ln1_nombre asc", "tablanodos")
        consultarsublineas1("select* from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1) order by sl1_nombre asc", "tablanodoshijo1")
        consultarsublineas2("select* from sublinea2 where sl2_sl1codigo in (select sl1_codigo from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1) order by sl2_nombre asc) ", "tablanodoshijo2")
        consultarsublineas3("select* from sublinea3 where sl3_sl2codigo in (select sl2_codigo from sublinea2 where sl2_sl1codigo in (select sl1_codigo from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1))) order by sl3_nombre", "tablanodoshijo3")
        consultarsublineas4("select* from sublinea4 where sl4_sl3codigo in (select sl3_codigo from sublinea3 where sl3_sl2codigo in (select sl2_codigo from sublinea2 where sl2_sl1codigo in (select sl1_codigo from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1))) order by sl3_nombre) order by sl4_nombre asc", "tablanodoshijo4")
        dtlineas.Relations.Add("sublinea1", dtlineas.Tables("tablanodos").Columns("ln1_codigo"), dtlineas.Tables("tablanodoshijo1").Columns("sl1_ln1codigo"))
        dtlineas.Relations.Add("sublinea2", dtlineas.Tables("tablanodoshijo1").Columns("sl1_codigo"), dtlineas.Tables("tablanodoshijo2").Columns("sl2_sl1codigo"))
        dtlineas.Relations.Add("sublinea3", dtlineas.Tables("tablanodoshijo2").Columns("sl2_codigo"), dtlineas.Tables("tablanodoshijo3").Columns("sl3_sl2codigo"))
        dtlineas.Relations.Add("sublinea4", dtlineas.Tables("tablanodoshijo3").Columns("sl3_codigo"), dtlineas.Tables("tablanodoshijo4").Columns("sl4_sl3codigo"))

    End Sub

    Sub limpiar_dataset()
        dtlineas.Relations.Clear()
        dtlineas.Tables(4).Clear()
        dtlineas.Tables(3).Clear()
        dtlineas.Tables(2).Clear()
        dtlineas.Tables(1).Clear()
        dtlineas.Tables(0).Clear()
        clase.dt_global.Tables.Clear()
        dtlineas = Nothing
        dtlineas = New DataSet
        '  dtlineas.Tables.Clear()
    End Sub

    Dim dtlineas As DataSet = New DataSet
    Dim cmlineas As MySqlDataAdapter
    Dim cmsublineas1 As MySqlDataAdapter
    Dim cmsublineas2 As MySqlDataAdapter
    Dim cmsublineas3 As MySqlDataAdapter
    Dim cmsublineas4 As MySqlDataAdapter
   

    Private Sub consultarlineas(ByVal sql As String, ByVal tablaconsulta As String)
        clase.Conectar()
        '   dtlineas.Tables.Clear()
        'MsgBox(sql)
        Dim cm1 As New MySqlCommand(sql, clase.conex)
        cmlineas = New MySqlDataAdapter()
        cmlineas.SelectCommand = cm1
        cmlineas.Fill(dtlineas, tablaconsulta)
        clase.desconectar()
    End Sub

    Private Sub consultarsublineas1(ByVal sql As String, ByVal tablaconsulta As String)
        clase.Conectar()
        '   dtlineas.Tables.Clear()
        'MsgBox(sql)
        Dim cm1 As New MySqlCommand(sql, clase.conex)
        cmsublineas1 = New MySqlDataAdapter()
        cmsublineas1.SelectCommand = cm1
        cmsublineas1.Fill(dtlineas, tablaconsulta)
        clase.desconectar()
    End Sub

    Private Sub consultarsublineas2(ByVal sql As String, ByVal tablaconsulta As String)
        clase.Conectar()
        '   dtlineas.Tables.Clear()
        'MsgBox(sql)
        Dim cm1 As New MySqlCommand(sql, clase.conex)
        cmsublineas2 = New MySqlDataAdapter()
        cmsublineas2.SelectCommand = cm1
        cmsublineas2.Fill(dtlineas, tablaconsulta)
        clase.desconectar()
    End Sub

    Private Sub consultarsublineas3(ByVal sql As String, ByVal tablaconsulta As String)
        clase.Conectar()
        '   dtlineas.Tables.Clear()
        'MsgBox(sql)
        Dim cm1 As New MySqlCommand(sql, clase.conex)
        cmsublineas3 = New MySqlDataAdapter()
        cmsublineas3.SelectCommand = cm1
        cmsublineas3.Fill(dtlineas, tablaconsulta)
        clase.desconectar()
    End Sub

    Private Sub consultarsublineas4(ByVal sql As String, ByVal tablaconsulta As String)
        clase.Conectar()
        '   dtlineas.Tables.Clear()
        'MsgBox(sql)
        Dim cm1 As New MySqlCommand(sql, clase.conex)
        cmsublineas4 = New MySqlDataAdapter()
        cmsublineas4.SelectCommand = cm1
        cmsublineas4.Fill(dtlineas, tablaconsulta)
        clase.desconectar()
    End Sub

    Sub llenar_treeview()
        TreeView1.Nodes.Clear()
        Dim nodopadre As TreeNode
        For Each rw As DataRow In dtlineas.Tables("tablanodos").Rows
            nodopadre = New TreeNode()
            nodopadre.Text = rw("ln1_nombre")
            nodopadre.Name = rw("ln1_codigo")
            TreeView1.Nodes.Add(nodopadre)
            Dim nodohijo As TreeNode
            For Each rw1 As DataRow In rw.GetChildRows("sublinea1")
                nodohijo = New TreeNode
                nodohijo.Text = rw1("sl1_nombre")
                nodohijo.Name = rw1("sl1_codigo")
                nodopadre.Nodes.Add(nodohijo)
                Dim nodohijo2 As TreeNode
                For Each rw2 As DataRow In rw1.GetChildRows("sublinea2")
                    nodohijo2 = New TreeNode
                    nodohijo2.Text = rw2("sl2_nombre")
                    nodohijo2.Name = rw2("sl2_codigo")
                    nodohijo.Nodes.Add(nodohijo2)
                    Dim nodohijo3 As TreeNode
                    For Each rw3 As DataRow In rw2.GetChildRows("sublinea3")
                        nodohijo3 = New TreeNode
                        nodohijo3.Text = rw3("sl3_nombre")
                        nodohijo3.Name = rw3("sl3_codigo")
                        nodohijo2.Nodes.Add(nodohijo3)
                        Dim nodohijo4 As TreeNode
                        For Each rw4 As DataRow In rw3.GetChildRows("sublinea4")
                            nodohijo4 = New TreeNode
                            nodohijo4.Text = rw4("sl4_nombre")
                            nodohijo4.Name = rw4("sl4_codigo")
                            nodohijo3.Nodes.Add(nodohijo4)
                        Next
                    Next
                Next
            Next
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' MsgBox(TreeView1.SelectedNode.value)
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not TreeView1.SelectedNode Is Nothing Then
            Dim level1 As Integer = TreeView1.SelectedNode.Level
            If level1 = 4 Then
                MessageBox.Show("No se puede crear una clasificación de orden inferior a Sublínea 4.", "ERROR AL CREAR NODO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If
        
        frm_agregar_nodo.ShowDialog()
        frm_agregar_nodo.Dispose()
        'Dim v As String = InputBox("Digite el codigo para la subclase.", "DIGITAR VALOR")
        'If v.Trim = "" Then
        '    MessageBox.Show("La Descripción no puede ser vacía", "COMPLETAR CAMPOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Exit Sub
        'End If
       
        '    ' limpiar_dataset()
        '    '  inicializar_lineas()
        '    llenar_treeview()
        ' End If
    End Sub



    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        Dim nodoseleccionado As TreeNode = TreeView1.SelectedNode
        If nodoseleccionado Is Nothing Then
            MessageBox.Show("Debe haber un nodo seleccionar para realizar la eliminación.", "SELECCINAR NODO", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            Dim v As String = MessageBox.Show("¿Desea eliminar el nodo seleccionado?", "ELIMINAR NODO", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If v = 6 Then
                Dim level As Integer = TreeView1.SelectedNode.Level
                Select Case level
                    Case 0
                        clase.consultar("select* from articulos where ar_linea = " & nodoseleccionado.Name & "", "buscar")
                        If clase.dt.Tables("buscar").Rows.Count > 0 Then
                            MessageBox.Show("Se encontraron(ó) " & clase.dt.Tables("buscar").Rows.Count & " articulo(s) que pertenecen a esta Línea, por favor modifique el campo Línea de la ficha de este(os) articulo(s) antes de realizar la eliminación.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            clase.consultar1("select* from sublinea1 where sl1_ln1codigo = " & nodoseleccionado.Name & "", "buscar1")
                            If clase.dt1.Tables("buscar1").Rows.Count > 0 Then
                                MessageBox.Show("Debe eliminar primero las sublinea1 asociada a esta linea.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                clase.borradoautomatico("delete from linea1 where ln1_codigo = " & nodoseleccionado.Name & "")
                                MessageBox.Show("Se eliminó la linea exitosamente.", "ELIMINACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End If
                    Case 1
                        clase.consultar("select* from articulos where ar_sublinea1 = " & nodoseleccionado.Name & "", "buscar")
                        If clase.dt.Tables("buscar").Rows.Count > 0 Then
                            MessageBox.Show("Se encontraron(ó) " & clase.dt.Tables("buscar").Rows.Count & " articulo(s) que pertenecen a esta Sublínea1, por favor modifique el campo Sublinea1 de la ficha de este(os) articulo(s) antes de realizar la eliminación.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            clase.consultar1("select* from sublinea2 where sl2_sl1codigo = " & nodoseleccionado.Name & "", "buscar1")
                            If clase.dt1.Tables("buscar1").Rows.Count > 0 Then
                                MessageBox.Show("Debe eliminar primero las sublinea2 asociada a esta sublinea1.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                clase.borradoautomatico("delete from sublinea1 where sl1_codigo = " & nodoseleccionado.Name & "")
                                MessageBox.Show("Se eliminó la sublinea1 exitosamente.", "ELIMINACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End If
                    Case 2
                        clase.consultar("select* from articulos where ar_sublinea2 = " & nodoseleccionado.Name & "", "buscar")
                        If clase.dt.Tables("buscar").Rows.Count > 0 Then
                            MessageBox.Show("Se encontraron(ó) " & clase.dt.Tables("buscar").Rows.Count & " articulo(s) que pertenecen a esta Sublínea2, por favor modifique el campo Sublinea2 de la ficha de este(os) articulo(s) antes de realizar la eliminación.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            clase.consultar1("select* from sublinea3 where sl3_sl2codigo = " & nodoseleccionado.Name & "", "buscar1")
                            If clase.dt1.Tables("buscar1").Rows.Count > 0 Then
                                MessageBox.Show("Debe eliminar primero las sublinea3 asociada a esta sublinea2.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                clase.borradoautomatico("delete from sublinea2 where sl2_codigo = " & nodoseleccionado.Name & "")
                                MessageBox.Show("Se eliminó la sublinea2 exitosamente.", "ELIMINACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End If
                    Case 3
                        clase.consultar("select* from articulos where ar_sublinea3 = " & nodoseleccionado.Name & "", "buscar")
                        If clase.dt.Tables("buscar").Rows.Count > 0 Then
                            MessageBox.Show("Se encontraron(ó) " & clase.dt.Tables("buscar").Rows.Count & " articulo(s) que pertenecen a esta Sublínea3, por favor modifique el campo Sublinea3 de la ficha de este(os) articulo(s) antes de realizar la eliminación.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            clase.consultar1("select* from sublinea4 where sl4_sl3codigo = " & nodoseleccionado.Name & "", "buscar1")
                            If clase.dt1.Tables("buscar1").Rows.Count > 0 Then
                                MessageBox.Show("Debe eliminar primero las sublinea4 asociada a esta sublinea3.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                clase.borradoautomatico("delete from sublinea3 where sl3_codigo = " & nodoseleccionado.Name & "")
                                MessageBox.Show("Se eliminó la sublinea3 exitosamente.", "ELIMINACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End If
                    Case 4
                        clase.consultar("select* from articulos where ar_sublinea4 = " & nodoseleccionado.Name & "", "buscar")
                        If clase.dt.Tables("buscar").Rows.Count > 0 Then
                            MessageBox.Show("Se encontraron(ó) " & clase.dt.Tables("buscar").Rows.Count & " articulo(s) que pertenecen a esta Sublínea4, por favor modifique el campo Sublinea4 de la ficha de este(os) articulo(s) antes de realizar la eliminación.", "ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            clase.borradoautomatico("delete from sublinea4 where sl4_codigo = " & nodoseleccionado.Name & "")
                            MessageBox.Show("Se eliminó la sublinea4 exitosamente.", "ELIMINACIÓN", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        End If
                End Select
                limpiar_dataset()
                inicializar_lineas()
                llenar_treeview()
            End If
            End If
           
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TreeView1.SelectedNode Is Nothing Then
            MessageBox.Show("Debe seleccionar un nodo para realizar la edición.", "EDITAR NODOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        frm_editar_nodos.ShowDialog()
        frm_editar_nodos.Dispose()
    End Sub
End Class