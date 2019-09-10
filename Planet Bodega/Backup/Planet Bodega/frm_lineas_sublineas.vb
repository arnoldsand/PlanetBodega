Public Class frm_lineas_sublineas
    Dim clase As New class_library
    Private Sub frm_lineas_sublineas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim nodopadre As TreeNode
        clase.consultar_global("select* from linea1 order by ln1_nombre asc", "tablanodos")
        clase.consultar_global("select* from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1) order by sl1_nombre asc", "tablanodoshijo1")
        clase.consultar_global("select* from sublinea2 where sl2_sl1codigo in (select sl1_codigo from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1) order by sl2_nombre asc) ", "tablanodoshijo2")
        clase.consultar_global("select* from sublinea3 where sl3_sl2codigo in (select sl2_codigo from sublinea2 where sl2_sl1codigo in (select sl1_codigo from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1))) order by sl3_nombre", "tablanodoshijo3")
        clase.consultar_global("select* from sublinea4 where sl4_sl3codigo in (select sl3_codigo from sublinea3 where sl3_sl2codigo in (select sl2_codigo from sublinea2 where sl2_sl1codigo in (select sl1_codigo from sublinea1 where sl1_ln1codigo in (select ln1_codigo from linea1))) order by sl3_nombre) order by sl4_nombre asc", "tablanodoshijo4")
        clase.dt_global.Relations.Add("sublinea1", clase.dt_global.Tables("tablanodos").Columns("ln1_codigo"), clase.dt_global.Tables("tablanodoshijo1").Columns("sl1_ln1codigo"))
        clase.dt_global.Relations.Add("sublinea2", clase.dt_global.Tables("tablanodoshijo1").Columns("sl1_codigo"), clase.dt_global.Tables("tablanodoshijo2").Columns("sl2_sl1codigo"))
        clase.dt_global.Relations.Add("sublinea3", clase.dt_global.Tables("tablanodoshijo2").Columns("sl2_codigo"), clase.dt_global.Tables("tablanodoshijo3").Columns("sl3_sl2codigo"))
        clase.dt_global.Relations.Add("sublinea4", clase.dt_global.Tables("tablanodoshijo3").Columns("sl3_codigo"), clase.dt_global.Tables("tablanodoshijo4").Columns("sl4_sl3codigo"))
        For Each rw As DataRow In clase.dt_global.Tables("tablanodos").Rows
            nodopadre = New TreeNode(rw("ln1_nombre"))
            TreeView1.Nodes.Add(nodopadre)
            Dim nodohijo As TreeNode
            For Each rw1 As DataRow In rw.GetChildRows("sublinea1")
                nodohijo = New TreeNode(rw1("sl1_nombre"))
                nodopadre.Nodes.Add(nodohijo)
                Dim nodohijo2 As TreeNode
                For Each rw2 As DataRow In rw1.GetChildRows("sublinea2")
                    nodohijo2 = New TreeNode(rw2("sl2_nombre"))
                    nodohijo.Nodes.Add(nodohijo2)
                    Dim nodohijo3 As TreeNode
                    For Each rw3 As DataRow In rw2.GetChildRows("sublinea3")
                        nodohijo3 = New TreeNode(rw3("sl3_nombre"))
                        nodohijo2.Nodes.Add(nodohijo3)
                        Dim nodohijo4 As TreeNode
                        For Each rw4 As DataRow In rw3.GetChildRows("sublinea4")
                            nodohijo4 = New TreeNode(rw4("sl4_nombre"))
                            nodohijo3.Nodes.Add(nodohijo4)
                        Next
                    Next
                Next
            Next

        Next

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' MsgBox(TreeView1.SelectedNode.value)
    End Sub
End Class