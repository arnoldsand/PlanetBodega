Imports System.Math
Public Class frm_ajustes
    Dim clase As New class_library
    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_ajustes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox6.SelectedIndex = 0
        dataGridView1.ColumnCount = 8
        preparar_columnas()
    End Sub

    Sub preparar_columnas()
        With dataGridView1
            .Columns(0).HeaderText = "Codigo"
            .Columns(1).HeaderText = "Tipo de Ajuste"
            .Columns(2).HeaderText = "Fecha"
            .Columns(3).HeaderText = "Realizado Por"
            .Columns(4).HeaderText = "Precio Costo"
            .Columns(5).HeaderText = "Precio Venta1"
            .Columns(6).HeaderText = "Precio Venta2"
            .Columns(0).Width = 50
            .Columns(1).Width = 200
            .Columns(2).Width = 100
            .Columns(3).Width = 200
            .Columns(4).Width = 150
            .Columns(5).Width = 150
            .Columns(6).Width = 150
            .Columns(7).Width = 4
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_password_ajustes.ShowDialog()
        frm_password_ajustes.Dispose()
        If passtrue = True Then
            frm_nuevo_ajustes.ShowDialog()
            frm_nuevo_ajustes.Dispose()
            passtrue = False
        End If
    End Sub

    Sub llenar_ajustes(tipo As Short, inicio As Date, final As Date)
        Select Case tipo
            Case 0
                If Len(TextBox1.Text) = 13 Then
                    TextBox1.Text = convertir_codigobarra_a_codigo_normal(TextBox1.Text)
                End If
                If clase.validar_cajas_text(TextBox1, "Codigo Articulo") = False Then Exit Sub '                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              
                clase.consultar("SELECT DISTINCT cabajuste.cabaj_codigo, tipos_ajustes.tip_nombre, cabajuste.cabaj_fecha, cabajuste.cabaj_operario, cabajuste.cabaj_codigo FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE (detajuste.detaj_articulo =" & TextBox1.Text & " AND cabajuste.cabaj_procesado =TRUE AND cabajuste.cabaj_fecha BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "')", "resultados")
                If clase.dt.Tables("resultados").Rows.Count > 0 Then
                    Dim x As Short
                    dataGridView1.RowCount = 0
                    Dim fecha As Date
                    For x = 0 To clase.dt.Tables("resultados").Rows.Count - 1
                        With dataGridView1
                            .RowCount = .RowCount + 1
                            .Item(0, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")
                            .Item(1, x).Value = clase.dt.Tables("resultados").Rows(x)("tip_nombre")
                            fecha = clase.dt.Tables("resultados").Rows(x)("cabaj_fecha")
                            .Item(2, x).Value = fecha.ToString("dd/MM/yyyy")
                            .Item(3, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_operario")
                            .Item(4, x).Value = FormatCurrency(precio_costo(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 0)
                            .Item(5, x).Value = FormatCurrency(precio_venta1(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 0)
                            .Item(6, x).Value = FormatCurrency(precio_venta2(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 0)
                        End With
                    Next
                Else
                    dataGridView1.RowCount = 0
                    MessageBox.Show("No se encontraron ajustes realizadas en el intervalo de fechas especificados.", "AJUSTES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Case 1
                If clase.validar_cajas_text(TextBox1, "Referencia") = False Then Exit Sub
                frm_elegir_linea_sublinea.ShowDialog()
                frm_elegir_linea_sublinea.Dispose()
                If cadena_string.Trim <> "" Then
                    clase.consultar(cadena_string, "tablrest")
                    Dim x As Short
                    Dim ind As Boolean = False '                                                                                                                                                                                                                                                                                                                                                                                                   
                    Dim sql_generada As String = "SELECT DISTINCT cabajuste.cabaj_codigo, tipos_ajustes.tip_nombre, cabajuste.cabaj_fecha, cabajuste.cabaj_operario, cabajuste.cabaj_codigo FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE (((cabajuste.cabaj_procesado =TRUE) AND (cabajuste.cabaj_fecha BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "')) AND ("
                    For x = 0 To clase.dt.Tables("tablrest").Rows.Count - 1
                        If ind = False Then
                            sql_generada = sql_generada & "(detajuste.detaj_articulo =" & clase.dt.Tables("tablrest").Rows(x)("ar_codigo") & ")"
                            ind = True
                        Else
                            sql_generada = sql_generada & " OR (detajuste.detaj_articulo =" & clase.dt.Tables("tablrest").Rows(x)("ar_codigo") & ")"
                        End If
                    Next
                    sql_generada = sql_generada & "))"
                    clase.consultar(sql_generada, "resultados")
                    Dim fecha As Date
                    dataGridView1.RowCount = 0
                    If clase.dt.Tables("resultados").Rows.Count > 0 Then
                        For x = 0 To clase.dt.Tables("resultados").Rows.Count - 1
                            With dataGridView1
                                .RowCount = .RowCount + 1
                                .Item(0, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")
                                .Item(1, x).Value = clase.dt.Tables("resultados").Rows(x)("tip_nombre")

                                fecha = clase.dt.Tables("resultados").Rows(x)("cabaj_fecha")
                                .Item(2, x).Value = fecha.ToString("dd/MM/yyyy")
                                .Item(3, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_operario")
                                .Item(4, x).Value = FormatCurrency(precio_costo(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                                .Item(5, x).Value = FormatCurrency(precio_venta1(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                                .Item(6, x).Value = FormatCurrency(precio_venta2(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                            End With
                        Next
                    Else
                        dataGridView1.RowCount = 0
                        MessageBox.Show("No se encontraron ajustes realizadas en el intervalo de fechas especificados.", "AJUSTES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    cadena_string = "" 'restablezco la variable para evitar bugs
                End If
            Case 2 '                                                                                                                                                                                                                                                                  
                clase.consultar("SELECT tipos_ajustes.tip_nombre, cabajuste.cabaj_fecha, cabajuste.cabaj_operario, cabajuste.cabaj_codigo FROM cabajuste INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE (cabajuste.cabaj_procesado =FALSE)", "resultados")
                If clase.dt.Tables("resultados").Rows.Count > 0 Then
                    Dim x As Short
                    Dim fecha As Date
                    dataGridView1.RowCount = 0
                    For x = 0 To clase.dt.Tables("resultados").Rows.Count - 1
                        With dataGridView1
                            .RowCount = .RowCount + 1
                            .Item(0, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")
                            .Item(1, x).Value = clase.dt.Tables("resultados").Rows(x)("tip_nombre")
                            fecha = clase.dt.Tables("resultados").Rows(x)("cabaj_fecha")
                            .Item(2, x).Value = fecha.ToString("dd/MM/yyyy")
                            .Item(3, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_operario")
                            .Item(4, x).Value = FormatCurrency(precio_costo(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                            .Item(5, x).Value = FormatCurrency(precio_venta1(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                            .Item(6, x).Value = FormatCurrency(precio_venta2(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                        End With
                    Next
                Else
                    dataGridView1.RowCount = 0
                    MessageBox.Show("No se encontraron ajustes realizadas en el intervalo de fechas especificados.", "AJUSTES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Case 3  ' '                                                                                                                                                                                                                                                                                                                                                                                                                     
                clase.consultar("SELECT tipos_ajustes.tip_nombre, cabajuste.cabaj_fecha, cabajuste.cabaj_operario, cabajuste.cabaj_codigo FROM cabajuste INNER JOIN tipos_ajustes ON (cabajuste.cabaj_tipo_ajuste = tipos_ajustes.tip_codigo) WHERE ((cabajuste.cabaj_procesado =TRUE) AND (cabajuste.cabaj_fecha BETWEEN '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' AND '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "'))", "resultados")
                If clase.dt.Tables("resultados").Rows.Count > 0 Then
                    Dim x As Short
                    Dim fecha As Date
                    dataGridView1.RowCount = 0
                    For x = 0 To clase.dt.Tables("resultados").Rows.Count - 1
                        With dataGridView1
                            .RowCount = .RowCount + 1
                            .Item(0, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")
                            .Item(1, x).Value = clase.dt.Tables("resultados").Rows(x)("tip_nombre")
                            fecha = clase.dt.Tables("resultados").Rows(x)("cabaj_fecha")
                            .Item(2, x).Value = fecha.ToString("dd/MM/yyyy")
                            .Item(3, x).Value = clase.dt.Tables("resultados").Rows(x)("cabaj_operario")
                            .Item(4, x).Value = FormatCurrency(precio_costo(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                            .Item(5, x).Value = FormatCurrency(precio_venta1(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                            .Item(6, x).Value = FormatCurrency(precio_venta2(clase.dt.Tables("resultados").Rows(x)("cabaj_codigo")), 2)
                        End With
                    Next
                Else
                    dataGridView1.RowCount = 0
                    MessageBox.Show("No se encontraron ajustes realizadas en el intervalo de fechas especificados.", "AJUSTES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
        End Select
    End Sub



    Function precio_costo(ajuste As Integer) As Double
        clase.consultar1("SELECT SUM(detaj_precio_costo * detaj_cantidad) AS Costo FROM detajuste WHERE (detaj_codigo_ajuste =" & ajuste & ")", "tabla1")
        If clase.dt1.Tables("tabla1").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("tabla1").Rows(0)("Costo")) Then
                Return 0
                Exit Function
            End If
            Return clase.dt1.Tables("tabla1").Rows(0)("Costo")
        End If
    End Function
    Function precio_venta1(ajuste As Integer) As Double
        clase.consultar1("SELECT SUM(detaj_precio_venta1 * detaj_cantidad) AS Venta FROM detajuste WHERE (detaj_codigo_ajuste =" & ajuste & ")", "tabla1")
        If clase.dt1.Tables("tabla1").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("tabla1").Rows(0)("Venta")) Then
                Return 0
                Exit Function
            End If
            Return clase.dt1.Tables("tabla1").Rows(0)("Venta")
        End If
    End Function

    Function precio_venta2(ajuste As Integer) As Double
        clase.consultar1("SELECT SUM(detaj_precio_venta2 * detaj_cantidad) AS Venta FROM detajuste WHERE (detaj_codigo_ajuste =" & ajuste & ")", "tabla1")
        If clase.dt1.Tables("tabla1").Rows.Count > 0 Then
            If IsDBNull(clase.dt1.Tables("tabla1").Rows(0)("Venta")) Then
                Return 0
                Exit Function
            End If
            Return clase.dt1.Tables("tabla1").Rows(0)("Venta")
        End If
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dataGridView1.CurrentRow.Index) Then
            frm_detalle_ajuste.MdiParent = mdi_principal
            frm_detalle_ajuste.Show()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        llenar_ajustes(ComboBox6.SelectedIndex, DateTimePicker1.Value.ToString("yyyy-MM-dd"), DateTimePicker2.Value.ToString("yyyy-MM-dd"))
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Select Case ComboBox6.SelectedIndex
            Case 0
                Label3.Text = "Codigo:"
                Label3.Visible = True
                TextBox1.Visible = True
                TextBox1.Focus()
            Case 1
                Label3.Text = "Referencia:"
                Label3.Visible = True
                TextBox1.Visible = True
                TextBox1.Focus()
            Case 2
                Label3.Visible = False
                TextBox1.Visible = False
            Case 3
                Label3.Visible = False
                TextBox1.Visible = False
        End Select
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If ComboBox6.SelectedIndex = 0 Then
            clase.validar_numeros(e)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If IsNumeric(dataGridView1.CurrentRow.Index) Then
            Select Case ComboBox6.SelectedIndex
                Case 0
                    frm_consolidado_codigo.ShowDialog()
                    frm_consolidado_codigo.Dispose()
                Case 1
                    frm_consolidado_ajustes.ShowDialog()
                    frm_consolidado_ajustes.Dispose()
                Case 2
                    frm_consolidado_ajustes.ShowDialog()
                    frm_consolidado_ajustes.Dispose()
                Case 3
                    frm_consolidado_ajustes.ShowDialog()
                    frm_consolidado_ajustes.Dispose()
            End Select
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If dataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        Dim tipodeajuste As Short
        Dim naturalezaajuste As String
        If IsNumeric(dataGridView1.CurrentRow.Index) Then
            Dim v As String = MessageBox.Show("¿Desea alterar el inventario con los ajustes generados en este momento?", "GUARDAR AJUSTES", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If v = 6 Then
                'SELECT detajuste.detaj_articulo, cabajuste.cabaj_bodega, detajuste.detaj_gondola, SUM(detajuste.detaj_cantidad) AS cant FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) WHERE (detajuste.detaj_codigo_ajuste =" & dataGridView1.Item(7, dataGridView1.CurrentRow.Index).Value & ") GROUP BY detajuste.detaj_articulo, cabajuste.cabaj_bodega, detajuste.detaj_gondola
                clase.consultar1("SELECT detajuste.detaj_articulo, SUM(detajuste.detaj_cantidad) AS cant FROM detajuste INNER JOIN cabajuste ON (detajuste.detaj_codigo_ajuste = cabajuste.cabaj_codigo) WHERE (cabajuste.cabaj_codigo =" & dataGridView1.Item(0, dataGridView1.CurrentRow.Index).Value & ") GROUP BY detajuste.detaj_articulo", "consulta")
                If clase.dt1.Tables("consulta").Rows.Count > 0 Then
                    clase.consultar("select* from  cabajuste where cabaj_codigo = " & dataGridView1.Item(0, dataGridView1.CurrentRow.Index).Value & "", "tabla")
                    If clase.dt.Tables("tabla").Rows.Count > 0 Then
                        tipodeajuste = clase.dt.Tables("tabla").Rows(0)("cabaj_tipo_ajuste")
                        If clase.dt.Tables("tabla").Rows(0)("cabaj_procesado") = True Then
                            MessageBox.Show("Este ajuste ya fue procesado antes. No se puede volver a procesar.", "AJUSTE YA PROCESADO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    End If
                    naturalezaajuste = naturaleza_del_ajuste(clase.dt.Tables("tabla").Rows(0)("cabaj_codigo"))
                    If naturalezaajuste <> "R" And naturalezaajuste <> "S" Then
                        MessageBox.Show("La naturaleza de este ajuste no permite ser procesado de esta forma.", "NO SE PUEDE PROCESAR AJUSTE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    Dim tot, tot_ajuste, cant_pos, cant_neg As Integer
                    Dim x As Integer
                    For x = 0 To clase.dt1.Tables("consulta").Rows.Count - 1
                        With clase.dt1.Tables("consulta")
                            clase.actualizar("UPDATE detajuste SET detaj_cantidad_anterior = " & cant_anterior_antes_de_ajuste(.Rows(x)("detaj_articulo")) & " WHERE detaj_articulo = " & .Rows(x)("detaj_articulo") & " AND detaj_codigo_ajuste = " & dataGridView1.Item(0, dataGridView1.CurrentRow.Index).Value & "")
                            tot_ajuste = .Rows(x)("cant")
                            clase.consultar("SELECT inv_cantidad, inv_ajustado_pos, inv_ajustado_neg  FROM inventario_bodega WHERE (inv_codigoart =" & .Rows(x)("detaj_articulo") & ")", "tablita")
                            Dim sql As String = ""
                            If clase.dt.Tables("tablita").Rows.Count > 0 Then
                                tot = clase.dt.Tables("tablita").Rows(0)("inv_cantidad")
                                cant_pos = comprobar_nulidad_de_integer(clase.dt.Tables("tablita").Rows(0)("inv_ajustado_pos"))
                                cant_neg = comprobar_nulidad_de_integer(clase.dt.Tables("tablita").Rows(0)("inv_ajustado_neg"))
                                Select Case naturalezaajuste
                                    Case "R"
                                        sql = " inv_ajustado_neg = '" & cant_neg + Abs(tot_ajuste) & "'"
                                    Case "S"
                                        sql = " inv_ajustado_pos = '" & cant_pos + tot_ajuste & "'"
                                End Select
                                clase.actualizar("UPDATE inventario_bodega SET inv_cantidad = " & tot + tot_ajuste & ", " & sql & " WHERE (inv_codigoart =" & .Rows(x)("detaj_articulo") & ")")
                            Else
                                Select Case naturalezaajuste
                                    Case "R"
                                        sql = "inv_ajustado_neg"
                                    Case "S"
                                        sql = "inv_ajustado_pos"
                                End Select
                                clase.agregar_registro("INSERT INTO `inventario_bodega`(`inv_codigo`,`inv_codigoart`,`" & sql & "`,`inv_cantidad`) VALUES ( NULL,'" & .Rows(x)("detaj_articulo") & "','" & Abs(tot_ajuste) & "','" & tot_ajuste & "')")
                            End If

                        End With
                    Next
                    clase.actualizar("UPDATE cabajuste set cabaj_procesado = TRUE, cabaj_hora = '" & Now.ToString("HH:mm:ss") & "' where cabaj_codigo = " & dataGridView1.Item(0, dataGridView1.CurrentRow.Index).Value & "")
                    MessageBox.Show("Los ajustes fueron cargados exitosamente al inventario, consulte los nuevos saldos.", "AJUSTE PROCESADO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No se puede procesar ya que no hay ajustes para realizar.", "NO SE ENCONTRARON AJUSTES", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub

    Function cant_anterior_antes_de_ajuste(articulo As Long)
        clase.consultar2("select inv_cantidad from inventario_bodega where inv_codigoart = " & articulo & "", "inv")
        If clase.dt2.Tables("inv").Rows.Count > 0 Then
            Return clase.dt2.Tables("inv").Rows(0)("inv_cantidad")
        Else
            Return 0
        End If
    End Function

    Function total_items_ajustados(ajuste As Short) As Integer
        clase.consultar("SELECT SUM(detaj_cantidad) AS cant FROM detajuste WHERE (detaj_codigo_ajuste =" & ajuste & ")", "rest")
        If clase.dt.Tables("rest").Rows.Count > 0 Then
            If IsDBNull(clase.dt.Tables("rest").Rows(0)("cant")) Then
                Return 0
            End If
            Return clase.dt.Tables("rest").Rows(0)("cant")
        End If
    End Function

    Function naturaleza_del_ajuste(ajuste As Short) As String
        clase.consultar("SELECT tipos_ajustes.tip_tipo FROM tipos_ajustes INNER JOIN cabajuste ON (tipos_ajustes.tip_codigo = cabajuste.cabaj_tipo_ajuste) WHERE (cabajuste.cabaj_codigo = " & ajuste & ")", "aju")
        If clase.dt.Tables("aju").Rows.Count > 0 Then
            Return clase.dt.Tables("aju").Rows(0)("tip_tipo")
        Else
            Return ""
        End If
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub
End Class