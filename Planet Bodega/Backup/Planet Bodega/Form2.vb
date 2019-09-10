Public Class Form2
    Dim clase As New class_library
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        clase.llenar_combo_datagrid(cmbcolores, "select* from colores order by colornombre asc", "colornombre", "cod_color")
        clase.llenar_combo_datagrid(tallas, "select* from tallas order by nombretalla asc", "nombretalla", "codigo_talla")
    End Sub
    Function generar_cadena_sql(ByVal referencia As String) As String ' esta funcion me genera una consulta sql evaluando los campos nulos de la _
        ' clasificacion por familias para mostrarlo en el combo que me permite seleccionar cuantos productor difentes estan asociados a la misma referencia
        generar_cadena_sql = ""
        clase.consultar("SELECT  DISTINCT articulos.ar_referencia, articulos.ar_linea, articulos.ar_sublinea1, articulos.ar_sublinea2, articulos.ar_sublinea3, articulos.ar_sublinea4  FROM articulos LEFT JOIN sublinea1 ON (articulos.ar_sublinea1 = sublinea1.sl1_codigo) LEFT JOIN sublinea2 ON (sublinea2.sl2_sl1codigo = sublinea1.sl1_codigo) AND (articulos.ar_sublinea2 = sublinea2.sl2_codigo) LEFT JOIN linea1 ON (articulos.ar_linea = linea1.ln1_codigo) AND (sublinea1.sl1_ln1codigo = linea1.ln1_codigo) LEFT JOIN sublinea3 ON (sublinea3.sl3_sl2codigo = sublinea2.sl2_codigo) AND (articulos.ar_sublinea3 = sublinea3.sl3_codigo) LEFT JOIN sublinea4 ON (sublinea4.sl4_sl3codigo = sublinea3.sl3_codigo) AND (articulos.ar_sublinea4 = sublinea4.sl4_codigo) WHERE (articulos.ar_referencia = '" & referencia & "')", "tbl123")
        If clase.dt.Tables("tbl123").Rows.Count > 0 Then
            Dim a, ab As Short

            For a = 3 To 5
                If IsDBNull(clase.dt.Tables("tbl123").Rows(0)(a)) Then
                    ab = a
                    Exit For
                End If
            Next
            Select Case ab
                Case 3 'sublinea2 esta vacio y por ende los demas
                    generar_cadena_sql = ""
                Case 4 'sublinea3 esta vacio y por ende los demas
                    generar_cadena_sql = ", sublinea2.sl2_nombre"
                Case 5 'sublinea4 esta vacio y por ende los demas
                    generar_cadena_sql = ", sublinea2.sl2_nombre, sublinea3.sl3_nombre"
                Case 6  ' no hay campos nulos
                    generar_cadena_sql = ", sublinea2.sl2_nombre, sublinea3.sl3_nombre, sublinea4.sl4_nombre"
            End Select
        End If
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsDBNull(DataGridView1.Item(1, 5).Value) Then
            MsgBox("nulo")
        Else
            MsgBox("no nulo")
        End If
        
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox(DataGridView1.RowCount)
    End Sub
End Class