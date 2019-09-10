Public Class frm_impuestos1
    Dim clase As New class_library
    Dim indcolumna As Short
    Dim indcodigocolumna As Short

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Me.Close()
    End Sub

    Private Sub frm_impuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clase.consultar2("SELECT imp_codigo, imp_impuesto, imp_porcentaje FROM impuestos ORDER BY imp_codigo ASC", "detalle")
        Dim a As Short
        Dim b As Short
        For a = 0 To clase.dt2.Tables("detalle").Rows.Count - 1
            DataGridView1.RowCount = DataGridView1.RowCount + 1
            DataGridView1.Item(0, a).Value = False
            DataGridView1.Item(1, a).Value = clase.dt2.Tables("detalle").Rows(a)("imp_codigo")
            DataGridView1.Item(2, a).Value = clase.dt2.Tables("detalle").Rows(a)("imp_impuesto")
            DataGridView1.Item(3, a).Value = clase.dt2.Tables("detalle").Rows(a)("imp_porcentaje")
            For b = 0 To impuestos1.Length - 1
                If impuestos1(b) = DataGridView1.Item(1, a).Value Then
                    DataGridView1.Item(0, a).Value = True
                    Exit For
                End If
            Next
        Next
        DataGridView1.Columns(0).ReadOnly = True
        DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).ReadOnly = True
        indcolumna = hallar_columna_de_chekeo()
        indcodigocolumna = hallar_columna_de_codigo_impuesto()
        DataGridView1.Columns(indcolumna).ReadOnly = False
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim a As Short
        Dim cont As Short = 0
        For a = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(indcolumna, a).Value = True Then
                cont += 1
            End If
        Next
        ReDim impuestos1(cont - 1)
        Dim cont1 As Short = 0
        For a = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(indcolumna, a).Value = True Then
                impuestos1(cont1) = DataGridView1.Item(indcodigocolumna, a).Value
                cont1 += 1
            End If
        Next
        Me.Close()
    End Sub

    'es necesario crear una funcion para que halle el indicador de columna ya que la columna pueda cambiar de lugar al cargar y descargar el formulario
    Function hallar_columna_de_chekeo() As Short
        Dim a As Short
        Dim ind As Boolean = False
        For a = 0 To DataGridView1.ColumnCount - 1
            If DataGridView1.Columns(a).HeaderText = "V" Then
                ind = True
                Exit For
            End If
        Next
        If ind = True Then
            Return a
        End If
    End Function

    Function hallar_columna_de_codigo_impuesto() As Short
        Dim a As Short
        Dim ind As Boolean = False
        For a = 0 To DataGridView1.ColumnCount - 1
            If DataGridView1.Columns(a).HeaderText = "Codigo" Then
                ind = True
                Exit For
            End If
        Next
        If ind = True Then
            Return a
        End If
    End Function
End Class