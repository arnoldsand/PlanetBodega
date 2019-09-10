Imports Microsoft.VisualBasic
Module Module1
    Public cod_importacion As Short
    Public ind_cargador_lote As Boolean 'variable que me dice si el archivo lo cargue antes o lo acabo de hacer
    Public colores() As Integer
    Public cantidades() As Integer
    Public arraytallas() As Integer
    Public ind_creacion_colores As Boolean = False ' variable que me dice si ya cree los colores con cantidades para las referencias
    Public ind_edicion_colores As Boolean 'esta variable me dice si permites eliminar filas del datagridview
    Public codigos_de_barra() As Integer
    Public codigos() As String
    Dim clase As New class_library

    'precioactu => precio en la base de datos
    'precioafijar => precio en los textbox
    Function precios_anteriores(ByVal precioactu() As Double, ByVal dtable As DataTable, ByVal precioafijar() As Double) As String
        Dim i As Integer
        precios_anteriores = ""
        Dim campospreciosanteriores1() As String = {"ar_precioanterior", "ar_precioanttixy1"}
        Dim campospreciosanteriores2() As String = {"ar_precioanterior2", "ar_precioanttixy2"}
        Dim precioant1(2) As Double
        For i = 0 To 1
            If IsDBNull(dtable.Rows(0)(campospreciosanteriores1(i))) Then
                precioant1(i) = "0"
            Else
                precioant1(i) = dtable.Rows(0)(campospreciosanteriores1(i))
            End If
        Next
        For i = 0 To 1
            If IsDBNull(precioactu(i)) Then
                precioactu(i) = "0"
            End If
        Next
        For i = 0 To 1
            If IsDBNull(precioant1(i)) Then
                precioant1(i) = "0"
            End If
        Next
        For i = 0 To 1
            If precioafijar(i) <> precioactu(i) Then
                If IsDBNull(campospreciosanteriores1(i)) Then
                    precios_anteriores = precios_anteriores & ", " & campospreciosanteriores1(i) & "=" & precioactu(i) & " "
                Else
                    precios_anteriores = precios_anteriores & ", " & campospreciosanteriores1(i) & "=" & precioactu(i) & ", " & campospreciosanteriores2(i) & "=" & precioant1(i) & " "
                End If
            End If
        Next
    End Function

    Function generar_codigo_ean13(ByVal codigo As String) As String
        'calcular codigo de verificacion
        Dim matriz() As Short = {1, 3, 1, 3, 1, 3, 1, 3, 1, 3, 1, 3}
        Dim digito As Short
        Dim i As Short
        Dim sum As Short = 0
        Dim val1 As Short
        Dim cod_verificacion As Short
        Dim prefijo As Long
        Select Case Len(codigo)
            Case 1
                prefijo = "77000000000"
            Case 2
                prefijo = "7700000000"
            Case 3
                prefijo = "770000000"
            Case 4
                prefijo = "77000000"
            Case 5
                prefijo = "7700000"
            Case 6
                prefijo = "770000"
            Case 7
                prefijo = "77000"
            Case 8
                prefijo = "7700"
            Case 9
                prefijo = "770"
        End Select
        For i = 0 To 11
            digito = Mid(prefijo & codigo, i + 1, 1)
            sum = sum + Val(digito) * matriz(i)
        Next
        val1 = sum
        Do While (val1 Mod 10) <> 0
            val1 = val1 + 1
        Loop
        cod_verificacion = val1 - sum
        generar_codigo_ean13 = prefijo & codigo & cod_verificacion
    End Function

    Function convertir_null_combobox(ByVal parametro As String) As Short
        If IsDBNull(parametro) Then
            convertir_null_combobox = -1
        Else
            convertir_null_combobox = parametro
        End If

    End Function

    Function verificar_nulidad_nulo(ByVal valor As Object) As String
        If IsDBNull(valor) Then
            Return "NULL"
        Else
            Return valor
        End If
    End Function

    Function verificar_nulidad_vacio(ByVal valor As Object) As String
        If IsDBNull(valor) Then
            Return ""
        Else
            Return valor
        End If
    End Function

    Public Class celda_grid
        Dim dt7 As DataRow
        Dim mostar As String

        Public Sub New(ByVal fila As System.Data.DataRow)
            dt7 = fila
        End Sub

        Public Function sublinea1() As String
            Return dt7("ar_sublinea1")
        End Function

        Public Function sublinea2() As String
            If IsDBNull(dt7("ar_sublinea2")) Then
                Return Nothing
            Else
                Return dt7("ar_sublinea2")
            End If
        End Function

        Public Function sublinea3() As String
            If IsDBNull(dt7("ar_sublinea3")) Then
                Return Nothing
            Else
                Return dt7("ar_sublinea3")
            End If
        End Function

        Public Function sublinea4() As String
            If IsDBNull(dt7("ar_sublinea4")) Then
                Return Nothing
            Else
                Return dt7("ar_sublinea4")
            End If
        End Function

        Public Function linea() As String
            Return dt7("ar_linea")
        End Function

        Public Overrides Function ToString() As String
            Return dt7("refart")
        End Function
    End Class


    Function generar_cadena_sql_apartir_del_combobox_de_coincidencias(ByVal sublinea2 As String, ByVal sublinea3 As String, ByVal sublinea4 As String) As String
        Dim complementosql2, complementosql3, complementosql4 As String
        If sublinea2 = "" Then
            complementosql2 = " AND ar_sublinea2 IS NULL"
        Else
            complementosql2 = " AND ar_sublinea2 = " & sublinea2
        End If
        If sublinea3 = "" Then
            complementosql3 = " AND ar_sublinea3 IS NULL"
        Else
            complementosql3 = " AND ar_sublinea3 = " & sublinea3
        End If
        If sublinea4 = "" Then
            complementosql4 = " AND ar_sublinea4 IS NULL"
        Else
            complementosql4 = " AND ar_sublinea4 = " & sublinea4
        End If
        Return complementosql2 & complementosql3 & complementosql4
    End Function

    Public Sub construir_matrices_colores_tallas_codigosdebarra(ByVal dtable As DataTable)
        ReDim colores(dtable.Rows.Count - 1)
        ReDim arraytallas(dtable.Rows.Count - 1)
        ReDim cantidades(dtable.Rows.Count - 1)
        ReDim codigos_de_barra(dtable.Rows.Count - 1)
        Dim c As Short
        For c = 0 To dtable.Rows.Count - 1
            colores(c) = dtable.Rows(c)("ar_color")
            arraytallas(c) = dtable.Rows(c)("ar_talla")
            cantidades(c) = 0
            codigos_de_barra(c) = dtable.Rows(c)("ar_codigo")
        Next
        ind_creacion_colores = True
    End Sub

    Public Function comprobar_existencia_fila_nuevoregistro(ByVal dtgrid As DataGridView) As Boolean
        Dim t As Short
        Dim cont As Short = 0
        For t = 0 To dtgrid.ColumnCount - 1
            If dtgrid.Item(t, dtgrid.RowCount - 1).Value = Nothing Then
                cont = cont + 1
            End If
        Next
        If cont = dtgrid.ColumnCount Then
            Return True
        Else
            Return False
        End If
    End Function

    Function color_apartirdecodigo(ByVal codigo As Integer) As String
        clase.consultar("select* from colores where cod_color = " & codigo & "", "tabla")
        If clase.dt.Tables("tabla").Rows.Count > 0 Then
            Return clase.dt.Tables("tabla").Rows(0)("colornombre")
        Else
            Return ""
        End If
    End Function
End Module
