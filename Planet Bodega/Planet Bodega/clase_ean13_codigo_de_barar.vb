Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class clase_ean13_codigo_de_barar
    Dim clase As New class_library
    Dim printer As New Printer

    Private Sub impresora_a_utlilizar(nombreimpresora As String)
        Dim prt As Printer
        For Each prt In Printers
            If prt.DeviceName = nombreimpresora Then
                printer = prt
                Exit For
            End If
        Next
    End Sub

#Region "Metodos para Impresora Zebra ZM4203dp1"
    Sub imprimir_codigos_de_barra_33_x_25(codarticulo As Long, cantplanet As Short, impresora As String)
        Dim cod_referencia As Long
        Dim a As Short = 0
        Dim p As Short
        cod_referencia = codarticulo
        impresora_a_utlilizar(impresora)
        For p = 1 To cantplanet
            If a = 0 Then printer.Print("^XA")
            clase.consultar("select* from registros_de_importaciones where articulo = " & codarticulo & "", "busqueda")
            If clase.dt.Tables("busqueda").Rows.Count > 0 Then
                printer.Print("^FO" & 25 + (a * 280) & ",10")
                printer.Print("^A0N,9,13,10^FD" & encabezado_codigos_barras() & "^FS")
                printer.Print("^FO" & 25 + (a * 280) & ",25")
                printer.Print("^A0N,14,11^FDRegistro Importacion: " & numero_registro_dian(cod_referencia) & "^FS")
                printer.Print("^FO" & 25 + (a * 280) & ",40")
                printer.Print("^A0N,11,11^FDFecha:  " & fecha_registro_dian(cod_referencia) & "^FS")
            End If
            printer.Print("^FO" & 25 + (a * 280) & ",65")
            printer.Print("^A0N,22,20^FD" & obtener_descripcion(codarticulo) & "^FS")
            printer.Print("^FO" & 25 + (a * 280) & ",85")
            printer.Print("^A0N,18,18^FD" & obtener_referencia(codarticulo) & "^FS")
            printer.Print("^FO" & 43 + (a * 280) & ",105^BY2")
            printer.Print("^BEN,25,Y,N")
            printer.Print("^FD" & generar_codigo_ean13(codarticulo) & "^FS")
            printer.Print("^FO" & 120 + (a * 280) & ",155")
            printer.Print("^A0N,25,30^FD" & FormatCurrency(hallar_precio_planet(codarticulo), 0) & "^FS")
            If a = 2 Then
                printer.Print("^XZ")
            Else
                If p = cantplanet Then
                    printer.Print("^XZ")
                End If
            End If
            a = a + 1
            If a > 2 Then
                a = 0
            End If
        Next
        printer.EndDoc()

    End Sub

    Sub imprimir_codigos_de_barra_19_x_15_203dpi(codarticulo As Long, cantplanet As Short, impresora As String)
        Dim cod_referencia As Long
        Dim a As Short = 0
        Dim p As Short
        cod_referencia = codarticulo
        impresora_a_utlilizar(impresora)
        For p = 1 To cantplanet
            If a = 0 Then printer.Print("^XA")
            clase.consultar("select* from registros_de_importaciones where articulo = " & codarticulo & "", "busqueda")
            If clase.dt.Tables("busqueda").Rows.Count > 0 Then
                printer.Print("^FO" & 35 + (a * 255) & ",2")
                printer.Print("^A0N,10,10,10^FD" & encabezado_codigos_barras() & "^FS")
                printer.Print("^FO" & 15 + (a * 255) & ",12")
                printer.Print("^A0N,10,10^FDREG IMP: ^FS")
                printer.Print("^FO" & 60 + (a * 255) & ",12")                                                                          '  Printer.Print("^FO15,30")
                printer.Print("^A0N,10,8^FD" & numero_registro_dian(cod_referencia) & "^FS")
                printer.Print("^FO" & 170 + (a * 255) & ",12")
                printer.Print("^A0N,10,10^FD" & fecha_registro_dian(cod_referencia) & "^FS")
            Else
                printer.Print("^FO" & 35 + (a * 255) & ",2")
                printer.Print("^A0N,10,10,10^FD COMERCIALIZADO POR PLANET LOVE LTDA ^FS")
                printer.Print("^FO" & 170 + (a * 255) & ",12")
                printer.Print("^A0N,10,10^FD" & fecha_creacion(cod_referencia) & "^FS")
            End If
            printer.Print("^FO" & 15 + (a * 255) & ",25")
            printer.Print("^A0N,13,18^FD" & obtener_descripcion(codarticulo) & "^FS")
            printer.Print("^FO" & 15 + (a * 255) & ",38")
            printer.Print("^A0N,13,18^FD" & obtener_referencia(codarticulo) & "^FS")
            printer.Print("^FO" & 30 + (a * 255) & ",52^BY2")
            printer.Print("^BEN,25,Y,N")
            printer.Print("^FD" & generar_codigo_ean13(codarticulo) & "^FS")
            printer.Print("^FO" & 110 + (a * 255) & ",97")
            printer.Print("^A0N,20,30^FD" & FormatCurrency(hallar_precio_planet(codarticulo), 0) & "^FS")
            If a = 2 Then
                printer.Print("^XZ")
            Else
                If p = cantplanet Then
                    printer.Print("^XZ")
                End If
            End If
            a = a + 1
            If a > 2 Then
                a = 0
            End If
        Next
        printer.EndDoc()
        'a = 0
        'For p = 1 To canttixy
        '    If a = 0 Then printer.Print("^XA")
        '    clase.consultar("select* from registros_de_importaciones where articulo = " & codarticulo & "", "busqueda")
        '    If clase.dt.Tables("busqueda").Rows.Count > 0 Then
        '        printer.Print("^FO" & 35 + (a * 255) & ",2")
        '        printer.Print("^A0N,10,10,10^FD" & encabezado_codigos_barras() & "^FS")
        '        printer.Print("^FO" & 15 + (a * 255) & ",12")
        '        printer.Print("^A0N,10,10^FDREG IMP: ^FS")
        '        printer.Print("^FO" & 60 + (a * 255) & ",12")                                                                          '  Printer.Print("^FO15,30")
        '        printer.Print("^A0N,10,8^FD" & numero_registro_dian(cod_referencia) & "^FS")
        '        printer.Print("^FO" & 170 + (a * 255) & ",12")
        '        printer.Print("^A0N,10,10^FD" & fecha_registro_dian(cod_referencia) & "^FS")
        '    Else
        '        printer.Print("^FO" & 35 + (a * 255) & ",2")
        '        printer.Print("^A0N,10,10,10^FD COMERCIALIZADO POR PLANET LOVE LTDA ^FS")
        '        printer.Print("^FO" & 170 + (a * 255) & ",12")
        '        printer.Print("^A0N,10,10^FD" & fecha_creacion(cod_referencia) & "^FS")
        '    End If
        '    printer.Print("^FO" & 15 + (a * 255) & ",25")
        '    printer.Print("^A0N,13,18^FD" & obtener_descripcion(codarticulo) & "^FS")
        '    printer.Print("^FO" & 15 + (a * 255) & ",38")
        '    printer.Print("^A0N,13,18^FD" & obtener_referencia(codarticulo) & "^FS")
        '    printer.Print("^FO" & 30 + (a * 255) & ",52^BY2")
        '    printer.Print("^BEN,25,Y,N")
        '    printer.Print("^FD" & generar_codigo_ean13(codarticulo) & "^FS")
        '    printer.Print("^FO" & 110 + (a * 255) & ",97")
        '    printer.Print("^A0N,20,30^FD" & FormatCurrency(hallar_precio_planet(codarticulo), 0) & "^FS")
        '    If a = 2 Then
        '        printer.Print("^XZ")
        '    Else
        '        If p = canttixy Then
        '            printer.Print("^XZ")
        '        End If
        '    End If
        '    a = a + 1
        '    If a > 2 Then
        '        a = 0
        '    End If
        'Next
        'printer.EndDoc()
    End Sub
#End Region

#Region "Metodos para impresora ZM4 300dpi"
    Sub imprimir_codigos_de_barra_19_x_15_300dpi(codarticulo As Long, cantplanet As Short, impresora As String)
        Dim cod_referencia As Long
        Dim a As Short = 0
        Dim p As Short

        cod_referencia = codarticulo
        impresora_a_utlilizar(impresora)

        For p = 1 To cantplanet
            If a = 0 Then printer.Print("^XA")
            clase.consultar("select* from registros_de_importaciones where articulo = " & codarticulo & "", "busqueda")
            If clase.dt.Tables("busqueda").Rows.Count > 0 Then
                printer.Print("^FO" & 90 + (a * 375) & ",3")
                printer.Print("^A0N,15,15,15^FD" & encabezado_codigos_barras() & "^FS")
                printer.Print("^FO" & 22 + (a * 410) & ",22")
                printer.Print("^A0N,15,17^FDREG IMP: ^FS")
                printer.Print("^FO" & 130 + (a * 375) & ",22")                                                                          '  Printer.Print("^FO15,30")
                printer.Print("^A0N,15,17^FD" & numero_registro_dian(cod_referencia) & "^FS")
                printer.Print("^FO" & 300 + (a * 375) & ",22")
                printer.Print("^A0N,15,17^FD" & fecha_registro_dian(cod_referencia) & "^FS")
            Else
                printer.Print("^FO" & 85 + (a * 375) & ",3")
                printer.Print("^A0N,15,15,15^FD COMERCIALIZADO POR PLANET LOVE LTDA ^FS")
                printer.Print("^FO" & 300 + (a * 375) & ",22")
                printer.Print("^A0N,15,15^FD" & fecha_creacion(cod_referencia) & "^FS")
            End If

            printer.Print("^FO" & 22 + (a * 410) & ",45")
            printer.Print("^A0N,20,27^FD" & obtener_descripcion(codarticulo) & "^FS")
            printer.Print("^FO" & 22 + (a * 410) & ",67")
            printer.Print("^A0N,20,25^FD" & obtener_referencia(codarticulo) & "^FS")
            printer.Print("^FO" & 85 + (a * 375) & ",90^BY3")
            printer.Print("^BEN,55,Y,N")
            printer.Print("^FD" & generar_codigo_ean13(codarticulo) & "^FS")
            printer.Print("^FO" & 238 + (a * 375) & ",180")
            printer.Print("^A0N,35,44^FD" & FormatCurrency(hallar_precio_planet(codarticulo), 0) & "^FS")
            If a = 2 Then
                printer.Print("^XZ")
            Else
                If p = cantplanet Then
                    printer.Print("^XZ")
                End If
            End If
            a = a + 1
            If a > 2 Then
                a = 0
            End If
        Next
        printer.EndDoc()
        'a = 0
        'For p = 1 To canttixy
        '    If a = 0 Then printer.Print("^XA")
        '    clase.consultar("select* from registros_de_importaciones where articulo = " & codarticulo & "", "busqueda")
        '    If clase.dt.Tables("busqueda").Rows.Count > 0 Then
        '        printer.Print("^FO" & 90 + (a * 375) & ",3")
        '        printer.Print("^A0N,15,15,15^FD" & encabezado_codigos_barras() & "^FS")
        '        printer.Print("^FO" & 82 + (a * 375) & ",22")
        '        printer.Print("^A0N,15,17^FDREG IMP: ^FS")
        '        printer.Print("^FO" & 130 + (a * 375) & ",22")                                                                          '  Printer.Print("^FO15,30")
        '        printer.Print("^A0N,15,17^FD" & numero_registro_dian(cod_referencia) & "^FS")
        '        printer.Print("^FO" & 300 + (a * 375) & ",22")
        '        printer.Print("^A0N,15,17^FD" & fecha_registro_dian(cod_referencia) & "^FS")
        '    Else
        '        printer.Print("^FO" & 85 + (a * 375) & ",3")
        '        printer.Print("^A0N,15,15,15^FD COMERCIALIZADO POR PLANET LOVE LTDA ^FS")
        '        printer.Print("^FO" & 300 + (a * 375) & ",22")
        '        printer.Print("^A0N,15,15^FD" & fecha_creacion(cod_referencia) & "^FS")
        '    End If
        '    printer.Print("^FO" & 82 + (a * 375) & ",45")
        '    printer.Print("^A0N,20,27^FD" & obtener_descripcion(codarticulo) & "^FS")
        '    printer.Print("^FO" & 82 + (a * 375) & ",67")
        '    printer.Print("^A0N,20,25^FD" & obtener_referencia(codarticulo) & "^FS")
        '    printer.Print("^FO" & 120 + (a * 375) & ",90^BY2")
        '    printer.Print("^BEN,55,Y,N")
        '    printer.Print("^FD" & generar_codigo_ean13(codarticulo) & "^FS")
        '    printer.Print("^FO" & 238 + (a * 375) & ",173")
        '    printer.Print("^A0N,30,44^FD" & FormatCurrency(hallar_precio_planet(codarticulo), 0) & "^FS")
        '    If a = 2 Then
        '        printer.Print("^XZ")
        '    Else
        '        If p = canttixy Then
        '            printer.Print("^XZ")
        '        End If
        '    End If
        '    a = a + 1
        '    If a > 2 Then
        '        a = 0
        '    End If
        'Next
        'printer.EndDoc()
    End Sub
#End Region

    Private Function obtener_referencia(articulo As Long) As String
        clase.consultar("select ar_referencia from articulos where ar_codigo = " & articulo & "", "referencia")
        Return clase.dt.Tables("referencia").Rows(0)("ar_referencia")
    End Function

    Private Function obtener_descripcion(articulo As Long) As String
        clase.consultar("select ar_descripcion from articulos where ar_codigo = " & articulo & "", "descripcion")
        Return clase.dt.Tables("descripcion").Rows(0)("ar_descripcion")
    End Function

    Private Function encabezado_marca_planetlove() As String
        clase.consultar("select* from informacion", "encabezado")
        Return clase.dt.Tables("encabezado").Rows(0)("encabezado_marca_planet")
    End Function

    Private Function encabezado_marca_tixy() As String
        clase.consultar("select* from informacion", "encabezado")
        Return clase.dt.Tables("encabezado").Rows(0)("encabezado_marca_tixy")
    End Function

    Private Function hallar_precio_planet(codigos As Long) As Double
        clase.consultar("SELECT ar_precio1 FROM articulos WHERE (ar_codigo =" & codigos & ")", "precio1")
        Return clase.dt.Tables("precio1").Rows(0)("ar_precio1")
    End Function

    Private Function hallar_precio_tixy(codigos As Long) As Double
        clase.consultar("SELECT ar_precio2 FROM articulos WHERE (ar_codigo =" & codigos & ")", "precio1")
        Return clase.dt.Tables("precio1").Rows(0)("ar_precio2")
    End Function

    Private Function obtener_precodigo_apartir_de_postcodigo(codigo As Long) As Long
        clase.consultar("SELECT asc_precodref FROM asociaciones_codigos WHERE (asc_postcodart =" & codigo & ")", "precodigo")
        If clase.dt.Tables("precodigo").Rows.Count > 0 Then
            Return clase.dt.Tables("precodigo").Rows(0)("asc_precodref")
        Else
            Return vbEmpty
        End If
    End Function

    Private Function encabezado_codigos_barras() As String
        clase.consultar("select* from informacion", "encabezado")
        Return clase.dt.Tables("encabezado").Rows(0)("encabezado_codigo_barras")
    End Function

    'Private Function numero_registro_dian(item_de_ref As Short) As String
    '    clase.consultar("SELECT detcab_registrodian FROM  detalle_importacion_detcajas WHERE (detcab_coditem =" & item_de_ref & ")", "numeroregistro")
    '    If clase.dt.Tables("numeroregistro").Rows.Count > 0 Then
    '        If IsDBNull(clase.dt.Tables("numeroregistro").Rows(clase.dt.Tables("numeroregistro").Rows.Count - 1)("detcab_registrodian")) Then
    '            Return ""
    '        Else
    '            Return clase.dt.Tables("numeroregistro").Rows(clase.dt.Tables("numeroregistro").Rows.Count - 1)("detcab_registrodian")
    '        End If
    '    Else
    '        Return ""
    '    End If
    'End Function

    Private Function numero_registro_dian(item_de_ref As Long) As String
        clase.consultar("SELECT registro FROM registros_de_importaciones WHERE (articulo =" & item_de_ref & ")", "numeroregistro")
        If clase.dt.Tables("numeroregistro").Rows.Count > 0 Then
            Return clase.dt.Tables("numeroregistro").Rows(clase.dt.Tables("numeroregistro").Rows.Count - 1)("registro")
        Else
            Return ""
        End If
    End Function
    '(LO de abajo de se puede eliminar una vez se pruebe la funcion que sigue) 
    'Private Function fecha_registro_dian(item_de_ref As Short) As String
    '    clase.consultar("SELECT detcab_fecharegistrodian FROM  detalle_importacion_detcajas WHERE (detcab_coditem =" & item_de_ref & ")", "numeroregistro")
    '    If clase.dt.Tables("numeroregistro").Rows.Count > 0 Then
    '        If IsDBNull(clase.dt.Tables("numeroregistro").Rows(clase.dt.Tables("numeroregistro").Rows.Count - 1)("detcab_fecharegistrodian")) Then
    '            Return ""
    '        Else
    '            Return clase.dt.Tables("numeroregistro").Rows(clase.dt.Tables("numeroregistro").Rows.Count - 1)("detcab_fecharegistrodian")
    '        End If
    '    Else
    '        Return ""
    '    End If
    'End Function

    Private Function fecha_registro_dian(item_de_ref As Long) As String
        clase.consultar("SELECT fecha FROM registros_de_importaciones WHERE (articulo =" & item_de_ref & ")", "fecharegistro")
        If clase.dt.Tables("fecharegistro").Rows.Count > 0 Then
            Dim fech As Date = clase.dt.Tables("fecharegistro").Rows(clase.dt.Tables("fecharegistro").Rows.Count - 1)("fecha")
            Return fech.ToString("dd/MM/yyyy")
        Else
            Return ""
        End If
    End Function

    Private Function fecha_creacion(articulo As Long) As String
        clase.consultar("SELECT ar_fechaingreso FROM articulos WHERE (ar_codigo =" & articulo & ")", "fecha")
        If clase.dt.Tables("fecha").Rows.Count > 0 Then
            Dim fech As Date = clase.dt.Tables("fecha").Rows(clase.dt.Tables("fecha").Rows.Count - 1)("ar_fechaingreso")
            Return fech.ToString("dd/MM/yyyy")
        Else
            Return ""
        End If
    End Function


    Private Function hallar_marca(item_de_ref As Short) As String
        clase.consultar("SELECT detcab_marca FROM detalle_importacion_detcajas WHERE (detcab_coditem =" & item_de_ref & ")", "numeroregistro")
        If clase.dt.Tables("numeroregistro").Rows.Count > 0 Then
            Return clase.dt.Tables("numeroregistro").Rows(0)("detcab_marca")
        Else
            Return "SIN MARCA"
        End If

    End Function

    Public Sub ImprimirCodigoPreispeccion(CodBarra As String, Impresora As String, Page As String)
        impresora_a_utlilizar(Impresora)
        printer.Print("^XA")
        printer.Print("^FO" & 180 & ",50^BY3")
        printer.Print("^AGN,^FDPLANET LOVE^FS")
        printer.Print("^FO" & 250 & ",110^BY3")
        printer.Print("^BCN,150,N,N,N")
        printer.Print("^FD" & CodBarra & "^FS")
        printer.Print("^FO" & 240 & ",270^BY3")
        printer.Print("^AGN,200,20")
        printer.Print("^FD" & CodBarra & "^FS")
        printer.Print("^FO" & 230 & ",430^BY3")
        printer.Print("^AGN,100,20")
        printer.Print("^FD" & Page & "^FS")
        printer.Print("^XZ")
        printer.EndDoc()
    End Sub
End Class
