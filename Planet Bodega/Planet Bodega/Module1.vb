Imports Microsoft.VisualBasic
Imports System.Drawing.Drawing2D
Imports System.Web.Mail
Module Module1
    Public ind_formulario As Boolean = False
    Public passtrue As Boolean = False
    Public cod_importacion As Short
    Public ind_cargador_lote As Boolean 'variable que me dice si el archivo lo cargue antes o lo acabo de hacer
    Public colores() As Integer
    Public cantidades() As Integer
    Public impuestos() As Short
    Public impuestos1() As Short
    Public impuestos2() As Short
    Public arraytallas() As Integer
    Public ind_creacion_colores As Boolean = False ' variable que me dice si ya cree los colores con cantidades para las referencias
    Public ind_edicion_colores As Boolean 'esta variable me dice si permites eliminar filas del datagridview
    Public codigos_de_barra() As Integer
    Public codigos() As String
    Dim clase As New class_library
    Public importacion As Integer
    Public salidamovimiento As Integer
    Public codiimportation As Integer
    Public cadena_string As String
    Public bodegaalmacenaje As Short
    Public tipoajuste As Integer
    Public conexion As New OleDb.OleDbConnection
    Public host As String
    Public user As String
    Public pass As String
    Public database As String
    Public indicador_de_formulario As Boolean = False ' esta variable me indica si el formulario crear articulos lo carga otro formaulario
    Public indicador_de_formulario2 As Boolean = False ' esta variable me indica si el formulario crear articulos lo carga otro formaulario
    Public ind_formulario_patron_distribucion = False ' este variable me indica si el formulario crear patrones de distribucion lo carga desde el formulario ingreso_de_mercancia_no_importada o desde el menu datos
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

    Function verificar_existencia_articulo(articulo As String) As Boolean
        clase.consultar("SELECT ar_codigo FROM articulos WHERE (ar_codigobarras ='" & articulo & "')", "articulo")
        If clase.dt.Tables("articulo").Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Sub inicializar_cadena_access()
        conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\detalle_entrada.mdb;Jet OLEDB:Database Password=PL456324;"
    End Sub

    Function campo_codigobarra_corto(tienda As Short) As String
        clase.consultar("SELECT tiposcodigosdebarra.campocodigocorto FROM tiendas INNER JOIN tiposcodigosdebarra ON (tiendas.tipocodigobarras = tiposcodigosdebarra.tipo) WHERE (tiendas.id =" & tienda & ")", "codigotienda")
        If clase.dt.Tables("codigotienda").Rows.Count > 0 Then
            Return clase.dt.Tables("codigotienda").Rows(0)("campocodigocorto")
        Else
            Return ""
        End If
    End Function

    Function campo_codigobarra_largo(tienda As Short) As String
        clase.consultar("SELECT tiposcodigosdebarra.campocodigolargo FROM tiendas INNER JOIN tiposcodigosdebarra ON (tiendas.tipocodigobarras = tiposcodigosdebarra.tipo) WHERE (tiendas.id =" & tienda & ")", "codigotienda")
        If clase.dt.Tables("codigotienda").Rows.Count > 0 Then
            Return clase.dt.Tables("codigotienda").Rows(0)("campocodigolargo")
        Else
            Return ""
        End If
    End Function

    

    Function hallar_email(tienda As Short) As String
        clase.consultar1("select email from tiendas where id = " & tienda & "", "email")
        Return clase.dt1.Tables("email").Rows(0)("email")
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

    Function empresa_tienda(tienda As Short) As Short
        clase.consultar("select empresa from tiendas where id = " & tienda & "", "tabla")
        Return clase.dt.Tables("tabla").Rows(0)("empresa")
    End Function

    Function consecutivo_empresa(empresa As Short) As Integer
        Dim complemento As String = ""
        Select Case empresa
            Case 1
                complemento = "tr_cons_planetlove"
            Case 2
                complemento = "tr_cons_inversiones"
            Case 3
                complemento = "tr_cons_comercializadora"
            Case 4
                complemento = "tr_cons_mivtrading"
        End Select
        clase.consultar("SELECT MAX(cabtransferencia." & complemento & ") AS maximo FROM cabtransferencia INNER JOIN tiendas ON (cabtransferencia.tr_destino = tiendas.id) INNER JOIN empresas ON (empresas.cod_empresa = tiendas.empresa) WHERE (empresas.cod_empresa =" & empresa & ")", "tabla1")
        If IsDBNull(clase.dt.Tables("tabla1").Rows(0)("maximo")) Then
            Return 0
        End If
        Return clase.dt.Tables("tabla1").Rows(0)("maximo")
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

    Function comprobar_nulidad(parametro As Object) As String ' esta funcion si es vacío devuelve nulo
        If parametro.trim = "" Then
            Return "NULL"
        Else
            Return "'" & parametro & "'"
        End If
    End Function

    Function comprobar_nulidad_de_integer(parametro As Object) As Integer
        If IsDBNull(parametro) Then
            Return 0
        Else
            Return parametro
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

    'esta funcion se puede utilizar para generar una consulta sql no necesariamente 
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
            ' MsgBox(codigos_de_barra(c))
        Next
        ind_creacion_colores = True
    End Sub

    Function nombre_talla(codigotalla) As String
        clase.consultar("select* from tallas where codigo_talla = " & codigotalla & "", "tablita")
        Return clase.dt.Tables("tablita").Rows(0)("nombretalla")
    End Function

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


    'Generate new image dimensions
    Public Function GenerateImageDimensions(ByVal currW As Integer, ByVal currH As Integer, ByVal destW As Integer, ByVal destH As Integer) As Size
        'double to hold the final multiplier to use when scaling the image
        Dim multiplier As Double = 0

        'string for holding layout
        Dim layout As String

        'determine if it's Portrait or Landscape
        If currH > currW Then
            layout = "portrait"
        Else
            layout = "landscape"
        End If

        Select Case layout.ToLower()
            Case "portrait"
                'calculate multiplier on heights
                If destH > destW Then
                    multiplier = CDbl(destW) / CDbl(currW)
                Else

                    multiplier = CDbl(destH) / CDbl(currH)
                End If
                Exit Select
            Case "landscape"
                'calculate multiplier on widths
                If destH > destW Then
                    multiplier = CDbl(destW) / CDbl(currW)
                Else

                    multiplier = CDbl(destH) / CDbl(currH)
                End If
                Exit Select
        End Select

        'return the new image dimensions
        Return New Size(CInt((currW * multiplier)), CInt((currH * multiplier)))
    End Function

    'Resize the image
    Public Sub SetImage(ByVal pb As PictureBox)
        Try
            'create a temp image
            Dim img As Image = pb.Image

            'calculate the size of the image
            Dim imgSize As Size = GenerateImageDimensions(img.Width, img.Height, pb.Width, pb.Height)

            'create a new Bitmap with the proper dimensions
            Dim finalImg As New Bitmap(img, imgSize.Width, imgSize.Height)

            'create a new Graphics object from the image
            Dim gfx As Graphics = Graphics.FromImage(img)

            'clean up the image (take care of any image loss from resizing)
            gfx.InterpolationMode = InterpolationMode.HighQualityBicubic

            'empty the PictureBox
            pb.Image = Nothing

            'center the new image
            pb.SizeMode = PictureBoxSizeMode.CenterImage

            'set the new image
            pb.Image = finalImg
        Catch e As System.Exception
            MessageBox.Show(e.Message)
        End Try
    End Sub
    'Sample usage

    Public Sub SetImage1(ByVal pb As DataGridViewImageCell)
        Try
            'create a temp image
            Dim img As Image = pb.Value
            'calculate the size of the image
            Dim imgSize As Size = GenerateImageDimensions(img.Width, img.Height, pb.Size.Width, 120)
            'create a new Bitmap with the proper dimensions
            Dim finalImg As New Bitmap(img, imgSize.Width, imgSize.Height)
            'create a new Graphics object from the image
            Dim gfx As Graphics = Graphics.FromImage(img)
            'clean up the image (take care of any image loss from resizing)
            gfx.InterpolationMode = InterpolationMode.HighQualityBicubic
            'empty the PictureBox
            pb.Value = Nothing
            'center the new image
            pb.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            'set the new image
            pb.Value = finalImg
        Catch e As System.Exception
            MessageBox.Show(e.Message)
        End Try
    End Sub


    'esta funcion me a genera una consulta, que me va a buscar las referencias en las ordenes de produccion para ver cuanto fue transferido
    Function buscar_codigos_a_apartir_de_la_referencia(ref As String, descri As String, linea As Integer, sublinea1 As Integer, sublinea2 As Object, sublinea3 As Object, sublinea4 As Object) As String
        Dim sub2 As String, sub3 As String, sub4 As String
        If IsDBNull(sublinea2) Then
            sub2 = " AND ar_sublinea2 IS NULL "
        Else
            sub2 = " AND ar_sublinea2 = " & sublinea2
        End If
        If IsDBNull(sublinea3) Then
            sub3 = " AND ar_sublinea3 IS NULL"
        Else
            sub3 = " AND ar_sublinea3 = " & sublinea3
        End If
        If IsDBNull(sublinea4) Then
            sub4 = " AND ar_sublinea4 IS NULL"
        Else
            sub4 = " AND ar_sublinea4 = " & sublinea4
        End If
        Return "select ar_codigo from articulos where ar_referencia = '" & ref & "' AND ar_descripcion = '" & descri & "' AND ar_linea = " & linea & " AND ar_sublinea1 = " & sublinea1 & "" & sub2 & sub3 & sub4
    End Function

    Function calcular_dañado_mercancia_no_almacenada_apartir_referencia(lugar As Short, importacion As Integer, ref As String, descri As String, linea As Integer, sublinea1 As Integer, sublinea2 As Object, sublinea3 As Object, sublinea4 As Object) As Integer
        clase.consultar(buscar_codigos_a_apartir_de_la_referencia(ref, descri, linea, sublinea1, sublinea2, sublinea3, sublinea4), "tabla3")
        If clase.dt.Tables("tabla3").Rows.Count > 0 Then
            Dim x As Short
            Dim sql As String = "SELECT SUM(detalle_baja_inventario.detbaja_cantidad) AS Canti FROM detalle_baja_inventario INNER JOIN cab_baja_inventario_mercancia_sin_almacenar ON (detalle_baja_inventario.detbaja_codigobaja = cab_baja_inventario_mercancia_sin_almacenar.binv_cod) WHERE ((cab_baja_inventario_mercancia_sin_almacenar.binv_importacion =" & importacion & ") AND ("
            Dim ind As Boolean = True
            For x = 0 To clase.dt.Tables("tabla3").Rows.Count - 1
                If ind = True Then
                    sql = sql & " (detalle_baja_inventario.detbaja_articulo = " & clase.dt.Tables("tabla3")(x)("ar_codigo") & ")"
                    ind = False
                Else
                    sql = sql & " OR (detalle_baja_inventario.detbaja_articulo = " & clase.dt.Tables("tabla3")(x)("ar_codigo") & ")"
                End If
            Next
            sql = sql & ") AND (cab_baja_inventario_mercancia_sin_almacenar.binv_lugar=" & lugar & "))"
            clase.consultar(sql, "tabla4")
            If clase.dt.Tables("tabla4").Rows.Count > 0 Then
                If IsDBNull(clase.dt.Tables("tabla4").Rows(0)("Canti")) Then
                    Return 0
                    Exit Function
                End If
                Return clase.dt.Tables("tabla4").Rows(0)("Canti")
            Else
                Return 0
            End If
        End If
    End Function

    Function convertir_codigobarra_a_codigo_normal(codigobarras As String) As Long
        Dim codigo As String = Mid(codigobarras, 1, 12)
        Dim x As Short
        Dim split As Integer
        For x = 1 To Len(codigobarras) - 3
            split = Mid(codigo, 2 + x, 1)
            If split <> 0 Then
                Exit For
            End If
        Next
        Return Mid(codigo, 2 + x, 12)
    End Function


    Function ruta_foto() As String
        clase.consultar("select* from informacion", "tbl")
        Return clase.dt.Tables("tbl").Rows(0)("carpeta_foto")
    End Function

    Function ruta_foto_factura() As String
        clase.consultar("select* from informacion", "tbl")
        Return clase.dt.Tables("tbl").Rows(0)("foto_factura")
    End Function

    Function ruta_google_drive() As String
        clase.consultar("select* from informacion", "tbl")
        Return clase.dt.Tables("tbl").Rows(0)("carpeta_google_drive")
    End Function

    Sub construir_archivo(transferencia As String)
        Try
            ' Donde guardamos los paths de los archivos que vamos a estar utilizando ..
            Dim PathArchivo As String = ""

            If System.IO.Directory.Exists("C:\Data") = False Then ' si no existe la carpeta se crea
                System.IO.Directory.CreateDirectory("C:\Data")
            End If
            clase.consultar("SELECT cabtransferencia.tr_fecha, cabtransferencia.tr_destino, tiendas.codigonegocio, tiendas.email FROM cabtransferencia INNER JOIN tiendas  ON (cabtransferencia.tr_destino = tiendas.id) WHERE (cabtransferencia.tr_numero =" & transferencia & ")", "tabla1")
            Dim destinofoxpos As String = clase.dt.Tables("tabla1").Rows(0)("codigonegocio")
            Dim emailtienda As String = clase.dt.Tables("tabla1").Rows(0)("email")
            Dim fechatransferencia As Date = clase.dt.Tables("tabla1").Rows(0)("tr_fecha")
            Dim tiendapdv As String = clase.dt.Tables("tabla1").Rows(0)("tr_destino")
            clase.consultar("SELECT empresas.codigo_bodega FROM empresas INNER JOIN tiendas ON (empresas.cod_empresa = tiendas.empresa) WHERE (tiendas.id =" & tiendapdv & ")", "tabla2")
            Dim bodegaempresa As String = clase.dt.Tables("tabla2").Rows(0)("codigo_bodega")
            Select Case Len(destinofoxpos)
                Case 1
                    PathArchivo = "F00" & destinofoxpos
                Case 2
                    PathArchivo = "F0" & destinofoxpos
                Case 3
                    PathArchivo = "F" & destinofoxpos
            End Select
            Dim transfor As String = ""
            Select Case Len(transferencia)
                Case 1
                    transfor = "000" & transferencia
                Case 2
                    transfor = "00" & transferencia
                Case 3
                    transfor = "0" & transferencia
                Case Is > 3
                    transfor = Microsoft.VisualBasic.Right(transferencia, 4)
            End Select
            PathArchivo = PathArchivo & transfor & ".TXT"

            Dim sql, sql1 As String
            If empresa_tienda(tiendapdv) = 3 Then
                sql = "SELECT dt_codarticulo, dt_cantidad, dt_costo, dt_venta2 AS precioventa FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ")"
                sql1 = "SELECT SUM(dt_costo * dt_cantidad) AS costo, SUM(dt_venta2 * dt_cantidad) AS venta FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ")"
            Else
                sql = "SELECT dt_codarticulo, dt_cantidad, dt_costo, dt_venta1  AS precioventa FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ")"
                sql1 = "SELECT SUM(dt_costo * dt_cantidad) AS costo, SUM(dt_venta1 * dt_cantidad) AS venta FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ")"
            End If
            clase.consultar(sql1, "tabla3")
            Dim totalcostotransferencia As String
            If IsDBNull(clase.dt.Tables("tabla3").Rows(0)("costo")) Then
                totalcostotransferencia = 0
            Else
                totalcostotransferencia = clase.dt.Tables("tabla3").Rows(0)("costo")
            End If
            Dim totalventatransferencia As String
            If IsDBNull(clase.dt.Tables("tabla3").Rows(0)("venta")) Then
                totalventatransferencia = 0
            Else
                totalventatransferencia = clase.dt.Tables("tabla3").Rows(0)("venta")
            End If
            clase.consultar(sql, "tabla4")
            If clase.dt.Tables("tabla4").Rows.Count > 0 Then
                Dim x As Short
                clase.borradoautomatico("delete from transferencia_foxpos where nro_transferencia = '" & colocar_espacios_vacios(transferencia, 7) & "'")
                For x = 0 To clase.dt.Tables("tabla4").Rows.Count - 1
                    With clase.dt.Tables("tabla4")
                        clase.agregar_registro("INSERT INTO `transferencia_foxpos`(`nro_transferencia`,`fecha`,`negenvia`,`negrecibe`,`costo`,`venta`,`codarticulo`,`cant`,`costounitario`,`ventaunitario`) VALUES ( '" & colocar_espacios_vacios(transferencia, 7) & "','" & fechatransferencia.ToString("yyyyMMdd") & "','" & colocar_espacios_vacios(bodegaempresa, 3) & "','" & colocar_espacios_vacios(destinofoxpos, 3) & "','" & colocar_espacios_vacios(Str(totalcostotransferencia), 11) & "','" & colocar_espacios_vacios(Str(totalventatransferencia), 9) & "','" & colocar_espacios_vacios(.Rows(x)("dt_codarticulo"), 7) & "','" & colocar_espacios_vacios(.Rows(x)("dt_cantidad") & ".00", 9) & "','" & colocar_espacios_vacios(Str(.Rows(x)("dt_costo")), 10) & "','" & colocar_espacios_vacios(Str(.Rows(x)("precioventa")), 8) & "')")
                    End With
                Next
                If System.IO.File.Exists("C:\Data\ARCHIVO.TXT") = True Then 'verifico si existe
                    System.IO.File.Delete("C:\Data\ARCHIVO.TXT") 'si exite lo elimino
                End If

                clase.consultaraccess("SELECT* INTO [TEXT;DATABASE=C:\Data].ARCHIVO.TXT FROM transferencia_foxpos where nro_transferencia = '" & colocar_espacios_vacios(transferencia, 7) & "'", "tablita") 'aqui creo el archivo
                If System.IO.File.Exists("C:\Data\" & PathArchivo) = True Then
                    System.IO.File.Delete("C:\Data\" & PathArchivo)
                End If
                System.IO.File.Move("C:\Data\ARCHIVO.TXT", "C:\Data\" & PathArchivo)
                enviar_transferencia_x_correo(emailtienda, "NOVEDAD DE TRANSFERENCIA", "auxiliarbodega@planetloveonline.com", "NUMERO DE TRANSFERENCIA: " & transferencia, "C:\Data\" & PathArchivo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function colocar_espacios_vacios(expresion As String, max As Short) As String
        Dim cant As Short = Len(expresion)
        If cant > max Then
            cant = max
        End If
        Dim i As Short = max - cant
        If i > 0 Then
            Dim a As Short
            For a = 1 To i
                expresion = " " & expresion
            Next
            Return expresion
        Else
            Return expresion
        End If
    End Function

    'hay 2 funciones para el envío de archivos este es uno y acontinuacion esta el otro
    Sub enviar_transferencia_x_correo(destinatario As String, Asunto As String, remitente As String, mensaje As String, file As String)
        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        clase.consultar("select* from informacion", "inf")
        With smtp
            .Port = clase.dt.Tables("inf").Rows(0)("port")
            .Host = clase.dt.Tables("inf").Rows(0)("host")
            .Credentials = New System.Net.NetworkCredential(clase.dt.Tables("inf").Rows(0)("user"), clase.dt.Tables("inf").Rows(0)("password"))
            .EnableSsl = False
        End With
        adjunto = New System.Net.Mail.Attachment(file)
        With correo
            .From = New System.Net.Mail.MailAddress(remitente)
            .To.Add(destinatario)
            .Subject = Asunto
            .Body = "<strong>" & mensaje & "</strong>"
            .IsBodyHtml = True
            .Priority = System.Net.Mail.MailPriority.Normal
            .Attachments.Add(adjunto)
        End With
        Try
            smtp.Send(correo)
            '   MessageBox.Show("Su mensaje de correo ha sido enviado.", _
            '      "Correo enviado", _
            ' MessageBoxButtons.OK)
        Catch ex As Exception
            ' MessageBox.Show("Error: " & ex.Message, _
            '                "Error al enviar correo", _
            '                 MessageBoxButtons.OK)
        End Try

    End Sub

    Sub enviar_actualizacion_x_correo(destinatario As String, Asunto As String, remitente As String, mensaje As String, file As String)
        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        clase.consultar("select* from informacion", "inf")
        With smtp
            .Port = clase.dt.Tables("inf").Rows(0)("port")
            .Host = clase.dt.Tables("inf").Rows(0)("host")
            .Credentials = New System.Net.NetworkCredential(clase.dt.Tables("inf").Rows(0)("user"), clase.dt.Tables("inf").Rows(0)("password"))
            .EnableSsl = False
        End With
        adjunto = New System.Net.Mail.Attachment(file)
        With correo
            .From = New System.Net.Mail.MailAddress(remitente)
            .To.Add(destinatario)
            .Subject = Asunto
            .Body = "<strong>" & mensaje & "</strong>"
            .IsBodyHtml = True
            .Priority = System.Net.Mail.MailPriority.Normal
            .Attachments.Add(adjunto)
        End With
        Try
            smtp.Send(correo)
            'MessageBox.Show("Su mensaje de correo ha sido enviado.", _
            '              "Correo enviado", _
            '             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, _
                            "Error al enviar correo", _
                            MessageBoxButtons.OK)
        End Try
    End Sub


    Sub imprimir_hoja_transferencia(transferencia As Integer)
        Dim m_excel As Microsoft.Office.Interop.Excel.Application
        m_excel = CreateObject("Excel.Application")
        m_excel.Workbooks.Open("C:\Data\transferencia.xls")
        m_excel.Visible = False
        clase.consultar("select* from cabtransferencia where tr_numero = " & transferencia & "", "tabla4")
        m_excel.Worksheets("Hoja1").cells(3, 1).value = "TRANSFERENCIA DE BODEGA No: " & clase.dt.Tables("tabla4").Rows(0)("tr_codigolovepos") & "                          CODIGO BODEGA: " & transferencia
        Dim destino As Short = clase.dt.Tables("tabla4").Rows(0)("tr_destino")
        Dim operador As String = clase.dt.Tables("tabla4").Rows(0)("tr_operador")
        Dim fecha As Date = clase.dt.Tables("tabla4").Rows(0)("tr_fecha")
        m_excel.Worksheets("Hoja1").cells(4, 1).value = "REMITENTE: BODEGA " & empresa(destino)
        m_excel.Worksheets("Hoja1").cells(5, 1).value = "DESTINATARIO: " & nombretienda(destino)
        m_excel.Worksheets("Hoja1").cells(6, 1).value = "GENERADO POR: " & operador
        m_excel.Worksheets("Hoja1").cells(1, 5).value = "FECHA: " & fecha.ToString("dd/MM/yyyy")
        Dim sql, sql1 As String
        Dim campocodigo As String = campo_codigobarra_corto(destino)
        If empresa_tienda(destino) = 3 Then
            sql = "SELECT SUM(dt_cantidad * dt_costo) AS Costo, SUM(dt_cantidad * dt_venta2) AS Venta FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ")"
            sql1 = "SELECT articulos." & campocodigo & ", articulos.ar_referencia, articulos.ar_descripcion, dettransferencia.dt_cantidad, ar_costo * dt_cantidad AS Costo, dt_venta2 * dt_cantidad AS Venta FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) WHERE (dettransferencia.dt_trnumero =" & transferencia & ") ORDER BY dettransferencia.dt_codarticulo ASC "
        Else
            sql = "SELECT SUM(dt_cantidad * dt_costo) AS Costo, SUM(dt_cantidad * dt_venta1) AS Venta FROM dettransferencia WHERE (dt_trnumero =" & transferencia & ")"
            sql1 = "SELECT articulos." & campocodigo & ", articulos.ar_referencia, articulos.ar_descripcion, dettransferencia.dt_cantidad, ar_costo * dt_cantidad AS Costo, dt_venta1 * dt_cantidad AS Venta FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) WHERE (dettransferencia.dt_trnumero =" & transferencia & ") ORDER BY dettransferencia.dt_codarticulo ASC "
        End If
        clase.consultar(sql, "tabla3")
        If IsDBNull(clase.dt.Tables("tabla3").Rows(0)("Costo")) Then
            m_excel.Worksheets("Hoja1").cells(4, 5).value = "TOTAL COSTO: " & FormatCurrency(0, 0)
        Else
            m_excel.Worksheets("Hoja1").cells(4, 5).value = "TOTAL COSTO: " & FormatCurrency(clase.dt.Tables("tabla3").Rows(0)("Costo"), 0)
        End If
        If IsDBNull(clase.dt.Tables("tabla3").Rows(0)("Venta")) Then
            m_excel.Worksheets("Hoja1").cells(5, 5).value = "TOTAL VENTA: " & FormatCurrency(0, 0)
        Else
            m_excel.Worksheets("Hoja1").cells(5, 5).value = "TOTAL VENTA: " & FormatCurrency(clase.dt.Tables("tabla3").Rows(0)("Venta"), 0)
        End If
        clase.consultar(sql1, "tabla5")
        If clase.dt.Tables("tabla5").Rows.Count > 0 Then
            With clase.dt.Tables("tabla5")
                Dim x As Integer
                For x = 0 To clase.dt.Tables("tabla5").Rows.Count - 1
                    m_excel.Worksheets("Hoja1").cells(9 + x, 1).value = .Rows(x)(campocodigo)
                    m_excel.Worksheets("Hoja1").cells(9 + x, 2).value = .Rows(x)("ar_referencia")
                    m_excel.Worksheets("Hoja1").cells(9 + x, 3).value = .Rows(x)("ar_descripcion")
                    m_excel.Worksheets("Hoja1").cells(9 + x, 4).value = .Rows(x)("dt_cantidad")
                    m_excel.Worksheets("Hoja1").cells(9 + x, 5).value = FormatCurrency(.Rows(x)("Costo"), 0)
                    m_excel.Worksheets("Hoja1").cells(9 + x, 6).value = FormatCurrency(.Rows(x)("Venta"), 0)
                Next
                m_excel.Worksheets("Hoja1").cells(9 + x + 5, 1).value = "                                        Firma Digitador"
                m_excel.Worksheets("Hoja1").cells(9 + x + 5, 3).value = "             Firma Revisor Bodega"
                m_excel.Worksheets("Hoja1").cells(9 + x + 5, 5).value = "                        Firma Revisor Tienda"
                m_excel.Worksheets("Hoja1").cells(9 + x + 5, 1).Font.Bold = True
                m_excel.Worksheets("Hoja1").cells(9 + x + 5, 3).Font.Bold = True
                m_excel.Worksheets("Hoja1").cells(9 + x + 5, 5).Font.Bold = True
                m_excel.Application.ActiveWorkbook.PrintOutEx()
                If Not m_excel Is Nothing Then
                    m_excel.Application.ActiveWorkbook.Saved = True
                    m_excel.Quit()
                    m_excel = Nothing
                End If
            End With
        End If
    End Sub

    Function empresa(tienda As Short) As String
        clase.consultar("SELECT empresas.nombre_empresa FROM empresas INNER JOIN tiendas ON (empresas.cod_empresa = tiendas.empresa) WHERE (tiendas.id =" & tienda & ")", "tabla2")
        Return clase.dt.Tables("tabla2").Rows(0)("nombre_empresa")
    End Function

    Function nombretienda(codigo As Short) As String
        clase.consultar("select* from tiendas where id = " & codigo & "", "tabla2")
        Return clase.dt.Tables("tabla2").Rows(0)("tienda")
    End Function

    Sub crear_foto_grande(origen As String, destino As String)

        If System.IO.File.Exists(destino) = True Then
            System.IO.File.Delete(destino)
        End If
        System.IO.File.Copy(origen, destino)
    End Sub

    Sub crear_foto_pequeño(rutafotogrande As String, rutafotopequeña As String)
        If System.IO.File.Exists(rutafotopequeña) = True Then
            System.IO.File.Delete(rutafotopequeña)
        End If
        Dim ruta2 As String = rutafotopequeña
        Dim ruta As String = rutafotogrande
        Dim imagen As New Bitmap(New Bitmap(ruta), 200, 149)
        imagen.Save(ruta2, System.Drawing.Imaging.ImageFormat.Jpeg)
    End Sub

    Sub sumar_al_inventario(articulo As Long, cant As Integer)
        Dim totinventario As Integer
        Dim transferido As Integer
        clase.consultar2("SELECT inv_cantidad, inv_transferido FROM inventario_bodega WHERE (inv_codigoart =" & articulo & ")", "sumar")
        If clase.dt2.Tables("sumar").Rows.Count > 0 Then
            totinventario = clase.dt2.Tables("sumar").Rows(0)("inv_cantidad")
            transferido = comprobar_nulidad_de_integer(clase.dt2.Tables("sumar").Rows(0)("inv_transferido"))
            clase.actualizar("UPDATE inventario_bodega SET inv_cantidad = " & totinventario + cant & ", inv_transferido = " & transferido - cant & " WHERE (inv_codigoart =" & articulo & ")")
        Else
            clase.agregar_registro("INSERT INTO `inventario_bodega`(`inv_codigo`,`inv_codigoart`,`inv_transferido`,`inv_cantidad`) VALUES ( NULL,'" & articulo & "','" & -1 * cant & "','" & cant & "')")
        End If
    End Sub


    Function CalcularCodigoTransferenciaSaliente(IdEmpresaDestino As Integer, IdTiendaDestino As Integer) As String
        clase.consultar("Select MigradoLovePos From tiendas Where id = " & IdTiendaDestino & "", "migrado")
        If clase.dt.Tables("migrado").Rows(0)("MigradoLovePos") = True Then
            Dim CodigoTransferencia As String = ""
            clase.ConsultarSQLServer("SELECT [Consecutivo] FROM [dbo].[ConsecutivosDocInventarioTienda] WHERE IdTienda = 99 AND IdDocumentoInventario = 'TRS' AND IdEmpresaDestino = " & IdEmpresaDestino & " AND IdTiendaDestino = " & IdTiendaDestino & "", "trans")
            If clase.dtSql.Tables("trans").Rows.Count > 0 Then
                CodigoTransferencia = "TRS-099" & IdTiendaDestino.ToString("000") & FormatearValor(clase.dtSql.Tables("trans").Rows(0)("Consecutivo") + 1)
            Else
                CodigoTransferencia = "TRS-099" & IdTiendaDestino.ToString("000") & "0001"
            End If
            Return CodigoTransferencia
        Else
            Return ""
        End If
    End Function



    Private Function FormatearValor(Valor As String)
        Select Case Valor.Length
            Case 1
                Return "000" + Valor
            Case 2
                Return "00" + Valor
            Case 3
                Return "0" + Valor
            Case Is < 3
                Return Valor
        End Select
    End Function


    Function GetCodigoEmpresaByTienda(IdTienda As Integer)
        clase.ConsultarSQLServer("Select IdEmpresa From [dbo].[Tienda] Where IdTienda = " & IdTienda & "", "emp")
        Return clase.dtSql.Tables("emp").Rows(0)("IdEmpresa")
    End Function

    Function VerificarExistenciaArticuloLovePOS(Codigo As Integer) As Boolean
        clase.ConsultarSQLServer("SELECT        IdArticulo FROM Articulo WHERE (IdArticulo = " & Codigo & ")", "articulos")
        If (clase.dtSql.Tables("articulos").Rows.Count > 0) Then
            Return True
        Else
            Return False
        End If
    End Function

    Sub ActualizarConsetuvosDocumentosInventarios(IdEmpresaDestino As Integer, IdTiendaDestino As Integer, CodigoLovePOS As String)
        clase.ConsultarSQLServer("SELECT [Consecutivo] FROM [dbo].[ConsecutivosDocInventarioTienda] WHERE IdTienda = 99 AND IdDocumentoInventario = 'TRS' AND IdEmpresaDestino = " & IdEmpresaDestino & " AND IdTiendaDestino = " & IdTiendaDestino & "", "trans")
        If clase.dtSql.Tables("trans").Rows.Count = 0 Then
            clase.AgregarSQLServer("INSERT INTO [dbo].[ConsecutivosDocInventarioTienda] ([IdEmpresa],[IdTienda],[IdDocumentoInventario],[IdCodMovimiento],[IdEmpresaDestino],[IdTiendaDestino],[Consecutivo],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('100','99','TRS','" & "TRS-099" & IdTiendaDestino.ToString("000") & "0002" & "','" & IdEmpresaDestino & "','" & IdTiendaDestino & "','2','AUTOMATICO','Planet Bodega','AUTOMATICO','" & Now & "')")
        Else
            Dim consecutivo As Integer = clase.dtSql.Tables("trans").Rows(0)("Consecutivo")
            clase.ActualizarSQLServer("UPDATE [dbo].[ConsecutivosDocInventarioTienda] SET Consecutivo = " & consecutivo + 1 & ", IdCodMovimiento = '" & CodigoLovePOS & "' WHERE IdTienda = 99 AND IdDocumentoInventario = 'TRS' AND IdEmpresaDestino = " & IdEmpresaDestino & " AND IdTiendaDestino = " & IdTiendaDestino & "")
        End If

    End Sub

    Function GuardarMovimientoInventarioEnLovePOS(IdTransferencia As Integer, IdCodMovimientoLovePOS As String) As Boolean
        clase.consultar("SELECT* FROM cabtransferencia WHERE tr_numero = " & IdTransferencia & "", "encabezado")
        clase.consultar1("SELECT* FROM dettransferencia WHERE dt_trnumero = " & IdTransferencia & "", "detalle")
        With clase.dt.Tables("encabezado")
            'encabezado de transferecia saliente
            clase.AgregarSQLServer("INSERT INTO [dbo].[CabMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[FechaOperativa],[IdEmpresaDestino]," &
                                   "[IdTiendaDestino],[IdBodegaDestino],[Operario],[Revisor],[Finalizado],[Revisada],[Recibido],[FechaRecibo],[UsuarioModif],[ProgramaModif],[Observaciones]," &
                                   "[EquipoModif],[FechaModif],[Anulado]) VALUES('100','99','1','TRS','" & IdCodMovimientoLovePOS & "','" & .Rows(0)("tr_fecha") & "'," &
                                   "'" & GetCodigoEmpresaByTienda(.Rows(0)("tr_destino")) & "','" & .Rows(0)("tr_destino") & "','1','" & .Rows(0)("tr_operador") & "','" & .Rows(0)("tr_revisor") & "','TRUE','TRUE'," &
                                   "'FALSE','01/01/1900 00:00:00','Admin','Planet Bodega','','MAQUINAUBUNTU'," &
                                   "'" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','0')")
            'encabezado de transferenca de entrante
            clase.AgregarSQLServer("INSERT INTO [dbo].[CabMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[FechaOperativa],[IdEmpresaDestino]," &
                                   "[IdTiendaDestino],[IdBodegaDestino],[Operario],[Revisor],[Finalizado],[Revisada],[Recibido],[FechaRecibo],[UsuarioModif],[ProgramaModif],[Observaciones]," &
                                   "[EquipoModif],[FechaModif],[Anulado]) VALUES('" & GetCodigoEmpresaByTienda(.Rows(0)("tr_destino")) & "','" & .Rows(0)("tr_destino") & "','1','TRR','" & IdCodMovimientoLovePOS & "','" & .Rows(0)("tr_fecha") & "'," &
                                   "'0','0','0','" & .Rows(0)("tr_operador") & "','" & .Rows(0)("tr_revisor") & "','TRUE','TRUE'," &
                                   "'FALSE','01/01/1900 00:00:00','Admin','Planet Bodega','','MAQUINAUBUNTU'," &
                                   "'" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','0')")
        End With
        With clase.dt1.Tables("detalle")
            Dim a As Integer
            For a = 0 To .Rows.Count - 1
                'detalle de transferencia saliente
                clase.AgregarSQLServer("INSERT INTO [dbo].[DetMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[IdArticulo],[Cantidad],[PrecioCosto],[PrecioVenta]" &
                                       ",[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif],[PrecioCostoFiscal]) VALUES ('100','99','1','TRS','" & IdCodMovimientoLovePOS & "','" & .Rows(a)("dt_codarticulo") & "'," &
                                       "'" & .Rows(a)("dt_cantidad") & "','" & Str(.Rows(a)("dt_costo2")) & "','" & Str(.Rows(a)("dt_venta1")) & "','Admin','Planet Bodega','MAQUINAUBUNTU'" &
                                       ",'" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Str(.Rows(a)("dt_costo")) & "')")
                'detalle de transferencia entrante
                clase.AgregarSQLServer("INSERT INTO [dbo].[DetMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[IdArticulo],[Cantidad],[PrecioCosto],[PrecioVenta]" &
                                 ",[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif],[PrecioCostoFiscal]) VALUES ('" & GetCodigoEmpresaByTienda(clase.dt.Tables("encabezado").Rows(0)("tr_destino")) & "','" & clase.dt.Tables("encabezado").Rows(0)("tr_destino") & "','1','TRR','" & IdCodMovimientoLovePOS & "','" & .Rows(a)("dt_codarticulo") & "'," &
                                 "'" & .Rows(a)("dt_cantidad") & "','" & Str(.Rows(a)("dt_costo2")) & "','" & Str(.Rows(a)("dt_venta1")) & "','Admin','Planet Bodega','MAQUINAUBUNTU'" &
                                 ",'" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Str(.Rows(a)("dt_costo")) & "')")
            Next
            ConsolidarSaldos(100, 99, 1, "TRS", IdCodMovimientoLovePOS)
        End With
        Return True
    End Function

    Function GetDocumentoDeAnulacion(IdTipoDocumento As String) As String
        clase.ConsultarSQLServer("SELECT [CodDocumento] FROM [dbo].[TipoDocumentosInventarios] WHERE DocQueAnula = '" & IdTipoDocumento & "'", "sql")
        Return clase.dtSql.Tables("sql").Rows(0)("CodDocumento")
    End Function

    Public Sub AnularMovimientoDeInventario(IdEmpresa As Integer, IdTienda As Integer, IdBodega As Integer, IdTipoMovimiento As String, IdCodMovimiento As String)
        Dim cadena As String = "SELECT        dt.IdArticulo, (dt.Cantidad * tp.OperacionDocumento) As Cantidad, dt.PrecioCosto, dt.PrecioCostoFiscal, dt.PrecioVenta, cb.Operario, cb.Revisor FROM  CabMovimientoInventario cb INNER JOIN DetMovimientoInventario dt On cb.IdEmpresa = dt.IdEmpresa And cb.IdTienda = dt.IdTienda And " &
          "cb.IdBodega = dt.IdBodega And cb.IdTipoMovimiento = dt.IdTipoMovimiento And cb.IdCodMovimiento = dt.IdCodMovimiento INNER JOIN TipoDocumentosInventarios tp On cb.IdTipoMovimiento = tp.CodDocumento " &
          "WHERE (cb.IdEmpresa = " & IdEmpresa & ") And (cb.IdTienda = " & IdTienda & ") And (cb.IdBodega = " & IdBodega & ") And (cb.IdTipoMovimiento = '" & IdTipoMovimiento & "') AND (cb.IdCodMovimiento = '" & IdCodMovimiento & "')"
        clase.ConsultarSQLServer1(cadena, "saldos")
        If clase.dtSql1.Tables("saldos").Rows.Count > 0 Then
            Dim Cant As Integer = 0
            Dim x As Integer
            Dim CodTipoAnulacion As String = GetDocumentoDeAnulacion(IdTipoMovimiento)

            clase.AgregarSQLServer("INSERT INTO [dbo].[CabMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[FechaOperativa],[IdEmpresaDestino]," &
                                   "[IdTiendaDestino],[IdBodegaDestino],[Operario],[Revisor],[Finalizado],[Revisada],[Recibido],[FechaRecibo],[UsuarioModif],[ProgramaModif],[Observaciones]," &
                                   "[EquipoModif],[FechaModif],[Anulado]) VALUES('" & IdEmpresa & "','" & IdTienda & "','" & IdBodega & "','" & CodTipoAnulacion & "','" & IdCodMovimiento & "','" & Now.ToString("dd/MM/yyyy") & "'," &
                                   "'0','0','0','" & clase.dtSql1.Tables("saldos").Rows(0)("Operario") & "','" & clase.dtSql1.Tables("saldos").Rows(0)("Revisor") & "','TRUE','TRUE'," &
                                   "'FALSE','01/01/1900 00:00:00','Admin','Planet Bodega','','MAQUINAUBUNTU'," &
                                   "'" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','0')")

            For x = 0 To clase.dtSql1.Tables("saldos").Rows.Count - 1
                With clase.dtSql1.Tables("saldos")
                    Cant = clase.dtSql1.Tables("saldos").Rows(x)("Cantidad") * -1
                    clase.AgregarSQLServer("INSERT INTO [dbo].[DetMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[IdArticulo],[Cantidad],[PrecioCosto],[PrecioVenta]" &
                                 ",[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif],[PrecioCostoFiscal]) VALUES ('" & IdEmpresa & "','" & IdTienda & "','" & IdBodega & "','" & CodTipoAnulacion & "','" & IdCodMovimiento & "','" & .Rows(x)("IdArticulo") & "'," &
                                 "'" & Cant & "','" & Str(.Rows(x)("PrecioCosto")) & "','" & Str(.Rows(x)("PrecioVenta")) & "','Admin','Planet Bodega','MAQUINAUBUNTU'" &
                                 ",'" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Str(.Rows(x)("PrecioCostoFiscal")) & "')")
                End With
            Next
            ConsolidarSaldos(IdEmpresa, IdTienda, IdBodega, CodTipoAnulacion, IdCodMovimiento)
            clase.ActualizarSQLServer("UPDATE [dbo].[CabMovimientoInventario] SET Anulado = '1', FechaModif = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE IdEmpresa = " & IdEmpresa & " AND IdTienda = " & IdTienda & " AND IdBodega = " & IdBodega & " AND IdTipoMovimiento = '" & IdTipoMovimiento & "' AND IdCodMovimiento = '" & IdCodMovimiento & "'")
            clase.ActualizarSQLServer("DELETE FROM DetMovimientoInventario Where IdTipoMovimiento = 'TRR' and IdCodMovimiento = '" & IdCodMovimiento & "'")
            clase.ActualizarSQLServer("DELETE FROM CabMovimientoInventario Where IdTipoMovimiento = 'TRR' and IdCodMovimiento = '" & IdCodMovimiento & "'")
        End If
    End Sub

    Sub SubirTransferenciaLovePos(revisor As String, codTransferencia As Integer, CodDestino As Integer)
        Dim CodigoMovLovePOS As String = CalcularCodigoTransferenciaSaliente(GetCodigoEmpresaByTienda(CodDestino), CodDestino)
        If CodigoMovLovePOS <> "" Then
            clase.actualizar("update cabtransferencia set tr_revisor = '" & revisor & "', tr_revisada = TRUE, tr_finalizada = TRUE, tr_codigolovepos = '" & CodigoMovLovePOS & "' where tr_numero = " & codTransferencia & "")
            GuardarMovimientoInventarioEnLovePOS(codTransferencia, CodigoMovLovePOS)
            ActualizarConsetuvosDocumentosInventarios(GetCodigoEmpresaByTienda(CodDestino), CodDestino, CodigoMovLovePOS)
        Else
            clase.actualizar("update cabtransferencia set tr_revisor = '" & revisor & "', tr_revisada = TRUE, tr_finalizada = TRUE  where tr_numero = " & codTransferencia & "")
        End If
    End Sub


    Public Sub CrearMovimientoCompraLovePOS1(IdEntrada As Integer, Operario As String)
        clase.consultar1("select* from detalleimportacioncompras where IdImportacion = " & IdEntrada & "", "importacion")
        If clase.dt1.Tables("importacion").Rows.Count > 0 Then
            'Cuando la tabla detalleimportacioncompras esta llena

        Else
            'cuando la tabla detalleimportacioncompras esta vacía

        End If
    End Sub


    Function ObtenerCostoReal(IdArticulo As Integer) As Double
        clase.consultar1("SELECT ar_costo2 FROM articulos where ar_codigo = " & IdArticulo & "", "costo")
        Return Math.Round(clase.dt1.Tables("costo").Rows(0)("ar_costo2"), 2)
    End Function

    Function ObtenerCostoFiscal(IdArticulo As Integer) As Double
        clase.consultar1("SELECT ar_costo FROM articulos where ar_codigo = " & IdArticulo & "", "costo")
        Return Math.Round(clase.dt1.Tables("costo").Rows(0)("ar_costo"), 2)
    End Function

    Function ObtenerPrecioVenta(IdArticulo As Integer) As Double
        clase.consultar1("SELECT ar_precio1 FROM articulos where ar_codigo = " & IdArticulo & "", "precio")
        Return clase.dt1.Tables("precio").Rows(0)("ar_precio1")
    End Function

    Public Sub CrearMovimientoCompraLovePOSImportacion(IdImportacion As Integer, IdProveedor As Integer, IdArticulo As Integer, Cant As Integer, Operario As String)
        MigrarArticulosLineasProveedores()
        clase.consultar1("SELECT* FROM detalleimportacioncompras WHERE IdImportacion = " & IdImportacion & " AND IdProveedor = " & IdProveedor & "", "importacion")
        Dim codigoLovePosMovimiento As String
        If clase.dt1.Tables("importacion").Rows.Count = 0 Then
            codigoLovePosMovimiento = CalcularCodigoLovePOSCompra()

            clase.AgregarSQLServer("INSERT INTO [dbo].[CabMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[FechaOperativa],[Operario],[Revisor]" &
                                            ",[Finalizado],[Revisada],[Recibido],[UsuarioModif],[ProgramaModif],[Observaciones],[EquipoModif],[FechaModif],[IdEmpresaDestino],[IdTiendaDestino],[IdBodegaDestino],[FechaRecibo],[IdProveedor],[IdEntrada])" &
        "VALUES ('100','99','1','COM','" & codigoLovePosMovimiento & "','" & Now.ToString("dd/MM/yyyy") & "'" &
        ",'" & Operario & "','" & Operario & "','TRUE','TRUE','TRUE','Automatico'" &
        ",'Automatico','','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','0','0','0','01/01/1900 00:00:00','" & IdProveedor & "','" & IdImportacion & "')")

            GuardarDetalleMovimiento(codigoLovePosMovimiento, IdArticulo, Cant)

            clase.ConsultarSQLServer("SELECT* FROM [dbo].[ConsecutivosDocInventarioTienda] WHERE IdEmpresa = 100 AND IdTienda = 99 AND IdDocumentoInventario = 'COM'", "consecutivo")
            If clase.dtSql.Tables("consecutivo").Rows.Count > 0 Then
                clase.AgregarSQLServer("UPDATE [dbo].[ConsecutivosDocInventarioTienda] SET [Consecutivo] = [Consecutivo] + 1, IdCodMovimiento = '" & codigoLovePosMovimiento & "', FechaModif = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE IdEmpresa = 100 and IdTienda = 99  and IdDocumentoInventario = 'COM'") 'consulta de actualizacion
            Else
                Dim cadenainsercion As String = "INSERT INTO [dbo].[ConsecutivosDocInventarioTienda] ([IdEmpresa],[IdTienda],[IdDocumentoInventario],[IdCodMovimiento],[IdTiendaDestino],[Consecutivo],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) " &
                        "VALUES ('100','99','COM','COM-0990001','0','2','automatico','automatico','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"

                clase.AgregarSQLServer(cadenainsercion)
            End If
            clase.agregar_registro("INSERT INTO `detalleimportacioncompras`(`IdImportacion`,`IdCompra`,`IdProveedor`) VALUES ( '" & IdImportacion & "','" & codigoLovePosMovimiento & "','" & IdProveedor & "')")
        Else
            codigoLovePosMovimiento = clase.dt1.Tables("importacion").Rows(0)("IdCompra")
            GuardarDetalleMovimiento(codigoLovePosMovimiento, IdArticulo, Cant)
        End If
        ConsolidarSaldos(100, 99, 1, "COM", codigoLovePosMovimiento)
    End Sub

    Public Sub GuardarDetalleMovimiento(IdCodigoCompra As String, IdArticulo As Integer, Cant As Integer)
        Dim detalleMovimiento As String = "INSERT INTO [dbo].[DetMovimientoInventario]([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[IdArticulo],[Cantidad],[PrecioCosto],[PrecioVenta],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif],[PrecioCostoFiscal]) VALUES " &
        "('100','99','1','COM','" & IdCodigoCompra & "','" & IdArticulo & "','" & Cant & "','" & Str(ObtenerCostoReal(IdArticulo)) & "','" & Str(ObtenerPrecioVenta(IdArticulo)) & "','automatico','automatico','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Str(ObtenerCostoFiscal(IdArticulo)) & "')"
        clase.AgregarSQLServer(detalleMovimiento)
    End Sub

    Public Sub CrearMovimientoCompraLovePOSNacional(IdImportacion As Integer, IdProveedor As Integer, Operario As String)
        MigrarArticulosLineasProveedores()
        clase.consultar1("SELECT* FROM detalleimportacioncompras WHERE IdImportacion = " & IdImportacion & " AND IdProveedor = " & IdProveedor & "", "importacion")
        If clase.dt1.Tables("importacion").Rows.Count = 0 Then
            '  Try
            'crear encabezado
            Dim codigoLovePosMovimiento As String = CalcularCodigoLovePOSCompra()

            clase.AgregarSQLServer("INSERT INTO [dbo].[CabMovimientoInventario] ([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[FechaOperativa],[Operario],[Revisor]" &
                                            ",[Finalizado],[Revisada],[Recibido],[UsuarioModif],[ProgramaModif],[Observaciones],[EquipoModif],[FechaModif],[IdEmpresaDestino],[IdTiendaDestino],[IdBodegaDestino],[FechaRecibo],[IdProveedor],[IdEntrada])" &
        "VALUES ('100','99','1','COM','" & codigoLovePosMovimiento & "','" & Now.ToString("dd/MM/yyyy") & "'" &
        ",'" & Operario & "','" & Operario & "','TRUE','TRUE','TRUE','Automatico'" &
        ",'Automatico','','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','0','0','0','01/01/1900 00:00:00','" & IdProveedor & "','" & IdImportacion & "')")

            'crear detalle
            clase.consultar("SELECT cb.cabmer_proveedor, dt.detmer_articulo, dt.detmer_cantidad, ar.ar_precio1, dt.detmer_costounitario FROM detmer_noimportada dt INNER JOIN articulos ar ON (dt.detmer_articulo = ar.ar_codigo) INNER JOIN cabmer_noimportada cb ON (dt.detmer_codigo_imp = cb.cabmer_codigo) WHERE (cb.cabmer_importacion =" & IdImportacion & " AND cb.cabmer_proveedor =" & IdProveedor & ")", "detalle-entrada")
            Dim a As Integer
            Dim cadenadetalle As String = "INSERT INTO [dbo].[DetMovimientoInventario]([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[IdArticulo],[Cantidad],[PrecioCosto],[PrecioVenta],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif],[PrecioCostoFiscal]) VALUES "
            Dim ind As Boolean = False
            For a = 0 To clase.dt.Tables("detalle-entrada").Rows.Count - 1
                With clase.dt.Tables("detalle-entrada")
                    If ind = True Then
                        cadenadetalle = cadenadetalle & ",('100','99','1','COM','" & codigoLovePosMovimiento & "','" & .Rows(a)("detmer_articulo") & "','" & .Rows(a)("detmer_cantidad") & "','" & Str(.Rows(a)("detmer_costounitario")) & "','" & Str(.Rows(a)("ar_precio1")) & "','automatico','automatico','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Str(.Rows(a)("detmer_costounitario")) & "')"
                    Else
                        cadenadetalle = cadenadetalle & "('100','99','1','COM','" & codigoLovePosMovimiento & "','" & .Rows(a)("detmer_articulo") & "','" & .Rows(a)("detmer_cantidad") & "','" & Str(.Rows(a)("detmer_costounitario")) & "','" & Str(.Rows(a)("ar_precio1")) & "','automatico','automatico','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Str(.Rows(a)("detmer_costounitario")) & "')"
                        ind = True
                    End If
                    clase.agregar_registro("UPDATE articulos SET ar_migrado = FALSE WHERE ar_codigo = " & .Rows(a)("detmer_articulo") & "")
                End With
            Next
            clase.AgregarSQLServer(cadenadetalle)
            clase.ConsultarSQLServer("SELECT* FROM [dbo].[ConsecutivosDocInventarioTienda] WHERE IdEmpresa = 100 AND IdTienda = 99 AND IdDocumentoInventario = 'COM'", "consecutivo")
            If clase.dtSql.Tables("consecutivo").Rows.Count > 0 Then
                clase.AgregarSQLServer("UPDATE [dbo].[ConsecutivosDocInventarioTienda] SET [Consecutivo] = [Consecutivo] + 1, IdCodMovimiento = '" & codigoLovePosMovimiento & "', FechaModif = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE IdEmpresa = 100 and IdTienda = 99  and IdDocumentoInventario = 'COM'") 'consulta de actualizacion
            Else
                Dim cadenainsercion As String = "INSERT INTO [dbo].[ConsecutivosDocInventarioTienda] ([IdEmpresa],[IdTienda],[IdDocumentoInventario],[IdCodMovimiento],[IdTiendaDestino],[Consecutivo],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) " &
                        "VALUES ('100','99','COM','COM-0990001','0','2','automatico','automatico','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"

                clase.AgregarSQLServer(cadenainsercion)
            End If
            ConsolidarSaldos(100, 99, 1, "COM", codigoLovePosMovimiento)
            clase.agregar_registro("INSERT INTO `detalleimportacioncompras`(`IdImportacion`,`IdCompra`,`IdProveedor`) VALUES ( '" & IdImportacion & "','" & codigoLovePosMovimiento & "','" & IdProveedor & "')")

            'guardar proveedores


            '  Catch ex As Exception
            '  MessageBox.Show(ex.Message.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '   End Try
        Else
            Dim CodigoCOmpraLovepOS As String = clase.dt1.Tables("importacion").Rows(0)("IdCompra")
            clase.consultar("SELECT cb.cabmer_proveedor, dt.detmer_articulo, dt.detmer_cantidad, ar.ar_precio1, dt.detmer_costounitario FROM detmer_noimportada dt INNER JOIN articulos ar ON (dt.detmer_articulo = ar.ar_codigo) INNER JOIN cabmer_noimportada cb ON (dt.detmer_codigo_imp = cb.cabmer_codigo) WHERE (cb.cabmer_importacion =" & IdImportacion & " AND cb.cabmer_proveedor =" & IdProveedor & ")", "imp")
            If clase.dt.Tables("imp").Rows.Count > 0 Then
                clase.ConsultarSQLServer("Select* from [dbo].[DetMovimientoInventario] WHERE IdEmpresa = 100 AND IdTienda = 99 AND IdBodega = 1 AND IdTipoMovimiento = 'COM' AND IdCodMovimiento = '" & CodigoCOmpraLovepOS & "'", "sql")
                Dim z As Integer
                Dim x As Integer
                Dim indbus As Boolean = False
                For x = 0 To clase.dt.Tables("imp").Rows.Count - 1
                    indbus = False
                    For z = 0 To clase.dtSql.Tables("sql").Rows.Count - 1
                        If clase.dt.Tables("imp").Rows(x)("detmer_articulo") = clase.dtSql.Tables("sql").Rows(z)("IdArticulo") Then
                            If clase.dt.Tables("imp").Rows(x)("detmer_cantidad") <> clase.dtSql.Tables("sql").Rows(z)("Cantidad") Then
                                clase.AgregarSQLServer("UPDATE [dbo].[DetMovimientoInventario] SET Cantidad = '" & clase.dt.Tables("imp").Rows(x)("detmer_cantidad") & "', FechaModif = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE IdEmpresa = 100 AND idTienda = 99 AND IdBodega = 1 AND IdTipoMovimiento = 'COM' AND IdCodMovimiento = '" & CodigoCOmpraLovepOS & "' AND IdArticulo = '" & clase.dt.Tables("imp").Rows(x)("detmer_articulo") & "'")
                                Dim cantidadactualizada As Integer = clase.dt.Tables("imp").Rows(x)("detmer_cantidad") - clase.dtSql.Tables("sql").Rows(z)("Cantidad")
                                Dim costo As Double = clase.dtSql.Tables("sql").Rows(z)("PrecioCosto")
                                'actualizar saldos en Existencias Articulos de sql server
                                ConsolidarArticulo(100, 99, 1, clase.dt.Tables("imp").Rows(x)("detmer_articulo"), cantidadactualizada, costo * cantidadactualizada, costo * cantidadactualizada)
                            End If
                            indbus = True
                            Exit For
                        End If
                    Next
                    If indbus = False Then
                        Dim cad As String = "INSERT INTO [dbo].[DetMovimientoInventario]([IdEmpresa],[IdTienda],[IdBodega],[IdTipoMovimiento],[IdCodMovimiento],[IdArticulo],[Cantidad],[PrecioCosto],[PrecioVenta],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif],[PrecioCostoFiscal]) VALUES " &
                            "('100','99','1','COM','" & CodigoCOmpraLovepOS & "','" & clase.dt.Tables("imp").Rows(x)("detmer_articulo") & "','" & clase.dt.Tables("imp").Rows(x)("detmer_cantidad") & "','" & Str(clase.dt.Tables("imp").Rows(x)("detmer_costounitario")) & "','" & Str(clase.dt.Tables("imp").Rows(x)("ar_precio1")) & "','automatico','automatico','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Str(clase.dt.Tables("imp").Rows(x)("detmer_costounitario")) & "')"
                        clase.AgregarSQLServer(cad)
                        ConsolidarArticulo(100, 99, 1, clase.dt.Tables("imp").Rows(x)("detmer_articulo"), clase.dt.Tables("imp").Rows(x)("detmer_cantidad"), clase.dt.Tables("imp").Rows(x)("detmer_costounitario") * clase.dt.Tables("imp").Rows(x)("detmer_cantidad"), clase.dt.Tables("imp").Rows(x)("detmer_costounitario") * clase.dt.Tables("imp").Rows(x)("detmer_cantidad"))
                        clase.agregar_registro("UPDATE articulos SET ar_migrado = FALSE WHERE ar_codigo = " & clase.dt.Tables("imp").Rows(x)("detmer_articulo") & "")
                    End If
                Next

            End If
        End If
    End Sub

    Public Sub ConsolidarSaldos(IdEmpresa As Integer, IdTienda As Integer, IdBodega As Integer, IdTipoMovimiento As String, IdMovimiento As String)
        clase.ConsultarSQLServer("SELECT [det].[IdArticulo],([det].[Cantidad]*[tip].[OperacionDocumento]) AS Cantidad,[det].[PrecioCosto],[PrecioCostoFiscal] FROM [dbo].[DetMovimientoInventario] det INNER JOIN [dbo].[TipoDocumentosInventarios] tip ON (det.IdTipoMovimiento = tip.CodDocumento)  WHERE det.IdEmpresa = " & IdEmpresa & " AND det.IdTienda = " & IdTienda & " AND det.IdBodega = " & IdBodega & " AND det.IdTipoMovimiento = '" & IdTipoMovimiento & "' AND det.IdCodMovimiento = '" & IdMovimiento & "'", "saldos")
        If clase.dtSql.Tables("saldos").Rows.Count > 0 Then
            Dim Cant As Integer = 0
            Dim CostoReal As Double = 0
            Dim CostoFiscal As Double = 0
            Dim x As Integer
            For x = 0 To clase.dtSql.Tables("saldos").Rows.Count - 1
                With clase.dtSql.Tables("saldos")
                    Cant = clase.dtSql.Tables("saldos").Rows(x)("Cantidad")
                    CostoReal = clase.dtSql.Tables("saldos").Rows(x)("PrecioCosto")
                    CostoFiscal = clase.dtSql.Tables("saldos").Rows(x)("PrecioCostoFiscal")
                    ConsolidarArticulo(IdEmpresa, IdTienda, IdBodega, .Rows(x)("IdArticulo"), Cant, CostoReal * Cant, CostoFiscal * Cant)
                End With
            Next
        End If
    End Sub



    Public Sub ConsolidarArticulo(IdEmpresa As Integer, IdTienda As Integer, IdBodega As Integer, IdArticulo As String, Cant As Integer, CostoReal As Double, CostoFiscal As Double)
        clase.ConsultarSQLServer1("SELECT [CostoInicial],[MovimientoPeriodo],[StockFinal],[CostoPeriodo],[CostoFinal],[CostoPeriodoFiscal],[CostoInicialFiscal] FROM [dbo].[ExistenciasArticulos] WHERE IdEmpresa = " & IdEmpresa & " AND IdTienda = " & IdTienda & " AND IdBodega = " & IdBodega & " AND IdArticulo = " & IdArticulo & "", "sald")
        If clase.dtSql1.Tables("sald").Rows.Count > 0 Then
            Dim MovPeriodo As Integer = clase.dtSql1.Tables("sald").Rows(0)("MovimientoPeriodo") + Cant
            Dim StockFinal As Integer = clase.dtSql1.Tables("sald").Rows(0)("StockFinal") + Cant
            Dim CostoPeriodoReal As Double = clase.dtSql1.Tables("sald").Rows(0)("CostoPeriodo") + CostoReal
            Dim CostoPeriodoFiscal As Double = clase.dtSql1.Tables("sald").Rows(0)("CostoPeriodoFiscal") + CostoFiscal
            'CostoFinalReal
            Dim CostoFinalReal As Double
            If (clase.dtSql1.Tables("sald").Rows(0)("CostoInicial") = 0) Then
                CostoFinalReal = CostoPeriodoReal
            Else
                CostoFinalReal = CostoPeriodoReal + clase.dtSql1.Tables("sald").Rows(0)("CostoInicial")
            End If
            'CostoFinalFiscal
            Dim CostoFinalFiscal As Double
            If (clase.dtSql1.Tables("sald").Rows(0)("CostoInicialFiscal") = 0) Then
                CostoFinalFiscal = CostoPeriodoFiscal
            Else
                CostoFinalFiscal = CostoPeriodoFiscal + clase.dtSql1.Tables("sald").Rows(0)("CostoInicialFiscal")
            End If

            clase.ActualizarSQLServer("UPDATE [dbo].[ExistenciasArticulos] SET [UsuarioModif] = 'automatico', [ProgramaModif] = 'planet bodega',[EquipoModif] = '" & Environment.MachineName & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "',[MovimientoPeriodo] = " & MovPeriodo & ",[StockFinal] = " & StockFinal & ", [CostoPeriodo] = " & Str(CostoPeriodoReal) & ",[CostoFinal] = " & Str(CostoFinalReal) & ",[CostoPeriodoFiscal] = " & Str(CostoPeriodoFiscal) & ",[CostoFinalFiscal] = " & Str(CostoFinalFiscal) & " WHERE IdEmpresa = " & IdEmpresa & " AND IdTienda = " & IdTienda & " AND IdBodega = " & IdBodega & " AND IdArticulo = " & IdArticulo & "")
        Else
            Dim cadenainsercion As String = "INSERT INTO [dbo].[ExistenciasArticulos] ([IdEmpresa],[IdTienda],[IdBodega],[IdArticulo],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif],[StockInicial],[MovimientoPeriodo],[StockFinal],[CostoInicial],[CostoPeriodo],[CostoFinal],[CostoInicialFiscal],[CostoPeriodoFiscal],[CostoFinalFiscal]) VALUES " &
                "('100','99','1','" & IdArticulo & "','automatico','planet bodega','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','0','" & Cant & "','" & Cant & "','0','" & Str(CostoReal) & "','" & Str(CostoReal) & "','0','" & Str(CostoFiscal) & "','" & Str(CostoFiscal) & "')"
            clase.AgregarSQLServer(cadenainsercion)
        End If
    End Sub

    Public Sub MigrarArticulosLineasProveedores()
        Dim cadena As String
        Dim i As Integer
        ''cabimportacion
        'clase.consultar("SELECT * FROM cabimportacion WHERE imp_migrado=FALSE", "imp")
        'If clase.dt.Tables("imp").Rows.Count > 0 Then
        '    Dim cadenalineas As String = ""
        '    For i = 0 To clase.dt.Tables("imp").Rows.Count - 1
        '        Application.DoEvents()
        '        With clase.dt.Tables("imp")
        '            clase.ConsultarSQLServer("SELECT * FROM [dbo].[RelacionEntradaMercancia] WHERE IdEntrada = '" & .Rows(i)("imp_codigo") & "'", "imp")
        '            If clase.dtSql.Tables("imp").Rows.Count > 0 Then
        '                cadenalineas = "UPDATE [dbo].[RelacionEntradaMercancia] SET [Fecha] = '" & .Rows(i)("imp_fecha") & "',[DescripcionCorta] = '" & .Rows(i)("imp_nombrefecha") & "',[DescripcionLarga] = '" & .Rows(i)("imp_descripcion") & "',[Importado] = '0',[Procedencia] = '" & .Rows(i)("imp_lugar") & "',[UsuarioModif] = 'automatico',[ProgramaModif] = 'planetbodega',[EquipoModif] = '" & Environment.MachineName & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE IdEntrada = '" & .Rows(i)("imp_codigo") & "'"
        '            Else
        '                cadenalineas = "INSERT INTO [dbo].[ColorArticulo]([CodigoColor],[NombreColor],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("cod_color") & "','" & .Rows(i)("colornombre") & "','PlanetBodega','PlanetBodega','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
        '                cadenalineas = "INSERT INTO [dbo].[RelacionEntradaMercancia]([IdEntrada],[Fecha],[DescripcionCorta],[DescripcionLarga],[Importado],[Procedencia],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("imp_codigo") & "','" & .Rows(i)("imp_fecha") & "',<DescripcionCorta, varchar(50),>,<DescripcionLarga, varchar(200),>,<Importado, bit,>,<Procedencia, varchar(50),>,<UsuarioModif, varchar(20),>,<ProgramaModif, varchar(50),>,<EquipoModif, varchar(20),>         ,<FechaModif, smalldatetime,>)"
        '            End If
        '            clase.AgregarSQLServer(cadenalineas)

        '        End With
        '    Next
        '    clase.agregar_registro("UPDATE colores SET cod_migrado = '1' WHERE cod_migrado=FALSE")
        'End If
        'colores
        clase.consultar("SELECT * FROM colores WHERE cod_migrado=FALSE", "colores")
        If clase.dt.Tables("colores").Rows.Count > 0 Then
            Dim cadenalineas As String = ""
            For i = 0 To clase.dt.Tables("colores").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("colores")
                    clase.ConsultarSQLServer("SELECT * FROM [dbo].[ColorArticulo] WHERE CodigoColor = '" & .Rows(i)("cod_color") & "'", "colores")
                    If clase.dtSql.Tables("colores").Rows.Count > 0 Then
                        cadenalineas = "UPDATE [dbo].[ColorArticulo] SET [NombreColor] = '" & .Rows(i)("colornombre") & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "'  WHERE CodigoColor = '" & .Rows(i)("cod_color") & "'"
                    Else
                        cadenalineas = "INSERT INTO [dbo].[ColorArticulo]([CodigoColor],[NombreColor],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("cod_color") & "','" & .Rows(i)("colornombre") & "','PlanetBodega','PlanetBodega','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                    End If
                    clase.AgregarSQLServer(cadenalineas)

                End With
            Next
            clase.agregar_registro("UPDATE colores SET cod_migrado = '1' WHERE cod_migrado=FALSE")
        End If
        'tallas
        clase.consultar("SELECT * FROM tallas WHERE migrado=FALSE", "tallas")
        If clase.dt.Tables("tallas").Rows.Count > 0 Then
            Dim cadenalineas As String = ""
            For i = 0 To clase.dt.Tables("tallas").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("tallas")
                    clase.ConsultarSQLServer("SELECT * FROM [dbo].[Talla] WHERE CodigoTalla = '" & .Rows(i)("codigo_talla") & "'", "tallas")
                    If clase.dtSql.Tables("tallas").Rows.Count > 0 Then
                        cadenalineas = "UPDATE [dbo].[Talla] SET [ValorTalla] = '" & .Rows(i)("nombretalla") & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "'  WHERE CodigoTalla = '" & .Rows(i)("codigo_talla") & "'"
                    Else
                        cadenalineas = "INSERT INTO [dbo].[Talla]([CodigoTalla],[ValorTalla],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("codigo_talla") & "','" & .Rows(i)("nombretalla") & "','PlanetBodega','PlanetBodega','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                    End If
                    clase.AgregarSQLServer(cadenalineas)

                End With
            Next
            clase.agregar_registro("UPDATE tallas SET migrado = '1' WHERE migrado=FALSE")
        End If
        'linea1
        clase.consultar("SELECT * FROM linea1 WHERE migrado=FALSE", "linea1")
        If clase.dt.Tables("linea1").Rows.Count > 0 Then
            Dim cadenalineas As String = ""
            For i = 0 To clase.dt.Tables("linea1").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("linea1")
                    clase.ConsultarSQLServer("SELECT * FROM [dbo].[Linea1] WHERE CodigoLinea1 = '" & .Rows(i)("ln1_codigo") & "'", "linea1")
                    If clase.dtSql.Tables("linea1").Rows.Count > 0 Then
                        cadenalineas = "UPDATE [dbo].[Linea1] SET [NombreLinea] = '" & .Rows(i)("ln1_nombre") & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "'  WHERE CodigoLinea1 = '" & .Rows(i)("ln1_codigo") & "'"
                    Else
                        cadenalineas = "INSERT INTO [dbo].[Linea1]([CodigoLinea1],[NombreLinea],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("ln1_codigo") & "','" & .Rows(i)("ln1_nombre") & "','ConsultasVentasPlanet','ConsultasVentasPlanet','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                    End If
                    clase.AgregarSQLServer(cadenalineas)

                End With
            Next
            clase.agregar_registro("UPDATE linea1 SET migrado = '1' WHERE migrado=FALSE")
        End If
        'sublinea1
        clase.consultar("SELECT * FROM sublinea1 WHERE migrado=FALSE", "sublinea1")
        If clase.dt.Tables("sublinea1").Rows.Count > 0 Then
            Dim cadenasublinea2 As String = ""
            For i = 0 To clase.dt.Tables("sublinea1").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("sublinea1")
                    clase.ConsultarSQLServer("SELECT * FROM [dbo].[Linea2] WHERE CodigoLinea2 = '" & .Rows(i)("sl1_codigo") & "'", "linea2")
                    If clase.dtSql.Tables("linea2").Rows.Count > 0 Then
                        cadenasublinea2 = "UPDATE [dbo].[Linea2] SET [CodigoLinea1] = '" & .Rows(i)("sl1_ln1codigo") & "',[NombreLinea] = '" & .Rows(i)("sl1_nombre") & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE CodigoLinea2 = '" & .Rows(i)("sl1_codigo") & "'"
                    Else
                        cadenasublinea2 = "INSERT INTO [dbo].[Linea2]([CodigoLinea2],[CodigoLinea1],[NombreLinea],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("sl1_codigo") & "','" & .Rows(i)("sl1_ln1codigo") & "','" & .Rows(i)("sl1_nombre") & "','ConsultasVentasPlanet','ConsultasVentasPlanet','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                    End If
                    clase.AgregarSQLServer(cadenasublinea2)

                End With
            Next
            clase.agregar_registro("UPDATE sublinea1 SET migrado = '1' WHERE migrado=FALSE")
        End If
        'sublinea2
        clase.consultar("SELECT * FROM sublinea2 WHERE migrado=FALSE", "sublinea2")
        If clase.dt.Tables("sublinea2").Rows.Count > 0 Then
            Dim cadenasublinea3 As String = ""
            For i = 0 To clase.dt.Tables("sublinea2").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("sublinea2")
                    clase.ConsultarSQLServer("SELECT * FROM [dbo].[Linea3] WHERE CodigoLinea3 = '" & .Rows(i)("sl2_codigo") & "'", "linea3")
                    If clase.dtSql.Tables("linea3").Rows.Count > 0 Then
                        cadenasublinea3 = "UPDATE [dbo].[Linea3] SET [CodigoLinea2] = '" & .Rows(i)("sl2_sl1codigo") & "',[NombreLinea] = '" & .Rows(i)("sl2_nombre") & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE CodigoLinea2 = '" & .Rows(i)("sl2_codigo") & "'"
                    Else
                        cadenasublinea3 = "INSERT INTO [dbo].[Linea3]([CodigoLinea3],[CodigoLinea2],[NombreLinea],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("sl2_codigo") & "','" & .Rows(i)("sl2_sl1codigo") & "','" & .Rows(i)("sl2_nombre") & "','ConsultasVentasPlanet','ConsultasVentasPlanet','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                    End If
                    clase.AgregarSQLServer(cadenasublinea3)
                End With
            Next
            clase.agregar_registro("UPDATE sublinea2 SET migrado = '1' WHERE migrado=FALSE")
        End If

        clase.consultar("SELECT* FROM proveedores WHERE prv_migrado = FALSE", "prv")
        If clase.dt.Tables("prv").Rows.Count > 0 Then
            For i = 0 To clase.dt.Tables("prv").Rows.Count - 1
                clase.ConsultarSQLServer("SELECT* FROM [dbo].[Proveedor] WHERE IdProveedor = " & clase.dt.Tables("prv").Rows(i)("prv_codigo") & "", "prvsql")
                Dim cad As String
                If clase.dtSql.Tables("prvsql").Rows.Count > 0 Then
                    cad = "UPDATE [dbo].[Proveedor] SET [RazonSocial] = '" & clase.dt.Tables("prv").Rows(i)("prv_nombre") & "',[Ciudad] = '" & clase.dt.Tables("prv").Rows(i)("prv_ciudad") & "',[Pais] = 'N/A',[Direccion] = '" & clase.dt.Tables("prv").Rows(i)("prv_direccion") & "'" &
                        ",[Contacto] = '" & clase.dt.Tables("prv").Rows(i)("prv_contacto") & "',[Telefono1] = '" & clase.dt.Tables("prv").Rows(i)("prv_telefono1") & "',[Telefono2] = '" & clase.dt.Tables("prv").Rows(i)("prv_telefono2") & "',[Email1] = '" & clase.dt.Tables("prv").Rows(i)("prv_email1") & "',[Email2] = '" & clase.dt.Tables("prv").Rows(i)("prv_email2") & "',[PaginaWeb] = '" & clase.dt.Tables("prv").Rows(i)("prv_web") & "'" &
                         " , [CodigoAsignado] = '" & clase.dt.Tables("prv").Rows(i)("prv_codigoasignado") & "', [UsuarioModif] = 'automatico', [ProgramaModif] = 'Automatico', [EquipoModif]='" & Environment.MachineName & "', [FechaModif]='" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "'  WHERE IdProveedor = " & clase.dt.Tables("prv").Rows(i)("prv_codigo") & ""
                Else
                    With clase.dt.Tables("prv")
                        cad = "INSERT INTO [dbo].[Proveedor]([IdProveedor],[CodigoAsignado],[RazonSocial],[Ciudad],[Pais],[Direccion],[Contacto],[Telefono1],[Telefono2],[Email1],[Email2],[PaginaWeb],[UsuarioModif],[ProgramaModif]" &
                          ",[FechaModif],[EquipoModif]) VALUES ('" & .Rows(i)("prv_codigo") & "','" & .Rows(i)("prv_codigoasignado") & "','" & .Rows(i)("prv_nombre") & "','" & .Rows(i)("prv_ciudad") & "','N/A','" & .Rows(i)("prv_direccion") & "' " &
                          ",'" & .Rows(i)("prv_contacto") & "','" & .Rows(i)("prv_telefono1") & "','" & .Rows(i)("prv_telefono2") & "','" & .Rows(i)("prv_email1") & "','" & .Rows(i)("prv_email2") & "','" & .Rows(i)("prv_web") & "','automatico','planet bodega' " &
                          ",'" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "','" & Environment.MachineName & "')"
                    End With
                End If
                clase.AgregarSQLServer(cad)
                clase.agregar_registro("UPDATE proveedores SET prv_migrado = TRUE WHERE prv_codigo = " & clase.dt.Tables("prv").Rows(i)("prv_codigo") & "")
            Next
        End If

        'Clasificaciones para analisis
        clase.consultar("SELECT* FROM CabClasificacion WHERE migrado = '0'", "Clasificacion")
        If clase.dt.Tables("Clasificacion").Rows.Count > 0 Then
            Dim cadenasublinea3 As String = ""
            For i = 0 To clase.dt.Tables("Clasificacion").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("Clasificacion")
                    clase.ConsultarSQLServer("select* from Clasificacion where IdClasificacion =  '" & .Rows(i)("IdClasificacion") & "'", "Clasificacion")
                    If clase.dtSql.Tables("Clasificacion").Rows.Count > 0 Then
                        cadenasublinea3 = "UPDATE [dbo].[Clasificacion]  SET [NombreClasificacion] = '" & .Rows(i)("Clasificacion") & "',[UsuarioModif] = 'automatico',[ProgramaModif] = 'automatico',[EquipoModif] = '" & Environment.MachineName & "',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE IdClasificacion = " & .Rows(i)("IdClasificacion") & ""
                    Else
                        cadenasublinea3 = "INSERT INTO [dbo].[Clasificacion] ([IdClasificacion],[NombreClasificacion],[UsuarioModif],[ProgramaModif],[EquipoModif],[FechaModif]) VALUES ('" & .Rows(i)("IdClasificacion") & "','" & .Rows(i)("Clasificacion") & "','automatico','automatico','" & Environment.MachineName & "','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                    End If
                    clase.AgregarSQLServer(cadenasublinea3)
                End With
            Next
            clase.agregar_registro("UPDATE CabClasificacion SET migrado = '1' WHERE migrado=FALSE")
        End If
        'articulos
        clase.consultar("SELECT* FROM articulos WHERE ar_migrado = FALSE", "articulos")
        If clase.dt.Tables("articulos").Rows.Count > 0 Then
            'Dim cadena As String =   "INSERT INTO dbo.Articulo (IdArticulo, Referencia, CodigoBarra1, CodigoBarra2, DescripcionCorta, DescripcioNext, Linea1, Linea2, Linea3, CodigoTalla, Codigocolor, Precioimporte1, Precioimporte2, Preciominimo, Preciomaximo, PermitirDescuento, Fechaentrada, Fechacaducidad, RutaFoto, FechaTemporada, IdImpuesto, UsuarioModif, ProgramaModif, EquipoModif, FechaModif) VALUES " ' ('" & .Rows(0)("ar_codigo") & "', '" & .Rows(0)("ar_referencia") & "', '" & .Rows(0)("ar_codigobarras") & "', '" & .Rows(0)("ar_codigobarras2") & "', '" & .Rows(0)("ar_descripcion") & "', '" & .Rows(0)("ar_descripcion") & "', '" & .Rows(0)("ar_linea") & "', '" & .Rows(0)("ar_sublinea1") & "', '" & .Rows(0)("ar_sublinea2") & "','" & .Rows(0)("ar_talla") & "','" & .Rows(0)("ar_color") & "','" & .Rows(0)("ar_costo") & "','" & .Rows(0)("ar_costo2") & "','" & .Rows(0)("ar_precio1") & "','" & .Rows(0)("ar_precio2") & "','TRUE', '" & .Rows(0)("ar_fechaingreso") & "','" & .Rows(0)("ar_fechaingreso") & "','" & .Rows(0)("ar_foto") & "','" & .Rows(0)("ar_fechaingreso") & "','1','Admin', 'inicial','MaquinaDefecto','" & Now.ToString("dd/MM/yyyy") & "')"
            Dim ind As Boolean = True
            For i = 0 To clase.dt.Tables("articulos").Rows.Count - 1
                Application.DoEvents()
                With clase.dt.Tables("articulos")
                    clase.ConsultarSQLServer("SELECT* FROM dbo.Articulo WHERE IdArticulo = " & .Rows(i)("ar_codigo") & "", "articulos")
                    If clase.dtSql.Tables("articulos").Rows.Count > 0 Then
                        clase.ActualizarSQLServer("UPDATE [dbo].[Articulo] SET [Referencia] = '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_referencia"), 50) & "', [CodigoBarra1] ='" & .Rows(i)("ar_codigobarras") & "',[CodigoBarra2] = '" & .Rows(i)("ar_codigobarras2") & "',[DescripcionCorta] = '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 20) & "',[DescripcioNext] = '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 50) & "',[Linea1] = '" & .Rows(i)("ar_linea") & "',[Linea2] = '" & .Rows(i)("ar_sublinea1") & "',[Linea3] = '" & .Rows(i)("ar_sublinea2") & "',[CodigoTalla] = '" & .Rows(i)("ar_talla") & "',[Codigocolor] = '" & Str(.Rows(i)("ar_color")) & "',[Precioimporte1] = '" & Str(.Rows(i)("ar_costo")) & "',[Precioimporte2] = '" & Str(.Rows(i)("ar_costo2")) & "',[Preciominimo] = '" & Str(.Rows(i)("ar_precio1")) & "',[Preciomaximo] = '" & Str(.Rows(i)("ar_precio2")) & "', [Fechaentrada] = '" & .Rows(i)("ar_fechaingreso") & "',[Fechacaducidad] = '" & .Rows(i)("ar_fechaingreso") & "',[RutaFoto] = '" & .Rows(i)("ar_foto") & "',[FechaTemporada] = '" & .Rows(i)("ar_fechaingreso") & "',[IdImpuesto] = '2',[UsuarioModif] = 'Admin',[ProgramaModif] = 'Inicial',[EquipoModif] = 'MaquinaDefecto',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "', [NombreArchivoFotografico] = '" & DeterminarNombreArchivo(.Rows(i)("ar_foto")) & "', IdProveedor = '" & ObtenerProveedor(.Rows(i)("ar_codigo")) & "' WHERE IdArticulo = " & .Rows(i)("ar_codigo") & "")
                        ' clase.ActualizarSQLServer("UPDATE [dbo].[Articulo] SET [Referencia] = '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_referencia"), 50) & "', [CodigoBarra1] ='" & .Rows(i)("ar_codigobarras") & "',[CodigoBarra2] = '" & .Rows(i)("ar_codigobarras2") & "',[DescripcionCorta] = '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 20) & "',[DescripcioNext] = '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 50) & "',[Linea1] = '" & .Rows(i)("ar_linea") & "',[Linea2] = '" & .Rows(i)("ar_sublinea1") & "',[Linea3] = '" & .Rows(i)("ar_sublinea2") & "',[CodigoTalla] = '" & .Rows(i)("ar_talla") & "',[Codigocolor] = '" & Str(.Rows(i)("ar_color")) & "',[Precioimporte1] = '" & Str(.Rows(i)("ar_costo")) & "',[Precioimporte2] = '" & Str(.Rows(i)("ar_costo2")) & "',[Preciominimo] = '" & Str(.Rows(i)("ar_precio1")) & "',[Preciomaximo] = '" & Str(.Rows(i)("ar_precio2")) & "', [Fechaentrada] = '" & .Rows(i)("ar_fechaingreso") & "',[Fechacaducidad] = '" & .Rows(i)("ar_fechaingreso") & "',[RutaFoto] = '" & .Rows(i)("ar_foto") & "',[FechaTemporada] = '" & .Rows(i)("ar_fechaingreso") & "',[IdImpuesto] = '1',[UsuarioModif] = 'Admin',[ProgramaModif] = 'Inicial',[EquipoModif] = 'MaquinaDefecto',[FechaModif] = '" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "' WHERE IdArticulo = " & .Rows(i)("ar_codigo") & "")
                    Else
                        cadena = "INSERT INTO dbo.Articulo (IdArticulo, Referencia, CodigoBarra1, CodigoBarra2, DescripcionCorta, DescripcioNext, Linea1, Linea2, Linea3, CodigoTalla, Codigocolor, Precioimporte1, Precioimporte2, Preciominimo, Preciomaximo, PermitirDescuento, Fechaentrada, Fechacaducidad, RutaFoto, FechaTemporada, IdImpuesto, UsuarioModif, ProgramaModif, EquipoModif, FechaModif, TieneFotoLocal, NombreArchivoFotografico, IdProveedor) VALUES ('" & .Rows(i)("ar_codigo") & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_referencia"), 50) & "', '" & .Rows(i)("ar_codigobarras") & "', '" & .Rows(i)("ar_codigobarras2") & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 20) & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 50) & "', '" & .Rows(i)("ar_linea") & "', '" & .Rows(i)("ar_sublinea1") & "', '" & .Rows(i)("ar_sublinea2") & "','" & .Rows(i)("ar_talla") & "','" & Str(.Rows(i)("ar_color")) & "','" & Str(.Rows(i)("ar_costo")) & "','" & Str(.Rows(i)("ar_costo2")) & "','" & Str(.Rows(i)("ar_precio1")) & "','" & Str(.Rows(i)("ar_precio2")) & "','TRUE', '" & .Rows(i)("ar_fechaingreso") & "','" & .Rows(i)("ar_fechaingreso") & "','" & .Rows(i)("ar_foto") & "','" & .Rows(i)("ar_fechaingreso") & "','2','Admin', 'Inicial','MaquinaDefecto','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "', 'FALSE', '" & DeterminarNombreArchivo(.Rows(i)("ar_foto")) & "', '" & ObtenerProveedor(.Rows(i)("ar_codigo")) & "')"
                        'cadena = "INSERT INTO dbo.Articulo (IdArticulo, Referencia, CodigoBarra1, CodigoBarra2, DescripcionCorta, DescripcioNext, Linea1, Linea2, Linea3, CodigoTalla, Codigocolor, Precioimporte1, Precioimporte2, Preciominimo, Preciomaximo, PermitirDescuento, Fechaentrada, Fechacaducidad, RutaFoto, FechaTemporada, IdImpuesto, UsuarioModif, ProgramaModif, EquipoModif, FechaModif) VALUES ('" & .Rows(i)("ar_codigo") & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_referencia"), 50) & "', '" & .Rows(i)("ar_codigobarras") & "', '" & .Rows(i)("ar_codigobarras2") & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 20) & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 50) & "', '" & .Rows(i)("ar_linea") & "', '" & .Rows(i)("ar_sublinea1") & "', '" & .Rows(i)("ar_sublinea2") & "','" & .Rows(i)("ar_talla") & "','" & Str(.Rows(i)("ar_color")) & "','" & Str(.Rows(i)("ar_costo")) & "','" & Str(.Rows(i)("ar_costo2")) & "','" & Str(.Rows(i)("ar_precio1")) & "','" & Str(.Rows(i)("ar_precio2")) & "','FALSE', '" & .Rows(i)("ar_fechaingreso") & "','" & .Rows(i)("ar_fechaingreso") & "','" & .Rows(i)("ar_foto") & "','" & .Rows(i)("ar_fechaingreso") & "','1','Admin', 'Inicial','MaquinaDefecto','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                        '  cadena = "INSERT INTO dbo.Articulo (IdArticulo, Referencia, CodigoBarra1, CodigoBarra2, DescripcionCorta, DescripcioNext, Linea1, Linea2, Linea3, CodigoTalla, Codigocolor, Precioimporte1, Precioimporte2, Preciominimo, Preciomaximo, PermitirDescuento, Fechaentrada, Fechacaducidad, RutaFoto, FechaTemporada, IdImpuesto, UsuarioModif, ProgramaModif, EquipoModif, FechaModif) VALUES ('" & .Rows(i)("ar_codigo") & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_referencia"), 50) & "', '" & .Rows(i)("ar_codigobarras") & "', '" & .Rows(i)("ar_codigobarras2") & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 20) & "', '" & Microsoft.VisualBasic.Left(.Rows(i)("ar_descripcion"), 50) & "', '" & .Rows(i)("ar_linea") & "', '" & .Rows(i)("ar_sublinea1") & "', '" & .Rows(i)("ar_sublinea2") & "','" & .Rows(i)("ar_talla") & "','" & Str(.Rows(i)("ar_color")) & "','" & Str(.Rows(i)("ar_costo")) & "','" & Str(.Rows(i)("ar_costo2")) & "','" & Str(.Rows(i)("ar_precio1")) & "','" & Str(.Rows(i)("ar_precio2")) & "','TRUE', '" & .Rows(i)("ar_fechaingreso") & "','" & .Rows(i)("ar_fechaingreso") & "','" & .Rows(i)("ar_foto") & "','" & .Rows(i)("ar_fechaingreso") & "','1','Admin', 'Inicial','MaquinaDefecto','" & Now.ToString("dd/MM/yyyy HH:mm:ss") & "')"
                        clase.AgregarSQLServer(cadena)
                    End If
                End With
                ' ProgressBar1.Increment(1)
            Next
            clase.agregar_registro("UPDATE articulos SET ar_migrado = TRUE")
        End If
    End Sub


    Function ObtenerProveedor(CodArticulo As Integer)
        clase.consultar1("SELECT* FROM detalle_proveedores_articulos WHERE codigo_articulo = " & CodArticulo & "", "art")
        If clase.dt1.Tables("art").Rows.Count > 0 Then
            Return clase.dt1.Tables("art").Rows(clase.dt1.Tables("art").Rows.Count - 1)("codigo_proveedor")
        Else
            Return 3251
        End If
    End Function

    Private Function DeterminarNombreArchivo(ruta As String) As String
        Dim cade As String()
        cade = Split(ruta, "\")
        Return cade(4)
    End Function

    Public Function RecuperarCostoReal(IdArticulo As Integer) As Double
        clase.ConsultarSQLServer("SELECT [Ex].[CostoFinal],[Ex].[StockFinal] FROM [dbo].[ExistenciasArticulos] Ex  WHERE  [Ex].[IdEmpresa] = 100 AND [Ex].[IdTienda]=99 AND [Ex].[IdBodega]=1 AND [Ex].[IdArticulo] = " & IdArticulo & "", "Costo")
        If (clase.dtSql.Tables("Costo").Rows.Count > 0) Then
            If (clase.dtSql.Tables("Costo").Rows(0)("StockFinal") <> 0) Then
                Return Math.Round(clase.dtSql.Tables("Costo").Rows(0)("CostoFinal") / clase.dtSql.Tables("Costo").Rows(0)("StockFinal"), 2)
            Else
                clase.consultar("SELECT ar_costo2 FROM articulos WHERE ar_codigo = " & IdArticulo & "", "art")
                Return Math.Round(clase.dt.Tables("art").Rows(0)("ar_costo2"), 2)
            End If
        Else
            clase.consultar("SELECT ar_costo2 FROM articulos WHERE ar_codigo = " & IdArticulo & "", "art")
            Return Math.Round(clase.dt.Tables("art").Rows(0)("ar_costo2"), 2)
        End If
    End Function

    Public Function RecuperarCostoFiscal(IdArticulo As Integer) As Double
        clase.ConsultarSQLServer("SELECT [Ex].[CostoFinalFiscal],[Ex].[StockFinal] FROM [dbo].[ExistenciasArticulos] Ex  WHERE  [Ex].[IdEmpresa] = 100 AND [Ex].[IdTienda]=99 AND [Ex].[IdBodega]=1 AND [Ex].[IdArticulo] = " & IdArticulo & "", "Costo")
        If (clase.dtSql.Tables("Costo").Rows.Count > 0) Then
            If (clase.dtSql.Tables("Costo").Rows(0)("StockFinal") <> 0) Then
                Return Math.Round(clase.dtSql.Tables("Costo").Rows(0)("CostoFinalFiscal") / clase.dtSql.Tables("Costo").Rows(0)("StockFinal"), 2)
            Else
                clase.consultar("SELECT ar_costo FROM articulos WHERE ar_codigo = " & IdArticulo & "", "art")
                Return Math.Round(clase.dt.Tables("art").Rows(0)("ar_costo"), 2)
            End If
        Else
            clase.consultar("SELECT ar_costo FROM articulos WHERE ar_codigo = " & IdArticulo & "", "art")
            Return Math.Round(clase.dt.Tables("art").Rows(0)("ar_costo"), 2)
        End If
    End Function

    Function CalcularCodigoLovePOSCompra() As String
        clase.ConsultarSQLServer("SELECT [Consecutivo]  FROM [dbo].[ConsecutivosDocInventarioTienda]  where IdEmpresa = 100 and IdTienda = 99  and IdDocumentoInventario = 'COM'", "codigo")
        If clase.dtSql.Tables("codigo").Rows.Count > 0 Then
            Dim consecutivo As Integer = clase.dtSql.Tables("codigo").Rows(0)("Consecutivo")
            Select Case consecutivo.ToString().Length
                Case 1
                    Return "COM-099" & "000" & consecutivo.ToString()
                Case 2
                    Return "COM-099" & "00" & consecutivo.ToString()
                Case 3
                    Return "COM-099" & "0" & consecutivo.ToString()
                Case < 3
                    Return "COM-099" & consecutivo.ToString()
            End Select

        Else
            Return "COM-0990001"
        End If
    End Function

End Module
