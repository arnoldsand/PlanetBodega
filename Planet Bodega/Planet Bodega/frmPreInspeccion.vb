Public Class frmPreInspeccion
    Dim clase As New class_library
    Public Ruta, ConsImport As String
    Public Archivo As Boolean = False
    Dim DsExcel As New DataSet
    Dim Validado As Boolean = False
    Dim ColRef, ColDesc, ColCant, ColCaja, ColMarca, ColProv, ColMaterial, ColSubpartida, ColOrigen, ColMedida, ColPrecio, ColTotal, ColCosto, ColFechaReg, ColRegistro As Integer
    Dim FechaRecepcion As Date
    Dim Hoja As String = ""

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        'SOLICITAMOS EL ARCHIVO DE EXCEL DE LA PRE INSPECCION
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Title = "SELECIONE ARCHIVO EXCEL"
        openFileDialog1.Filter = "Archivo Pre Inspeccion|*.xls;*.xlsx"

        'SE OBTIENE LA RUTA DEL ARCHIVO SELECCIONADO
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Ruta = openFileDialog1.FileName
            txtRuta.Text = openFileDialog1.FileName
            LlenarGrid()
            btnActualizar.Enabled = True
            Archivo = True
        End If
        openFileDialog1.Dispose()
    End Sub

    Public Sub LlenarGrid()
        grdPreInspeccion.DataSource = Nothing

        'CONSULTAR EL ARCHIVO DE EXCEL Y LLENAR EL DATASET CON LA INFORMACION DEL MISMO
        Hoja = InputBox("POR FAVOR DIGITE EL NOMBRE DE LA HOJA DEL DOCUMENTO DE EXCEL", "PLANET LOVE")

        If clase.consultarExcel("select * from [" & Hoja & "$]", "excel", Ruta) = True Then
            grdPreInspeccion.DataSource = clase.Ds.Tables("excel")
            DsExcel = clase.Ds
            grdPreInspeccion.RowHeadersVisible = False
        Else
            Limpiar()
        End If
    End Sub

    Public Sub LlenarGrid1()
        grdPreInspeccion.DataSource = Nothing

        If clase.consultarExcel("select * from [" & Hoja & "$]", "excel", Ruta) = True Then
            grdPreInspeccion.DataSource = clase.Ds.Tables("excel")
            DsExcel = clase.Ds
            grdPreInspeccion.RowHeadersVisible = False
        Else
            Limpiar()
        End If
    End Sub

    Private Sub btnImportacion_Click(sender As Object, e As EventArgs) Handles btnImportacion.Click
        frm_listado_importacion.ShowDialog()
        frm_listado_importacion.Dispose()
    End Sub

    Private Sub btnValidar_Click(sender As Object, e As EventArgs) Handles btnValidar.Click
        Dim Fecha As String = dtpFecha.Value.Date.ToString("yyyy-MM-dd")
        If clase.validar_cajas_text(txtImportacion, "Importación") = False Then Exit Sub
        btnValidar.Enabled = False
        'VALIDAR NOMBRE DE LAS COLUMNAS REQUERIDAS
        For j = 0 To grdPreInspeccion.Columns.Count - 1
            'If grdPreInspeccion.Columns(j).HeaderText = "# DE CAJA" Then
            '    ColCaja = j + 1
            'End If
            If grdPreInspeccion.Columns(j).HeaderText = "REFERENCIA" Then
                ColRef = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "DESCRIPCION" Then
                ColDesc = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "CANT" Then
                ColCant = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "MARCA" Then
                ColMarca = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "IDPROVEEDOR" Then
                ColProv = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "MATERIAL" Then
                ColMaterial = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "SUBPARTIDA" Then
                ColSubpartida = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "ORIGEN" Then
                ColOrigen = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "MEDIDA" Then
                ColMedida = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "PRECIO" Then
                ColPrecio = j + 1
            End If
            'If grdPreInspeccion.Columns(j).HeaderText = "TOTAL" Then
            '    ColTotal = j + 1
            'End If
            If grdPreInspeccion.Columns(j).HeaderText = "COSTO" Then
                ColCosto = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "FECHA REGISTRO IMPORTACION" Then
                ColFechaReg = j + 1
            End If
            If grdPreInspeccion.Columns(j).HeaderText = "REGISTRO IMPORTACION" Then
                ColRegistro = j + 1
            End If
        Next

        'If ColCaja = 0 Then
        '    MessageBox.Show("No se encuentra columna # DE CAJA, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    '  Limpiar()
        '    Exit Sub
        'Else
        '    ColCaja = ColCaja - 1
        'End If
        If ColRef = 0 Then
            MessageBox.Show("No se encuentra columna REFERENCIA, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColRef = ColRef - 1
        End If
        If ColDesc = 0 Then
            MessageBox.Show("No se encuentra columna DESCRIPCION, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColDesc = ColDesc - 1
        End If
        If ColCant = 0 Then
            MessageBox.Show("No se encuentra columna CANT, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Limpiar()
            Exit Sub
        Else
            ColCant = ColCant - 1
        End If
        If ColProv = 0 Then
            MessageBox.Show("No se encuentra columna IDPROVEEDOR, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '   Limpiar()
            Exit Sub
        Else
            ColProv = ColProv - 1
        End If
        If ColMarca = 0 Then
            MessageBox.Show("No se encuentra columna MARCA, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColMarca = ColMarca - 1
        End If
        If ColMaterial = 0 Then
            MessageBox.Show("No se encuentra columna MATERIAL, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColMaterial = ColMaterial - 1
        End If
        If ColSubpartida = 0 Then
            MessageBox.Show("No se encuentra columna SUBPARTIDA, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Limpiar()
            Exit Sub
        Else
            ColSubpartida = ColSubpartida - 1
        End If
        If ColOrigen = 0 Then
            MessageBox.Show("No se encuentra columna ORIGEN, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColOrigen = ColOrigen - 1
        End If
        If ColMedida = 0 Then
            MessageBox.Show("No se encuentra columna MEDIDA, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColMedida = ColMedida - 1
        End If
        If ColPrecio = 0 Then
            MessageBox.Show("No se encuentra columna PRECIO, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Limpiar()
            Exit Sub
        Else
            ColPrecio = ColPrecio - 1
        End If
        'If ColTotal = 0 Then
        '    MessageBox.Show("No se encuentra columna TOTAL, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    '  Limpiar()
        '    Exit Sub
        'Else
        '    ColTotal = ColTotal - 1
        'End If
        If ColCosto = 0 Then
            MessageBox.Show("No se encuentra columna COSTO, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColCosto = ColCosto - 1
        End If
        If ColFechaReg = 0 Then
            MessageBox.Show("No se encuentra columna FECHA REGISTRO IMPORTACION, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColFechaReg = ColFechaReg - 1
        End If
        If ColRegistro = 0 Then
            MessageBox.Show("No se encuentra columna REGISTRO IMPORTACION, revise el archivo y vuelva a intentarlo", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  Limpiar()
            Exit Sub
        Else
            ColRegistro = ColRegistro - 1
        End If

        'VALIDAR CONTENIDO DEL GRID ANTES DE GUARDARLO EN LA TABLA DET_CARGAR_PREINSPECCION 
        For j = 0 To grdPreInspeccion.Columns.Count - 1
            For i = 0 To grdPreInspeccion.RowCount - 1
                If grdPreInspeccion(j, i).Value.ToString = "" Then
                    If grdPreInspeccion.Columns(j).HeaderText = "FECHA REGISTRO IMPORTACION" Or grdPreInspeccion.Columns(j).HeaderText = "REGISTRO IMPORTACION" Then
                        Continue For
                    Else
                        MessageBox.Show("No debe haber campos en blanco, corregir la fila " & i + 2 & " en la columna " & grdPreInspeccion.Columns(j).HeaderText & ",  e intente nuevamente.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        '     Limpiar()
                        Exit Sub
                    End If

                End If
                'VALIDAR UNIDADES DE MEDIDA
                If grdPreInspeccion.Columns(j).HeaderText = "MEDIDA" Then
                    If grdPreInspeccion(j, i).Value <> "PAIR" And grdPreInspeccion(j, i).Value <> "PCS" And grdPreInspeccion(j, i).Value <> "DZ" Then
                        MessageBox.Show("No se reconoce tipo de medida en la casilla " & i + 2 & " de la columna unidad de medida, revise e intente nuevamente", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        '  Limpiar()
                        Exit Sub
                    End If
                    If grdPreInspeccion(j, i).Value = "PCS" Then
                        grdPreInspeccion(j, i).Value = "Unidad(s)"
                    End If
                    If grdPreInspeccion(j, i).Value = "PAIR" Then
                        grdPreInspeccion(j, i).Value = "Unidad(s)"
                    End If
                    If grdPreInspeccion(j, i).Value = "DZ" Then
                        grdPreInspeccion(j, i).Value = "Docena(s)"
                    End If
                End If

                'VALIDAR PROVEEDOR
                If grdPreInspeccion.Columns(j).HeaderText = "IDPROVEEDOR" Then
                    clase.consultar("SELECT * FROM proveedores WHERE prv_codigoasignado='" & grdPreInspeccion(j, i).Value & "'", "prov")
                    If clase.dt.Tables("prov").Rows.Count = 0 Then
                        MessageBox.Show("No se reconoce el proveedor en la fila " & i + 2 & " de la columna PROVEEDOR, revise e intente nuevamente", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        '  Limpiar()
                        Exit Sub
                    End If
                End If

                'VALIDAR NUMEROS 
                If grdPreInspeccion.Columns(j).HeaderText = "CANT" Or grdPreInspeccion.Columns(j).HeaderText = "TOTAL" Or grdPreInspeccion.Columns(j).HeaderText = "COSTO" Then
                    If IsNumeric(grdPreInspeccion(j, i).Value) = False Then
                        MessageBox.Show("El valor debe ser numerico en la fila " & i + 2 & " de la columna CANT, revise e intente nuevamente", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            Next
        Next

        'VALIDAR # DE CAJA
        'Dim Cad As String = ""
        'For i = 0 To grdPreInspeccion.RowCount - 1
        '    Cad = Mid(grdPreInspeccion(ColCaja, i).Value, 1, 5)
        '    If Cad <> grdPreInspeccion(ColProv, i).Value Then
        '        MessageBox.Show("# DE CAJA no valido en la fila " & i + 2 & " de la columna # DE CAJA, revise e intente nuevamente", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        '  Limpiar()
        '        Exit Sub
        '    End If
        '    Cad = Mid(grdPreInspeccion(ColCaja, i).Value, 6)
        '    If IsNumeric(Cad) = False Then
        '        MessageBox.Show("# DE CAJA no valido en la fila " & i + 2 & " de la columna # DE CAJA, revise e intente nuevamente", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        '  Limpiar()
        '        Exit Sub
        '    End If
        'Next

        Validado = True
        btnGuardar.Enabled = True
        MessageBox.Show("El archivo se encuentra listo para guardar.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
        btnActualizar.Enabled = False
    End Sub
    Public Sub Limpiar()
        txtRuta.Text = ""
        txtImportacion.Text = ""
        grdPreInspeccion.DataSource = Nothing
        Validado = False
        OpenFileDialog1.Dispose()
        ColCaja = 0
        ColCant = 0
        ColCosto = 0
        ColDesc = 0
        ColFechaReg = 0
        ColMarca = 0
        ColMaterial = 0
        ColMedida = 0
        ColOrigen = 0
        ColPrecio = 0
        ColProv = 0
        ColRef = 0
        ColRegistro = 0
        ColSubpartida = 0
        ColTotal = 0
        Hoja = ""
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Dim IdImportacion As String = ""
        Dim IdCabPreinspeccion As String = ""

        If Validado = False Then
            MessageBox.Show("El archivo no ha pasado el proceso de validación", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        clase.borradoautomatico("DELETE FROM cab_cargar_preinspeccion")

        'CONSULTAR CODIGO IMPORTACION
        clase.consultar("SELECT * FROM cabimportacion WHERE imp_nombrefecha='" & txtImportacion.Text & "'", "importacion")
        IdImportacion = clase.dt.Tables("importacion").Rows(0)("imp_codigo")

        'GUARDAR ENCABEZADO EXCEL EN TABLA CAB_CARGAR_PREINSPECCION
        clase.agregar_registro("INSERT INTO cab_cargar_preinspeccion(cab_fecha_preinspeccion,cab_importacion) VALUES('" & Date.Today.ToString("yyyy-MM-dd") & "','" & IdImportacion & "')")
        clase.consultar("SELECT MAX(cab_id_preinspeccion) AS ID FROM cab_cargar_preinspeccion", "consecutivo")
        IdCabPreinspeccion = clase.dt.Tables("consecutivo").Rows(0)("ID")

        clase.consultar1("select * from consecutivo_codigobarra_importacion", "ultimo")
        Dim ultimocodigo As Integer
        ultimocodigo = clase.dt1.Tables("ultimo").Rows(0)("codigo_ultimo") + 1

        'GUARDAR DETALLE EN TABLA DET_CARGAR_PREINSPECCION
        For i = 0 To grdPreInspeccion.RowCount - 1
            If (IsDBNull(grdPreInspeccion(ColRegistro, i).Value) And IsDBNull(grdPreInspeccion(ColFechaReg, i).Value)) Then
                clase.agregar_registro("INSERT INTO det_cargar_preinspeccion(pre_idcab,pre_numero_caja,pre_referencia,pre_descripcion,pre_subpartida,pre_marca,pre_material,pre_cant,pre_und_medida,pre_idproveedor,pre_origen,pre_precio,pre_costo,pre_fecha_registro_importacion,pre_registro_importacion) VALUES('" & IdCabPreinspeccion & "','" & ultimocodigo & "','" & grdPreInspeccion(ColRef, i).Value & "','" & grdPreInspeccion(ColDesc, i).Value & "','" & grdPreInspeccion(ColSubpartida, i).Value & "','" & grdPreInspeccion(ColMarca, i).Value & "','" & grdPreInspeccion(ColMaterial, i).Value & "','" & grdPreInspeccion(ColCant, i).Value & "','" & grdPreInspeccion(ColMedida, i).Value & "','" & grdPreInspeccion(ColProv, i).Value & "','" & grdPreInspeccion(ColOrigen, i).Value & "','" & Str(grdPreInspeccion(ColPrecio, i).Value) & "','" & Str(grdPreInspeccion(ColCosto, i).Value) & "',NULL,NULL)")
            Else
                If IsDBNull(grdPreInspeccion(ColFechaReg, i).Value) Then
                    clase.agregar_registro("INSERT INTO det_cargar_preinspeccion(pre_idcab,pre_numero_caja,pre_referencia,pre_descripcion,pre_subpartida,pre_marca,pre_material,pre_cant,pre_und_medida,pre_idproveedor,pre_origen,pre_precio,pre_total,pre_costo,pre_fecha_registro_importacion,pre_registro_importacion) VALUES('" & IdCabPreinspeccion & "','" & ultimocodigo & "','" & grdPreInspeccion(ColRef, i).Value & "','" & grdPreInspeccion(ColDesc, i).Value & "','" & grdPreInspeccion(ColSubpartida, i).Value & "','" & grdPreInspeccion(ColMarca, i).Value & "','" & grdPreInspeccion(ColMaterial, i).Value & "','" & grdPreInspeccion(ColCant, i).Value & "','" & grdPreInspeccion(ColMedida, i).Value & "','" & grdPreInspeccion(ColProv, i).Value & "','" & grdPreInspeccion(ColOrigen, i).Value & "','" & Str(grdPreInspeccion(ColPrecio, i).Value) & "',NULL,'" & Str(grdPreInspeccion(ColCosto, i).Value) & "',NULL,'" & grdPreInspeccion(ColRegistro, i).Value & "')")
                Else
                    Dim Fbase As Date = grdPreInspeccion(ColFechaReg, i).Value
                    Dim Fecha As String = Fbase.ToString("yyyy-MM-dd")
                  
                    clase.agregar_registro("INSERT INTO det_cargar_preinspeccion(pre_idcab,pre_numero_caja,pre_referencia,pre_descripcion,pre_subpartida,pre_marca,pre_material,pre_cant,pre_und_medida,pre_idproveedor,pre_origen,pre_precio,pre_total,pre_costo,pre_fecha_registro_importacion,pre_registro_importacion) VALUES('" & IdCabPreinspeccion & "','" & ultimocodigo & "','" & grdPreInspeccion(ColRef, i).Value & "','" & grdPreInspeccion(ColDesc, i).Value & "','" & grdPreInspeccion(ColSubpartida, i).Value & "','" & grdPreInspeccion(ColMarca, i).Value & "','" & grdPreInspeccion(ColMaterial, i).Value & "','" & grdPreInspeccion(ColCant, i).Value & "','" & grdPreInspeccion(ColMedida, i).Value & "','" & grdPreInspeccion(ColProv, i).Value & "','" & grdPreInspeccion(ColOrigen, i).Value & "','" & Str(grdPreInspeccion(ColPrecio, i).Value) & "',NULL,'" & Str(grdPreInspeccion(ColCosto, i).Value) & "','" & Fecha & "','" & grdPreInspeccion(ColRegistro, i).Value & "')")
                End If
            End If
            ultimocodigo += 1
        Next
        clase.actualizar("UPDATE consecutivo_codigobarra_importacion SET codigo_ultimo = " & ultimocodigo & "")


        Dim Prov As String = ""
        Dim fechas As Date = dtpFecha.Value
        clase.consultar("SELECT pre_numero_caja,pre_idproveedor FROM det_cargar_preinspeccion WHERE pre_idcab='" & IdCabPreinspeccion & "' GROUP BY pre_numero_caja", "cabcajas")
        For i = 0 To clase.dt.Tables("cabcajas").Rows.Count - 1
            'GUARDAR ENCABEZADO
            clase.consultar1("SELECT * FROM proveedores WHERE prv_codigoasignado='" & clase.dt.Tables("cabcajas").Rows(i)("pre_idproveedor") & "'", "prov")
            Prov = clase.dt1.Tables("prov").Rows(0)("prv_codigo")

            clase.agregar_registro("INSERT INTO detalle_importacion_cabcajas(det_codigoimportacion,det_caja,det_codigoproveedor,det_peso,det_hoja,det_recepcion,det_fecharecepcion) VALUES('" & IdImportacion & "','" & clase.dt.Tables("cabcajas").Rows(i)("pre_numero_caja") & "','" & Prov & "',NULL,NULL,TRUE,'" & fechas.ToString("yyyy-MM-dd") & "')")

            clase.consultar1("SELECT * FROM det_cargar_preinspeccion WHERE pre_numero_caja='" & clase.dt.Tables("cabcajas").Rows(i)("pre_numero_caja") & "'", "detcajas")
            For a = 0 To clase.dt1.Tables("detcajas").Rows.Count - 1
                'GUARDAR DETALLE
                If IsDBNull(clase.dt1.Tables("detcajas").Rows(a)("pre_registro_importacion")) And IsDBNull(clase.dt1.Tables("detcajas").Rows(a)("pre_fecha_registro_importacion")) Then
                    clase.agregar_registro("INSERT INTO detalle_importacion_detcajas(detcab_codigocaja,detcab_referencia,detcab_descripcion,detcab_marca,detcab_composicion,detcab_costodolares,detcab_costopesos_x_pieza,detcab_cantidad,detcab_unimedida,detcab_registrodian,detcab_fecharegistrodian,detcab_subpartida,detcab_origen) VALUES('" & clase.dt1.Tables("detcajas").Rows(a)("pre_numero_caja") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_referencia") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_descripcion") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_marca") & "','" & Replace(clase.dt1.Tables("detcajas").Rows(a)("pre_material"), "'", "") & "','" & Str(clase.dt1.Tables("detcajas").Rows(a)("pre_precio")) & "','" & Str(clase.dt1.Tables("detcajas").Rows(a)("pre_costo")) & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_cant") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_und_medida") & "',NULL,NULL,'" & clase.dt1.Tables("detcajas").Rows(a)("pre_subpartida") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_origen") & "')")
                Else
                    If IsDBNull(clase.dt1.Tables("detcajas").Rows(a)("pre_fecha_registro_importacion")) Then
                        clase.agregar_registro("INSERT INTO detalle_importacion_detcajas(detcab_codigocaja,detcab_referencia,detcab_descripcion,detcab_marca,detcab_composicion,detcab_costodolares,detcab_costopesos_x_pieza,detcab_cantidad,detcab_unimedida,detcab_registrodian,detcab_fecharegistrodian,detcab_subpartida,detcab_origen) VALUES('" & clase.dt1.Tables("detcajas").Rows(a)("pre_numero_caja") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_referencia") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_descripcion") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_marca") & "','" & Replace(clase.dt1.Tables("detcajas").Rows(a)("pre_material"), "'", "") & "','" & Str(clase.dt1.Tables("detcajas").Rows(a)("pre_precio")) & "','" & Str(clase.dt1.Tables("detcajas").Rows(a)("pre_costo")) & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_cant") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_und_medida") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_registro_importacion") & "',NULL,'" & clase.dt1.Tables("detcajas").Rows(a)("pre_subpartida") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_origen") & "')")
                    Else
                        Dim FechaBase As Date = clase.dt1.Tables("detcajas").Rows(a)("pre_fecha_registro_importacion")
                        Dim FechaFin As String = FechaBase.ToString("yyyy-MM-dd")
                        clase.agregar_registro("INSERT INTO detalle_importacion_detcajas(detcab_codigocaja,detcab_referencia,detcab_descripcion,detcab_marca,detcab_composicion,detcab_costodolares,detcab_costopesos_x_pieza,detcab_cantidad,detcab_unimedida,detcab_registrodian,detcab_fecharegistrodian,detcab_subpartida,detcab_origen) VALUES('" & clase.dt1.Tables("detcajas").Rows(a)("pre_numero_caja") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_referencia") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_descripcion") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_marca") & "','" & Replace(clase.dt1.Tables("detcajas").Rows(a)("pre_material"), "'", "") & "','" & Str(clase.dt1.Tables("detcajas").Rows(a)("pre_precio")) & "','" & Str(clase.dt1.Tables("detcajas").Rows(a)("pre_costo")) & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_cant") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_und_medida") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_registro_importacion") & "','" & FechaFin & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_subpartida") & "','" & clase.dt1.Tables("detcajas").Rows(a)("pre_origen") & "')")
                    End If
                End If

            Next
        Next

        clase.borradoautomatico("DELETE FROM cab_cargar_preinspeccion")

        MessageBox.Show("Importación guardada con éxito.", "PLANET LOVE", MessageBoxButtons.OK, MessageBoxIcon.Information)
        btnGuardar.Enabled = False
        Limpiar()
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Limpiar()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnActualizar.Click
        LlenarGrid1()
        btnValidar.Enabled = True
    End Sub
End Class
