Public Class frm_diferencias_revision
    Dim clase As New class_library
    Public validado As Boolean


    Private Sub frm_diferencias_revision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisableCloseButton(Me.Handle)
        diferencias.Columns(0).ReadOnly = True
        diferencias.Columns(1).ReadOnly = True
        diferencias.Columns(2).ReadOnly = True
        diferencias.Columns(3).ReadOnly = False
        validado = False
        Dim x As Short
        For x = 0 To frm_revision.grdRevision.RowCount - 1
            If (frm_revision.grdRevision.Item(4, x).Value = "Inconsistencia") Or (frm_revision.grdRevision.Item(4, x).Value = "Agregado") Then
                diferencias.RowCount = diferencias.RowCount + 1
                diferencias.Item(0, diferencias.RowCount - 1).Value = frm_revision.grdRevision.Item(0, x).Value
                diferencias.Item(1, diferencias.RowCount - 1).Value = frm_revision.grdRevision.Item(1, x).Value
                diferencias.Item(2, diferencias.RowCount - 1).Value = frm_revision.grdRevision.Item(2, x).Value
                diferencias.Item(3, diferencias.RowCount - 1).Value = frm_revision.grdRevision.Item(3, x).Value
            End If
        Next
        'Dim p As Short
        'For p = 0 To frm_revision.grdRevision.RowCount - 1
        '    If frm_revision.grdRevision.Item(4, p).Value = "Agregado" Then
        '        clase.agregar_registro("INSERT INTO `dettransferencia` (`dt_trnumero`, `dt_codarticulo`, `dt_cantidad`, `dt_costo`, `dt_venta1`, `dt_venta2`) VALUES ('" & frm_revision.textBox2.Text & "', '" & frm_revision.grdRevision.Item(0, p).Value & "', '" & frm_revision.grdRevision.Item(3, p).Value & "', '" & Str(precio_costo_articulo(frm_revision.grdRevision.Item(0, p).Value)) & "', '" & Str(precio_venta1_articulo(frm_revision.grdRevision.Item(0, p).Value)) & "', '" & Str(precio_venta2_articulo(frm_revision.grdRevision.Item(0, p).Value)) & "')")
        '    End If
        'Next
    End Sub

    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As IntPtr, ByVal nPosition As Integer, ByVal wFlags As Long) As IntPtr
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As IntPtr, ByVal bRevert As Boolean) As IntPtr
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As IntPtr) As Integer
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As IntPtr) As Boolean
    Private Const MF_BYPOSITION = &H400
    Private Const MF_REMOVE = &H1000
    Private Const MF_DISABLED = &H2

    Public Sub DisableCloseButton(ByVal hwnd As IntPtr)
        Dim hMenu As IntPtr
        Dim menuItemCount As Integer
        hMenu = GetSystemMenu(hwnd, False)
        menuItemCount = GetMenuItemCount(hMenu)
        Call RemoveMenu(hMenu, menuItemCount - 1, MF_DISABLED Or MF_BYPOSITION)
        Call RemoveMenu(hMenu, menuItemCount - 2, MF_DISABLED Or MF_BYPOSITION)
        Call DrawMenuBar(hwnd)
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        Dim x As Short
        Dim ind1 As Boolean = False
        For x = 0 To diferencias.RowCount - 1
            If Val(diferencias.Item(3, x).Value) <= 0 Then
                ind1 = True
                diferencias.CurrentCell = diferencias.Item(3, x)
                MessageBox.Show("Hay algunos valores que no son validos en el campo cantidad, revise y vuelva a intentarlo.", "VALORES INVALIDOS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Next
        frm_contrasena_revision.ShowDialog()
        validado = frm_contrasena_revision.validado
        If frm_contrasena_revision.validado = True Then
            Dim faltante As Short
            For x = 0 To diferencias.RowCount - 1
                faltante = frm_revision.obtener_cant_total_transferencia(frm_revision.textBox2.Text, diferencias.Item(0, x).Value) - diferencias.Item(3, x).Value
                If faltante <> 0 Then
                    clase.agregar_registro("insert into faltante_transferencia (rev_transferencia, rev_articulo, rev_faltante) values ('" & frm_revision.textBox2.Text & "', '" & diferencias.Item(0, x).Value & "', '" & faltante & "') ")
                    clase.actualizar("update dettransferencia set dt_cantidad = " & diferencias.Item(3, x).Value & " where dt_trnumero = " & frm_revision.textBox2.Text & " and dt_codarticulo = " & diferencias.Item(0, x).Value)
                End If
            Next
        End If
        frm_contrasena_revision.Dispose()
        Me.Close()
    End Sub

    Function precio_costo_articulo(articulo As Long) As Double
        clase.consultar2("SELECT ar_costo FROM articulos WHERE (ar_codigo =" & articulo & ")", "costo")
        Return clase.dt2.Tables("costo").Rows(0)("ar_costo")
    End Function

    Function precio_venta1_articulo(articulo As Long) As Double
        clase.consultar2("SELECT ar_precio1 FROM articulos WHERE (ar_codigo =" & articulo & ")", "precio1")
        Return clase.dt2.Tables("precio1").Rows(0)("ar_precio1")
    End Function

    Function precio_venta2_articulo(articulo As Long) As Double
        clase.consultar2("SELECT ar_precio2 FROM articulos WHERE (ar_codigo =" & articulo & ")", "precio2")
        Return clase.dt2.Tables("precio2").Rows(0)("ar_precio2")
    End Function


End Class

