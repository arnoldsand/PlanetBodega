Public Class frm_articulos_faltantes_x_revision
    Dim clase As New class_library
    Public indicador_no_agregado As Boolean
    Private Sub frm_articulos_faltantes_x_revision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisableCloseButton(Me.Handle)
        grillafaltantes.Columns(3).Visible = False 'OJO la columna 3 donde esta la cantidad esta oculta, es decir con la propiedad visible = false
        indicador_no_agregado = False
        clase.consultar1("SELECT articulos.ar_referencia, articulos.ar_descripcion, dettransferencia.dt_codarticulo, dettransferencia.dt_cantidad FROM dettransferencia INNER JOIN articulos ON (dettransferencia.dt_codarticulo = articulos.ar_codigo) WHERE (dettransferencia.dt_trnumero =" & frm_revision.textBox2.Text & ")", "detalle")
        If clase.dt1.Tables("detalle").Rows.Count > 0 Then
            Dim x As Short
            Dim z As Short
            Dim ind As Boolean
            For x = 0 To clase.dt1.Tables("detalle").Rows.Count - 1
                ind = False
                For z = 0 To frm_revision.grdRevision.RowCount - 1
                    If frm_revision.grdRevision.Item(0, z).Value = clase.dt1.Tables("detalle").Rows(x)("dt_codarticulo") Then
                        ind = True
                    End If
                Next
                If ind = False Then
                    grillafaltantes.RowCount = grillafaltantes.RowCount + 1
                    grillafaltantes.Item(0, grillafaltantes.RowCount - 1).Value = clase.dt1.Tables("detalle").Rows(x)("dt_codarticulo")
                    grillafaltantes.Item(1, grillafaltantes.RowCount - 1).Value = clase.dt1.Tables("detalle").Rows(x)("ar_referencia")
                    grillafaltantes.Item(2, grillafaltantes.RowCount - 1).Value = clase.dt1.Tables("detalle").Rows(x)("ar_descripcion")
                    grillafaltantes.Item(3, grillafaltantes.RowCount - 1).Value = clase.dt1.Tables("detalle").Rows(x)("dt_cantidad")
                End If
            Next
        End If
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        indicador_no_agregado = True
        Me.Close()
    End Sub

    
    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Dim z As Short
        For z = 0 To grillafaltantes.RowCount - 1
            clase.agregar_registro("insert into faltante_transferencia (rev_transferencia, rev_articulo, rev_faltante) values ('" & frm_revision.textBox2.Text & "', '" & grillafaltantes.Item(0, z).Value & "', '" & grillafaltantes.Item(3, z).Value & "')")
            clase.borradoautomatico("delete from dettransferencia where dt_trnumero = " & frm_revision.textBox2.Text & " and dt_codarticulo = " & grillafaltantes.Item(0, z).Value & "")
            clase.agregar_registro("insert into eliminados_transferencia (trelim_idtransferencia, trelim_articulo, trelim_cantidad) values ('" & frm_revision.textBox2.Text & "', '" & grillafaltantes.Item(0, z).Value & "', '" & grillafaltantes.Item(3, z).Value & "')")
        Next
        indicador_no_agregado = False
        Me.Close()
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
End Class