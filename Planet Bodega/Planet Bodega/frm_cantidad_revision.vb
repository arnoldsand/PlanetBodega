Public Class frm_cantidad_revision
    Dim clase As New class_library
    Public cantidadrevisada As Integer
    Public TituloForm As String
    Public FormularioCerrado As Boolean
    Dim indCerrad As Boolean

    Private Sub frm_cantidad_revision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisableCloseButton(Me.Handle)
        txtcant.Focus()
        cantidadrevisada = vbEmpty
        Me.cantidadrevisada = 0
        Me.FormularioCerrado = False
        indCerrad = False
        ' Me.AcceptButton = btnAceptar
        If Me.TituloForm <> "" Then
            Me.Text = Me.TituloForm
        End If
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.FormularioCerrado = True
        Me.Close()
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        If clase.validar_cajas_text(txtcant, "Cantidad") = False Then Exit Sub
        If Val(txtcant.Text) = 0 Then
            MessageBox.Show("Debe digitar una cantidad valida.", "CANTIDAD NO VALIDA", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If Len(txtcant.Text) > 3 Then
            MessageBox.Show("Cantidad invalida, por favor revisela", "Cantidad No Valida", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtcant.Text = ""
            txtcant.Focus()
        Else
            cantidadrevisada = txtcant.Text
            indCerrad = True
            Me.Close()
        End If
    End Sub

    Private Sub txtcant_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtcant.KeyPress
        clase.validar_numeros(e)
        If Len(txtcant.Text) > 2 Then
            If Asc(e.KeyChar) = 8 Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
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

    Private Sub frm_cantidad_revision_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If indCerrad = False Then
            Me.FormularioCerrado = True
        End If
    End Sub
End Class