Public Class frmPreTransferencia
    Dim clase As New class_library
    Dim IdTienda As String

    Private Sub frmPreTransferencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LlenarGridTiendas()
    End Sub

    Public Sub LlenarGridTiendas()
        grdTiendas.DataSource = Nothing
        clase.consultar("SELECT id AS IDTIENDA,tienda AS TIENDAS FROM tiendas ORDER BY tienda ASC", "tiendas")
        grdTiendas.DataSource = clase.dt.Tables("tiendas")
        grdTiendas.Columns(0).Visible = False
        grdTiendas.Columns(1).Width = 200
        grdTiendas.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        grdTiendas.RowHeadersWidth = 4
    End Sub

    Public Sub LlenarGridPreTransfer()
        grdPreTransfer.DataSource = Nothing
        clase.consultar("SELECT distribucion_orden_de_produccion.articulo AS CODIGO, articulos.ar_referencia AS REFERENCIA, articulos.ar_descripcion AS DESCRIPCION, distribucion_orden_de_produccion.cantidad AS CANT FROM distribucion_orden_de_produccion INNER JOIN articulos ON (distribucion_orden_de_produccion.articulo = articulos.ar_codigo) WHERE (distribucion_orden_de_produccion.tienda ='" & IdTienda & "' AND distribucion_orden_de_produccion.ordendeproduccion ='" & frmAsignar.txtOrden.Text & "');", "pretransfer")
        grdPreTransfer.DataSource = clase.dt.Tables("pretransfer")
        grdPreTransfer.RowHeadersWidth = 4
        grdPreTransfer.Columns(0).Width = 70
        grdPreTransfer.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdPreTransfer.Columns(1).Width = 200
        grdPreTransfer.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        grdPreTransfer.Columns(2).Width = 200
        grdPreTransfer.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        grdPreTransfer.Columns(3).Width = 50
        grdPreTransfer.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        frmAsignar.Enabled = True
        Me.Close()
    End Sub

    Private Sub grdTiendas_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdTiendas.CellClick
        Dim Y As Integer = grdTiendas.CurrentCell.RowIndex
        IdTienda = grdTiendas(0, [Y]).Value.ToString
        LlenarGridPreTransfer()
    End Sub
End Class