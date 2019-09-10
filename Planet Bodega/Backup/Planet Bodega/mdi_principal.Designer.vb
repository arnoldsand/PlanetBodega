<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mdi_principal
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(mdi_principal))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip
        Me.FileMenu = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PuntosDeVentaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.LineasYSublineasDeArticuloToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TallasYColoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BodegasBloquesGondolasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PatroneDeDistribuciónToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SalirToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TareasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ModificacionesDePrecioToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.InventarioToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.StocksDeArticulosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EntradaDeArticuloToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AjustesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TrasladosEntreBodegasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MermasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TransferenciasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MantenientoDeTransferenciasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NuevaTrasnferenciaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ConsultasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStrip = New System.Windows.Forms.ToolStrip
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.StatusStrip = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImportacionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImprimirCodigosDeBarraParaCajasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NewToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.OpenToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.PrintToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.PrintPreviewToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.HelpToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip.SuspendLayout()
        Me.ToolStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileMenu, Me.TareasToolStripMenuItem, Me.InventarioToolStripMenuItem, Me.TransferenciasToolStripMenuItem, Me.ConsultasToolStripMenuItem, Me.ImportacionToolStripMenuItem})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(632, 24)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        '
        'FileMenu
        '
        Me.FileMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.ExitToolStripMenuItem, Me.PuntosDeVentaToolStripMenuItem, Me.LineasYSublineasDeArticuloToolStripMenuItem, Me.TallasYColoresToolStripMenuItem, Me.ProveedoresToolStripMenuItem, Me.BodegasBloquesGondolasToolStripMenuItem, Me.PatroneDeDistribuciónToolStripMenuItem, Me.SalirToolStripMenuItem})
        Me.FileMenu.ImageTransparentColor = System.Drawing.SystemColors.ActiveBorder
        Me.FileMenu.Name = "FileMenu"
        Me.FileMenu.Size = New System.Drawing.Size(49, 20)
        Me.FileMenu.Text = "&Datos"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.ExitToolStripMenuItem.Text = "Mantenimiento de Articulos"
        '
        'PuntosDeVentaToolStripMenuItem
        '
        Me.PuntosDeVentaToolStripMenuItem.Name = "PuntosDeVentaToolStripMenuItem"
        Me.PuntosDeVentaToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.PuntosDeVentaToolStripMenuItem.Text = "Puntos de Venta"
        '
        'LineasYSublineasDeArticuloToolStripMenuItem
        '
        Me.LineasYSublineasDeArticuloToolStripMenuItem.Name = "LineasYSublineasDeArticuloToolStripMenuItem"
        Me.LineasYSublineasDeArticuloToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.LineasYSublineasDeArticuloToolStripMenuItem.Text = "Lineas y Sublineas de Articulo"
        '
        'TallasYColoresToolStripMenuItem
        '
        Me.TallasYColoresToolStripMenuItem.Name = "TallasYColoresToolStripMenuItem"
        Me.TallasYColoresToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.TallasYColoresToolStripMenuItem.Text = "Tallas y Colores"
        '
        'ProveedoresToolStripMenuItem
        '
        Me.ProveedoresToolStripMenuItem.Name = "ProveedoresToolStripMenuItem"
        Me.ProveedoresToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.ProveedoresToolStripMenuItem.Text = "Proveedores"
        '
        'BodegasBloquesGondolasToolStripMenuItem
        '
        Me.BodegasBloquesGondolasToolStripMenuItem.Name = "BodegasBloquesGondolasToolStripMenuItem"
        Me.BodegasBloquesGondolasToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.BodegasBloquesGondolasToolStripMenuItem.Text = "Bodegas - Bloques - Góndolas"
        '
        'PatroneDeDistribuciónToolStripMenuItem
        '
        Me.PatroneDeDistribuciónToolStripMenuItem.Name = "PatroneDeDistribuciónToolStripMenuItem"
        Me.PatroneDeDistribuciónToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.PatroneDeDistribuciónToolStripMenuItem.Text = "Patrones de Distribución"
        '
        'SalirToolStripMenuItem
        '
        Me.SalirToolStripMenuItem.Name = "SalirToolStripMenuItem"
        Me.SalirToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.SalirToolStripMenuItem.Text = "Salir"
        '
        'TareasToolStripMenuItem
        '
        Me.TareasToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ModificacionesDePrecioToolStripMenuItem})
        Me.TareasToolStripMenuItem.Name = "TareasToolStripMenuItem"
        Me.TareasToolStripMenuItem.Size = New System.Drawing.Size(53, 20)
        Me.TareasToolStripMenuItem.Text = "Tareas"
        '
        'ModificacionesDePrecioToolStripMenuItem
        '
        Me.ModificacionesDePrecioToolStripMenuItem.Name = "ModificacionesDePrecioToolStripMenuItem"
        Me.ModificacionesDePrecioToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.ModificacionesDePrecioToolStripMenuItem.Text = "Modificaciones de Precio"
        '
        'InventarioToolStripMenuItem
        '
        Me.InventarioToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StocksDeArticulosToolStripMenuItem, Me.EntradaDeArticuloToolStripMenuItem, Me.AjustesToolStripMenuItem, Me.TrasladosEntreBodegasToolStripMenuItem, Me.MermasToolStripMenuItem})
        Me.InventarioToolStripMenuItem.Name = "InventarioToolStripMenuItem"
        Me.InventarioToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.InventarioToolStripMenuItem.Text = "Bodega"
        '
        'StocksDeArticulosToolStripMenuItem
        '
        Me.StocksDeArticulosToolStripMenuItem.Name = "StocksDeArticulosToolStripMenuItem"
        Me.StocksDeArticulosToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.StocksDeArticulosToolStripMenuItem.Text = "Stock(s) de Articulos"
        '
        'EntradaDeArticuloToolStripMenuItem
        '
        Me.EntradaDeArticuloToolStripMenuItem.Name = "EntradaDeArticuloToolStripMenuItem"
        Me.EntradaDeArticuloToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.EntradaDeArticuloToolStripMenuItem.Text = "Almacenaje de Mercancía"
        '
        'AjustesToolStripMenuItem
        '
        Me.AjustesToolStripMenuItem.Name = "AjustesToolStripMenuItem"
        Me.AjustesToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.AjustesToolStripMenuItem.Text = "Ajustes "
        '
        'TrasladosEntreBodegasToolStripMenuItem
        '
        Me.TrasladosEntreBodegasToolStripMenuItem.Name = "TrasladosEntreBodegasToolStripMenuItem"
        Me.TrasladosEntreBodegasToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.TrasladosEntreBodegasToolStripMenuItem.Text = "Traslados entre Bodegas"
        '
        'MermasToolStripMenuItem
        '
        Me.MermasToolStripMenuItem.Name = "MermasToolStripMenuItem"
        Me.MermasToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.MermasToolStripMenuItem.Text = "Mermas"
        '
        'TransferenciasToolStripMenuItem
        '
        Me.TransferenciasToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MantenientoDeTransferenciasToolStripMenuItem, Me.NuevaTrasnferenciaToolStripMenuItem})
        Me.TransferenciasToolStripMenuItem.Name = "TransferenciasToolStripMenuItem"
        Me.TransferenciasToolStripMenuItem.Size = New System.Drawing.Size(95, 20)
        Me.TransferenciasToolStripMenuItem.Text = "Transferencias"
        '
        'MantenientoDeTransferenciasToolStripMenuItem
        '
        Me.MantenientoDeTransferenciasToolStripMenuItem.Name = "MantenientoDeTransferenciasToolStripMenuItem"
        Me.MantenientoDeTransferenciasToolStripMenuItem.Size = New System.Drawing.Size(237, 22)
        Me.MantenientoDeTransferenciasToolStripMenuItem.Text = "Manteniento de Transferencias"
        '
        'NuevaTrasnferenciaToolStripMenuItem
        '
        Me.NuevaTrasnferenciaToolStripMenuItem.Name = "NuevaTrasnferenciaToolStripMenuItem"
        Me.NuevaTrasnferenciaToolStripMenuItem.Size = New System.Drawing.Size(237, 22)
        Me.NuevaTrasnferenciaToolStripMenuItem.Text = "Nueva Trasnferencia"
        '
        'ConsultasToolStripMenuItem
        '
        Me.ConsultasToolStripMenuItem.Name = "ConsultasToolStripMenuItem"
        Me.ConsultasToolStripMenuItem.Size = New System.Drawing.Size(71, 20)
        Me.ConsultasToolStripMenuItem.Text = "Consultas"
        '
        'ToolStrip
        '
        Me.ToolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.ToolStripSeparator1, Me.PrintToolStripButton, Me.PrintPreviewToolStripButton, Me.ToolStripSeparator2, Me.HelpToolStripButton})
        Me.ToolStrip.Location = New System.Drawing.Point(0, 24)
        Me.ToolStrip.Name = "ToolStrip"
        Me.ToolStrip.Size = New System.Drawing.Size(632, 25)
        Me.ToolStrip.TabIndex = 6
        Me.ToolStrip.Text = "ToolStrip"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 431)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(632, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(42, 17)
        Me.ToolStripStatusLabel.Text = "Estado"
        '
        'ImportacionToolStripMenuItem
        '
        Me.ImportacionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImprimirCodigosDeBarraParaCajasToolStripMenuItem})
        Me.ImportacionToolStripMenuItem.Name = "ImportacionToolStripMenuItem"
        Me.ImportacionToolStripMenuItem.Size = New System.Drawing.Size(84, 20)
        Me.ImportacionToolStripMenuItem.Text = "Importación"
        '
        'ImprimirCodigosDeBarraParaCajasToolStripMenuItem
        '
        Me.ImprimirCodigosDeBarraParaCajasToolStripMenuItem.Name = "ImprimirCodigosDeBarraParaCajasToolStripMenuItem"
        Me.ImprimirCodigosDeBarraParaCajasToolStripMenuItem.Size = New System.Drawing.Size(270, 22)
        Me.ImprimirCodigosDeBarraParaCajasToolStripMenuItem.Text = "Imprimir Codigos de Barra Para Cajas"
        '
        'NewToolStripButton
        '
        Me.NewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.NewToolStripButton.Image = CType(resources.GetObject("NewToolStripButton.Image"), System.Drawing.Image)
        Me.NewToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.NewToolStripButton.Name = "NewToolStripButton"
        Me.NewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.NewToolStripButton.Text = "Nuevo"
        '
        'OpenToolStripButton
        '
        Me.OpenToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.OpenToolStripButton.Image = CType(resources.GetObject("OpenToolStripButton.Image"), System.Drawing.Image)
        Me.OpenToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.OpenToolStripButton.Name = "OpenToolStripButton"
        Me.OpenToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.OpenToolStripButton.Text = "Abrir"
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SaveToolStripButton.Text = "Guardar"
        '
        'PrintToolStripButton
        '
        Me.PrintToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintToolStripButton.Image = CType(resources.GetObject("PrintToolStripButton.Image"), System.Drawing.Image)
        Me.PrintToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.PrintToolStripButton.Name = "PrintToolStripButton"
        Me.PrintToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintToolStripButton.Text = "Imprimir"
        '
        'PrintPreviewToolStripButton
        '
        Me.PrintPreviewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintPreviewToolStripButton.Image = CType(resources.GetObject("PrintPreviewToolStripButton.Image"), System.Drawing.Image)
        Me.PrintPreviewToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.PrintPreviewToolStripButton.Name = "PrintPreviewToolStripButton"
        Me.PrintPreviewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintPreviewToolStripButton.Text = "Vista previa de impresión"
        '
        'HelpToolStripButton
        '
        Me.HelpToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.HelpToolStripButton.Image = CType(resources.GetObject("HelpToolStripButton.Image"), System.Drawing.Image)
        Me.HelpToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.HelpToolStripButton.Name = "HelpToolStripButton"
        Me.HelpToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.HelpToolStripButton.Text = "Ayuda"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Image = Global.WindowsApplication1.My.Resources.Resources.application_form_edit
        Me.NewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Black
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(269, 22)
        Me.NewToolStripMenuItem.Text = "&Creación de Articulos x Lotes"
        '
        'mdi_principal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(632, 453)
        Me.Controls.Add(Me.ToolStrip)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "mdi_principal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Planet Bodega"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.ToolStrip.ResumeLayout(False)
        Me.ToolStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents HelpToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents PrintPreviewToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents PrintToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents NewToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStrip As System.Windows.Forms.ToolStrip
    Friend WithEvents OpenToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents SaveToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents PuntosDeVentaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LineasYSublineasDeArticuloToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TallasYColoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TareasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ModificacionesDePrecioToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InventarioToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransferenciasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StocksDeArticulosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EntradaDeArticuloToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AjustesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ConsultasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProveedoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TrasladosEntreBodegasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MermasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalirToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MantenientoDeTransferenciasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NuevaTrasnferenciaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BodegasBloquesGondolasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PatroneDeDistribuciónToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportacionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImprimirCodigosDeBarraParaCajasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
