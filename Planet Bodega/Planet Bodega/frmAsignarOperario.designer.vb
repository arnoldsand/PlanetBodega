<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAsignarOperario
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.txtOperario = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btn6 = New System.Windows.Forms.Button()
        Me.btn7 = New System.Windows.Forms.Button()
        Me.btn8 = New System.Windows.Forms.Button()
        Me.btn9 = New System.Windows.Forms.Button()
        Me.btn10 = New System.Windows.Forms.Button()
        Me.btn5 = New System.Windows.Forms.Button()
        Me.btn4 = New System.Windows.Forms.Button()
        Me.btn3 = New System.Windows.Forms.Button()
        Me.btn2 = New System.Windows.Forms.Button()
        Me.btn1 = New System.Windows.Forms.Button()
        Me.txtCantidad = New System.Windows.Forms.TextBox()
        Me.txtReferencia = New System.Windows.Forms.TextBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.txtCodigo = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnAsignar = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkTodos = New System.Windows.Forms.CheckBox()
        Me.grdReserva = New System.Windows.Forms.DataGridView()
        Me.txtReserva = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.grdPatron = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnAsigReserva = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnDeshacer = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.grdReserva, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdPatron, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCerrar
        '
        Me.btnCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCerrar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCerrar.ForeColor = System.Drawing.Color.Navy
        Me.btnCerrar.Image = Global.WindowsApplication1.My.Resources.Resources.remove
        Me.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCerrar.Location = New System.Drawing.Point(435, 12)
        Me.btnCerrar.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(97, 33)
        Me.btnCerrar.TabIndex = 6
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.ForeColor = System.Drawing.Color.Navy
        Me.btnClear.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClear.Location = New System.Drawing.Point(107, 206)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(52, 50)
        Me.btnClear.TabIndex = 23
        Me.btnClear.Text = "C"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'txtOperario
        '
        Me.txtOperario.BackColor = System.Drawing.Color.White
        Me.txtOperario.Location = New System.Drawing.Point(72, 20)
        Me.txtOperario.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtOperario.Name = "txtOperario"
        Me.txtOperario.ReadOnly = True
        Me.txtOperario.Size = New System.Drawing.Size(52, 23)
        Me.txtOperario.TabIndex = 35
        Me.txtOperario.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 15)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Operario:"
        '
        'btn6
        '
        Me.btn6.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn6.Location = New System.Drawing.Point(107, 107)
        Me.btn6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn6.Name = "btn6"
        Me.btn6.Size = New System.Drawing.Size(52, 50)
        Me.btn6.TabIndex = 29
        Me.btn6.Text = "6"
        Me.btn6.UseVisualStyleBackColor = True
        '
        'btn7
        '
        Me.btn7.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn7.Location = New System.Drawing.Point(5, 157)
        Me.btn7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn7.Name = "btn7"
        Me.btn7.Size = New System.Drawing.Size(52, 50)
        Me.btn7.TabIndex = 30
        Me.btn7.Text = "7"
        Me.btn7.UseVisualStyleBackColor = True
        '
        'btn8
        '
        Me.btn8.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn8.Location = New System.Drawing.Point(56, 157)
        Me.btn8.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn8.Name = "btn8"
        Me.btn8.Size = New System.Drawing.Size(52, 50)
        Me.btn8.TabIndex = 31
        Me.btn8.Text = "8"
        Me.btn8.UseVisualStyleBackColor = True
        '
        'btn9
        '
        Me.btn9.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn9.Location = New System.Drawing.Point(107, 157)
        Me.btn9.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn9.Name = "btn9"
        Me.btn9.Size = New System.Drawing.Size(52, 50)
        Me.btn9.TabIndex = 32
        Me.btn9.Text = "9"
        Me.btn9.UseVisualStyleBackColor = True
        '
        'btn10
        '
        Me.btn10.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn10.Location = New System.Drawing.Point(56, 206)
        Me.btn10.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn10.Name = "btn10"
        Me.btn10.Size = New System.Drawing.Size(52, 50)
        Me.btn10.TabIndex = 33
        Me.btn10.Text = "0"
        Me.btn10.UseVisualStyleBackColor = True
        '
        'btn5
        '
        Me.btn5.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn5.Location = New System.Drawing.Point(56, 107)
        Me.btn5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn5.Name = "btn5"
        Me.btn5.Size = New System.Drawing.Size(52, 50)
        Me.btn5.TabIndex = 28
        Me.btn5.Text = "5"
        Me.btn5.UseVisualStyleBackColor = True
        '
        'btn4
        '
        Me.btn4.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn4.Location = New System.Drawing.Point(5, 107)
        Me.btn4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn4.Name = "btn4"
        Me.btn4.Size = New System.Drawing.Size(52, 50)
        Me.btn4.TabIndex = 27
        Me.btn4.Text = "4"
        Me.btn4.UseVisualStyleBackColor = True
        '
        'btn3
        '
        Me.btn3.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn3.Location = New System.Drawing.Point(107, 58)
        Me.btn3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn3.Name = "btn3"
        Me.btn3.Size = New System.Drawing.Size(52, 50)
        Me.btn3.TabIndex = 26
        Me.btn3.Text = "3"
        Me.btn3.UseVisualStyleBackColor = True
        '
        'btn2
        '
        Me.btn2.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn2.Location = New System.Drawing.Point(56, 58)
        Me.btn2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn2.Name = "btn2"
        Me.btn2.Size = New System.Drawing.Size(52, 50)
        Me.btn2.TabIndex = 25
        Me.btn2.Text = "2"
        Me.btn2.UseVisualStyleBackColor = True
        '
        'btn1
        '
        Me.btn1.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn1.Location = New System.Drawing.Point(5, 58)
        Me.btn1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn1.Name = "btn1"
        Me.btn1.Size = New System.Drawing.Size(52, 50)
        Me.btn1.TabIndex = 24
        Me.btn1.Text = "1"
        Me.btn1.UseVisualStyleBackColor = True
        '
        'txtCantidad
        '
        Me.txtCantidad.BackColor = System.Drawing.Color.White
        Me.txtCantidad.Location = New System.Drawing.Point(385, 18)
        Me.txtCantidad.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtCantidad.Name = "txtCantidad"
        Me.txtCantidad.ReadOnly = True
        Me.txtCantidad.Size = New System.Drawing.Size(62, 23)
        Me.txtCantidad.TabIndex = 43
        Me.txtCantidad.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtReferencia
        '
        Me.txtReferencia.BackColor = System.Drawing.Color.White
        Me.txtReferencia.Location = New System.Drawing.Point(257, 41)
        Me.txtReferencia.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtReferencia.Name = "txtReferencia"
        Me.txtReferencia.ReadOnly = True
        Me.txtReferencia.Size = New System.Drawing.Size(154, 23)
        Me.txtReferencia.TabIndex = 42
        '
        'txtDescripcion
        '
        Me.txtDescripcion.BackColor = System.Drawing.Color.White
        Me.txtDescripcion.Location = New System.Drawing.Point(257, 14)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.ReadOnly = True
        Me.txtDescripcion.Size = New System.Drawing.Size(154, 23)
        Me.txtDescripcion.TabIndex = 41
        '
        'txtCodigo
        '
        Me.txtCodigo.BackColor = System.Drawing.Color.White
        Me.txtCodigo.Location = New System.Drawing.Point(63, 15)
        Me.txtCodigo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.ReadOnly = True
        Me.txtCodigo.Size = New System.Drawing.Size(88, 23)
        Me.txtCodigo.TabIndex = 40
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(182, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 15)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Referencia:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(320, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 15)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "Cantidad:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(178, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 15)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "Descripcion:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 15)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "Codigo:"
        '
        'btnAsignar
        '
        Me.btnAsignar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnAsignar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAsignar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAsignar.ForeColor = System.Drawing.Color.Navy
        Me.btnAsignar.Image = Global.WindowsApplication1.My.Resources.Resources.accept2
        Me.btnAsignar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAsignar.Location = New System.Drawing.Point(13, 12)
        Me.btnAsignar.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnAsignar.Name = "btnAsignar"
        Me.btnAsignar.Size = New System.Drawing.Size(97, 33)
        Me.btnAsignar.TabIndex = 44
        Me.btnAsignar.Text = "Asignar"
        Me.btnAsignar.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtReferencia)
        Me.GroupBox1.Controls.Add(Me.txtCodigo)
        Me.GroupBox1.Controls.Add(Me.txtDescripcion)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 55)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(520, 87)
        Me.GroupBox1.TabIndex = 45
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkTodos)
        Me.GroupBox2.Controls.Add(Me.grdReserva)
        Me.GroupBox2.Controls.Add(Me.txtReserva)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.grdPatron)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.btn1)
        Me.GroupBox2.Controls.Add(Me.btn2)
        Me.GroupBox2.Controls.Add(Me.txtCantidad)
        Me.GroupBox2.Controls.Add(Me.btn3)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.btn4)
        Me.GroupBox2.Controls.Add(Me.btnClear)
        Me.GroupBox2.Controls.Add(Me.btn5)
        Me.GroupBox2.Controls.Add(Me.txtOperario)
        Me.GroupBox2.Controls.Add(Me.btn10)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.btn9)
        Me.GroupBox2.Controls.Add(Me.btn6)
        Me.GroupBox2.Controls.Add(Me.btn8)
        Me.GroupBox2.Controls.Add(Me.btn7)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 146)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(520, 275)
        Me.GroupBox2.TabIndex = 46
        Me.GroupBox2.TabStop = False
        '
        'chkTodos
        '
        Me.chkTodos.AutoSize = True
        Me.chkTodos.Location = New System.Drawing.Point(453, 24)
        Me.chkTodos.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkTodos.Name = "chkTodos"
        Me.chkTodos.Size = New System.Drawing.Size(61, 19)
        Me.chkTodos.TabIndex = 49
        Me.chkTodos.Text = "Todos:"
        Me.chkTodos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkTodos.UseVisualStyleBackColor = True
        '
        'grdReserva
        '
        Me.grdReserva.AllowUserToAddRows = False
        Me.grdReserva.AllowUserToDeleteRows = False
        Me.grdReserva.AllowUserToResizeColumns = False
        Me.grdReserva.AllowUserToResizeRows = False
        Me.grdReserva.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.grdReserva.BackgroundColor = System.Drawing.Color.White
        Me.grdReserva.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdReserva.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.grdReserva.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdReserva.Location = New System.Drawing.Point(165, 59)
        Me.grdReserva.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grdReserva.Name = "grdReserva"
        Me.grdReserva.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdReserva.Size = New System.Drawing.Size(349, 197)
        Me.grdReserva.TabIndex = 48
        Me.grdReserva.Visible = False
        '
        'txtReserva
        '
        Me.txtReserva.Location = New System.Drawing.Point(252, 20)
        Me.txtReserva.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtReserva.Name = "txtReserva"
        Me.txtReserva.Size = New System.Drawing.Size(62, 23)
        Me.txtReserva.TabIndex = 47
        Me.txtReserva.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(149, 23)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(97, 15)
        Me.Label6.TabIndex = 46
        Me.Label6.Text = "Asignar Reserva:"
        '
        'grdPatron
        '
        Me.grdPatron.AllowUserToAddRows = False
        Me.grdPatron.AllowUserToDeleteRows = False
        Me.grdPatron.AllowUserToResizeColumns = False
        Me.grdPatron.AllowUserToResizeRows = False
        Me.grdPatron.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.grdPatron.BackgroundColor = System.Drawing.Color.White
        Me.grdPatron.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdPatron.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.grdPatron.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdPatron.Location = New System.Drawing.Point(165, 59)
        Me.grdPatron.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grdPatron.Name = "grdPatron"
        Me.grdPatron.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdPatron.Size = New System.Drawing.Size(349, 197)
        Me.grdPatron.TabIndex = 45
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(5, 206)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(52, 50)
        Me.Button1.TabIndex = 44
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnAsigReserva
        '
        Me.btnAsigReserva.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnAsigReserva.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAsigReserva.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAsigReserva.ForeColor = System.Drawing.Color.Navy
        Me.btnAsigReserva.Image = Global.WindowsApplication1.My.Resources.Resources.accept
        Me.btnAsigReserva.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAsigReserva.Location = New System.Drawing.Point(13, 12)
        Me.btnAsigReserva.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnAsigReserva.Name = "btnAsigReserva"
        Me.btnAsigReserva.Size = New System.Drawing.Size(97, 33)
        Me.btnAsigReserva.TabIndex = 48
        Me.btnAsigReserva.Text = "Asig Reserva"
        Me.btnAsigReserva.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAsigReserva.UseVisualStyleBackColor = True
        Me.btnAsigReserva.Visible = False
        '
        'btnGuardar
        '
        Me.btnGuardar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGuardar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGuardar.ForeColor = System.Drawing.Color.Navy
        Me.btnGuardar.Image = Global.WindowsApplication1.My.Resources.Resources.save
        Me.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGuardar.Location = New System.Drawing.Point(111, 12)
        Me.btnGuardar.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(97, 33)
        Me.btnGuardar.TabIndex = 49
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        Me.btnGuardar.Visible = False
        '
        'btnDeshacer
        '
        Me.btnDeshacer.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnDeshacer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDeshacer.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeshacer.ForeColor = System.Drawing.Color.Navy
        Me.btnDeshacer.Image = Global.WindowsApplication1.My.Resources.Resources.repeat
        Me.btnDeshacer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDeshacer.Location = New System.Drawing.Point(209, 12)
        Me.btnDeshacer.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnDeshacer.Name = "btnDeshacer"
        Me.btnDeshacer.Size = New System.Drawing.Size(97, 33)
        Me.btnDeshacer.TabIndex = 50
        Me.btnDeshacer.Text = "Deshacer"
        Me.btnDeshacer.UseVisualStyleBackColor = True
        Me.btnDeshacer.Visible = False
        '
        'frmAsignarOperario
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(545, 442)
        Me.Controls.Add(Me.btnDeshacer)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnAsigReserva)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.btnAsignar)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Navy
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.Name = "frmAsignarOperario"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ASIGNAR OPERARIO"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.grdReserva, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdPatron, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents txtOperario As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btn6 As System.Windows.Forms.Button
    Friend WithEvents btn7 As System.Windows.Forms.Button
    Friend WithEvents btn8 As System.Windows.Forms.Button
    Friend WithEvents btn9 As System.Windows.Forms.Button
    Friend WithEvents btn10 As System.Windows.Forms.Button
    Friend WithEvents btn5 As System.Windows.Forms.Button
    Friend WithEvents btn4 As System.Windows.Forms.Button
    Friend WithEvents btn3 As System.Windows.Forms.Button
    Friend WithEvents btn2 As System.Windows.Forms.Button
    Friend WithEvents btn1 As System.Windows.Forms.Button
    Friend WithEvents txtCantidad As System.Windows.Forms.TextBox
    Friend WithEvents txtReferencia As System.Windows.Forms.TextBox
    Friend WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Friend WithEvents txtCodigo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnAsignar As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents grdPatron As System.Windows.Forms.DataGridView
    Friend WithEvents txtReserva As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnAsigReserva As System.Windows.Forms.Button
    Friend WithEvents grdReserva As System.Windows.Forms.DataGridView
    Friend WithEvents btnGuardar As System.Windows.Forms.Button
    Friend WithEvents chkTodos As System.Windows.Forms.CheckBox
    Friend WithEvents btnDeshacer As System.Windows.Forms.Button
End Class
