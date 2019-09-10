<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_crear_importacion
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
        Me.button2 = New System.Windows.Forms.Button
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.textBox3 = New System.Windows.Forms.TextBox
        Me.textBox2 = New System.Windows.Forms.TextBox
        Me.dateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.label4 = New System.Windows.Forms.Label
        Me.label3 = New System.Windows.Forms.Label
        Me.label2 = New System.Windows.Forms.Label
        Me.button1 = New System.Windows.Forms.Button
        Me.groupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'button2
        '
        Me.button2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button2.ForeColor = System.Drawing.Color.Navy
        Me.button2.Location = New System.Drawing.Point(216, 233)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(90, 27)
        Me.button2.TabIndex = 5
        Me.button2.Text = "Aceptar"
        Me.button2.UseVisualStyleBackColor = True
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.textBox3)
        Me.groupBox1.Controls.Add(Me.textBox2)
        Me.groupBox1.Controls.Add(Me.dateTimePicker1)
        Me.groupBox1.Controls.Add(Me.label4)
        Me.groupBox1.Controls.Add(Me.label3)
        Me.groupBox1.Controls.Add(Me.label2)
        Me.groupBox1.Location = New System.Drawing.Point(6, 2)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(396, 225)
        Me.groupBox1.TabIndex = 4
        Me.groupBox1.TabStop = False
        '
        'textBox3
        '
        Me.textBox3.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox3.Location = New System.Drawing.Point(13, 111)
        Me.textBox3.Multiline = True
        Me.textBox3.Name = "textBox3"
        Me.textBox3.Size = New System.Drawing.Size(375, 101)
        Me.textBox3.TabIndex = 17
        '
        'textBox2
        '
        Me.textBox2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox2.Location = New System.Drawing.Point(71, 46)
        Me.textBox2.Name = "textBox2"
        Me.textBox2.Size = New System.Drawing.Size(250, 26)
        Me.textBox2.TabIndex = 16
        '
        'dateTimePicker1
        '
        Me.dateTimePicker1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dateTimePicker1.Location = New System.Drawing.Point(71, 15)
        Me.dateTimePicker1.Name = "dateTimePicker1"
        Me.dateTimePicker1.Size = New System.Drawing.Size(143, 26)
        Me.dateTimePicker1.TabIndex = 15
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label4.ForeColor = System.Drawing.Color.Navy
        Me.label4.Location = New System.Drawing.Point(10, 54)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(45, 18)
        Me.label4.TabIndex = 13
        Me.label4.Text = "Lugar:"
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label3.ForeColor = System.Drawing.Color.Navy
        Me.label3.Location = New System.Drawing.Point(10, 90)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(88, 18)
        Me.label3.TabIndex = 12
        Me.label3.Text = "Descripicion:"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label2.ForeColor = System.Drawing.Color.Navy
        Me.label2.Location = New System.Drawing.Point(10, 15)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(48, 18)
        Me.label2.TabIndex = 11
        Me.label2.Text = "Fecha:"
        '
        'button1
        '
        Me.button1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button1.ForeColor = System.Drawing.Color.Navy
        Me.button1.Location = New System.Drawing.Point(312, 233)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(90, 27)
        Me.button1.TabIndex = 3
        Me.button1.Text = "Cerrar"
        Me.button1.UseVisualStyleBackColor = True
        '
        'frm_crear_importacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(408, 262)
        Me.Controls.Add(Me.button2)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_crear_importacion"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Nueva Importación"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents button2 As System.Windows.Forms.Button
    Private WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents textBox3 As System.Windows.Forms.TextBox
    Private WithEvents textBox2 As System.Windows.Forms.TextBox
    Private WithEvents dateTimePicker1 As System.Windows.Forms.DateTimePicker
    Private WithEvents label4 As System.Windows.Forms.Label
    Private WithEvents label3 As System.Windows.Forms.Label
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents button1 As System.Windows.Forms.Button
End Class
