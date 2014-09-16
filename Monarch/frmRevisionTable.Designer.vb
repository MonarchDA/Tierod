<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRevisionTable
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LabelGradient5 = New LabelGradient.LabelGradient
        Me.LabelGradient2 = New LabelGradient.LabelGradient
        Me.LabelGradient3 = New LabelGradient.LabelGradient
        Me.LabelGradient4 = New LabelGradient.LabelGradient
        Me.dgvRevisionTable = New System.Windows.Forms.DataGridView
        Me.btnUpdateRevision = New System.Windows.Forms.Button
        CType(Me.dgvRevisionTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LabelGradient5
        '
        Me.LabelGradient5.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient5.Dock = System.Windows.Forms.DockStyle.Top
        Me.LabelGradient5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient5.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient5.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient5.Location = New System.Drawing.Point(0, 0)
        Me.LabelGradient5.Name = "LabelGradient5"
        Me.LabelGradient5.Size = New System.Drawing.Size(874, 11)
        Me.LabelGradient5.TabIndex = 112
        Me.LabelGradient5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient2
        '
        Me.LabelGradient2.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.LabelGradient2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient2.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient2.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient2.Location = New System.Drawing.Point(0, 317)
        Me.LabelGradient2.Name = "LabelGradient2"
        Me.LabelGradient2.Size = New System.Drawing.Size(874, 11)
        Me.LabelGradient2.TabIndex = 113
        Me.LabelGradient2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient3
        '
        Me.LabelGradient3.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient3.Dock = System.Windows.Forms.DockStyle.Left
        Me.LabelGradient3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient3.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient3.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient3.Location = New System.Drawing.Point(0, 11)
        Me.LabelGradient3.Name = "LabelGradient3"
        Me.LabelGradient3.Size = New System.Drawing.Size(10, 306)
        Me.LabelGradient3.TabIndex = 114
        Me.LabelGradient3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient4
        '
        Me.LabelGradient4.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient4.Dock = System.Windows.Forms.DockStyle.Right
        Me.LabelGradient4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient4.GradientColorOne = System.Drawing.Color.Honeydew
        Me.LabelGradient4.GradientColorTwo = System.Drawing.Color.Olive
        Me.LabelGradient4.Location = New System.Drawing.Point(864, 11)
        Me.LabelGradient4.Name = "LabelGradient4"
        Me.LabelGradient4.Size = New System.Drawing.Size(10, 306)
        Me.LabelGradient4.TabIndex = 115
        Me.LabelGradient4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dgvRevisionTable
        '
        Me.dgvRevisionTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRevisionTable.Location = New System.Drawing.Point(16, 19)
        Me.dgvRevisionTable.Name = "dgvRevisionTable"
        Me.dgvRevisionTable.Size = New System.Drawing.Size(842, 266)
        Me.dgvRevisionTable.TabIndex = 116
        '
        'btnUpdateRevision
        '
        Me.btnUpdateRevision.Location = New System.Drawing.Point(774, 291)
        Me.btnUpdateRevision.Name = "btnUpdateRevision"
        Me.btnUpdateRevision.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdateRevision.TabIndex = 117
        Me.btnUpdateRevision.Text = "Update"
        Me.btnUpdateRevision.UseVisualStyleBackColor = True
        '
        'frmRevisionTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(874, 328)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnUpdateRevision)
        Me.Controls.Add(Me.dgvRevisionTable)
        Me.Controls.Add(Me.LabelGradient4)
        Me.Controls.Add(Me.LabelGradient3)
        Me.Controls.Add(Me.LabelGradient2)
        Me.Controls.Add(Me.LabelGradient5)
        Me.Name = "frmRevisionTable"
        Me.Text = "Revision Status "
        CType(Me.dgvRevisionTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelGradient5 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient2 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient3 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient4 As LabelGradient.LabelGradient
    Friend WithEvents dgvRevisionTable As System.Windows.Forms.DataGridView
    Friend WithEvents btnUpdateRevision As System.Windows.Forms.Button
End Class
