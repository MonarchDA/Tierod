<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMonarch
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMonarch))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LVCustomer = New IFLCustomUILayer.IFLListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.txtListViewSearchObject = New System.Windows.Forms.TextBox
        Me.lvwContractDetails = New IFLCustomUILayer.IFLListView
        Me.LabelGradient4 = New LabelGradient.LabelGradient
        Me.LabelGradient5 = New LabelGradient.LabelGradient
        Me.LabelGradient1 = New LabelGradient.LabelGradient
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.LabelGradient2 = New LabelGradient.LabelGradient
        Me.LabelGradient3 = New LabelGradient.LabelGradient
        Me.LabelGradient6 = New LabelGradient.LabelGradient
        Me.LabelGradient7 = New LabelGradient.LabelGradient
        Me.btnRelease = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Black
        Me.GroupBox1.Controls.Add(Me.LVCustomer)
        Me.GroupBox1.Controls.Add(Me.txtListViewSearchObject)
        Me.GroupBox1.Controls.Add(Me.lvwContractDetails)
        Me.GroupBox1.Controls.Add(Me.LabelGradient4)
        Me.GroupBox1.Controls.Add(Me.LabelGradient5)
        Me.GroupBox1.Controls.Add(Me.LabelGradient1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Ivory
        Me.GroupBox1.Location = New System.Drawing.Point(16, 14)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(878, 650)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Project Action"
        '
        'LVCustomer
        '
        Me.LVCustomer.BackColor = System.Drawing.Color.Honeydew
        Me.LVCustomer.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.LVCustomer.DisplayHeaders = CType(resources.GetObject("LVCustomer.DisplayHeaders"), System.Collections.ArrayList)
        Me.LVCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LVCustomer.FullRowSelect = True
        Me.LVCustomer.GridLines = True
        Me.LVCustomer.IFLDataTag = Nothing
        Me.LVCustomer.IsCheckBoxEnabled = False
        Me.LVCustomer.IsFilterOptionEnabled = False
        Me.LVCustomer.IsTypeSearchEnable = True
        Me.LVCustomer.Location = New System.Drawing.Point(9, 43)
        Me.LVCustomer.MultiSelect = False
        Me.LVCustomer.Name = "LVCustomer"
        Me.LVCustomer.SearchObject = Nothing
        Me.LVCustomer.Size = New System.Drawing.Size(863, 275)
        Me.LVCustomer.SourceTable = Nothing
        Me.LVCustomer.TabIndex = 15
        Me.LVCustomer.UseCompatibleStateImageBehavior = False
        Me.LVCustomer.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Customer Name"
        Me.ColumnHeader1.Width = 812
        '
        'txtListViewSearchObject
        '
        Me.txtListViewSearchObject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtListViewSearchObject.Location = New System.Drawing.Point(90, 681)
        Me.txtListViewSearchObject.Name = "txtListViewSearchObject"
        Me.txtListViewSearchObject.Size = New System.Drawing.Size(100, 20)
        Me.txtListViewSearchObject.TabIndex = 14
        Me.txtListViewSearchObject.Visible = False
        '
        'lvwContractDetails
        '
        Me.lvwContractDetails.BackColor = System.Drawing.Color.Honeydew
        Me.lvwContractDetails.DisplayHeaders = CType(resources.GetObject("lvwContractDetails.DisplayHeaders"), System.Collections.ArrayList)
        Me.lvwContractDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvwContractDetails.FullRowSelect = True
        Me.lvwContractDetails.GridLines = True
        Me.lvwContractDetails.IFLDataTag = Nothing
        Me.lvwContractDetails.IsCheckBoxEnabled = False
        Me.lvwContractDetails.IsFilterOptionEnabled = False
        Me.lvwContractDetails.IsTypeSearchEnable = True
        Me.lvwContractDetails.Location = New System.Drawing.Point(9, 348)
        Me.lvwContractDetails.MultiSelect = False
        Me.lvwContractDetails.Name = "lvwContractDetails"
        Me.lvwContractDetails.SearchObject = Nothing
        Me.lvwContractDetails.Size = New System.Drawing.Size(863, 290)
        Me.lvwContractDetails.SourceTable = Nothing
        Me.lvwContractDetails.TabIndex = 13
        Me.lvwContractDetails.UseCompatibleStateImageBehavior = False
        Me.lvwContractDetails.View = System.Windows.Forms.View.Details
        '
        'LabelGradient4
        '
        Me.LabelGradient4.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.LabelGradient4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient4.ForeColor = System.Drawing.Color.White
        Me.LabelGradient4.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient4.GradientColorTwo = System.Drawing.Color.Azure
        Me.LabelGradient4.Location = New System.Drawing.Point(9, 323)
        Me.LabelGradient4.Name = "LabelGradient4"
        Me.LabelGradient4.Size = New System.Drawing.Size(863, 26)
        Me.LabelGradient4.TabIndex = 16
        Me.LabelGradient4.Text = "Assembly Details"
        Me.LabelGradient4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelGradient5
        '
        Me.LabelGradient5.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.LabelGradient5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient5.ForeColor = System.Drawing.Color.White
        Me.LabelGradient5.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient5.GradientColorTwo = System.Drawing.Color.Azure
        Me.LabelGradient5.Location = New System.Drawing.Point(9, 19)
        Me.LabelGradient5.Name = "LabelGradient5"
        Me.LabelGradient5.Size = New System.Drawing.Size(863, 24)
        Me.LabelGradient5.TabIndex = 17
        Me.LabelGradient5.Text = "Customer Details"
        Me.LabelGradient5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelGradient1
        '
        Me.LabelGradient1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelGradient1.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient1.GradientColorOne = System.Drawing.Color.DarkGoldenrod
        Me.LabelGradient1.GradientColorTwo = System.Drawing.Color.DarkGoldenrod
        Me.LabelGradient1.Location = New System.Drawing.Point(6, 14)
        Me.LabelGradient1.Name = "LabelGradient1"
        Me.LabelGradient1.Size = New System.Drawing.Size(868, 631)
        Me.LabelGradient1.TabIndex = 41
        Me.LabelGradient1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'LabelGradient2
        '
        Me.LabelGradient2.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient2.Dock = System.Windows.Forms.DockStyle.Top
        Me.LabelGradient2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient2.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient2.GradientColorTwo = System.Drawing.Color.Olive
        Me.LabelGradient2.Location = New System.Drawing.Point(10, 0)
        Me.LabelGradient2.Name = "LabelGradient2"
        Me.LabelGradient2.Size = New System.Drawing.Size(1178, 11)
        Me.LabelGradient2.TabIndex = 115
        Me.LabelGradient2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient3
        '
        Me.LabelGradient3.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient3.Dock = System.Windows.Forms.DockStyle.Right
        Me.LabelGradient3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient3.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient3.GradientColorTwo = System.Drawing.Color.Olive
        Me.LabelGradient3.Location = New System.Drawing.Point(1188, 0)
        Me.LabelGradient3.Name = "LabelGradient3"
        Me.LabelGradient3.Size = New System.Drawing.Size(10, 835)
        Me.LabelGradient3.TabIndex = 114
        Me.LabelGradient3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient6
        '
        Me.LabelGradient6.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient6.Dock = System.Windows.Forms.DockStyle.Left
        Me.LabelGradient6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient6.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient6.GradientColorTwo = System.Drawing.Color.Olive
        Me.LabelGradient6.Location = New System.Drawing.Point(0, 0)
        Me.LabelGradient6.Name = "LabelGradient6"
        Me.LabelGradient6.Size = New System.Drawing.Size(10, 835)
        Me.LabelGradient6.TabIndex = 113
        Me.LabelGradient6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient7
        '
        Me.LabelGradient7.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient7.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.LabelGradient7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient7.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient7.GradientColorTwo = System.Drawing.Color.Olive
        Me.LabelGradient7.Location = New System.Drawing.Point(0, 835)
        Me.LabelGradient7.Name = "LabelGradient7"
        Me.LabelGradient7.Size = New System.Drawing.Size(1198, 11)
        Me.LabelGradient7.TabIndex = 112
        Me.LabelGradient7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnRelease
        '
        Me.btnRelease.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.btnRelease.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRelease.Location = New System.Drawing.Point(25, 730)
        Me.btnRelease.Name = "btnRelease"
        Me.btnRelease.Size = New System.Drawing.Size(75, 23)
        Me.btnRelease.TabIndex = 116
        Me.btnRelease.Text = "Release"
        Me.btnRelease.UseVisualStyleBackColor = False
        Me.btnRelease.Visible = False
        '
        'frmMonarch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(1198, 846)
        Me.Controls.Add(Me.btnRelease)
        Me.Controls.Add(Me.LabelGradient2)
        Me.Controls.Add(Me.LabelGradient3)
        Me.Controls.Add(Me.LabelGradient6)
        Me.Controls.Add(Me.LabelGradient7)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Location = New System.Drawing.Point(100, 100)
        Me.Name = "frmMonarch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    'Friend WithEvents Label1 As LabelGradient.LabelGradient
    'Friend WithEvents Label2 As LabelGradient.LabelGradient
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    'Friend WithEvents lvwProjectInformation As IFLCustomUILayer.IFLListView
    Friend WithEvents lvwContractDetails As IFLCustomUILayer.IFLListView
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    'Friend WithEvents Label1 As LabelGradient.LabelGradient

    'Friend WithEvents Label2 As LabelGradient.LabelGradient
    Friend WithEvents txtListViewSearchObject As System.Windows.Forms.TextBox
    Friend WithEvents LVCustomer As IFLCustomUILayer.IFLListView
    Friend WithEvents LabelGradient4 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient5 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient1 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient2 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient3 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient6 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient7 As LabelGradient.LabelGradient
    Friend WithEvents btnRelease As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
End Class
