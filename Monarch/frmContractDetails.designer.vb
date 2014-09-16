<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmContractDetails
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
        Me.grpContractDetails = New System.Windows.Forms.GroupBox
        Me.btnChangePartNumber = New System.Windows.Forms.Button
        Me.chkManageCustomers = New System.Windows.Forms.CheckBox
        Me.txtlPartCode = New IFLCustomUILayer.IFLTextBox
        Me.cmbAssemblyType = New IFLCustomUILayer.IFLComboBox
        Me.lblPartCode = New System.Windows.Forms.Label
        Me.lblAssemblyType = New System.Windows.Forms.Label
        Me.lblCustomerName = New System.Windows.Forms.Label
        Me.lblGradientPrimaryInformation = New LabelGradient.LabelGradient
        Me.cmbCustomerName = New IFLCustomUILayer.IFLComboBox
        Me.pnlManageCustomerDetails = New System.Windows.Forms.Panel
        Me.txtCustomerName_Add = New System.Windows.Forms.TextBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.cmbCustomerName_Delete = New IFLCustomUILayer.IFLComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.lblGradientContractDetails = New LabelGradient.LabelGradient
        Me.LabelGradient5 = New LabelGradient.LabelGradient
        Me.LabelGradient4 = New LabelGradient.LabelGradient
        Me.LabelGradient3 = New LabelGradient.LabelGradient
        Me.LabelGradient2 = New LabelGradient.LabelGradient
        Me.LabelGradient1 = New LabelGradient.LabelGradient
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LabelGradient6 = New LabelGradient.LabelGradient
        Me.lblBackBrowse = New LabelGradient.LabelGradient
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lblBrowse = New System.Windows.Forms.Label
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.LabelGradient8 = New LabelGradient.LabelGradient
        Me.lVLogInformation = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.lblLogInformation = New LabelGradient.LabelGradient
        Me.grpContractDetails.SuspendLayout()
        Me.pnlManageCustomerDetails.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpContractDetails
        '
        Me.grpContractDetails.BackColor = System.Drawing.Color.Ivory
        Me.grpContractDetails.Controls.Add(Me.btnChangePartNumber)
        Me.grpContractDetails.Controls.Add(Me.chkManageCustomers)
        Me.grpContractDetails.Controls.Add(Me.txtlPartCode)
        Me.grpContractDetails.Controls.Add(Me.cmbAssemblyType)
        Me.grpContractDetails.Controls.Add(Me.lblPartCode)
        Me.grpContractDetails.Controls.Add(Me.lblAssemblyType)
        Me.grpContractDetails.Controls.Add(Me.lblCustomerName)
        Me.grpContractDetails.Controls.Add(Me.lblGradientPrimaryInformation)
        Me.grpContractDetails.Controls.Add(Me.cmbCustomerName)
        Me.grpContractDetails.Location = New System.Drawing.Point(19, 29)
        Me.grpContractDetails.Name = "grpContractDetails"
        Me.grpContractDetails.Size = New System.Drawing.Size(469, 127)
        Me.grpContractDetails.TabIndex = 0
        Me.grpContractDetails.TabStop = False
        '
        'btnChangePartNumber
        '
        Me.btnChangePartNumber.Location = New System.Drawing.Point(349, 89)
        Me.btnChangePartNumber.Name = "btnChangePartNumber"
        Me.btnChangePartNumber.Size = New System.Drawing.Size(114, 23)
        Me.btnChangePartNumber.TabIndex = 25
        Me.btnChangePartNumber.Text = "Change Part Number"
        Me.btnChangePartNumber.UseVisualStyleBackColor = True
        '
        'chkManageCustomers
        '
        Me.chkManageCustomers.AutoSize = True
        Me.chkManageCustomers.Location = New System.Drawing.Point(349, 33)
        Me.chkManageCustomers.Name = "chkManageCustomers"
        Me.chkManageCustomers.Size = New System.Drawing.Size(112, 17)
        Me.chkManageCustomers.TabIndex = 23
        Me.chkManageCustomers.Text = "Manage Customer"
        Me.chkManageCustomers.UseVisualStyleBackColor = True
        '
        'txtlPartCode
        '
        Me.txtlPartCode.AcceptEnterKeyAsTab = True
        Me.txtlPartCode.ApplyIFLColor = True
        Me.txtlPartCode.AssociateLabel = Nothing
        Me.txtlPartCode.IFLDataTag = Nothing
        Me.txtlPartCode.InvalidInputCharacters = ""
        Me.txtlPartCode.Location = New System.Drawing.Point(101, 91)
        Me.txtlPartCode.MaxLength = 32
        Me.txtlPartCode.Name = "txtlPartCode"
        Me.txtlPartCode.Size = New System.Drawing.Size(238, 20)
        Me.txtlPartCode.StatusMessage = ""
        Me.txtlPartCode.StatusObject = Nothing
        Me.txtlPartCode.TabIndex = 3
        '
        'cmbAssemblyType
        '
        Me.cmbAssemblyType.AcceptEnterKeyAsTab = True
        Me.cmbAssemblyType.AssociateLabel = Nothing
        Me.cmbAssemblyType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAssemblyType.FormattingEnabled = True
        Me.cmbAssemblyType.IFLDataTag = Nothing
        Me.cmbAssemblyType.ItemHeight = 13
        Me.cmbAssemblyType.Items.AddRange(New Object() {"Tie Rod Cylinder Assembly"})
        Me.cmbAssemblyType.Location = New System.Drawing.Point(101, 61)
        Me.cmbAssemblyType.Name = "cmbAssemblyType"
        Me.cmbAssemblyType.Size = New System.Drawing.Size(238, 21)
        Me.cmbAssemblyType.StatusMessage = Nothing
        Me.cmbAssemblyType.StatusObject = Nothing
        Me.cmbAssemblyType.TabIndex = 2
        '
        'lblPartCode
        '
        Me.lblPartCode.AutoSize = True
        Me.lblPartCode.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!)
        Me.lblPartCode.Location = New System.Drawing.Point(26, 96)
        Me.lblPartCode.Name = "lblPartCode"
        Me.lblPartCode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.lblPartCode.Size = New System.Drawing.Size(75, 15)
        Me.lblPartCode.TabIndex = 4
        Me.lblPartCode.Text = "Part Number"
        '
        'lblAssemblyType
        '
        Me.lblAssemblyType.AutoSize = True
        Me.lblAssemblyType.Location = New System.Drawing.Point(67, 61)
        Me.lblAssemblyType.Name = "lblAssemblyType"
        Me.lblAssemblyType.Size = New System.Drawing.Size(31, 13)
        Me.lblAssemblyType.TabIndex = 3
        Me.lblAssemblyType.Text = "Type"
        '
        'lblCustomerName
        '
        Me.lblCustomerName.AutoSize = True
        Me.lblCustomerName.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCustomerName.Location = New System.Drawing.Point(3, 33)
        Me.lblCustomerName.Name = "lblCustomerName"
        Me.lblCustomerName.Size = New System.Drawing.Size(95, 15)
        Me.lblCustomerName.TabIndex = 0
        Me.lblCustomerName.Text = "Customer Name"
        '
        'lblGradientPrimaryInformation
        '
        Me.lblGradientPrimaryInformation.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblGradientPrimaryInformation.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.lblGradientPrimaryInformation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGradientPrimaryInformation.ForeColor = System.Drawing.Color.White
        Me.lblGradientPrimaryInformation.GradientColorOne = System.Drawing.Color.Olive
        Me.lblGradientPrimaryInformation.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.lblGradientPrimaryInformation.Location = New System.Drawing.Point(-3, 0)
        Me.lblGradientPrimaryInformation.Name = "lblGradientPrimaryInformation"
        Me.lblGradientPrimaryInformation.Size = New System.Drawing.Size(472, 21)
        Me.lblGradientPrimaryInformation.TabIndex = 20
        Me.lblGradientPrimaryInformation.Text = "Preliminary Information"
        Me.lblGradientPrimaryInformation.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbCustomerName
        '
        Me.cmbCustomerName.AcceptEnterKeyAsTab = True
        Me.cmbCustomerName.AssociateLabel = Nothing
        Me.cmbCustomerName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCustomerName.FormattingEnabled = True
        Me.cmbCustomerName.IFLDataTag = Nothing
        Me.cmbCustomerName.ItemHeight = 13
        Me.cmbCustomerName.Location = New System.Drawing.Point(101, 31)
        Me.cmbCustomerName.Name = "cmbCustomerName"
        Me.cmbCustomerName.Size = New System.Drawing.Size(238, 21)
        Me.cmbCustomerName.StatusMessage = Nothing
        Me.cmbCustomerName.StatusObject = Nothing
        Me.cmbCustomerName.TabIndex = 2
        '
        'pnlManageCustomerDetails
        '
        Me.pnlManageCustomerDetails.Controls.Add(Me.txtCustomerName_Add)
        Me.pnlManageCustomerDetails.Controls.Add(Me.btnDelete)
        Me.pnlManageCustomerDetails.Controls.Add(Me.btnAdd)
        Me.pnlManageCustomerDetails.Controls.Add(Me.cmbCustomerName_Delete)
        Me.pnlManageCustomerDetails.Controls.Add(Me.Label2)
        Me.pnlManageCustomerDetails.Controls.Add(Me.Label1)
        Me.pnlManageCustomerDetails.Controls.Add(Me.ListBox1)
        Me.pnlManageCustomerDetails.Location = New System.Drawing.Point(9, 21)
        Me.pnlManageCustomerDetails.Name = "pnlManageCustomerDetails"
        Me.pnlManageCustomerDetails.Size = New System.Drawing.Size(450, 101)
        Me.pnlManageCustomerDetails.TabIndex = 22
        Me.pnlManageCustomerDetails.Visible = False
        '
        'txtCustomerName_Add
        '
        Me.txtCustomerName_Add.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerName_Add.Location = New System.Drawing.Point(105, 18)
        Me.txtCustomerName_Add.Name = "txtCustomerName_Add"
        Me.txtCustomerName_Add.Size = New System.Drawing.Size(228, 20)
        Me.txtCustomerName_Add.TabIndex = 23
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(356, 57)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 22
        Me.btnDelete.Text = "Delete Customer"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(356, 16)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 21
        Me.btnAdd.Text = "Add Customer"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'cmbCustomerName_Delete
        '
        Me.cmbCustomerName_Delete.AcceptEnterKeyAsTab = True
        Me.cmbCustomerName_Delete.AssociateLabel = Nothing
        Me.cmbCustomerName_Delete.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCustomerName_Delete.FormattingEnabled = True
        Me.cmbCustomerName_Delete.IFLDataTag = Nothing
        Me.cmbCustomerName_Delete.ItemHeight = 13
        Me.cmbCustomerName_Delete.Location = New System.Drawing.Point(104, 59)
        Me.cmbCustomerName_Delete.Name = "cmbCustomerName_Delete"
        Me.cmbCustomerName_Delete.Size = New System.Drawing.Size(229, 21)
        Me.cmbCustomerName_Delete.StatusMessage = Nothing
        Me.cmbCustomerName_Delete.StatusObject = Nothing
        Me.cmbCustomerName_Delete.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 15)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Customer Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer Name"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(106, 39)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(228, 69)
        Me.ListBox1.TabIndex = 112
        Me.ListBox1.Visible = False
        '
        'lblGradientContractDetails
        '
        Me.lblGradientContractDetails.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.lblGradientContractDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGradientContractDetails.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGradientContractDetails.GradientColorOne = System.Drawing.Color.DarkGoldenrod
        Me.lblGradientContractDetails.GradientColorTwo = System.Drawing.Color.DarkGoldenrod
        Me.lblGradientContractDetails.Location = New System.Drawing.Point(16, 20)
        Me.lblGradientContractDetails.Name = "lblGradientContractDetails"
        Me.lblGradientContractDetails.Size = New System.Drawing.Size(478, 144)
        Me.lblGradientContractDetails.TabIndex = 40
        Me.lblGradientContractDetails.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient5
        '
        Me.LabelGradient5.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient5.Dock = System.Windows.Forms.DockStyle.Top
        Me.LabelGradient5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient5.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient5.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient5.Location = New System.Drawing.Point(10, 0)
        Me.LabelGradient5.Name = "LabelGradient5"
        Me.LabelGradient5.Size = New System.Drawing.Size(1016, 11)
        Me.LabelGradient5.TabIndex = 111
        Me.LabelGradient5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient4
        '
        Me.LabelGradient4.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient4.Dock = System.Windows.Forms.DockStyle.Right
        Me.LabelGradient4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient4.GradientColorOne = System.Drawing.Color.Honeydew
        Me.LabelGradient4.GradientColorTwo = System.Drawing.Color.Olive
        Me.LabelGradient4.Location = New System.Drawing.Point(1026, 0)
        Me.LabelGradient4.Name = "LabelGradient4"
        Me.LabelGradient4.Size = New System.Drawing.Size(10, 694)
        Me.LabelGradient4.TabIndex = 110
        Me.LabelGradient4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient3
        '
        Me.LabelGradient3.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient3.Dock = System.Windows.Forms.DockStyle.Left
        Me.LabelGradient3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient3.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient3.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient3.Location = New System.Drawing.Point(0, 0)
        Me.LabelGradient3.Name = "LabelGradient3"
        Me.LabelGradient3.Size = New System.Drawing.Size(10, 694)
        Me.LabelGradient3.TabIndex = 109
        Me.LabelGradient3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient2
        '
        Me.LabelGradient2.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.LabelGradient2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient2.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient2.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient2.Location = New System.Drawing.Point(0, 694)
        Me.LabelGradient2.Name = "LabelGradient2"
        Me.LabelGradient2.Size = New System.Drawing.Size(1036, 11)
        Me.LabelGradient2.TabIndex = 108
        Me.LabelGradient2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LabelGradient1
        '
        Me.LabelGradient1.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient1.GradientColorOne = System.Drawing.Color.DarkGoldenrod
        Me.LabelGradient1.GradientColorTwo = System.Drawing.Color.DarkGoldenrod
        Me.LabelGradient1.Location = New System.Drawing.Point(510, 30)
        Me.LabelGradient1.Name = "LabelGradient1"
        Me.LabelGradient1.Size = New System.Drawing.Size(478, 127)
        Me.LabelGradient1.TabIndex = 112
        Me.LabelGradient1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.LabelGradient1.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Ivory
        Me.GroupBox1.Controls.Add(Me.LabelGradient6)
        Me.GroupBox1.Controls.Add(Me.pnlManageCustomerDetails)
        Me.GroupBox1.Location = New System.Drawing.Point(513, 37)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(469, 114)
        Me.GroupBox1.TabIndex = 113
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Visible = False
        '
        'LabelGradient6
        '
        Me.LabelGradient6.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelGradient6.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient6.ForeColor = System.Drawing.Color.White
        Me.LabelGradient6.GradientColorOne = System.Drawing.Color.Olive
        Me.LabelGradient6.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient6.Location = New System.Drawing.Point(-3, -5)
        Me.LabelGradient6.Name = "LabelGradient6"
        Me.LabelGradient6.Size = New System.Drawing.Size(472, 26)
        Me.LabelGradient6.TabIndex = 20
        Me.LabelGradient6.Text = "Customer Information"
        Me.LabelGradient6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblBackBrowse
        '
        Me.lblBackBrowse.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.lblBackBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBackBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBackBrowse.GradientColorOne = System.Drawing.Color.DarkGoldenrod
        Me.lblBackBrowse.GradientColorTwo = System.Drawing.Color.DarkGoldenrod
        Me.lblBackBrowse.Location = New System.Drawing.Point(16, 175)
        Me.lblBackBrowse.Name = "lblBackBrowse"
        Me.lblBackBrowse.Size = New System.Drawing.Size(478, 60)
        Me.lblBackBrowse.TabIndex = 114
        Me.lblBackBrowse.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Ivory
        Me.GroupBox2.Controls.Add(Me.lblBrowse)
        Me.GroupBox2.Controls.Add(Me.btnBrowse)
        Me.GroupBox2.Location = New System.Drawing.Point(24, 184)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(463, 42)
        Me.GroupBox2.TabIndex = 116
        Me.GroupBox2.TabStop = False
        '
        'lblBrowse
        '
        Me.lblBrowse.AutoSize = True
        Me.lblBrowse.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!)
        Me.lblBrowse.Location = New System.Drawing.Point(54, 16)
        Me.lblBrowse.Name = "lblBrowse"
        Me.lblBrowse.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.lblBrowse.Size = New System.Drawing.Size(97, 15)
        Me.lblBrowse.TabIndex = 122
        Me.lblBrowse.Text = "Browse Excel File"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(176, 9)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(95, 28)
        Me.btnBrowse.TabIndex = 121
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'LabelGradient8
        '
        Me.LabelGradient8.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient8.GradientColorOne = System.Drawing.Color.DarkGoldenrod
        Me.LabelGradient8.GradientColorTwo = System.Drawing.Color.DarkGoldenrod
        Me.LabelGradient8.Location = New System.Drawing.Point(16, 251)
        Me.LabelGradient8.Name = "LabelGradient8"
        Me.LabelGradient8.Size = New System.Drawing.Size(478, 316)
        Me.LabelGradient8.TabIndex = 117
        Me.LabelGradient8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lVLogInformation
        '
        Me.lVLogInformation.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.lVLogInformation.GridLines = True
        Me.lVLogInformation.Location = New System.Drawing.Point(0, 26)
        Me.lVLogInformation.Name = "lVLogInformation"
        Me.lVLogInformation.Size = New System.Drawing.Size(478, 290)
        Me.lVLogInformation.TabIndex = 118
        Me.lVLogInformation.UseCompatibleStateImageBehavior = False
        Me.lVLogInformation.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Error Details"
        Me.ColumnHeader1.Width = 973
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Ivory
        Me.GroupBox3.Controls.Add(Me.lblLogInformation)
        Me.GroupBox3.Controls.Add(Me.lVLogInformation)
        Me.GroupBox3.Location = New System.Drawing.Point(16, 251)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(478, 316)
        Me.GroupBox3.TabIndex = 120
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "GroupBox3"
        '
        'lblLogInformation
        '
        Me.lblLogInformation.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblLogInformation.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.lblLogInformation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLogInformation.ForeColor = System.Drawing.Color.White
        Me.lblLogInformation.GradientColorOne = System.Drawing.Color.Olive
        Me.lblLogInformation.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.lblLogInformation.Location = New System.Drawing.Point(-3, 0)
        Me.lblLogInformation.Name = "lblLogInformation"
        Me.lblLogInformation.Size = New System.Drawing.Size(481, 26)
        Me.lblLogInformation.TabIndex = 119
        Me.lblLogInformation.Text = "Log Information"
        Me.lblLogInformation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmContractDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(1036, 705)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.LabelGradient8)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.lblBackBrowse)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.LabelGradient1)
        Me.Controls.Add(Me.LabelGradient5)
        Me.Controls.Add(Me.LabelGradient4)
        Me.Controls.Add(Me.LabelGradient3)
        Me.Controls.Add(Me.LabelGradient2)
        Me.Controls.Add(Me.grpContractDetails)
        Me.Controls.Add(Me.lblGradientContractDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmContractDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.grpContractDetails.ResumeLayout(False)
        Me.grpContractDetails.PerformLayout()
        Me.pnlManageCustomerDetails.ResumeLayout(False)
        Me.pnlManageCustomerDetails.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    'Friend WithEvents LabelGradient2 As LabelGradient.LabelGradient
    'Friend WithEvents lblMachineNameText As System.Windows.Forms.Label
    'Friend WithEvents LabelGradient1 As LabelGradient.LabelGradient
    'Friend WithEvents lblServer As System.Windows.Forms.Label
    'Friend WithEvents Label3 As System.Windows.Forms.Label
    'Friend WithEvents lblDB As System.Windows.Forms.Label
    'Friend WithEvents labledb As System.Windows.Forms.Label
    'Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    'Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    'Friend WithEvents LabelGradient3 As LabelGradient.LabelGradient
    'Friend WithEvents lblMachineName As System.Windows.Forms.Label
    'Friend WithEvents lblUserNameText As System.Windows.Forms.Label
    'Friend WithEvents lblUserName As System.Windows.Forms.Label
    'Friend WithEvents gbUserLogin As System.Windows.Forms.GroupBox
    'Friend WithEvents gbLoginDetails As System.Windows.Forms.GroupBox
    'Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    'Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents grpContractDetails As System.Windows.Forms.GroupBox
    Friend WithEvents lblGradientPrimaryInformation As LabelGradient.LabelGradient
    Friend WithEvents cmbAssemblyType As IFLCustomUILayer.IFLComboBox
    Friend WithEvents lblPartCode As System.Windows.Forms.Label
    Friend WithEvents lblAssemblyType As System.Windows.Forms.Label
    Friend WithEvents lblCustomerName As System.Windows.Forms.Label
    Friend WithEvents txtlPartCode As IFLCustomUILayer.IFLTextBox

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Friend WithEvents lblGradientContractDetails As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient5 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient4 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient3 As LabelGradient.LabelGradient
    Friend WithEvents LabelGradient2 As LabelGradient.LabelGradient
    Friend WithEvents chkManageCustomers As System.Windows.Forms.CheckBox
    Friend WithEvents pnlManageCustomerDetails As System.Windows.Forms.Panel
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents txtCustomerName_Add As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbCustomerName_Delete As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents LabelGradient1 As LabelGradient.LabelGradient
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelGradient6 As LabelGradient.LabelGradient
    Friend WithEvents cmbCustomerName As IFLCustomUILayer.IFLComboBox
    Friend WithEvents btnChangePartNumber As System.Windows.Forms.Button
    Friend WithEvents lblBackBrowse As LabelGradient.LabelGradient
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents LabelGradient8 As LabelGradient.LabelGradient
    Friend WithEvents lVLogInformation As System.Windows.Forms.ListView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblBrowse As System.Windows.Forms.Label
    Friend WithEvents lblLogInformation As LabelGradient.LabelGradient
End Class
