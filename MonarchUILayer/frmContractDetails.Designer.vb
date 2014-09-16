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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmContractDetails))
        Me.LabelGradient2 = New LabelGradient.LabelGradient
        Me.lblMachineNameText = New System.Windows.Forms.Label
        Me.LabelGradient1 = New LabelGradient.LabelGradient
        Me.lblServer = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblDB = New System.Windows.Forms.Label
        Me.labledb = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.LabelGradient3 = New LabelGradient.LabelGradient
        Me.lblMachineName = New System.Windows.Forms.Label
        Me.lblUserNameText = New System.Windows.Forms.Label
        Me.lblUserName = New System.Windows.Forms.Label
        Me.gbUserLogin = New System.Windows.Forms.GroupBox
        Me.gbLoginDetails = New System.Windows.Forms.GroupBox
        Me.Label2 = New LabelGradient.LabelGradient
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.grpDeckPlateDetails = New System.Windows.Forms.GroupBox
        Me.txtlPartCode = New IFLCustomUILayer.IFLTextBox
        Me.txtContractNumber = New IFLCustomUILayer.IFLNumericBox
        Me.cmbAssemblyType = New IFLCustomUILayer.IFLComboBox
        Me.lblNoofFlats = New System.Windows.Forms.Label
        Me.lblDPFlatThickness = New System.Windows.Forms.Label
        Me.lblDPFlatWidth = New System.Windows.Forms.Label
        Me.txtCustomerName = New IFLCustomUILayer.IFLTextBox
        Me.lblDPInsulationThk = New System.Windows.Forms.Label
        Me.LabelGradient4 = New LabelGradient.LabelGradient
        Me.btnNext = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbUserLogin.SuspendLayout()
        Me.gbLoginDetails.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDeckPlateDetails.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'LabelGradient2
        '
        Me.LabelGradient2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelGradient2.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient2.ForeColor = System.Drawing.Color.Black
        Me.LabelGradient2.GradientColorTwo = System.Drawing.Color.White
        Me.LabelGradient2.Location = New System.Drawing.Point(30, 0)
        Me.LabelGradient2.Name = "LabelGradient2"
        Me.LabelGradient2.Size = New System.Drawing.Size(179, 15)
        Me.LabelGradient2.TabIndex = 23
        Me.LabelGradient2.Text = "Login details"
        Me.LabelGradient2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMachineNameText
        '
        Me.lblMachineNameText.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblMachineNameText.AutoSize = True
        Me.lblMachineNameText.ForeColor = System.Drawing.Color.Black
        Me.lblMachineNameText.Location = New System.Drawing.Point(97, 39)
        Me.lblMachineNameText.Name = "lblMachineNameText"
        Me.lblMachineNameText.Size = New System.Drawing.Size(0, 13)
        Me.lblMachineNameText.TabIndex = 21
        '
        'LabelGradient1
        '
        Me.LabelGradient1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelGradient1.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient1.ForeColor = System.Drawing.Color.Black
        Me.LabelGradient1.GradientColorTwo = System.Drawing.Color.White
        Me.LabelGradient1.Location = New System.Drawing.Point(42, 2)
        Me.LabelGradient1.Name = "LabelGradient1"
        Me.LabelGradient1.Size = New System.Drawing.Size(190, 15)
        Me.LabelGradient1.TabIndex = 24
        Me.LabelGradient1.Text = "Database details"
        Me.LabelGradient1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblServer
        '
        Me.lblServer.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(97, 36)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(0, 13)
        Me.lblServer.TabIndex = 21
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(6, 42)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Server Name:"
        '
        'lblDB
        '
        Me.lblDB.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblDB.AutoSize = True
        Me.lblDB.Location = New System.Drawing.Point(97, 17)
        Me.lblDB.Name = "lblDB"
        Me.lblDB.Size = New System.Drawing.Size(0, 13)
        Me.lblDB.TabIndex = 19
        '
        'labledb
        '
        Me.labledb.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.labledb.AutoSize = True
        Me.labledb.BackColor = System.Drawing.Color.Transparent
        Me.labledb.ForeColor = System.Drawing.Color.Black
        Me.labledb.Location = New System.Drawing.Point(6, 23)
        Me.labledb.Name = "labledb"
        Me.labledb.Size = New System.Drawing.Size(87, 13)
        Me.labledb.TabIndex = 18
        Me.labledb.Text = "Database Name:"
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(1082, -206)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(207, 56)
        Me.PictureBox2.TabIndex = 38
        Me.PictureBox2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(-319, -206)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(52, 56)
        Me.PictureBox3.TabIndex = 37
        Me.PictureBox3.TabStop = False
        '
        'LabelGradient3
        '
        Me.LabelGradient3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelGradient3.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.LabelGradient3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient3.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.LabelGradient3.GradientColorTwo = System.Drawing.Color.Azure
        Me.LabelGradient3.Location = New System.Drawing.Point(-270, -206)
        Me.LabelGradient3.Name = "LabelGradient3"
        Me.LabelGradient3.Size = New System.Drawing.Size(1364, 56)
        Me.LabelGradient3.TabIndex = 39
        Me.LabelGradient3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMachineName
        '
        Me.lblMachineName.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblMachineName.AutoSize = True
        Me.lblMachineName.BackColor = System.Drawing.Color.Transparent
        Me.lblMachineName.ForeColor = System.Drawing.Color.Black
        Me.lblMachineName.Location = New System.Drawing.Point(6, 44)
        Me.lblMachineName.Name = "lblMachineName"
        Me.lblMachineName.Size = New System.Drawing.Size(82, 13)
        Me.lblMachineName.TabIndex = 20
        Me.lblMachineName.Text = "Machine Name:"
        '
        'lblUserNameText
        '
        Me.lblUserNameText.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblUserNameText.AutoSize = True
        Me.lblUserNameText.Location = New System.Drawing.Point(97, 20)
        Me.lblUserNameText.Name = "lblUserNameText"
        Me.lblUserNameText.Size = New System.Drawing.Size(0, 13)
        Me.lblUserNameText.TabIndex = 19
        '
        'lblUserName
        '
        Me.lblUserName.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.lblUserName.AutoSize = True
        Me.lblUserName.BackColor = System.Drawing.Color.Transparent
        Me.lblUserName.ForeColor = System.Drawing.Color.Black
        Me.lblUserName.Location = New System.Drawing.Point(6, 25)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(63, 13)
        Me.lblUserName.TabIndex = 18
        Me.lblUserName.Text = "User Name:"
        '
        'gbUserLogin
        '
        Me.gbUserLogin.BackColor = System.Drawing.Color.Transparent
        Me.gbUserLogin.Controls.Add(Me.LabelGradient2)
        Me.gbUserLogin.Controls.Add(Me.lblMachineNameText)
        Me.gbUserLogin.Controls.Add(Me.lblMachineName)
        Me.gbUserLogin.Controls.Add(Me.lblUserNameText)
        Me.gbUserLogin.Controls.Add(Me.lblUserName)
        Me.gbUserLogin.Location = New System.Drawing.Point(-303, -132)
        Me.gbUserLogin.Name = "gbUserLogin"
        Me.gbUserLogin.Size = New System.Drawing.Size(264, 61)
        Me.gbUserLogin.TabIndex = 36
        Me.gbUserLogin.TabStop = False
        '
        'gbLoginDetails
        '
        Me.gbLoginDetails.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbLoginDetails.BackColor = System.Drawing.Color.Transparent
        Me.gbLoginDetails.Controls.Add(Me.LabelGradient1)
        Me.gbLoginDetails.Controls.Add(Me.lblServer)
        Me.gbLoginDetails.Controls.Add(Me.Label3)
        Me.gbLoginDetails.Controls.Add(Me.lblDB)
        Me.gbLoginDetails.Controls.Add(Me.labledb)
        Me.gbLoginDetails.Location = New System.Drawing.Point(1021, -133)
        Me.gbLoginDetails.Name = "gbLoginDetails"
        Me.gbLoginDetails.Size = New System.Drawing.Size(264, 58)
        Me.gbLoginDetails.TabIndex = 33
        Me.gbLoginDetails.TabStop = False
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label2.AutoSize = True
        Me.Label2.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.Label2.GradientColorTwo = System.Drawing.Color.Azure
        Me.Label2.Location = New System.Drawing.Point(395, 552)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(191, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Copyrights © 2008 Idola Fori Ltd"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Location = New System.Drawing.Point(-310, -65)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1607, 10)
        Me.PictureBox1.TabIndex = 31
        Me.PictureBox1.TabStop = False
        '
        'btnExit
        '
        Me.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.OutlineButton
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.btnExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ImageIndex = 0
        Me.btnExit.Location = New System.Drawing.Point(1170, 547)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(119, 23)
        Me.btnExit.TabIndex = 35
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'grpDeckPlateDetails
        '
        Me.grpDeckPlateDetails.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.grpDeckPlateDetails.BackColor = System.Drawing.Color.WhiteSmoke
        Me.grpDeckPlateDetails.Controls.Add(Me.txtlPartCode)
        Me.grpDeckPlateDetails.Controls.Add(Me.txtContractNumber)
        Me.grpDeckPlateDetails.Controls.Add(Me.cmbAssemblyType)
        Me.grpDeckPlateDetails.Controls.Add(Me.lblNoofFlats)
        Me.grpDeckPlateDetails.Controls.Add(Me.lblDPFlatThickness)
        Me.grpDeckPlateDetails.Controls.Add(Me.lblDPFlatWidth)
        Me.grpDeckPlateDetails.Controls.Add(Me.txtCustomerName)
        Me.grpDeckPlateDetails.Controls.Add(Me.lblDPInsulationThk)
        Me.grpDeckPlateDetails.Controls.Add(Me.LabelGradient4)
        Me.grpDeckPlateDetails.Location = New System.Drawing.Point(493, 89)
        Me.grpDeckPlateDetails.Name = "grpDeckPlateDetails"
        Me.grpDeckPlateDetails.Size = New System.Drawing.Size(450, 174)
        Me.grpDeckPlateDetails.TabIndex = 44
        Me.grpDeckPlateDetails.TabStop = False
        '
        'txtlPartCode
        '
        Me.txtlPartCode.AcceptEnterKeyAsTab = True
        Me.txtlPartCode.ApplyIFLColor = True
        Me.txtlPartCode.AssociateLabel = Nothing
        Me.txtlPartCode.IFLDataTag = Nothing
        Me.txtlPartCode.InvalidInputCharacters = ""
        Me.txtlPartCode.Location = New System.Drawing.Point(105, 126)
        Me.txtlPartCode.MaxLength = 32
        Me.txtlPartCode.Name = "txtlPartCode"
        Me.txtlPartCode.Size = New System.Drawing.Size(321, 20)
        Me.txtlPartCode.StatusMessage = ""
        Me.txtlPartCode.StatusObject = Nothing
        Me.txtlPartCode.TabIndex = 22
        Me.txtlPartCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtContractNumber
        '
        Me.txtContractNumber.AcceptEnterKeyAsTab = True
        Me.txtContractNumber.ApplyIFLColor = True
        Me.txtContractNumber.AssociateLabel = Nothing
        Me.txtContractNumber.DecimalValue = 0
        Me.txtContractNumber.IFLDataTag = Nothing
        Me.txtContractNumber.InvalidInputCharacters = ""
        Me.txtContractNumber.IsAllowNegative = False
        Me.txtContractNumber.LengthValue = 6
        Me.txtContractNumber.Location = New System.Drawing.Point(105, 62)
        Me.txtContractNumber.MaximumValue = 99999
        Me.txtContractNumber.MaxLength = 32
        Me.txtContractNumber.MinimumValue = 0
        Me.txtContractNumber.Name = "txtContractNumber"
        Me.txtContractNumber.Size = New System.Drawing.Size(321, 20)
        Me.txtContractNumber.StatusMessage = ""
        Me.txtContractNumber.StatusObject = Nothing
        Me.txtContractNumber.TabIndex = 21
        Me.txtContractNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmbAssemblyType
        '
        Me.cmbAssemblyType.AcceptEnterKeyAsTab = True
        Me.cmbAssemblyType.AssociateLabel = Nothing
        Me.cmbAssemblyType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAssemblyType.FormattingEnabled = True
        Me.cmbAssemblyType.IFLDataTag = Nothing
        Me.cmbAssemblyType.Items.AddRange(New Object() {"", "Tie Rod Cylinder Assembly", "Welded Cylinder Assembly"})
        Me.cmbAssemblyType.Location = New System.Drawing.Point(105, 93)
        Me.cmbAssemblyType.Name = "cmbAssemblyType"
        Me.cmbAssemblyType.Size = New System.Drawing.Size(321, 21)
        Me.cmbAssemblyType.StatusMessage = Nothing
        Me.cmbAssemblyType.StatusObject = Nothing
        Me.cmbAssemblyType.TabIndex = 3
        '
        'lblNoofFlats
        '
        Me.lblNoofFlats.AutoSize = True
        Me.lblNoofFlats.Location = New System.Drawing.Point(45, 130)
        Me.lblNoofFlats.Name = "lblNoofFlats"
        Me.lblNoofFlats.Size = New System.Drawing.Size(54, 13)
        Me.lblNoofFlats.TabIndex = 4
        Me.lblNoofFlats.Text = "Part Code"
        '
        'lblDPFlatThickness
        '
        Me.lblDPFlatThickness.AutoSize = True
        Me.lblDPFlatThickness.Location = New System.Drawing.Point(68, 97)
        Me.lblDPFlatThickness.Name = "lblDPFlatThickness"
        Me.lblDPFlatThickness.Size = New System.Drawing.Size(31, 13)
        Me.lblDPFlatThickness.TabIndex = 3
        Me.lblDPFlatThickness.Text = "Type"
        '
        'lblDPFlatWidth
        '
        Me.lblDPFlatWidth.AutoSize = True
        Me.lblDPFlatWidth.Location = New System.Drawing.Point(12, 66)
        Me.lblDPFlatWidth.Name = "lblDPFlatWidth"
        Me.lblDPFlatWidth.Size = New System.Drawing.Size(87, 13)
        Me.lblDPFlatWidth.TabIndex = 2
        Me.lblDPFlatWidth.Text = "Contract Number"
        '
        'txtCustomerName
        '
        Me.txtCustomerName.AcceptEnterKeyAsTab = True
        Me.txtCustomerName.ApplyIFLColor = True
        Me.txtCustomerName.AssociateLabel = Nothing
        Me.txtCustomerName.IFLDataTag = Nothing
        Me.txtCustomerName.InvalidInputCharacters = ""
        Me.txtCustomerName.Location = New System.Drawing.Point(105, 31)
        Me.txtCustomerName.MaxLength = 32
        Me.txtCustomerName.Name = "txtCustomerName"
        Me.txtCustomerName.Size = New System.Drawing.Size(321, 20)
        Me.txtCustomerName.StatusMessage = ""
        Me.txtCustomerName.StatusObject = Nothing
        Me.txtCustomerName.TabIndex = 1
        Me.txtCustomerName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDPInsulationThk
        '
        Me.lblDPInsulationThk.AutoSize = True
        Me.lblDPInsulationThk.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDPInsulationThk.Location = New System.Drawing.Point(4, 34)
        Me.lblDPInsulationThk.Name = "lblDPInsulationThk"
        Me.lblDPInsulationThk.Size = New System.Drawing.Size(95, 15)
        Me.lblDPInsulationThk.TabIndex = 0
        Me.lblDPInsulationThk.Text = "Customer Name"
        '
        'LabelGradient4
        '
        Me.LabelGradient4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelGradient4.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient4.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.LabelGradient4.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient4.Location = New System.Drawing.Point(0, 0)
        Me.LabelGradient4.Name = "LabelGradient4"
        Me.LabelGradient4.Size = New System.Drawing.Size(450, 15)
        Me.LabelGradient4.TabIndex = 20
        Me.LabelGradient4.Text = "Priliminary Information"
        Me.LabelGradient4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNext.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.btnNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext.ImageIndex = 0
        Me.btnNext.Location = New System.Drawing.Point(882, 315)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(61, 23)
        Me.btnNext.TabIndex = 45
        Me.btnNext.TabStop = False
        Me.btnNext.Text = "&Next"
        Me.btnNext.UseVisualStyleBackColor = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripDropDownButton1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 341)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.ShowItemToolTips = True
        Me.StatusStrip1.Size = New System.Drawing.Size(977, 22)
        Me.StatusStrip1.TabIndex = 46
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(111, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripDropDownButton1.Image = CType(resources.GetObject("ToolStripDropDownButton1.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(29, 20)
        Me.ToolStripDropDownButton1.Text = "ToolStripDropDownButton1"
        '
        'frmContractDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(977, 363)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.grpDeckPlateDetails)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.LabelGradient3)
        Me.Controls.Add(Me.gbUserLogin)
        Me.Controls.Add(Me.gbLoginDetails)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmContractDetails"
        Me.Text = "Priliminary Information"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbUserLogin.ResumeLayout(False)
        Me.gbUserLogin.PerformLayout()
        Me.gbLoginDetails.ResumeLayout(False)
        Me.gbLoginDetails.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDeckPlateDetails.ResumeLayout(False)
        Me.grpDeckPlateDetails.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelGradient2 As LabelGradient.LabelGradient
    Friend WithEvents lblMachineNameText As System.Windows.Forms.Label
    Friend WithEvents LabelGradient1 As LabelGradient.LabelGradient
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblDB As System.Windows.Forms.Label
    Friend WithEvents labledb As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents LabelGradient3 As LabelGradient.LabelGradient
    Friend WithEvents lblMachineName As System.Windows.Forms.Label
    Friend WithEvents lblUserNameText As System.Windows.Forms.Label
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    Friend WithEvents gbUserLogin As System.Windows.Forms.GroupBox
    Friend WithEvents gbLoginDetails As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As LabelGradient.LabelGradient
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents grpDeckPlateDetails As System.Windows.Forms.GroupBox
    Friend WithEvents LabelGradient4 As LabelGradient.LabelGradient
    Friend WithEvents cmbAssemblyType As IFLCustomUILayer.IFLComboBox
    Friend WithEvents lblNoofFlats As System.Windows.Forms.Label
    Friend WithEvents lblDPFlatThickness As System.Windows.Forms.Label
    Friend WithEvents lblDPFlatWidth As System.Windows.Forms.Label
    Friend WithEvents txtCustomerName As IFLCustomUILayer.IFLTextBox
    Friend WithEvents lblDPInsulationThk As System.Windows.Forms.Label
    Friend WithEvents txtlPartCode As IFLCustomUILayer.IFLTextBox
    Friend WithEvents txtContractNumber As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripDropDownButton1 As System.Windows.Forms.ToolStripDropDownButton
End Class
