<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTieRod2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTieRod2))
        Me.GroupBox19 = New System.Windows.Forms.GroupBox
        Me.IflNumericBox14 = New IFLCustomUILayer.IFLNumericBox
        Me.Label61 = New System.Windows.Forms.Label
        Me.IflNumericBox15 = New IFLCustomUILayer.IFLNumericBox
        Me.Label62 = New System.Windows.Forms.Label
        Me.IflComboBox12 = New IFLCustomUILayer.IFLComboBox
        Me.Label63 = New System.Windows.Forms.Label
        Me.IflComboBox13 = New IFLCustomUILayer.IFLComboBox
        Me.Label64 = New System.Windows.Forms.Label
        Me.LabelGradient20 = New LabelGradient.LabelGradient
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.txtTieRodNutQty = New IFLCustomUILayer.IFLNumericBox
        Me.Label34 = New System.Windows.Forms.Label
        Me.txtTieRodNutSize = New IFLCustomUILayer.IFLNumericBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.txtTieRodSize = New IFLCustomUILayer.IFLNumericBox
        Me.cmbThreadProtected = New IFLCustomUILayer.IFLComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.LabelGradient8 = New LabelGradient.LabelGradient
        Me.btnBack = New System.Windows.Forms.Button
        Me.btnNext = New System.Windows.Forms.Button
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.txtClevisCap = New IFLCustomUILayer.IFLNumericBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.cmbRodClevis = New IFLCustomUILayer.IFLComboBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.cmbRodEndThread = New IFLCustomUILayer.IFLComboBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.txtRodCap = New IFLCustomUILayer.IFLNumericBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.LabelGradient6 = New LabelGradient.LabelGradient
        Me.cmbRodSealPackage = New IFLCustomUILayer.IFLComboBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.IflNumericBox11 = New IFLCustomUILayer.IFLNumericBox
        Me.IflNumericBox10 = New IFLCustomUILayer.IFLNumericBox
        Me.cmbClips = New IFLCustomUILayer.IFLComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.LVPinSizeDetails = New IFLCustomUILayer.IFLListView
        Me.optPinsNo = New System.Windows.Forms.RadioButton
        Me.optPinsYes = New System.Windows.Forms.RadioButton
        Me.cmbPinMaterial = New IFLCustomUILayer.IFLComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.LabelGradient2 = New LabelGradient.LabelGradient
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.LabelGradient7 = New LabelGradient.LabelGradient
        Me.cmbPistonSealPackage = New IFLCustomUILayer.IFLComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtWorkingPressure = New IFLCustomUILayer.IFLNumericBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtColumnLoad = New IFLCustomUILayer.IFLNumericBox
        Me.GroupBox19.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox19
        '
        Me.GroupBox19.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox19.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox19.Controls.Add(Me.IflNumericBox14)
        Me.GroupBox19.Controls.Add(Me.Label61)
        Me.GroupBox19.Controls.Add(Me.IflNumericBox15)
        Me.GroupBox19.Controls.Add(Me.Label62)
        Me.GroupBox19.Controls.Add(Me.IflComboBox12)
        Me.GroupBox19.Controls.Add(Me.Label63)
        Me.GroupBox19.Controls.Add(Me.IflComboBox13)
        Me.GroupBox19.Controls.Add(Me.Label64)
        Me.GroupBox19.Controls.Add(Me.LabelGradient20)
        Me.GroupBox19.Location = New System.Drawing.Point(502, 322)
        Me.GroupBox19.Name = "GroupBox19"
        Me.GroupBox19.Size = New System.Drawing.Size(436, 144)
        Me.GroupBox19.TabIndex = 85
        Me.GroupBox19.TabStop = False
        '
        'IflNumericBox14
        '
        Me.IflNumericBox14.AcceptEnterKeyAsTab = True
        Me.IflNumericBox14.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.IflNumericBox14.ApplyIFLColor = True
        Me.IflNumericBox14.AssociateLabel = Nothing
        Me.IflNumericBox14.DecimalValue = 2
        Me.IflNumericBox14.IFLDataTag = Nothing
        Me.IflNumericBox14.InvalidInputCharacters = ""
        Me.IflNumericBox14.IsAllowNegative = False
        Me.IflNumericBox14.LengthValue = 6
        Me.IflNumericBox14.Location = New System.Drawing.Point(132, 79)
        Me.IflNumericBox14.MaximumValue = 99999
        Me.IflNumericBox14.MaxLength = 6
        Me.IflNumericBox14.MinimumValue = 0
        Me.IflNumericBox14.Name = "IflNumericBox14"
        Me.IflNumericBox14.Size = New System.Drawing.Size(279, 20)
        Me.IflNumericBox14.StatusMessage = ""
        Me.IflNumericBox14.StatusObject = Nothing
        Me.IflNumericBox14.TabIndex = 59
        Me.IflNumericBox14.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label61
        '
        Me.Label61.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label61.AutoSize = True
        Me.Label61.Location = New System.Drawing.Point(60, 84)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(58, 13)
        Me.Label61.TabIndex = 58
        Me.Label61.Text = "Rod Wiper"
        '
        'IflNumericBox15
        '
        Me.IflNumericBox15.AcceptEnterKeyAsTab = True
        Me.IflNumericBox15.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.IflNumericBox15.ApplyIFLColor = True
        Me.IflNumericBox15.AssociateLabel = Nothing
        Me.IflNumericBox15.DecimalValue = 2
        Me.IflNumericBox15.IFLDataTag = Nothing
        Me.IflNumericBox15.InvalidInputCharacters = ""
        Me.IflNumericBox15.IsAllowNegative = False
        Me.IflNumericBox15.LengthValue = 6
        Me.IflNumericBox15.Location = New System.Drawing.Point(132, 53)
        Me.IflNumericBox15.MaximumValue = 99999
        Me.IflNumericBox15.MaxLength = 6
        Me.IflNumericBox15.MinimumValue = 0
        Me.IflNumericBox15.Name = "IflNumericBox15"
        Me.IflNumericBox15.Size = New System.Drawing.Size(279, 20)
        Me.IflNumericBox15.StatusMessage = ""
        Me.IflNumericBox15.StatusObject = Nothing
        Me.IflNumericBox15.TabIndex = 57
        Me.IflNumericBox15.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label62
        '
        Me.Label62.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label62.AutoSize = True
        Me.Label62.Location = New System.Drawing.Point(60, 58)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(58, 13)
        Me.Label62.TabIndex = 56
        Me.Label62.Text = "Packaging"
        '
        'IflComboBox12
        '
        Me.IflComboBox12.AcceptEnterKeyAsTab = True
        Me.IflComboBox12.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.IflComboBox12.AssociateLabel = Nothing
        Me.IflComboBox12.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.IflComboBox12.FormattingEnabled = True
        Me.IflComboBox12.IFLDataTag = Nothing
        Me.IflComboBox12.Location = New System.Drawing.Point(132, 26)
        Me.IflComboBox12.Name = "IflComboBox12"
        Me.IflComboBox12.Size = New System.Drawing.Size(280, 21)
        Me.IflComboBox12.StatusMessage = Nothing
        Me.IflComboBox12.StatusObject = Nothing
        Me.IflComboBox12.TabIndex = 71
        '
        'Label63
        '
        Me.Label63.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label63.AutoSize = True
        Me.Label63.Location = New System.Drawing.Point(89, 34)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(31, 13)
        Me.Label63.TabIndex = 70
        Me.Label63.Text = "Paint"
        '
        'IflComboBox13
        '
        Me.IflComboBox13.AcceptEnterKeyAsTab = True
        Me.IflComboBox13.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.IflComboBox13.AssociateLabel = Nothing
        Me.IflComboBox13.BackColor = System.Drawing.Color.White
        Me.IflComboBox13.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.IflComboBox13.FormattingEnabled = True
        Me.IflComboBox13.IFLDataTag = Nothing
        Me.IflComboBox13.Location = New System.Drawing.Point(132, 109)
        Me.IflComboBox13.Name = "IflComboBox13"
        Me.IflComboBox13.Size = New System.Drawing.Size(279, 21)
        Me.IflComboBox13.StatusMessage = Nothing
        Me.IflComboBox13.StatusObject = Nothing
        Me.IflComboBox13.TabIndex = 69
        '
        'Label64
        '
        Me.Label64.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label64.AutoSize = True
        Me.Label64.Location = New System.Drawing.Point(64, 114)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(56, 13)
        Me.Label64.TabIndex = 68
        Me.Label64.Text = "Tube Seal"
        '
        'LabelGradient20
        '
        Me.LabelGradient20.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelGradient20.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient20.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient20.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.LabelGradient20.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient20.Location = New System.Drawing.Point(-3, 4)
        Me.LabelGradient20.Name = "LabelGradient20"
        Me.LabelGradient20.Size = New System.Drawing.Size(439, 16)
        Me.LabelGradient20.TabIndex = 20
        Me.LabelGradient20.Text = "External Details"
        Me.LabelGradient20.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'GroupBox7
        '
        Me.GroupBox7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox7.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox7.Controls.Add(Me.txtTieRodNutQty)
        Me.GroupBox7.Controls.Add(Me.Label34)
        Me.GroupBox7.Controls.Add(Me.txtTieRodNutSize)
        Me.GroupBox7.Controls.Add(Me.Label33)
        Me.GroupBox7.Controls.Add(Me.txtTieRodSize)
        Me.GroupBox7.Controls.Add(Me.cmbThreadProtected)
        Me.GroupBox7.Controls.Add(Me.Label17)
        Me.GroupBox7.Controls.Add(Me.Label32)
        Me.GroupBox7.Controls.Add(Me.LabelGradient8)
        Me.GroupBox7.Location = New System.Drawing.Point(34, 323)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(462, 144)
        Me.GroupBox7.TabIndex = 84
        Me.GroupBox7.TabStop = False
        '
        'txtTieRodNutQty
        '
        Me.txtTieRodNutQty.AcceptEnterKeyAsTab = True
        Me.txtTieRodNutQty.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtTieRodNutQty.ApplyIFLColor = True
        Me.txtTieRodNutQty.AssociateLabel = Nothing
        Me.txtTieRodNutQty.DecimalValue = 2
        Me.txtTieRodNutQty.IFLDataTag = Nothing
        Me.txtTieRodNutQty.InvalidInputCharacters = ""
        Me.txtTieRodNutQty.IsAllowNegative = False
        Me.txtTieRodNutQty.LengthValue = 6
        Me.txtTieRodNutQty.Location = New System.Drawing.Point(109, 87)
        Me.txtTieRodNutQty.MaximumValue = 99999
        Me.txtTieRodNutQty.MaxLength = 6
        Me.txtTieRodNutQty.MinimumValue = 0
        Me.txtTieRodNutQty.Name = "txtTieRodNutQty"
        Me.txtTieRodNutQty.Size = New System.Drawing.Size(341, 20)
        Me.txtTieRodNutQty.StatusMessage = ""
        Me.txtTieRodNutQty.StatusObject = Nothing
        Me.txtTieRodNutQty.TabIndex = 59
        Me.txtTieRodNutQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label34
        '
        Me.Label34.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(18, 91)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(84, 13)
        Me.Label34.TabIndex = 58
        Me.Label34.Text = "Tie Rod Nut Qty"
        '
        'txtTieRodNutSize
        '
        Me.txtTieRodNutSize.AcceptEnterKeyAsTab = True
        Me.txtTieRodNutSize.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtTieRodNutSize.ApplyIFLColor = True
        Me.txtTieRodNutSize.AssociateLabel = Nothing
        Me.txtTieRodNutSize.DecimalValue = 2
        Me.txtTieRodNutSize.IFLDataTag = Nothing
        Me.txtTieRodNutSize.InvalidInputCharacters = ""
        Me.txtTieRodNutSize.IsAllowNegative = False
        Me.txtTieRodNutSize.LengthValue = 6
        Me.txtTieRodNutSize.Location = New System.Drawing.Point(109, 61)
        Me.txtTieRodNutSize.MaximumValue = 99999
        Me.txtTieRodNutSize.MaxLength = 6
        Me.txtTieRodNutSize.MinimumValue = 0
        Me.txtTieRodNutSize.Name = "txtTieRodNutSize"
        Me.txtTieRodNutSize.Size = New System.Drawing.Size(341, 20)
        Me.txtTieRodNutSize.StatusMessage = ""
        Me.txtTieRodNutSize.StatusObject = Nothing
        Me.txtTieRodNutSize.TabIndex = 57
        Me.txtTieRodNutSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label33
        '
        Me.Label33.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(14, 65)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(88, 13)
        Me.Label33.TabIndex = 56
        Me.Label33.Text = "Tie Rod Nut Size"
        '
        'txtTieRodSize
        '
        Me.txtTieRodSize.AcceptEnterKeyAsTab = True
        Me.txtTieRodSize.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtTieRodSize.ApplyIFLColor = True
        Me.txtTieRodSize.AssociateLabel = Nothing
        Me.txtTieRodSize.DecimalValue = 2
        Me.txtTieRodSize.IFLDataTag = Nothing
        Me.txtTieRodSize.InvalidInputCharacters = ""
        Me.txtTieRodSize.IsAllowNegative = False
        Me.txtTieRodSize.LengthValue = 6
        Me.txtTieRodSize.Location = New System.Drawing.Point(109, 35)
        Me.txtTieRodSize.MaximumValue = 99999
        Me.txtTieRodSize.MaxLength = 6
        Me.txtTieRodSize.MinimumValue = 0
        Me.txtTieRodSize.Name = "txtTieRodSize"
        Me.txtTieRodSize.Size = New System.Drawing.Size(341, 20)
        Me.txtTieRodSize.StatusMessage = ""
        Me.txtTieRodSize.StatusObject = Nothing
        Me.txtTieRodSize.TabIndex = 55
        Me.txtTieRodSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmbThreadProtected
        '
        Me.cmbThreadProtected.AcceptEnterKeyAsTab = True
        Me.cmbThreadProtected.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmbThreadProtected.AssociateLabel = Nothing
        Me.cmbThreadProtected.BackColor = System.Drawing.Color.White
        Me.cmbThreadProtected.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbThreadProtected.FormattingEnabled = True
        Me.cmbThreadProtected.IFLDataTag = Nothing
        Me.cmbThreadProtected.Items.AddRange(New Object() {"", "Standard", "All Permenant"})
        Me.cmbThreadProtected.Location = New System.Drawing.Point(109, 110)
        Me.cmbThreadProtected.Name = "cmbThreadProtected"
        Me.cmbThreadProtected.Size = New System.Drawing.Size(341, 21)
        Me.cmbThreadProtected.StatusMessage = Nothing
        Me.cmbThreadProtected.StatusObject = Nothing
        Me.cmbThreadProtected.TabIndex = 69
        '
        'Label17
        '
        Me.Label17.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(12, 114)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(90, 13)
        Me.Label17.TabIndex = 68
        Me.Label17.Text = "Thread Protected"
        '
        'Label32
        '
        Me.Label32.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(34, 39)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(68, 13)
        Me.Label32.TabIndex = 54
        Me.Label32.Text = "Tie Rod Size"
        '
        'LabelGradient8
        '
        Me.LabelGradient8.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelGradient8.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient8.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.LabelGradient8.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient8.Location = New System.Drawing.Point(0, 0)
        Me.LabelGradient8.Name = "LabelGradient8"
        Me.LabelGradient8.Size = New System.Drawing.Size(462, 16)
        Me.LabelGradient8.TabIndex = 20
        Me.LabelGradient8.Text = "Tie Rod Details"
        Me.LabelGradient8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnBack
        '
        Me.btnBack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBack.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.btnBack.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.ImageIndex = 0
        Me.btnBack.Location = New System.Drawing.Point(810, 626)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(61, 23)
        Me.btnBack.TabIndex = 82
        Me.btnBack.TabStop = False
        Me.btnBack.Text = "&Back"
        Me.btnBack.UseVisualStyleBackColor = False
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNext.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.btnNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext.ImageIndex = 0
        Me.btnNext.Location = New System.Drawing.Point(877, 626)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(61, 23)
        Me.btnNext.TabIndex = 81
        Me.btnNext.TabStop = False
        Me.btnNext.Text = "&Next"
        Me.btnNext.UseVisualStyleBackColor = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox5.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox5.Controls.Add(Me.txtClevisCap)
        Me.GroupBox5.Controls.Add(Me.Label31)
        Me.GroupBox5.Controls.Add(Me.cmbRodClevis)
        Me.GroupBox5.Controls.Add(Me.Label23)
        Me.GroupBox5.Controls.Add(Me.cmbRodEndThread)
        Me.GroupBox5.Controls.Add(Me.Label25)
        Me.GroupBox5.Controls.Add(Me.txtRodCap)
        Me.GroupBox5.Controls.Add(Me.Label26)
        Me.GroupBox5.Controls.Add(Me.LabelGradient6)
        Me.GroupBox5.Controls.Add(Me.cmbRodSealPackage)
        Me.GroupBox5.Controls.Add(Me.Label27)
        Me.GroupBox5.Location = New System.Drawing.Point(502, 75)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(436, 171)
        Me.GroupBox5.TabIndex = 80
        Me.GroupBox5.TabStop = False
        '
        'txtClevisCap
        '
        Me.txtClevisCap.AcceptEnterKeyAsTab = True
        Me.txtClevisCap.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtClevisCap.ApplyIFLColor = True
        Me.txtClevisCap.AssociateLabel = Nothing
        Me.txtClevisCap.DecimalValue = 2
        Me.txtClevisCap.IFLDataTag = Nothing
        Me.txtClevisCap.InvalidInputCharacters = ""
        Me.txtClevisCap.IsAllowNegative = False
        Me.txtClevisCap.LengthValue = 6
        Me.txtClevisCap.Location = New System.Drawing.Point(132, 79)
        Me.txtClevisCap.MaximumValue = 99999
        Me.txtClevisCap.MaxLength = 6
        Me.txtClevisCap.MinimumValue = 0
        Me.txtClevisCap.Name = "txtClevisCap"
        Me.txtClevisCap.Size = New System.Drawing.Size(280, 20)
        Me.txtClevisCap.StatusMessage = ""
        Me.txtClevisCap.StatusObject = Nothing
        Me.txtClevisCap.TabIndex = 62
        Me.txtClevisCap.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label31
        '
        Me.Label31.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(63, 83)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(57, 13)
        Me.Label31.TabIndex = 61
        Me.Label31.Text = "Clevis Cap"
        '
        'cmbRodClevis
        '
        Me.cmbRodClevis.AcceptEnterKeyAsTab = True
        Me.cmbRodClevis.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmbRodClevis.AssociateLabel = Nothing
        Me.cmbRodClevis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbRodClevis.FormattingEnabled = True
        Me.cmbRodClevis.IFLDataTag = Nothing
        Me.cmbRodClevis.Items.AddRange(New Object() {"", "10", "12", "16"})
        Me.cmbRodClevis.Location = New System.Drawing.Point(132, 129)
        Me.cmbRodClevis.Name = "cmbRodClevis"
        Me.cmbRodClevis.Size = New System.Drawing.Size(280, 21)
        Me.cmbRodClevis.StatusMessage = Nothing
        Me.cmbRodClevis.StatusObject = Nothing
        Me.cmbRodClevis.TabIndex = 60
        '
        'Label23
        '
        Me.Label23.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(62, 133)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(58, 13)
        Me.Label23.TabIndex = 59
        Me.Label23.Text = "Rod Clevis"
        '
        'cmbRodEndThread
        '
        Me.cmbRodEndThread.AcceptEnterKeyAsTab = True
        Me.cmbRodEndThread.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmbRodEndThread.AssociateLabel = Nothing
        Me.cmbRodEndThread.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbRodEndThread.FormattingEnabled = True
        Me.cmbRodEndThread.IFLDataTag = Nothing
        Me.cmbRodEndThread.Location = New System.Drawing.Point(132, 102)
        Me.cmbRodEndThread.Name = "cmbRodEndThread"
        Me.cmbRodEndThread.Size = New System.Drawing.Size(280, 21)
        Me.cmbRodEndThread.StatusMessage = Nothing
        Me.cmbRodEndThread.StatusObject = Nothing
        Me.cmbRodEndThread.TabIndex = 51
        '
        'Label25
        '
        Me.Label25.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(11, 106)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(109, 13)
        Me.Label25.TabIndex = 50
        Me.Label25.Text = "Rod End Thread Size"
        '
        'txtRodCap
        '
        Me.txtRodCap.AcceptEnterKeyAsTab = True
        Me.txtRodCap.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtRodCap.ApplyIFLColor = True
        Me.txtRodCap.AssociateLabel = Nothing
        Me.txtRodCap.DecimalValue = 2
        Me.txtRodCap.IFLDataTag = Nothing
        Me.txtRodCap.InvalidInputCharacters = ""
        Me.txtRodCap.IsAllowNegative = False
        Me.txtRodCap.LengthValue = 6
        Me.txtRodCap.Location = New System.Drawing.Point(132, 56)
        Me.txtRodCap.MaximumValue = 99999
        Me.txtRodCap.MaxLength = 6
        Me.txtRodCap.MinimumValue = 0
        Me.txtRodCap.Name = "txtRodCap"
        Me.txtRodCap.Size = New System.Drawing.Size(280, 20)
        Me.txtRodCap.StatusMessage = ""
        Me.txtRodCap.StatusObject = Nothing
        Me.txtRodCap.TabIndex = 49
        Me.txtRodCap.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label26
        '
        Me.Label26.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(71, 60)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(49, 13)
        Me.Label26.TabIndex = 48
        Me.Label26.Text = "Rod Cap"
        '
        'LabelGradient6
        '
        Me.LabelGradient6.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelGradient6.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient6.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.LabelGradient6.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient6.Location = New System.Drawing.Point(0, 0)
        Me.LabelGradient6.Name = "LabelGradient6"
        Me.LabelGradient6.Size = New System.Drawing.Size(436, 15)
        Me.LabelGradient6.TabIndex = 20
        Me.LabelGradient6.Text = "Rod Seal, Rod thread and Rod Clevis Details"
        Me.LabelGradient6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmbRodSealPackage
        '
        Me.cmbRodSealPackage.AcceptEnterKeyAsTab = True
        Me.cmbRodSealPackage.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmbRodSealPackage.AssociateLabel = Nothing
        Me.cmbRodSealPackage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbRodSealPackage.FormattingEnabled = True
        Me.cmbRodSealPackage.IFLDataTag = Nothing
        Me.cmbRodSealPackage.Items.AddRange(New Object() {""})
        Me.cmbRodSealPackage.Location = New System.Drawing.Point(132, 29)
        Me.cmbRodSealPackage.Name = "cmbRodSealPackage"
        Me.cmbRodSealPackage.Size = New System.Drawing.Size(280, 21)
        Me.cmbRodSealPackage.StatusMessage = Nothing
        Me.cmbRodSealPackage.StatusObject = Nothing
        Me.cmbRodSealPackage.TabIndex = 3
        '
        'Label27
        '
        Me.Label27.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(23, 33)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(97, 13)
        Me.Label27.TabIndex = 3
        Me.Label27.Text = "Rod Seal Package"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox2.Controls.Add(Me.Label30)
        Me.GroupBox2.Controls.Add(Me.IflNumericBox11)
        Me.GroupBox2.Controls.Add(Me.IflNumericBox10)
        Me.GroupBox2.Controls.Add(Me.cmbClips)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.LVPinSizeDetails)
        Me.GroupBox2.Controls.Add(Me.optPinsNo)
        Me.GroupBox2.Controls.Add(Me.optPinsYes)
        Me.GroupBox2.Controls.Add(Me.cmbPinMaterial)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.LabelGradient2)
        Me.GroupBox2.Location = New System.Drawing.Point(34, 75)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(462, 242)
        Me.GroupBox2.TabIndex = 79
        Me.GroupBox2.TabStop = False
        '
        'Label30
        '
        Me.Label30.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(22, 106)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(80, 13)
        Me.Label30.TabIndex = 63
        Me.Label30.Text = "Pin Size Details"
        '
        'IflNumericBox11
        '
        Me.IflNumericBox11.AcceptEnterKeyAsTab = True
        Me.IflNumericBox11.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.IflNumericBox11.ApplyIFLColor = True
        Me.IflNumericBox11.AssociateLabel = Nothing
        Me.IflNumericBox11.DecimalValue = 2
        Me.IflNumericBox11.IFLDataTag = Nothing
        Me.IflNumericBox11.InvalidInputCharacters = ""
        Me.IflNumericBox11.IsAllowNegative = False
        Me.IflNumericBox11.LengthValue = 6
        Me.IflNumericBox11.Location = New System.Drawing.Point(108, 187)
        Me.IflNumericBox11.MaximumValue = 99999
        Me.IflNumericBox11.MaxLength = 6
        Me.IflNumericBox11.MinimumValue = 0
        Me.IflNumericBox11.Name = "IflNumericBox11"
        Me.IflNumericBox11.Size = New System.Drawing.Size(342, 20)
        Me.IflNumericBox11.StatusMessage = ""
        Me.IflNumericBox11.StatusObject = Nothing
        Me.IflNumericBox11.TabIndex = 62
        Me.IflNumericBox11.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'IflNumericBox10
        '
        Me.IflNumericBox10.AcceptEnterKeyAsTab = True
        Me.IflNumericBox10.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.IflNumericBox10.ApplyIFLColor = True
        Me.IflNumericBox10.AssociateLabel = Nothing
        Me.IflNumericBox10.DecimalValue = 2
        Me.IflNumericBox10.IFLDataTag = Nothing
        Me.IflNumericBox10.InvalidInputCharacters = ""
        Me.IflNumericBox10.IsAllowNegative = False
        Me.IflNumericBox10.LengthValue = 6
        Me.IflNumericBox10.Location = New System.Drawing.Point(108, 161)
        Me.IflNumericBox10.MaximumValue = 99999
        Me.IflNumericBox10.MaxLength = 6
        Me.IflNumericBox10.MinimumValue = 0
        Me.IflNumericBox10.Name = "IflNumericBox10"
        Me.IflNumericBox10.Size = New System.Drawing.Size(342, 20)
        Me.IflNumericBox10.StatusMessage = ""
        Me.IflNumericBox10.StatusObject = Nothing
        Me.IflNumericBox10.TabIndex = 61
        Me.IflNumericBox10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmbClips
        '
        Me.cmbClips.AcceptEnterKeyAsTab = True
        Me.cmbClips.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmbClips.AssociateLabel = Nothing
        Me.cmbClips.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbClips.FormattingEnabled = True
        Me.cmbClips.IFLDataTag = Nothing
        Me.cmbClips.Items.AddRange(New Object() {"", "Hair Pin", "Cotter Pin", "Cir Clips", "R Clip"})
        Me.cmbClips.Location = New System.Drawing.Point(108, 213)
        Me.cmbClips.Name = "cmbClips"
        Me.cmbClips.Size = New System.Drawing.Size(342, 21)
        Me.cmbClips.StatusMessage = Nothing
        Me.cmbClips.StatusObject = Nothing
        Me.cmbClips.TabIndex = 58
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(73, 217)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(29, 13)
        Me.Label9.TabIndex = 57
        Me.Label9.Text = "Clips"
        '
        'LVPinSizeDetails
        '
        Me.LVPinSizeDetails.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LVPinSizeDetails.BackColor = System.Drawing.Color.WhiteSmoke
        Me.LVPinSizeDetails.DisplayHeaders = CType(resources.GetObject("LVPinSizeDetails.DisplayHeaders"), System.Collections.ArrayList)
        Me.LVPinSizeDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LVPinSizeDetails.FullRowSelect = True
        Me.LVPinSizeDetails.GridLines = True
        Me.LVPinSizeDetails.HideSelection = False
        Me.LVPinSizeDetails.IFLDataTag = Nothing
        Me.LVPinSizeDetails.IsCheckBoxEnabled = False
        Me.LVPinSizeDetails.IsFilterOptionEnabled = False
        Me.LVPinSizeDetails.IsTypeSearchEnable = True
        Me.LVPinSizeDetails.Location = New System.Drawing.Point(108, 69)
        Me.LVPinSizeDetails.MultiSelect = False
        Me.LVPinSizeDetails.Name = "LVPinSizeDetails"
        Me.LVPinSizeDetails.SearchObject = Nothing
        Me.LVPinSizeDetails.Size = New System.Drawing.Size(342, 86)
        Me.LVPinSizeDetails.SourceTable = Nothing
        Me.LVPinSizeDetails.TabIndex = 56
        Me.LVPinSizeDetails.UseCompatibleStateImageBehavior = False
        Me.LVPinSizeDetails.View = System.Windows.Forms.View.Details
        '
        'optPinsNo
        '
        Me.optPinsNo.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.optPinsNo.AutoSize = True
        Me.optPinsNo.Location = New System.Drawing.Point(188, 21)
        Me.optPinsNo.Name = "optPinsNo"
        Me.optPinsNo.Size = New System.Drawing.Size(39, 17)
        Me.optPinsNo.TabIndex = 55
        Me.optPinsNo.TabStop = True
        Me.optPinsNo.Text = "No"
        Me.optPinsNo.UseVisualStyleBackColor = True
        '
        'optPinsYes
        '
        Me.optPinsYes.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.optPinsYes.AutoSize = True
        Me.optPinsYes.Location = New System.Drawing.Point(108, 21)
        Me.optPinsYes.Name = "optPinsYes"
        Me.optPinsYes.Size = New System.Drawing.Size(43, 17)
        Me.optPinsYes.TabIndex = 54
        Me.optPinsYes.TabStop = True
        Me.optPinsYes.Text = "Yes"
        Me.optPinsYes.UseVisualStyleBackColor = True
        '
        'cmbPinMaterial
        '
        Me.cmbPinMaterial.AcceptEnterKeyAsTab = True
        Me.cmbPinMaterial.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmbPinMaterial.AssociateLabel = Nothing
        Me.cmbPinMaterial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPinMaterial.FormattingEnabled = True
        Me.cmbPinMaterial.IFLDataTag = Nothing
        Me.cmbPinMaterial.Items.AddRange(New Object() {"", "Induction Hardened", "Standard"})
        Me.cmbPinMaterial.Location = New System.Drawing.Point(108, 44)
        Me.cmbPinMaterial.Name = "cmbPinMaterial"
        Me.cmbPinMaterial.Size = New System.Drawing.Size(342, 21)
        Me.cmbPinMaterial.StatusMessage = Nothing
        Me.cmbPinMaterial.StatusObject = Nothing
        Me.cmbPinMaterial.TabIndex = 22
        '
        'Label11
        '
        Me.Label11.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(75, 23)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(27, 13)
        Me.Label11.TabIndex = 53
        Me.Label11.Text = "Pins"
        '
        'Label12
        '
        Me.Label12.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(40, 48)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(62, 13)
        Me.Label12.TabIndex = 21
        Me.Label12.Text = "Pin Material"
        '
        'LabelGradient2
        '
        Me.LabelGradient2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelGradient2.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient2.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.LabelGradient2.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient2.Location = New System.Drawing.Point(-3, -1)
        Me.LabelGradient2.Name = "LabelGradient2"
        Me.LabelGradient2.Size = New System.Drawing.Size(468, 19)
        Me.LabelGradient2.TabIndex = 20
        Me.LabelGradient2.Text = "Pin Details"
        Me.LabelGradient2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'GroupBox6
        '
        Me.GroupBox6.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox6.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox6.Controls.Add(Me.LabelGradient7)
        Me.GroupBox6.Controls.Add(Me.cmbPistonSealPackage)
        Me.GroupBox6.Controls.Add(Me.Label10)
        Me.GroupBox6.Location = New System.Drawing.Point(502, 261)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(436, 55)
        Me.GroupBox6.TabIndex = 83
        Me.GroupBox6.TabStop = False
        '
        'LabelGradient7
        '
        Me.LabelGradient7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelGradient7.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust
        Me.LabelGradient7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient7.GradientColorOne = System.Drawing.Color.DeepSkyBlue
        Me.LabelGradient7.GradientColorTwo = System.Drawing.Color.Honeydew
        Me.LabelGradient7.Location = New System.Drawing.Point(0, -4)
        Me.LabelGradient7.Name = "LabelGradient7"
        Me.LabelGradient7.Size = New System.Drawing.Size(436, 21)
        Me.LabelGradient7.TabIndex = 20
        Me.LabelGradient7.Text = "Piston Seal Details"
        Me.LabelGradient7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmbPistonSealPackage
        '
        Me.cmbPistonSealPackage.AcceptEnterKeyAsTab = True
        Me.cmbPistonSealPackage.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmbPistonSealPackage.AssociateLabel = Nothing
        Me.cmbPistonSealPackage.BackColor = System.Drawing.Color.White
        Me.cmbPistonSealPackage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPistonSealPackage.FormattingEnabled = True
        Me.cmbPistonSealPackage.IFLDataTag = Nothing
        Me.cmbPistonSealPackage.Items.AddRange(New Object() {""})
        Me.cmbPistonSealPackage.Location = New System.Drawing.Point(132, 19)
        Me.cmbPistonSealPackage.Name = "cmbPistonSealPackage"
        Me.cmbPistonSealPackage.Size = New System.Drawing.Size(280, 21)
        Me.cmbPistonSealPackage.StatusMessage = Nothing
        Me.cmbPistonSealPackage.StatusObject = Nothing
        Me.cmbPistonSealPackage.TabIndex = 60
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(17, 26)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(103, 13)
        Me.Label10.TabIndex = 59
        Me.Label10.Text = "Piston Seal  Pakage"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripDropDownButton1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 652)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.ShowItemToolTips = True
        Me.StatusStrip1.Size = New System.Drawing.Size(995, 22)
        Me.StatusStrip1.TabIndex = 86
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
        'GroupBox1
        '
        Me.GroupBox1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox1.Controls.Add(Me.txtWorkingPressure)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtColumnLoad)
        Me.GroupBox1.Location = New System.Drawing.Point(34, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(904, 65)
        Me.GroupBox1.TabIndex = 103
        Me.GroupBox1.TabStop = False
        '
        'txtWorkingPressure
        '
        Me.txtWorkingPressure.AcceptEnterKeyAsTab = True
        Me.txtWorkingPressure.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtWorkingPressure.ApplyIFLColor = True
        Me.txtWorkingPressure.AssociateLabel = Nothing
        Me.txtWorkingPressure.DecimalValue = 2
        Me.txtWorkingPressure.Enabled = False
        Me.txtWorkingPressure.IFLDataTag = Nothing
        Me.txtWorkingPressure.InvalidInputCharacters = ""
        Me.txtWorkingPressure.IsAllowNegative = False
        Me.txtWorkingPressure.LengthValue = 6
        Me.txtWorkingPressure.Location = New System.Drawing.Point(107, 12)
        Me.txtWorkingPressure.MaximumValue = 99999
        Me.txtWorkingPressure.MaxLength = 6
        Me.txtWorkingPressure.MinimumValue = 0
        Me.txtWorkingPressure.Name = "txtWorkingPressure"
        Me.txtWorkingPressure.Size = New System.Drawing.Size(201, 20)
        Me.txtWorkingPressure.StatusMessage = ""
        Me.txtWorkingPressure.StatusObject = Nothing
        Me.txtWorkingPressure.TabIndex = 35
        Me.txtWorkingPressure.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(1, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(91, 13)
        Me.Label7.TabIndex = 34
        Me.Label7.Text = "Working Pressure"
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(23, 42)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(69, 13)
        Me.Label8.TabIndex = 36
        Me.Label8.Text = "Column Load"
        '
        'txtColumnLoad
        '
        Me.txtColumnLoad.AcceptEnterKeyAsTab = True
        Me.txtColumnLoad.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtColumnLoad.ApplyIFLColor = True
        Me.txtColumnLoad.AssociateLabel = Nothing
        Me.txtColumnLoad.DecimalValue = 2
        Me.txtColumnLoad.Enabled = False
        Me.txtColumnLoad.IFLDataTag = Nothing
        Me.txtColumnLoad.InvalidInputCharacters = ""
        Me.txtColumnLoad.IsAllowNegative = False
        Me.txtColumnLoad.LengthValue = 6
        Me.txtColumnLoad.Location = New System.Drawing.Point(107, 38)
        Me.txtColumnLoad.MaximumValue = 99999
        Me.txtColumnLoad.MaxLength = 6
        Me.txtColumnLoad.MinimumValue = 0
        Me.txtColumnLoad.Name = "txtColumnLoad"
        Me.txtColumnLoad.Size = New System.Drawing.Size(201, 20)
        Me.txtColumnLoad.StatusMessage = ""
        Me.txtColumnLoad.StatusObject = Nothing
        Me.txtColumnLoad.TabIndex = 37
        Me.txtColumnLoad.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'frmTieRod2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(995, 674)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.GroupBox19)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox6)
        Me.Name = "frmTieRod2"
        Me.Text = "TieRod Cylinder Assembly Details-Page2"
        Me.GroupBox19.ResumeLayout(False)
        Me.GroupBox19.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox19 As System.Windows.Forms.GroupBox
    Friend WithEvents IflNumericBox14 As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents IflNumericBox15 As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents IflComboBox12 As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents IflComboBox13 As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents LabelGradient20 As LabelGradient.LabelGradient
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents txtTieRodNutQty As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents txtTieRodNutSize As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents txtTieRodSize As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents cmbThreadProtected As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents LabelGradient8 As LabelGradient.LabelGradient
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents txtClevisCap As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents cmbRodClevis As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents cmbRodEndThread As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtRodCap As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents LabelGradient6 As LabelGradient.LabelGradient
    Friend WithEvents cmbRodSealPackage As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents IflNumericBox11 As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents IflNumericBox10 As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents cmbClips As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents LVPinSizeDetails As IFLCustomUILayer.IFLListView
    Friend WithEvents optPinsNo As System.Windows.Forms.RadioButton
    Friend WithEvents optPinsYes As System.Windows.Forms.RadioButton
    Friend WithEvents cmbPinMaterial As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents LabelGradient2 As LabelGradient.LabelGradient
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelGradient7 As LabelGradient.LabelGradient
    Friend WithEvents cmbPistonSealPackage As IFLCustomUILayer.IFLComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripDropDownButton1 As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtWorkingPressure As IFLCustomUILayer.IFLNumericBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtColumnLoad As IFLCustomUILayer.IFLNumericBox
End Class
