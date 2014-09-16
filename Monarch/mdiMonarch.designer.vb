<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mdiMonarch
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(mdiMonarch))
        Me.pnlInformationArea = New System.Windows.Forms.Panel()
        Me.mdiComponent = New System.Windows.Forms.ListView()
        Me.lvwGeneralInformation = New System.Windows.Forms.ListView()
        Me.lvwLoginDetails = New System.Windows.Forms.ListView()
        Me.pnlMonarchLogo = New System.Windows.Forms.Panel()
        Me.picMonarchLogo = New System.Windows.Forms.PictureBox()
        Me.LabelGradient5 = New LabelGradient.LabelGradient()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.AssemblyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuNewCylinder = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuRevison = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuReleased = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnHome = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.pnlBottom = New System.Windows.Forms.Panel()
        Me.btnGenerateFromExcel = New System.Windows.Forms.Button()
        Me.prb = New System.Windows.Forms.ProgressBar()
        Me.btnGenerateReport = New System.Windows.Forms.Button()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label2 = New LabelGradient.LabelGradient()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.pnlChildFormArea = New System.Windows.Forms.Panel()
        Me.toolTipInfo = New System.Windows.Forms.ToolTip(Me.components)
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.pnlInformationArea.SuspendLayout()
        Me.pnlMonarchLogo.SuspendLayout()
        CType(Me.picMonarchLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlInformationArea
        '
        Me.pnlInformationArea.BackColor = System.Drawing.SystemColors.Window
        Me.pnlInformationArea.BackgroundImage = CType(resources.GetObject("pnlInformationArea.BackgroundImage"), System.Drawing.Image)
        Me.pnlInformationArea.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pnlInformationArea.Controls.Add(Me.mdiComponent)
        Me.pnlInformationArea.Controls.Add(Me.lvwGeneralInformation)
        Me.pnlInformationArea.Controls.Add(Me.lvwLoginDetails)
        Me.pnlInformationArea.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlInformationArea.Location = New System.Drawing.Point(0, 100)
        Me.pnlInformationArea.Name = "pnlInformationArea"
        Me.pnlInformationArea.Size = New System.Drawing.Size(313, 646)
        Me.pnlInformationArea.TabIndex = 9
        '
        'mdiComponent
        '
        Me.mdiComponent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.mdiComponent.BackColor = System.Drawing.SystemColors.Info
        Me.mdiComponent.BackgroundImage = CType(resources.GetObject("mdiComponent.BackgroundImage"), System.Drawing.Image)
        Me.mdiComponent.BackgroundImageTiled = True
        Me.mdiComponent.FullRowSelect = True
        Me.mdiComponent.GridLines = True
        Me.mdiComponent.Location = New System.Drawing.Point(0, 279)
        Me.mdiComponent.Name = "mdiComponent"
        Me.mdiComponent.Size = New System.Drawing.Size(312, 275)
        Me.mdiComponent.TabIndex = 1
        Me.mdiComponent.UseCompatibleStateImageBehavior = False
        Me.mdiComponent.View = System.Windows.Forms.View.Details
        '
        'lvwGeneralInformation
        '
        Me.lvwGeneralInformation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lvwGeneralInformation.BackColor = System.Drawing.SystemColors.Info
        Me.lvwGeneralInformation.BackgroundImage = CType(resources.GetObject("lvwGeneralInformation.BackgroundImage"), System.Drawing.Image)
        Me.lvwGeneralInformation.BackgroundImageTiled = True
        Me.lvwGeneralInformation.FullRowSelect = True
        Me.lvwGeneralInformation.GridLines = True
        Me.lvwGeneralInformation.Location = New System.Drawing.Point(3, 3)
        Me.lvwGeneralInformation.Name = "lvwGeneralInformation"
        Me.lvwGeneralInformation.Size = New System.Drawing.Size(312, 160)
        Me.lvwGeneralInformation.TabIndex = 0
        Me.lvwGeneralInformation.UseCompatibleStateImageBehavior = False
        Me.lvwGeneralInformation.View = System.Windows.Forms.View.Details
        '
        'lvwLoginDetails
        '
        Me.lvwLoginDetails.BackColor = System.Drawing.SystemColors.Info
        Me.lvwLoginDetails.BackgroundImage = CType(resources.GetObject("lvwLoginDetails.BackgroundImage"), System.Drawing.Image)
        Me.lvwLoginDetails.BackgroundImageTiled = True
        Me.lvwLoginDetails.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lvwLoginDetails.GridLines = True
        Me.lvwLoginDetails.Location = New System.Drawing.Point(0, 552)
        Me.lvwLoginDetails.Name = "lvwLoginDetails"
        Me.lvwLoginDetails.Size = New System.Drawing.Size(313, 94)
        Me.lvwLoginDetails.TabIndex = 0
        Me.lvwLoginDetails.UseCompatibleStateImageBehavior = False
        Me.lvwLoginDetails.View = System.Windows.Forms.View.Details
        '
        'pnlMonarchLogo
        '
        Me.pnlMonarchLogo.BackColor = System.Drawing.Color.Transparent
        Me.pnlMonarchLogo.BackgroundImage = CType(resources.GetObject("pnlMonarchLogo.BackgroundImage"), System.Drawing.Image)
        Me.pnlMonarchLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pnlMonarchLogo.Controls.Add(Me.picMonarchLogo)
        Me.pnlMonarchLogo.Controls.Add(Me.LabelGradient5)
        Me.pnlMonarchLogo.Controls.Add(Me.MenuStrip1)
        Me.pnlMonarchLogo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMonarchLogo.Location = New System.Drawing.Point(0, 0)
        Me.pnlMonarchLogo.Name = "pnlMonarchLogo"
        Me.pnlMonarchLogo.Size = New System.Drawing.Size(1028, 100)
        Me.pnlMonarchLogo.TabIndex = 10
        '
        'picMonarchLogo
        '
        Me.picMonarchLogo.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picMonarchLogo.BackgroundImage = CType(resources.GetObject("picMonarchLogo.BackgroundImage"), System.Drawing.Image)
        Me.picMonarchLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.picMonarchLogo.Location = New System.Drawing.Point(289, 24)
        Me.picMonarchLogo.Name = "picMonarchLogo"
        Me.picMonarchLogo.Size = New System.Drawing.Size(451, 67)
        Me.picMonarchLogo.TabIndex = 111
        Me.picMonarchLogo.TabStop = False
        '
        'LabelGradient5
        '
        Me.LabelGradient5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelGradient5.AutoSize = True
        Me.LabelGradient5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelGradient5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LabelGradient5.GradientColorOne = System.Drawing.Color.MediumTurquoise
        Me.LabelGradient5.GradientColorTwo = System.Drawing.Color.White
        Me.LabelGradient5.Image = CType(resources.GetObject("LabelGradient5.Image"), System.Drawing.Image)
        Me.LabelGradient5.Location = New System.Drawing.Point(-3, 24)
        Me.LabelGradient5.Name = "LabelGradient5"
        Me.LabelGradient5.Size = New System.Drawing.Size(0, 15)
        Me.LabelGradient5.TabIndex = 110
        Me.LabelGradient5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AssemblyToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1028, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'AssemblyToolStripMenuItem
        '
        Me.AssemblyToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuNewCylinder, Me.MenuRevison, Me.MenuReleased})
        Me.AssemblyToolStripMenuItem.Name = "AssemblyToolStripMenuItem"
        Me.AssemblyToolStripMenuItem.Size = New System.Drawing.Size(74, 20)
        Me.AssemblyToolStripMenuItem.Text = "Assembly"
        '
        'MenuNewCylinder
        '
        Me.MenuNewCylinder.BackColor = System.Drawing.Color.DarkKhaki
        Me.MenuNewCylinder.Name = "MenuNewCylinder"
        Me.MenuNewCylinder.Size = New System.Drawing.Size(136, 22)
        Me.MenuNewCylinder.Text = "&New"
        '
        'MenuRevison
        '
        Me.MenuRevison.BackColor = System.Drawing.Color.DarkKhaki
        Me.MenuRevison.Name = "MenuRevison"
        Me.MenuRevison.Size = New System.Drawing.Size(136, 22)
        Me.MenuRevison.Text = "&Revision"
        '
        'MenuReleased
        '
        Me.MenuReleased.BackColor = System.Drawing.Color.DarkKhaki
        Me.MenuReleased.Name = "MenuReleased"
        Me.MenuReleased.Size = New System.Drawing.Size(136, 22)
        Me.MenuReleased.Text = "Release"
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNext.BackgroundImage = CType(resources.GetObject("btnNext.BackgroundImage"), System.Drawing.Image)
        Me.btnNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnNext.Location = New System.Drawing.Point(654, 70)
        Me.btnNext.Margin = New System.Windows.Forms.Padding(5)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(40, 40)
        Me.btnNext.TabIndex = 2
        Me.toolTipInfo.SetToolTip(Me.btnNext, "Next")
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnHome
        '
        Me.btnHome.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnHome.BackgroundImage = CType(resources.GetObject("btnHome.BackgroundImage"), System.Drawing.Image)
        Me.btnHome.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnHome.Location = New System.Drawing.Point(606, 70)
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(40, 40)
        Me.btnHome.TabIndex = 1
        Me.toolTipInfo.SetToolTip(Me.btnHome, "Home")
        Me.btnHome.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBack.BackgroundImage = CType(resources.GetObject("btnBack.BackgroundImage"), System.Drawing.Image)
        Me.btnBack.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnBack.Location = New System.Drawing.Point(560, 70)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(40, 40)
        Me.btnBack.TabIndex = 0
        Me.toolTipInfo.SetToolTip(Me.btnBack, "Previous page")
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'pnlBottom
        '
        Me.pnlBottom.BackColor = System.Drawing.Color.White
        Me.pnlBottom.BackgroundImage = CType(resources.GetObject("pnlBottom.BackgroundImage"), System.Drawing.Image)
        Me.pnlBottom.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pnlBottom.Controls.Add(Me.btnGenerateFromExcel)
        Me.pnlBottom.Controls.Add(Me.prb)
        Me.pnlBottom.Controls.Add(Me.btnGenerateReport)
        Me.pnlBottom.Controls.Add(Me.btnGenerate)
        Me.pnlBottom.Controls.Add(Me.btnCancel)
        Me.pnlBottom.Controls.Add(Me.Label2)
        Me.pnlBottom.Controls.Add(Me.PictureBox3)
        Me.pnlBottom.Controls.Add(Me.btnNext)
        Me.pnlBottom.Controls.Add(Me.btnBack)
        Me.pnlBottom.Controls.Add(Me.btnHome)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(313, 624)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(715, 122)
        Me.pnlBottom.TabIndex = 13
        '
        'btnGenerateFromExcel
        '
        Me.btnGenerateFromExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerateFromExcel.BackgroundImage = CType(resources.GetObject("btnGenerateFromExcel.BackgroundImage"), System.Drawing.Image)
        Me.btnGenerateFromExcel.Location = New System.Drawing.Point(652, 64)
        Me.btnGenerateFromExcel.Name = "btnGenerateFromExcel"
        Me.btnGenerateFromExcel.Size = New System.Drawing.Size(52, 52)
        Me.btnGenerateFromExcel.TabIndex = 32
        Me.toolTipInfo.SetToolTip(Me.btnGenerateFromExcel, "Clich here to generate Model")
        Me.btnGenerateFromExcel.UseVisualStyleBackColor = True
        Me.btnGenerateFromExcel.Visible = False
        '
        'prb
        '
        Me.prb.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.prb.BackColor = System.Drawing.Color.Silver
        Me.prb.ForeColor = System.Drawing.Color.Orange
        Me.prb.Location = New System.Drawing.Point(12, 18)
        Me.prb.Name = "prb"
        Me.prb.Size = New System.Drawing.Size(803, 23)
        Me.prb.TabIndex = 0
        Me.prb.UseWaitCursor = True
        Me.prb.Visible = False
        '
        'btnGenerateReport
        '
        Me.btnGenerateReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerateReport.BackgroundImage = CType(resources.GetObject("btnGenerateReport.BackgroundImage"), System.Drawing.Image)
        Me.btnGenerateReport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnGenerateReport.Location = New System.Drawing.Point(468, 70)
        Me.btnGenerateReport.Name = "btnGenerateReport"
        Me.btnGenerateReport.Size = New System.Drawing.Size(40, 40)
        Me.btnGenerateReport.TabIndex = 31
        Me.toolTipInfo.SetToolTip(Me.btnGenerateReport, "Click here to Generate Report")
        Me.btnGenerateReport.UseVisualStyleBackColor = True
        '
        'btnGenerate
        '
        Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerate.BackgroundImage = CType(resources.GetObject("btnGenerate.BackgroundImage"), System.Drawing.Image)
        Me.btnGenerate.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnGenerate.Location = New System.Drawing.Point(407, 64)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(55, 52)
        Me.btnGenerate.TabIndex = 30
        Me.toolTipInfo.SetToolTip(Me.btnGenerate, "Click here to generate model")
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.BackgroundImage = CType(resources.GetObject("btnCancel.BackgroundImage"), System.Drawing.Image)
        Me.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnCancel.Location = New System.Drawing.Point(514, 70)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(40, 40)
        Me.btnCancel.TabIndex = 29
        Me.toolTipInfo.SetToolTip(Me.btnCancel, "Click here to Close ")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.GradientColorOne = System.Drawing.Color.Orange
        Me.Label2.GradientColorTwo = System.Drawing.Color.Azure
        Me.Label2.Location = New System.Drawing.Point(205, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(269, 13)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Copyright © 2013, Invilogic Software Pvt. Ltd."
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'pnlChildFormArea
        '
        Me.pnlChildFormArea.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlChildFormArea.AutoScroll = True
        Me.pnlChildFormArea.BackColor = System.Drawing.Color.Black
        Me.pnlChildFormArea.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pnlChildFormArea.Location = New System.Drawing.Point(313, 100)
        Me.pnlChildFormArea.Name = "pnlChildFormArea"
        Me.pnlChildFormArea.Size = New System.Drawing.Size(715, 530)
        Me.pnlChildFormArea.TabIndex = 15
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(6, 18)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(192, 98)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 26
        Me.PictureBox3.TabStop = False
        '
        'mdiMonarch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1028, 746)
        Me.Controls.Add(Me.pnlChildFormArea)
        Me.Controls.Add(Me.pnlBottom)
        Me.Controls.Add(Me.pnlInformationArea)
        Me.Controls.Add(Me.pnlMonarchLogo)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "mdiMonarch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CYDAv2.0"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlInformationArea.ResumeLayout(False)
        Me.pnlMonarchLogo.ResumeLayout(False)
        Me.pnlMonarchLogo.PerformLayout()
        CType(Me.picMonarchLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlBottom.PerformLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlInformationArea As System.Windows.Forms.Panel
    Friend WithEvents pnlMonarchLogo As System.Windows.Forms.Panel
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnHome As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents lvwLoginDetails As System.Windows.Forms.ListView
    Friend WithEvents lvwGeneralInformation As System.Windows.Forms.ListView
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents AssemblyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuNewCylinder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label2 As LabelGradient.LabelGradient
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents pnlChildFormArea As System.Windows.Forms.Panel
    Friend WithEvents LabelGradient5 As LabelGradient.LabelGradient
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnGenerateReport As System.Windows.Forms.Button
    Friend WithEvents toolTipInfo As System.Windows.Forms.ToolTip
    Friend WithEvents prb As System.Windows.Forms.ProgressBar
    Friend WithEvents mdiComponent As System.Windows.Forms.ListView
    Friend WithEvents picMonarchLogo As System.Windows.Forms.PictureBox
    Friend WithEvents MenuRevison As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuReleased As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnGenerateFromExcel As System.Windows.Forms.Button
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox

End Class
