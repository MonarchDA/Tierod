<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mdiWeldedCylinder
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(mdiWeldedCylinder))
        Me.pnlInformationArea = New System.Windows.Forms.Panel
        Me.lvwGeneralInformation = New System.Windows.Forms.ListView
        Me.lvwLoginDetails = New System.Windows.Forms.ListView
        Me.pnlMonarchLogo = New System.Windows.Forms.Panel
        Me.pnlChildFormArea = New System.Windows.Forms.Panel
        Me.btnNext = New System.Windows.Forms.Button
        Me.btnHome = New System.Windows.Forms.Button
        Me.btnBack = New System.Windows.Forms.Button
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.pnlInformationArea.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlInformationArea
        '
        Me.pnlInformationArea.BackColor = System.Drawing.SystemColors.Window
        Me.pnlInformationArea.Controls.Add(Me.lvwGeneralInformation)
        Me.pnlInformationArea.Controls.Add(Me.lvwLoginDetails)
        Me.pnlInformationArea.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlInformationArea.Location = New System.Drawing.Point(0, 148)
        Me.pnlInformationArea.Name = "pnlInformationArea"
        Me.pnlInformationArea.Size = New System.Drawing.Size(313, 717)
        Me.pnlInformationArea.TabIndex = 9
        '
        'lvwGeneralInformation
        '
        Me.lvwGeneralInformation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lvwGeneralInformation.GridLines = True
        Me.lvwGeneralInformation.Location = New System.Drawing.Point(1, 1)
        Me.lvwGeneralInformation.Name = "lvwGeneralInformation"
        Me.lvwGeneralInformation.Size = New System.Drawing.Size(312, 625)
        Me.lvwGeneralInformation.TabIndex = 0
        Me.lvwGeneralInformation.UseCompatibleStateImageBehavior = False
        Me.lvwGeneralInformation.View = System.Windows.Forms.View.Details
        '
        'lvwLoginDetails
        '
        Me.lvwLoginDetails.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lvwLoginDetails.GridLines = True
        Me.lvwLoginDetails.Location = New System.Drawing.Point(0, 623)
        Me.lvwLoginDetails.Name = "lvwLoginDetails"
        Me.lvwLoginDetails.Size = New System.Drawing.Size(313, 94)
        Me.lvwLoginDetails.TabIndex = 0
        Me.lvwLoginDetails.UseCompatibleStateImageBehavior = False
        Me.lvwLoginDetails.View = System.Windows.Forms.View.Details
        '
        'pnlMonarchLogo
        '
        Me.pnlMonarchLogo.BackColor = System.Drawing.SystemColors.Window
        Me.pnlMonarchLogo.BackgroundImage = CType(resources.GetObject("pnlMonarchLogo.BackgroundImage"), System.Drawing.Image)
        Me.pnlMonarchLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pnlMonarchLogo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMonarchLogo.Location = New System.Drawing.Point(0, 0)
        Me.pnlMonarchLogo.Name = "pnlMonarchLogo"
        Me.pnlMonarchLogo.Size = New System.Drawing.Size(1229, 148)
        Me.pnlMonarchLogo.TabIndex = 10
        '
        'pnlChildFormArea
        '
        Me.pnlChildFormArea.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlChildFormArea.BackColor = System.Drawing.Color.WhiteSmoke
        Me.pnlChildFormArea.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.pnlChildFormArea.Location = New System.Drawing.Point(313, 148)
        Me.pnlChildFormArea.Name = "pnlChildFormArea"
        Me.pnlChildFormArea.Size = New System.Drawing.Size(914, 588)
        Me.pnlChildFormArea.TabIndex = 11
        '
        'btnNext
        '
        Me.btnNext.BackgroundImage = CType(resources.GetObject("btnNext.BackgroundImage"), System.Drawing.Image)
        Me.btnNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnNext.Location = New System.Drawing.Point(854, 49)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(46, 32)
        Me.btnNext.TabIndex = 2
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnHome
        '
        Me.btnHome.BackgroundImage = CType(resources.GetObject("btnHome.BackgroundImage"), System.Drawing.Image)
        Me.btnHome.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnHome.Location = New System.Drawing.Point(798, 49)
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(50, 34)
        Me.btnHome.TabIndex = 1
        Me.btnHome.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.BackgroundImage = CType(resources.GetObject("btnBack.BackgroundImage"), System.Drawing.Image)
        Me.btnBack.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnBack.Location = New System.Drawing.Point(746, 51)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(46, 32)
        Me.btnBack.TabIndex = 0
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'pnlBottom
        '
        Me.pnlBottom.BackColor = System.Drawing.Color.WhiteSmoke
        Me.pnlBottom.BackgroundImage = CType(resources.GetObject("pnlBottom.BackgroundImage"), System.Drawing.Image)
        Me.pnlBottom.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.pnlBottom.Controls.Add(Me.btnNext)
        Me.pnlBottom.Controls.Add(Me.btnBack)
        Me.pnlBottom.Controls.Add(Me.btnHome)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(313, 737)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(916, 128)
        Me.pnlBottom.TabIndex = 13
        '
        'mdiWeldedCylinder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1229, 865)
        Me.Controls.Add(Me.pnlBottom)
        Me.Controls.Add(Me.pnlInformationArea)
        Me.Controls.Add(Me.pnlChildFormArea)
        Me.Controls.Add(Me.pnlMonarchLogo)
        Me.IsMdiContainer = True
        Me.Name = "mdiWeldedCylinder"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Welded Cylinder"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlInformationArea.ResumeLayout(False)
        Me.pnlBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlInformationArea As System.Windows.Forms.Panel
    Friend WithEvents pnlMonarchLogo As System.Windows.Forms.Panel
    Friend WithEvents pnlChildFormArea As System.Windows.Forms.Panel
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnHome As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents lvwLoginDetails As System.Windows.Forms.ListView
    Friend WithEvents lvwGeneralInformation As System.Windows.Forms.ListView

End Class
